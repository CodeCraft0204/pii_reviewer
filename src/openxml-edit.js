// openxml-edit.js
import JSZip from "jszip";
import { XMLParser, XMLBuilder } from "fast-xml-parser";

// Utilities
const fpOpts = { ignoreAttributes: false, attributeNamePrefix: "" };
const parser = new XMLParser(fpOpts);
const builder = new XMLBuilder(fpOpts);

// docx text node QName
const W = {
  P: "w:p", R: "w:r", T: "w:t", RPR: "w:rPr", COLOR: "w:color",
  B: "w:b", I: "w:i", U: "w:u", VAL: "w:val"
};

function copyNonTextChildren(origRun) {
  const keep = {};
  for (const k of Object.keys(origRun)) {
    if (k === W.RPR || k === W.T) continue;
    keep[k] = clone(origRun[k]);
  }
  return keep;
}

// test original “bold + red”
function isBoldRed(rPr) {
  if (!rPr) return false;
  const hasBold = !!rPr[W.B];
  const color = rPr[W.COLOR]?.[W.VAL];
  // Many Word docs use "FF0000" for red; accept small variants if needed
  const redish = color && /^(FF0000|C00000|E00000)$/i.test(color);
  return !!hasBold && !!redish;
}

function clone(obj) {
  return JSON.parse(JSON.stringify(obj));
}

// Build a flat text view with per-character mapping to (pIndex, rIndex, charOffsetInRun)
export function indexRuns(docXml) {
  const doc = parser.parse(docXml);
  const body = doc["w:document"]["w:body"];
  const paragraphs = Array.isArray(body[W.P]) ? body[W.P] : (body[W.P] ? [body[W.P]] : []);

  const map = []; // [{p, r, o}] per character
  const plist = []; // keep structure reference
  let full = "";

  paragraphs.forEach((p, pIdx) => {
    const runs = Array.isArray(p[W.R]) ? p[W.R] : (p[W.R] ? [p[W.R]] : []);
    plist.push(runs);
    runs.forEach((r, rIdx) => {
      // account for tabs/breaks as single spaces in the flattened text
      const hasTab = Array.isArray(r["w:tab"]) ? r["w:tab"].length : (r["w:tab"] ? 1 : 0);
      const hasBr = Array.isArray(r["w:br"]) ? r["w:br"].length : (r["w:br"] ? 1 : 0);
      const wsCount = hasTab + hasBr;
      if (wsCount > 0) {
        // insert one space per tab/break for matching purposes
        for (let k = 0; k < wsCount; k++) {
          map.push({ p: pIdx, r: rIdx, o: -1 }); // o:-1 marks synthetic whitespace
          full += " ";
        }
      }

      const t = r[W.T];
      if (t == null) return;
      const text = typeof t === "string" ? t : (t?.["#text"] ?? "");
      
      // Add space between runs only if both are non-whitespace and not punctuation
      if (full.length > 0 && text.length > 0) {
        const lastChar = full[full.length - 1];
        const firstChar = text[0];
        
        // Only add space if:
        // 1. Last char is a letter/digit and first char is a letter/digit
        // 2. Or last char is a letter/digit and first char is not punctuation
        // 3. And neither is already whitespace
        const shouldAddSpace = !/\s/.test(lastChar) && !/\s/.test(firstChar) && 
                              ((/[a-zA-Z0-9]/.test(lastChar) && /[a-zA-Z0-9]/.test(firstChar)) ||
                               (/[a-zA-Z0-9]/.test(lastChar) && !/[.,;:!?()[\]{}<>$"'\-]/.test(firstChar)));
        
        if (shouldAddSpace) {
          map.push({ p: pIdx, r: rIdx, o: -2 }); // o:-2 marks synthetic space between runs
          full += " ";
        }
      }
      
      for (let i = 0; i < text.length; i++) {
        map.push({ p: pIdx, r: rIdx, o: i });
      }
      full += text;
    });
  });
  return { doc, body, paragraphs, full, map, plist };
}

// Given absolute [start,end) over full text, split the exact runs and apply a style overlay
function applyStyleToRange(struct, start, end, styler) {
  const { paragraphs, plist, map } = struct;

  if (start < 0 || end <= start || end > map.length) return;

  // Determine all (p,r,offset) positions covered
  const first = map[start];
  const last = map[end - 1];
  if (!first || !last) return;

  for (let i = start; i < end;) {
    const pos = map[i];
    const pRuns = plist[pos.p];
    let rNode = pRuns[pos.r];
    const tNode = rNode[W.T];
    if (pos.o === -1 || pos.o === -2) { i++; continue; } // synthetic whitespace (tab/br/space) — skip styling
    if (tNode == null) { i++; continue; }

    const txt = typeof tNode === "string" ? tNode : (tNode?.["#text"] ?? "");
    // Compute how much of this run we can consume
    const runStartAbs = i;
    // Find how far we can go within this run
    // Count remaining chars in this run from offset 'pos.o'
    const remainingInRun = txt.length - pos.o;
    const take = Math.min(remainingInRun, end - i);

    // Split the run into [before][target][after]
    const beforeText = txt.slice(0, pos.o);
    const midText = txt.slice(pos.o, pos.o + take);
    const afterText = txt.slice(pos.o + take);

    // Prepare three run nodes based on original styling
    const baseRPr = clone(rNode[W.RPR] || {});

    const mkRun = (text, pr = null) => {
      const out = {};
      if (pr && Object.keys(pr).length) out[W.RPR] = pr;
      // Always preserve spaces to avoid collapsing when we split runs
      out[W.T] = { "#text": text, "xml:space": "preserve" };
      return out;
    };

    const runsNew = [];
    if (beforeText) runsNew.push(mkRun(beforeText, baseRPr));

    // Styled middle
    const styledRPr = styler(clone(baseRPr), isBoldRed(baseRPr));
    runsNew.push(mkRun(midText, styledRPr));

    if (afterText) runsNew.push(mkRun(afterText, baseRPr));

    // Replace original run with new split runs
    if (runsNew.length) {
      const extras = copyNonTextChildren(rNode);
      Object.assign(runsNew[0], extras);
    }
    pRuns.splice(pos.r, 1, ...runsNew);

    // Update mapping for following characters:
    // We must also update 'map' to reflect the new run boundaries.
    // Easiest approach: recompute mapping for this paragraph from scratch.
    // (We only call this per affected range, scale is fine for typical docs.)
    // Re-index the paragraph runs into global 'map':
    const newRuns = pRuns;
    let paraStartIdx = null;
    // find first global index of this paragraph in map
    for (let gi = 0; gi < map.length; gi++) {
      if (map[gi].p === pos.p) { paraStartIdx = gi; break; }
    }
    // remove old entries for this paragraph
    for (let gi = map.length - 1; gi >= 0; gi--) {
      if (map[gi].p === pos.p) map.splice(gi, 1);
    }
    // rebuild paragraph mapping
    for (let rIdx = 0; rIdx < newRuns.length; rIdx++) {
      const rt = newRuns[rIdx][W.T];
      if (rt == null) continue;
      const rtxt = typeof rt === "string" ? rt : (rt?.["#text"] ?? "");
      for (let off = 0; off < rtxt.length; off++) {
        map.splice(paraStartIdx + (map[paraStartIdx]?.p === pos.p ? 0 : 0), 0, { p: pos.p, r: rIdx, o: off });
      }
    }

    // Advance i by 'take'
    i = runStartAbs + take;
  }
}

// Style overlays
function overlayExact(rPr /*orig*/, wasBoldRed) {
  // If original was bold+red (ground truth), we keep bold as is and add italic+underline and force red color.
  const out = rPr || {};
  out[W.I] = {};                 // italic
  out[W.U] = { [W.VAL]: "single" }; // underline
  out[W.COLOR] = { [W.VAL]: "FF0000" }; // red
  return out;
}

function overlayBrown(rPr /*orig*/) {
  const out = rPr || {};
  out[W.U] = { [W.VAL]: "single" }; // underline
  out[W.COLOR] = { [W.VAL]: "7B3F00" }; // brown
  return out;
}

// Find all occurrences of a value in the full text (non-overlapping)
function findAll(text, value) {
  if (!value) return [];
  const hits = [];
  let i = 0;
  while (i < text.length) {
    const j = text.indexOf(value, i);
    if (j < 0) break;
    hits.push([j, j + value.length]);
    i = j + value.length;
  }
  return hits;
}

// Decide if a span was originally “bold + red”
function spanWasBoldRed(struct, start, end) {
  const { map, paragraphs, plist } = struct;
  if (start < 0 || end <= start || end > map.length) return false;
  for (let k = start; k < end; k++) {
    const pos = map[k];
    const rNode = plist[pos.p][pos.r];
    const rPr = rNode[W.RPR];
    if (!isBoldRed(rPr)) return false; // every char of span must be bold red
  }
  return true;
}

/**
 * Annotate a DOCX in memory.
 * detections: [{ type, value }]
 * Rules:
 *  - If matching span exists and was originally bold+red => restyle to red+italic+underline
 *  - If matching span exists but NOT bold+red => restyle to brown+underline
 *  - If value not found in doc => leave document unchanged (document-only policy)
 */
export async function annotateDocxWithDetections(docxArrayBuffer, detections) {
  try {
    console.log("annotateDocxWithDetections: Starting");
    const zip = await JSZip.loadAsync(docxArrayBuffer);
    console.log("ZIP loaded");
    
    const docXml = await zip.file("word/document.xml")?.async("string");
    if (!docXml) throw new Error("word/document.xml not found");
    console.log("Document XML extracted");

    // Use mammoth for text extraction to get proper spacing
    const mammoth = await import("mammoth");
    const { value: docTextRaw } = await mammoth.extractRawText({ arrayBuffer: docxArrayBuffer });
    const full = (docTextRaw || "").replace(/\r/g, "");
    console.log(`Document text extracted with mammoth, length: ${full.length}`);

    // Also extract with indexRuns for XML structure mapping
    const struct = indexRuns(docXml);
    const { full: xmlText } = struct;
    console.log(`XML text extracted, length: ${xmlText.length}`);

    console.log(`Processing ${detections.length} detections...`);
    for (let i = 0; i < detections.length; i++) {
      const det = detections[i];
      console.log(`Processing detection ${i + 1}/${detections.length}: "${det.value}"`);
      
      const val = String(det.value || "");
      if (!val) {
        console.log("Skipping empty value");
        continue;
      }

      // Find ranges in the mammoth text (with proper spacing)
      const ranges = findAll(full, val);
      console.log(`Found ${ranges.length} ranges for "${val}" in mammoth text`);
      
      if (!ranges.length) {
        console.log("No ranges found, skipping");
        continue;
      }

      // For each range found in mammoth text, try to find corresponding range in XML text
      for (let j = 0; j < ranges.length; j++) {
        const [s, e] = ranges[j];
        console.log(`Processing range ${j + 1}/${ranges.length}: [${s}, ${e}]`);
        
        try {
          // Find the corresponding range in XML text
          const xmlRanges = findAll(xmlText, val);
          if (xmlRanges.length > j) {
            const [xmlS, xmlE] = xmlRanges[j];
            console.log(`Found corresponding XML range: [${xmlS}, ${xmlE}]`);
            
            const wasBR = spanWasBoldRed(struct, xmlS, xmlE);
            console.log(`Was originally bold+red: ${wasBR}`);
            
            applyStyleToRange(struct, xmlS, xmlE, (rPr) => (wasBR ? overlayExact(rPr, true) : overlayBrown(rPr)));
            console.log(`Applied styling to XML range [${xmlS}, ${xmlE}]`);
          }
        } catch (err) {
          console.warn(`Error processing range [${s}, ${e}] for value "${val}":`, err);
          // Continue with other ranges even if one fails
        }
      }
    }

    console.log("Rebuilding XML...");
    // Rebuild XML and zip
    const newDocXml = builder.build(struct.doc);
    zip.file("word/document.xml", newDocXml);
    
    console.log("Generating final document...");
    const outBuf = await zip.generateAsync({ type: "blob" });
    console.log("Document generated successfully");
    return outBuf;
  } catch (error) {
    console.error("Error in annotateDocxWithDetections:", error);
    throw error;
  }
}
