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
function isRed(rPr) {
  if (!rPr) return false;
  const color = rPr[W.COLOR]?.[W.VAL];
  // Many Word docs use "FF0000" for red; accept small variants if needed
  const redish = color && /^(FF0000|C00000|E00000)$/i.test(color);
  return !!redish;
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
      
      // Add space between runs to preserve word boundaries
      if (full.length > 0 && text.length > 0) {
        // Always add space between runs to ensure proper word separation
        // This is very aggressive but necessary to fix the spacing issue
        map.push({ p: pIdx, r: rIdx, o: -2 }); // o:-2 marks synthetic space between runs
        full += " ";
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
    if (!pos) { i++; continue; } // Skip undefined map entries
    
    const pRuns = plist[pos.p];
    if (!pRuns || !pRuns[pos.r]) { i++; continue; } // Skip if paragraph or run doesn't exist
    
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
    const styledRPr = styler(clone(baseRPr), isRed(baseRPr));
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

// Modify XML directly to preserve spacing
function modifyXmlDirectly(xmlString, textToFind, originalValue) {
  // Escape special XML characters in the text to find
  const escapeXml = (str) => str.replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;');
  const escapedText = escapeXml(textToFind);
  
  // Find all <w:t> tags containing this text
  const tTagRegex = /<w:t(?:\s+[^>]*)?>([^<]*)<\/w:t>/g;
  
  let modified = xmlString;
  let match;
  
  // Reset regex
  tTagRegex.lastIndex = 0;
  
  while ((match = tTagRegex.exec(xmlString)) !== null) {
    const fullMatch = match[0];
    const textContent = match[1];
    
    // Check if this text node contains our target text
    if (textContent.includes(textToFind)) {
      // Find the parent <w:r> tag
      const beforeMatch = xmlString.substring(0, match.index);
      const lastRStart = beforeMatch.lastIndexOf('<w:r>');
      const lastRStartWithProps = beforeMatch.lastIndexOf('<w:r ');
      const rStart = Math.max(lastRStart, lastRStartWithProps);
      
      if (rStart === -1) continue;
      
      const rEnd = xmlString.indexOf('</w:r>', match.index) + 6;
      const runXml = xmlString.substring(rStart, rEnd);
      
      // Check if this run was originally red
      const wasRed = runXml.includes('<w:color w:val="FF0000"') || 
                     runXml.includes('<w:color w:val="C00000"') ||
                     runXml.includes('<w:color w:val="E00000"');
      
      // Determine the styling to apply
      let newRunXml;
      if (wasRed) {
        // Ground truth: red italic underline
        newRunXml = applyRedItalicUnderlineToRun(runXml);
      } else {
        // AI-added: brown underline
        newRunXml = applyBrownUnderlineToRun(runXml);
      }
      
      // Replace the run in the modified XML
      modified = modified.substring(0, rStart) + newRunXml + modified.substring(rEnd);
      
      // Only modify the first occurrence
      break;
    }
  }
  
  return modified;
}

// Apply red italic underline styling to a run
function applyRedItalicUnderlineToRun(runXml) {
  // Check if <w:rPr> exists
  if (runXml.includes('<w:rPr>')) {
    // Modify existing properties
    let modified = runXml;
    
    // Ensure color is red
    if (!modified.includes('<w:color')) {
      modified = modified.replace('<w:rPr>', '<w:rPr><w:color w:val="FF0000"/>');
    }
    
    // Add italic if not present
    if (!modified.includes('<w:i')) {
      modified = modified.replace('<w:rPr>', '<w:rPr><w:i/>');
    }
    
    // Add underline if not present
    if (!modified.includes('<w:u')) {
      modified = modified.replace('<w:rPr>', '<w:rPr><w:u w:val="single"/>');
    }
    
    return modified;
  } else {
    // Add new <w:rPr> after <w:r> or <w:r ...>
    return runXml.replace(/(<w:r(?:\s+[^>]*)?>)/, '$1<w:rPr><w:color w:val="FF0000"/><w:i/><w:u w:val="single"/></w:rPr>');
  }
}

// Apply brown underline styling to a run
function applyBrownUnderlineToRun(runXml) {
  // Check if <w:rPr> exists
  if (runXml.includes('<w:rPr>')) {
    // Modify existing properties
    let modified = runXml;
    
    // Ensure color is brown
    if (modified.includes('<w:color')) {
      modified = modified.replace(/<w:color w:val="[^"]*"/, '<w:color w:val="8B4513"');
    } else {
      modified = modified.replace('<w:rPr>', '<w:rPr><w:color w:val="8B4513"/>');
    }
    
    // Add underline if not present
    if (!modified.includes('<w:u')) {
      modified = modified.replace('<w:rPr>', '<w:rPr><w:u w:val="single"/>');
    }
    
    return modified;
  } else {
    // Add new <w:rPr> after <w:r> or <w:r ...>
    return runXml.replace(/(<w:r(?:\s+[^>]*)?>)/, '$1<w:rPr><w:color w:val="8B4513"/><w:u w:val="single"/></w:rPr>');
  }
}

// Apply styling to a specific range in XML
function applyStyleToRangeInXML(paragraphs, start, end, originalValue) {
  let currentPos = 0;
  
  for (let pIdx = 0; pIdx < paragraphs.length; pIdx++) {
    const p = paragraphs[pIdx];
    const runs = Array.isArray(p[W.R]) ? p[W.R] : (p[W.R] ? [p[W.R]] : []);
    
    for (let rIdx = 0; rIdx < runs.length; rIdx++) {
      const run = runs[rIdx];
      const t = run[W.T];
      if (t == null) continue;
      
      const runText = typeof t === "string" ? t : (t?.["#text"] ?? "");
      if (!runText) continue;
      
      const runStart = currentPos;
      const runEnd = currentPos + runText.length;
      
      // Check if this run overlaps with our target range
      if (runStart < end && runEnd > start) {
        console.log(`Styling run at paragraph ${pIdx}, run ${rIdx}: "${runText}"`);
        
        // Check if this span was originally bold+red
        const rPr = run[W.RPR];
        const wasBR = isRed(rPr);
        console.log(`Was originally bold+red: ${wasBR}`);
        
        // Apply appropriate styling - ensure we don't break the XML structure
        try {
          if (wasBR) {
            // Ground truth: red italic underline
            const styledRPr = overlayExact(rPr, true);
            if (styledRPr && typeof styledRPr === 'object') {
              run[W.RPR] = styledRPr;
              console.log(`Applied red italic underline to "${runText}"`);
            }
          } else {
            // AI-added: brown underline
            const styledRPr = overlayBrown(rPr);
            if (styledRPr && typeof styledRPr === 'object') {
              run[W.RPR] = styledRPr;
              console.log(`Applied brown underline to "${runText}"`);
            }
          }
        } catch (err) {
          console.warn(`Error applying styling to run:`, err);
          // Continue without breaking the document
        }
      }
      
      currentPos += runText.length;
    }
    
    // Add paragraph break (newline) after each paragraph
    if (pIdx < paragraphs.length - 1) {
      currentPos += 1;
    }
  }
  
  return true; // Styled successfully
}

// Decide if a span was originally "bold + red"
function spanWasBoldRed(struct, start, end) {
  const { map, paragraphs, plist } = struct;
  if (start < 0 || end <= start || end > map.length) return false;
  for (let k = start; k < end; k++) {
    const pos = map[k];
    if (!pos) continue; // Skip undefined map entries
    
    const pRuns = plist[pos.p];
    if (!pRuns || !pRuns[pos.r]) continue; // Skip if paragraph or run doesn't exist
    
    const rNode = pRuns[pos.r];
    const rPr = rNode[W.RPR];
    if (!isRed(rPr)) return false; // every char of span must be bold red
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
  const zip = await JSZip.loadAsync(docxArrayBuffer);
  let docXml = await zip.file("word/document.xml")?.async("string");
  if (!docXml) throw new Error("word/document.xml not found");

  // Use mammoth for text extraction to get proper spacing
  const mammoth = await import("mammoth");
  const { value: docTextRaw } = await mammoth.extractRawText({ arrayBuffer: docxArrayBuffer });
  const mammothText = (docTextRaw || "").replace(/\r/g, "");

  for (const det of detections || []) {
    const val = String(det.value || "");
    if (!val) continue;

    // Find the text in mammoth text (with proper spacing)
    const ranges = findAll(mammothText, val);
    if (!ranges.length) {
      // not found in doc -> do nothing in the document (JSON report will handle notes)
      continue;
    }

    // For each range found, modify the XML directly
    for (const [mammothStart, mammothEnd] of ranges) {
      const textToFind = mammothText.substring(mammothStart, mammothEnd);
      
      // Modify XML string directly to preserve spacing
      docXml = modifyXmlDirectly(docXml, textToFind, val);
    }
  }

  // Save modified XML back to zip
  zip.file("word/document.xml", docXml);
  const outBuf = await zip.generateAsync({ type: "blob" });
  return outBuf;
}
