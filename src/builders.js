import { Document, Packer, Paragraph, TextRun, HeadingLevel, PageBreak } from "docx";

/**
 * Helper to create a styled TextRun with a default size (half-points).
 * The 'docx' lib expects size in half-points; we pass it consistently.
 */
function run(text, opts = {}, defaultHalfPts = null) {
  const base = { text };
  if (defaultHalfPts) base.size = defaultHalfPts;
  return new TextRun({ ...base, ...opts });
}

// Build runs with styles while keeping default font size from original
function partitionTextToRuns(text, ranges, defaultHalfPts) {
  const runs = [];
  let cursor = 0;
  const sorted = [...ranges].sort((a, b) => a.start - b.start);

  for (const r of sorted) {
    if (cursor < r.start) {
      runs.push(run(text.slice(cursor, r.start), {}, defaultHalfPts));
      cursor = r.start;
    }
    const style =
      r.kind === "exact"
        ? { color: "FF0000", italics: true, underline: {} } // red italic underline
        : { color: "7B3F00", underline: {} };               // brown underline for partial/“AI-added”
    runs.push(run(text.slice(r.start, r.end), style, defaultHalfPts));
    cursor = r.end;
  }
  if (cursor < text.length) runs.push(run(text.slice(cursor), {}, defaultHalfPts));
  return runs;
}

/**
 * <doc> REVIEWED.docx
 * - exact matches: red/italic/underline
 * - partial/normalized matches: brown/underline
 * - appendix: JSON items not found (brown)
 * - font size: preserved via original default half-points
 */
export function makeReviewedDocumentDocx({
  originalName,
  text,
  matchedHits,
  notFoundList,
  defaultFontHalfPts, // <= new
}) {
  const ranges = [];
  matchedHits.forEach(({ hits, status }) =>
    hits.forEach((h) => {
      if (h.start === -1) return; // normalized match—no precise span
      ranges.push({ start: h.start, end: h.end, kind: status === "exact" ? "exact" : "partial" });
    })
  );

  const runs = partitionTextToRuns(text, ranges, defaultFontHalfPts);

  const children = [
    new Paragraph({ text: `${originalName} — REVIEWED`, heading: HeadingLevel.HEADING_1 }),
    new Paragraph({}),
    new Paragraph({ children: runs }),
  ];

  if (notFoundList.length) {
    children.push(new Paragraph({ children: [new PageBreak()] }));
    children.push(
      new Paragraph({
        text: "Not found in document (detected but no exact match):",
        heading: HeadingLevel.HEADING_2,
      })
    );
    notFoundList.forEach((d) =>
      children.push(
        new Paragraph({
          children: [
            run(`${d.type}: `, { bold: true }, defaultFontHalfPts),
            run(`${d.value}`, { color: "7B3F00", italics: true }, defaultFontHalfPts),
          ],
        })
      )
    );
  }

  const doc = new Document({ sections: [{ children }] });
  return Packer.toBlob(doc);
}

/**
 * <json>.REVIEWED.docx
 * Requirement: keep the original JSON structure/content as-is,
 * and only color the "value" strings + append analyst comments
 * (does not modify the JSON string; comments appear inline visually).
 *
 * Strategy:
 * - We take the original JSON string.
 * - For each detection (in source order), we find the next occurrence of
 *   `"value": "<escaped>"` and mark the **value substring**.
 * - If issues exist, we append a red italic comment immediately after the value.
 */
export function makeReviewedJsonDocx({
  jsonName,
  originalJsonText,   // <= NEW: the exact JSON string
  detections,
  evalResult,
  defaultFontHalfPts, // use a monospace look but preserve size
}) {
  const { notFound, partials, duplicates, notes } = evalResult;

  const isNotFound = (d) => notFound.includes(d);
  const isPartial  = (d) => partials.includes(d);
  const dupFor = (d) => {
    const v = String(d.value || "");
    return duplicates.find((x) => x.value === v) || null;
  };

  // Build list of value spans and comments in the JSON string
  const json = originalJsonText; // unchanged
  const ranges = [];             // [{start,end,color,comment}]
  let cursor = 0;

  // How to find each "value": use a simple sequential search to the next `"value":"..."`
  function nextValueSpan(jsonStr, from, value) {
    const needle = `"value":`;
    const pos = jsonStr.indexOf(needle, from);
    if (pos === -1) return null;
    // find first quote after colon
    const q1 = jsonStr.indexOf('"', pos + needle.length);
    if (q1 === -1) return null;
    // find closing quote taking escaped quotes into account
    // naive but OK for typical content:
    let q2 = q1 + 1;
    while (q2 < jsonStr.length) {
      if (jsonStr[q2] === '"' && jsonStr[q2 - 1] !== '\\') break;
      q2++;
    }
    const foundValue = jsonStr.slice(q1 + 1, q2);
    return { start: q1 + 1, end: q2, foundValue, after: q2 + 1 };
  }

  // For each detection in order, find the next "value" span and color it
  for (const d of detections) {
    const span = nextValueSpan(json, cursor, d.value);
    if (!span) break; // no more "value" fields
    cursor = span.after;

    const issues = [];
    if (isNotFound(d)) issues.push("Not found in document.");
    if (isPartial(d)) issues.push("Partial/boundary mismatch.");
    const dup = dupFor(d);
    if (dup) issues.push(`Duplicate: JSON ${dup.countInJson} vs Doc ${dup.countInDoc}.`);

    const extra = (notes.get(d) || []).join(" ");
    const hasIssues = issues.length || extra.length;

    ranges.push({
      start: span.start,
      end: span.end,
      color: hasIssues ? "000000" : "0066CC", // blue if OK, black if issues
      comment: hasIssues ? `  — ${[...issues, extra].filter(Boolean).join(" ")}` : "",
    });
  }

  // Now render the entire JSON string verbatim, injecting style on ranges
  const runs = [];
  let i = 0;
  for (const r of ranges) {
    if (i < r.start) runs.push(run(json.slice(i, r.start), { font: "Courier New" }, defaultFontHalfPts));
    runs.push(run(json.slice(r.start, r.end), { color: r.color, font: "Courier New" }, defaultFontHalfPts));
    if (r.comment) runs.push(run(r.comment, { color: "FF0000", italics: true, font: "Courier New" }, defaultFontHalfPts));
    i = r.end;
  }
  if (i < json.length) {
    runs.push(run(json.slice(i), { font: "Courier New" }, defaultFontHalfPts));
  }

  const doc = new Document({
    sections: [
      {
        children: [
          new Paragraph({ text: `${jsonName} — REVIEWED`, heading: HeadingLevel.HEADING_1 }),
          new Paragraph({}),
          // Put the JSON text in paragraphs split by newline to keep docx happy
          ...json.split("\n").map((line, idx) => {
            // consume from 'runs' in order by slicing per line
            // simpler approach: rebuild line-by-line from original string
            // We'll re-scan ranges per line (compact but reliable)
            const lineStart = json.split("\n", idx).slice(0, idx).join("\n").length + (idx ? 1 : 0);
            const lineEnd = lineStart + line.length;
            // collect runs overlapping this line
            const lineRuns = [];
            let cursor = lineStart;

            // get all ranges overlapping this line
            const overlapped = ranges.filter(rr => rr.start < lineEnd && rr.end > lineStart);
            if (!overlapped.length) {
              return new Paragraph({ children: [run(line, { font: "Courier New" }, defaultFontHalfPts)] });
            }
            for (const rr of overlapped) {
              if (cursor < rr.start) {
                lineRuns.push(run(json.slice(cursor, rr.start), { font: "Courier New" }, defaultFontHalfPts));
                cursor = rr.start;
              }
              const segEnd = Math.min(rr.end, lineEnd);
              lineRuns.push(run(json.slice(cursor, segEnd), { color: rr.color, font: "Courier New" }, defaultFontHalfPts));
              if (segEnd === rr.end && rr.comment) {
                lineRuns.push(run(rr.comment, { color: "FF0000", italics: true, font: "Courier New" }, defaultFontHalfPts));
              }
              cursor = segEnd;
            }
            if (cursor < lineEnd) {
              lineRuns.push(run(json.slice(cursor, lineEnd), { font: "Courier New" }, defaultFontHalfPts));
            }
            return new Paragraph({ children: lineRuns });
          }),
        ],
      },
    ],
  });

  return Packer.toBlob(doc);
}
