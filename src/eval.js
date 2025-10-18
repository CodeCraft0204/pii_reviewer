// ---- File utils ----
export const readFileAsArrayBuffer = (file) =>
  new Promise((resolve, reject) => {
    const fr = new FileReader();
    fr.onload = () => resolve(fr.result);
    fr.onerror = reject;
    fr.readAsArrayBuffer(file);
  });

export const readFileAsText = (file) =>
  new Promise((resolve, reject) => {
    const fr = new FileReader();
    fr.onload = () => resolve(fr.result);
    fr.onerror = reject;
    fr.readAsText(file, "utf-8");
  });

// ---- Normalizers & detectors ----
export const normalizePhone = (s) => (s || "").replace(/[^\d+]/g, "");
export const isEmail = (s) => /@/.test(s);
export const isPhone = (s) => /^[+]?[\d ()-]{6,}$/.test(s);

// Simple NFKC normalization helps with accents/combining chars
export const nfkc = (s) => (s || "").normalize("NFKC");

// Non-overlapping occurrences
export function findAllOccurrences(text, needle) {
  if (!needle) return [];
  const hits = [];
  let idx = 0;
  while (idx < text.length) {
    const found = text.indexOf(needle, idx);
    if (found === -1) break;
    hits.push([found, found + needle.length]);
    idx = found + needle.length;
  }
  return hits;
}

// Token boundary check to avoid substrings inside words
export function hasTokenBoundary(full, start, end) {
  const left = start <= 0 ? " " : full[start - 1];
  const right = end >= full.length ? " " : full[end];
  const isBoundary = (c) => /\s|[.,;:!?()[\]{}<>$"'\-]|^$/.test(c);
  return isBoundary(left) && isBoundary(right);
}

// Levenshtein distance for near-miss hints
export function levenshtein(a, b) {
  a = a || ""; b = b || "";
  const dp = Array.from({ length: a.length + 1 }, (_, i) => [i]);
  for (let j = 1; j <= b.length; j++) dp[0][j] = j;
  for (let i = 1; i <= a.length; i++) {
    for (let j = 1; j <= b.length; j++) {
      dp[i][j] = Math.min(
        dp[i - 1][j] + 1,
        dp[i][j - 1] + 1,
        dp[i - 1][j - 1] + (a[i - 1] === b[j - 1] ? 0 : 1)
      );
    }
  }
  return dp[a.length][b.length];
}
export const similarity = (a, b) => {
  const maxLen = Math.max(a?.length || 0, b?.length || 0);
  if (!maxLen) return 0;
  const dist = levenshtein(nfkc(a), nfkc(b));
  return 1 - dist / maxLen;
};

// Build candidate substrings near a hit (for near-miss suggestion)
function windowAround(text, start, end, extra = 12) {
  const s = Math.max(0, start - extra);
  const e = Math.min(text.length, end + extra);
  return text.slice(s, e);
}

/**
 * Core evaluation:
 * returns {
 *   matched: [{ det, hits:[{start,end}|{start:-1,normalized:true}], status:'exact'|'partial' }],
 *   notFound:[det],
 *   duplicates:[{value,countInJson,countInDoc,extra}],
 *   partials:[det],
 *   occurrences: Map(value -> [{start,end}]),
 *   notes: Map(det -> string[])   // natural-language analyst comments
 * }
 */
export function evaluateDetections({ docText, detections }) {
  const text = nfkc(docText || "");
  const matched = [];
  const notFound = [];
  const partials = [];
  const duplicates = [];
  const occurrences = new Map();
  const notes = new Map();

  const byValue = new Map();
  detections.forEach((d) => {
    const key = nfkc(String(d.value || "").trim());
    if (!byValue.has(key)) byValue.set(key, []);
    byValue.get(key).push(d);
  });

  function addNote(d, msg) {
    const arr = notes.get(d) || [];
    arr.push(msg);
    notes.set(d, arr);
  }

  for (const d of detections) {
    const raw = String(d.value || "");
    const v = nfkc(raw);
    if (!v) { notFound.push(d); addNote(d, "Empty value in JSON."); continue; }

    // Debug: Check if the value exists in the text
    const directMatch = text.includes(v);
    console.log(`Checking "${v}": direct match = ${directMatch}`);

    // phone: normalized stream
    const hits = isPhone(v)
      ? (() => {
          const normDoc = normalizePhone(text);
          const normVal = normalizePhone(v);
          const idxes = findAllOccurrences(normDoc, normVal);
          return idxes.length ? [{ start: -1, end: -1, normalized: true }] : [];
        })()
      : findAllOccurrences(text, v).map(([s, e]) => ({ start: s, end: e }));
    
    console.log(`Found ${hits.length} hits for "${v}"`);

    if (!hits.length) {
      // Try near-miss suggestion (look around similar tokens)
      const near = [];
      for (let i = 0; i < text.length - v.length + 1; i++) {
        const segment = text.slice(i, i + v.length + 4); // slightly larger window
        if (similarity(v, segment) >= 0.7) {
          near.push(segment);
          if (near.length >= 3) break;
        }
      }
      if (near.length) addNote(d, `Not found; nearest text candidates: ${near.map(x => `"${x}"`).join(", ")}.`);
      notFound.push(d);
      continue;
    }

    const good = hits.filter(
      (h) => h.normalized || hasTokenBoundary(text, h.start, h.end) || isEmail(v)
    );
    if (!good.length) {
      partials.push(d);
      addNote(d, "Found only as a substring (token boundary mismatch).");
      continue;
    }

    occurrences.set(v, (occurrences.get(v) || []).concat(good));
    matched.push({ det: d, hits: good, status: good.some(h => h.start === -1) ? 'partial' : 'exact' });

    // extra analyst hints: if multiple occurrences but JSON listed once, or vice versa
    if (good.length > 1) addNote(d, `Appears ${good.length} times in document.`);
  }

  // duplicates: JSON count > doc occurrences
  for (const [val, jsonList] of byValue.entries()) {
    const occ = occurrences.get(val) || [];
    if (jsonList.length > occ.length) {
      duplicates.push({
        value: val,
        countInJson: jsonList.length,
        countInDoc: occ.length,
        extra: jsonList.length - occ.length,
      });
      // add notes to the extra JSON items
      jsonList.forEach((d, idx) => {
        if (idx >= occ.length) addNote(d, `Duplicate: JSON lists ${jsonList.length}, but document has ${occ.length}.`);
      });
    }
  }

  // classify partials as “near-miss” if a boundary window shows high similarity
  for (const d of partials) {
    const v = nfkc(String(d.value || ""));
    const occs = findAllOccurrences(text, v);
    for (const [s, e] of occs) {
      const w = windowAround(text, s, e, 14);
      if (similarity(v, w) >= 0.75) {
        addNote(d, `Partial match near “…${w}…”. Consider boundary/spacing/diacritics issues.`);
        break;
      }
    }
  }

  return { matched, notFound, partials, duplicates, occurrences, notes };
}
