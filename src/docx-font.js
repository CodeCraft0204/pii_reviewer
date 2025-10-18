import JSZip from "jszip";
import { XMLParser } from "fast-xml-parser";

// returns default run font size in half-points (w:sz @ w:val), or null if not found
export async function readDefaultFontSizeHalfPtsFromDocx(arrayBuffer) {
  try {
    const zip = await JSZip.loadAsync(arrayBuffer);
    const stylesXml = await zip.file("word/styles.xml")?.async("string");
    if (!stylesXml) return null;

    const parser = new XMLParser({
      ignoreAttributes: false,
      attributeNamePrefix: "",
    });
    const xml = parser.parse(stylesXml);

    // w:styles/w:docDefaults/w:rPrDefault/w:rPr/w:sz @ w:val
    const styles   = xml["w:styles"];
    const defaults = styles?.["w:docDefaults"];
    const rPrDef   = defaults?.["w:rPrDefault"];
    const rPr      = rPrDef?.["w:rPr"];
    const sz       = rPr?.["w:sz"];

    // w:sz has shape { "w:val": "22" } (half-points)
    const halfPts = Number(sz?.["w:val"]);
    return Number.isFinite(halfPts) ? halfPts : null;
  } catch {
    return null;
  }
}
