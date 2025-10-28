/**
 * Text masking module for DOC/DOCX/PDF documents
 * Masks detected PII with 'X' characters
 */

import JSZip from "jszip";
import { PDFDocument } from "pdf-lib";
import mammoth from "mammoth";

/**
 * Find all occurrences of a value in text
 * @param {string} text - The text to search in
 * @param {string} value - The value to find
 * @returns {Array<{start: number, end: number}>} - Array of ranges
 */
function findAllOccurrences(text, value) {
  const results = [];
  if (!value || !text) return results;
  
  // Try multiple variations of the value
  const variations = [
    value, // Original
    value.replace(/\s+/g, ''), // No spaces
    value.replace(/\s+/g, ' '), // Single spaces
    value.replace(/\s+/g, '  '), // Double spaces
  ];
  
  for (const variation of variations) {
    let index = 0;
    while ((index = text.indexOf(variation, index)) !== -1) {
      results.push({ start: index, end: index + variation.length });
      index += variation.length;
    }
  }
  
  // Remove duplicates and sort
  const uniqueResults = results.filter((item, index, self) => 
    index === self.findIndex(t => t.start === item.start && t.end === item.end)
  );
  
  return uniqueResults.sort((a, b) => a.start - b.start);
}

/**
 * Mask text by replacing characters with 'X'
 * @param {string} text - Original text
 * @param {Array<{start: number, end: number}>} ranges - Ranges to mask
 * @returns {string} - Masked text
 */
function maskTextRanges(text, ranges) {
  if (!ranges || ranges.length === 0) return text;
  
  // Sort ranges by start position
  const sortedRanges = [...ranges].sort((a, b) => a.start - b.start);
  
  let result = '';
  let lastEnd = 0;
  
  for (const range of sortedRanges) {
    // Add text before this range
    result += text.substring(lastEnd, range.start);
    
    // Add masked text (preserve whitespace)
    const originalText = text.substring(range.start, range.end);
    const masked = originalText.replace(/[^\s]/g, 'X');
    result += masked;
    
    lastEnd = range.end;
  }
  
  // Add remaining text
  result += text.substring(lastEnd);
  
  return result;
}

/**
 * Mask DOCX file by replacing PII with 'X' characters
 * @param {ArrayBuffer} docxArrayBuffer - Original DOCX file
 * @param {Array<{type: string, value: string}>} detections - PII detections
 * @returns {Promise<Blob>} - Masked DOCX file
 */
export async function maskDocx(docxArrayBuffer, detections) {
  const zip = await JSZip.loadAsync(docxArrayBuffer);
  let docXml = await zip.file("word/document.xml")?.async("string");
  if (!docXml) throw new Error("word/document.xml not found");

  // Extract text using mammoth for proper spacing
  const { value: docTextRaw } = await mammoth.extractRawText({ arrayBuffer: docxArrayBuffer });
  const docText = (docTextRaw || "").replace(/\r/g, "");

  console.log("Extracted text preview:", docText.substring(0, 500));
  console.log("Looking for:", detections.map(d => d.value));

  // Find all ranges to mask
  const allRanges = [];
  for (const det of detections || []) {
    const val = String(det.value || "").trim();
    if (!val) continue;
    
    console.log(`Searching for "${val}" in text...`);
    const ranges = findAllOccurrences(docText, val);
    console.log(`Found ${ranges.length} occurrences:`, ranges);
    allRanges.push(...ranges);
  }

  console.log("All ranges to mask:", allRanges);

  // Merge overlapping ranges
  const mergedRanges = mergeRanges(allRanges);
  console.log("Merged ranges:", mergedRanges);

  // Handle split text nodes by working with individual text nodes
  console.log("Starting masking process...");
  
  // Extract all text nodes and their positions
  const textNodes = [];
  docXml.replace(/<w:t[^>]*>(.*?)<\/w:t>/g, (match, content, offset) => {
    textNodes.push({ content, offset, match });
  });
  
  console.log("Found text nodes:", textNodes.map(n => `"${n.content}"`));
  
  // Reconstruct full text from text nodes
  const fullText = textNodes.map(n => n.content).join('');
  console.log("Reconstructed full text:", fullText.substring(0, 200));
  
  // Apply masking to the full text
  let maskedFullText = fullText;
  for (const det of detections || []) {
    const val = String(det.value || "").trim();
    if (!val) continue;
    
    const maskedVal = val.replace(/[^\s]/g, 'X');
    const beforeMask = maskedFullText;
    maskedFullText = maskedFullText.replace(new RegExp(escapeRegExp(val), 'g'), maskedVal);
    
    if (maskedFullText !== beforeMask) {
      console.log(`✅ Masked "${val}" -> "${maskedVal}"`);
    }
  }
  
  console.log("Masked full text:", maskedFullText.substring(0, 200));
  
  // Now distribute the masked text back to the text nodes
  let textIndex = 0;
  let resultXml = docXml;
  
  resultXml = resultXml.replace(/<w:t[^>]*>(.*?)<\/w:t>/g, (match, content) => {
    const originalLength = content.length;
    const maskedContent = maskedFullText.substring(textIndex, textIndex + originalLength);
    textIndex += originalLength;
    return match.replace(content, maskedContent);
  });
  
  docXml = resultXml;

  // Save modified XML back to zip
  zip.file("word/document.xml", docXml);
  const outBuf = await zip.generateAsync({ type: "blob" });
  return outBuf;
}

/**
 * Replace text in XML while preserving structure
 * @param {string} xmlString - Original XML
 * @param {string} originalText - Text to find
 * @param {string} maskedText - Replacement text
 * @returns {string} - Modified XML
 */
function replaceTextInXml(xmlString, originalText, maskedText) {
  console.log(`replaceTextInXml: Looking for "${originalText}" to replace with "${maskedText}"`);
  
  // For credit cards, use a nuclear approach - replace ALL digits in the document
  if (originalText.match(/\d{4}\s*\d{4}\s*\d{4}\s*\d{4}/)) {
    console.log(`Handling credit card with nuclear approach: "${originalText}"`);
    
    // Extract the specific digits from this credit card
    const ccDigits = originalText.replace(/\s+/g, '').split('');
    console.log(`Credit card digits to replace: ${ccDigits}`);
    
    // Only replace within <w:t> tags to avoid breaking XML structure
    const tTagRegex = /<w:t(?:\s+[^>]*)?>([^<]*)<\/w:t>/g;
    
    xmlString = xmlString.replace(tTagRegex, (match, textContent) => {
      let newContent = textContent;
      let modified = false;
      
      // Replace each specific digit from the credit card
      for (const digit of ccDigits) {
        if (textContent.includes(digit)) {
          console.log(`Replacing digit "${digit}" in text node: "${textContent}"`);
          newContent = newContent.replace(new RegExp(digit, 'g'), 'X');
          modified = true;
        }
      }
      
      if (modified) {
        console.log(`Modified text node: "${textContent}" -> "${newContent}"`);
        return match.replace(textContent, newContent);
      }
      
      return match;
    });
    
    return xmlString;
  }
  
  // For other text, use the original approach
  const tTagRegex = /<w:t(?:\s+[^>]*)?>([^<]*)<\/w:t>/g;
  
  let modified = xmlString;
  let match;
  let found = false;
  
  // Reset regex
  tTagRegex.lastIndex = 0;
  
  while ((match = tTagRegex.exec(xmlString)) !== null) {
    const fullMatch = match[0];
    const textContent = match[1];
    
    console.log(`Found text node: "${textContent}"`);
    
     // Check if this text node contains our target text
     if (textContent.includes(originalText)) {
       console.log(`Found exact match! Replacing "${originalText}" with "${maskedText}"`);
       const newContent = textContent.replace(originalText, maskedText);
       const newTag = fullMatch.replace(textContent, newContent);
       modified = modified.replace(fullMatch, newTag);
       found = true;
       break; // Only replace first occurrence
     }
     
     // Check for partial matches (for credit cards that might be truncated)
     if (originalText.length > 10 && textContent.length > 10) {
       // Try matching the first part of the original text
       const partialLength = Math.min(originalText.length - 1, textContent.length);
       const partialOriginal = originalText.substring(0, partialLength);
       const partialMasked = partialOriginal.replace(/[^\s]/g, 'X');
       
       if (textContent.includes(partialOriginal)) {
         console.log(`Found partial match! Replacing "${partialOriginal}" with "${partialMasked}"`);
         const newContent = textContent.replace(partialOriginal, partialMasked);
         const newTag = fullMatch.replace(textContent, newContent);
         modified = modified.replace(fullMatch, newTag);
         found = true;
         break; // Only replace first occurrence
       }
     }
  }
  
  if (!found) {
    console.log(`No XML text node found containing "${originalText}"`);
  }
  
  return modified;
}

/**
 * Merge overlapping ranges
 * @param {Array<{start: number, end: number}>} ranges
 * @returns {Array<{start: number, end: number}>}
 */
function mergeRanges(ranges) {
  if (ranges.length === 0) return [];
  
  const sorted = [...ranges].sort((a, b) => a.start - b.start);
  const merged = [sorted[0]];
  
  for (let i = 1; i < sorted.length; i++) {
    const current = sorted[i];
    const last = merged[merged.length - 1];
    
    if (current.start <= last.end) {
      // Overlapping, merge
      last.end = Math.max(last.end, current.end);
    } else {
      // Not overlapping, add new range
      merged.push(current);
    }
  }
  
  return merged;
}

/**
 * Mask PDF file by replacing PII with 'X' characters
 * @param {ArrayBuffer} pdfArrayBuffer - Original PDF file
 * @param {Array<{type: string, value: string}>} detections - PII detections
 * @param {string} extractedText - Pre-extracted text from PDF
 * @returns {Promise<Blob>} - Masked PDF file
 */
export async function maskPdf(pdfArrayBuffer, detections, extractedText) {
  // For text-based PDFs, we'll create a new PDF with masked text
  // This is complex, so for now we'll create a simple text-based PDF
  
  const pdfDoc = await PDFDocument.load(pdfArrayBuffer);
  
  // Find all ranges to mask
  const allRanges = [];
  for (const det of detections || []) {
    const val = String(det.value || "").trim();
    if (!val) continue;
    
    const ranges = findAllOccurrences(extractedText, val);
    allRanges.push(...ranges);
  }

  // Merge overlapping ranges
  const mergedRanges = mergeRanges(allRanges);
  
  // Mask the text
  const maskedText = maskTextRanges(extractedText, mergedRanges);
  
  // Create a new PDF with masked text
  const newPdfDoc = await PDFDocument.create();
  const pages = pdfDoc.getPages();
  
  // For simplicity, we'll draw the masked text on new pages
  // In a production system, you'd want to preserve the original layout
  for (let i = 0; i < pages.length; i++) {
    const originalPage = pages[i];
    const { width, height } = originalPage.getSize();
    const newPage = newPdfDoc.addPage([width, height]);
    
    // Draw masked text (simplified - in production, you'd need proper text layout)
    newPage.drawText(maskedText.substring(i * 1000, (i + 1) * 1000), {
      x: 50,
      y: height - 50,
      size: 12,
    });
  }
  
  const pdfBytes = await newPdfDoc.save();
  return new Blob([pdfBytes], { type: 'application/pdf' });
}

/**
 * Mask DOC file (older Word format)
 * @param {ArrayBuffer} docArrayBuffer - Original DOC file
 * @param {Array<{type: string, value: string}>} detections - PII detections
 * @returns {Promise<Blob>} - Masked DOC file (converted to DOCX)
 */
export async function maskDoc(docArrayBuffer, detections) {
  // Convert DOC to DOCX first using mammoth, then mask
  // Note: This will lose some formatting, but it's the simplest approach
  const { value: docTextRaw } = await mammoth.extractRawText({ arrayBuffer: docArrayBuffer });
  const docText = (docTextRaw || "").replace(/\r/g, "");
  
  // Find all ranges to mask
  const allRanges = [];
  for (const det of detections || []) {
    const val = String(det.value || "").trim();
    if (!val) continue;
    
    const ranges = findAllOccurrences(docText, val);
    allRanges.push(...ranges);
  }

  // Merge overlapping ranges
  const mergedRanges = mergeRanges(allRanges);
  
  // Mask the text
  const maskedText = maskTextRanges(docText, mergedRanges);
  
  // Create a simple DOCX with the masked text
  // For production, you'd want to use a proper DOCX library
  return createSimpleDocx(maskedText);
}

/**
 * Create a simple DOCX file with text
 * @param {string} text - Text content
 * @returns {Promise<Blob>}
 */
async function createSimpleDocx(text) {
  const zip = new JSZip();
  
  // Minimal DOCX structure
  const docXml = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
  <w:body>
    <w:p>
      <w:r>
        <w:t xml:space="preserve">${escapeXml(text)}</w:t>
      </w:r>
    </w:p>
  </w:body>
</w:document>`;

  const contentTypes = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
  <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
  <Default Extension="xml" ContentType="application/xml"/>
  <Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>
</Types>`;

  const rels = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/>
</Relationships>`;

  zip.file("[Content_Types].xml", contentTypes);
  zip.file("_rels/.rels", rels);
  zip.file("word/document.xml", docXml);
  
  return await zip.generateAsync({ type: "blob" });
}

/**
 * Escape XML special characters
 * @param {string} str
 * @returns {string}
 */
function escapeRegExp(string) {
  return string.replace(/[.*+?^${}()|[\]\\]/g, '\\$&');
}

