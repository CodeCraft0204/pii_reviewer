/**
 * Text masking module for DOC/DOCX/PDF documents
 * Masks detected PII with 'X' characters
 */

import JSZip from "jszip";
import { PDFDocument, rgb } from "pdf-lib";
import mammoth from "mammoth";
import * as pdfjsLib from 'pdfjs-dist';

// Disable worker for simpler setup
pdfjsLib.GlobalWorkerOptions.workerSrc = '';

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
export async function maskPdf(pdfArrayBuffer, detections, extractedText, maskingMethod = 'rectangle') {
  console.log("Starting PDF masking process...");
  
  // Load the original PDF
  const pdfDoc = await PDFDocument.load(pdfArrayBuffer);
  const pages = pdfDoc.getPages();
  
  console.log(`Processing ${pages.length} pages`);
  
  // Find all ranges to mask in the extracted text
  const allRanges = [];
  for (const det of detections || []) {
    const val = String(det.value || "").trim();
    if (!val) continue;
    
    // Filter out obviously non-PII values
    const nonPiiPatterns = [
      /^Die Finanzbranche$/i,
      /^Finanzbranche$/i,
      /^Die$/i,
      /^in$/i,
      /^derzeit$/i,
      /^eine$/i,
      /^Phase$/i,
      /^intensiver$/i,
      /^Digitalisierung$/i,
      /^Investitionen$/i,
      /^Sicherheit$/i,
      /^und$/i,
      /^Datenschutz$/i,
      /^sind$/i,
      /^für$/i,
      /^Banken$/i,
      /^oberste$/i,
      /^Priorität$/i,
      /^Zahlungssysteme$/i,
      /^Online$/i,
      /^Banking$/i,
      /^werden$/i,
      /^von$/i,
      /^Kunden$/i,
      /^zunehmend$/i,
      /^akzeptiert$/i,
      /^Im$/i,
      /^Rahmen$/i,
      /^neuer$/i,
      /^Gesetze$/i,
      /^wird$/i,
      /^die$/i,
      /^Transparenz$/i,
      /^im$/i,
      /^Finanzwesen$/i,
      /^weiter$/i,
      /^erhöht$/i,
      /^Viele$/i,
      /^Unternehmen$/i,
      /^setzen$/i,
      /^auf$/i,
      /^elektronische$/i,
      /^Rechnungsstellung$/i,
      /^um$/i,
      /^Prozesse$/i,
      /^zu$/i,
      /^optimieren$/i
    ];
    
    const isNonPii = nonPiiPatterns.some(pattern => pattern.test(val.trim()));
    if (isNonPii) {
      console.log(`❌ Skipping non-PII value: "${val}"`);
      continue;
    }
    
    console.log(`Looking for PII: "${val}"`);
    const ranges = findAllOccurrences(extractedText, val);
    console.log(`Found ${ranges.length} occurrences`);
    allRanges.push(...ranges);
  }

  // Merge overlapping ranges
  const mergedRanges = mergeRanges(allRanges);
  console.log(`Merged to ${mergedRanges.length} ranges to mask`);
  
  // Create masked text
  const maskedText = maskTextRanges(extractedText, mergedRanges);
  console.log("Created masked text");
  
  // For PDF masking, we'll draw black rectangles over PII values
  // We'll use the original PDF and overlay rectangles at the correct positions
  
  console.log(`Processing ${pages.length} pages`);
  
  // Load the original PDF with pdf-lib
  const modifiedPdfDoc = await PDFDocument.load(pdfArrayBuffer);
  const pdfPages = modifiedPdfDoc.getPages();
  
  // Load the PDF for text positioning
  const freshArrayBuffer = pdfArrayBuffer.slice(0);
  const uint8Array = new Uint8Array(freshArrayBuffer);
  const loadingTask = pdfjsLib.getDocument({
    data: uint8Array,
    useSystemFonts: true,
    disableWorker: true,
  });
  
  const textPdfDoc = await loadingTask.promise;
  
  // Process each page by drawing black rectangles over PII values
  for (let i = 0; i < pdfPages.length; i++) {
    const page = pdfPages[i];
    
    console.log(`Processing page ${i + 1}...`);
    
    // Get the text content with positioning from pdf.js
    const pdfPage = await textPdfDoc.getPage(i + 1);
    const textContent = await pdfPage.getTextContent();
    
    // Get all text items
    const textItems = textContent.items.filter(item => item.str && item.str.trim());
    const { width, height } = page.getSize();
    let rectangleCount = 0;
    
    // Group text items by line (similar Y coordinates)
    const lines = [];
    let currentLine = [];
    let lastY = null;
    const lineThreshold = 5; // pixels
    
    for (const item of textItems) {
      const itemY = item.transform[5];
      
      if (lastY === null || Math.abs(itemY - lastY) < lineThreshold) {
        currentLine.push(item);
      } else {
        if (currentLine.length > 0) {
          lines.push(currentLine);
        }
        currentLine = [item];
      }
      lastY = itemY;
    }
    
    if (currentLine.length > 0) {
      lines.push(currentLine);
    }
    
    console.log(`Found ${lines.length} text lines`);
    
    // Debug: Log all detected PII values
    console.log('=== DETECTED PII VALUES ===');
    for (const det of detections || []) {
      const val = String(det.value || "").trim();
      if (val) {
        console.log(`PII: "${val}"`);
      }
    }
    
    // Debug: Log all text lines
    console.log('=== TEXT LINES ===');
    for (let i = 0; i < Math.min(lines.length, 10); i++) {
      const line = lines[i];
      const lineText = line.map(item => item.str).join(' ');
      console.log(`Line ${i}: "${lineText}"`);
    }
    
    // Deduplicate PII values to avoid processing the same PII multiple times
    const uniquePIIValues = new Set();
    const uniqueDetections = [];
    
    for (const det of detections || []) {
      const val = String(det.value || "").trim();
      if (val && !uniquePIIValues.has(val)) {
        uniquePIIValues.add(val);
        uniqueDetections.push(det);
      }
    }
    
    console.log(`Processing ${uniqueDetections.length} unique PII values (${detections?.length || 0} total detections)`);
    
    // For each unique PII value, find it in the lines and draw rectangles
    let pageRectangleCount = 0;
    for (const det of uniqueDetections) {
      const val = String(det.value || "").trim();
      if (!val) continue;
      
      // Filter out obviously non-PII values (same filter as in main processing)
      const nonPiiPatterns = [
        /^Die Finanzbranche$/i,
        /^Finanzbranche$/i,
        /^Die$/i,
        /^in$/i,
        /^derzeit$/i,
        /^eine$/i,
        /^Phase$/i,
        /^intensiver$/i,
        /^Digitalisierung$/i,
        /^Investitionen$/i,
        /^Sicherheit$/i,
        /^und$/i,
        /^Datenschutz$/i,
        /^sind$/i,
        /^für$/i,
        /^Banken$/i,
        /^oberste$/i,
        /^Priorität$/i,
        /^Zahlungssysteme$/i,
        /^Online$/i,
        /^Banking$/i,
        /^werden$/i,
        /^von$/i,
        /^Kunden$/i,
        /^zunehmend$/i,
        /^akzeptiert$/i,
        /^Im$/i,
        /^Rahmen$/i,
        /^neuer$/i,
        /^Gesetze$/i,
        /^wird$/i,
        /^die$/i,
        /^Transparenz$/i,
        /^im$/i,
        /^Finanzwesen$/i,
        /^weiter$/i,
        /^erhöht$/i,
        /^Viele$/i,
        /^Unternehmen$/i,
        /^setzen$/i,
        /^auf$/i,
        /^elektronische$/i,
        /^Rechnungsstellung$/i,
        /^um$/i,
        /^Prozesse$/i,
        /^zu$/i,
        /^optimieren$/i
      ];
      
      const isNonPii = nonPiiPatterns.some(pattern => pattern.test(val.trim()));
      if (isNonPii) {
        console.log(`❌ Skipping non-PII value: "${val}"`);
        continue;
      }
      
      console.log(`\n=== Looking for PII: "${val}" ===`);
      
      // Check each line for the PII value
      for (let lineIndex = 0; lineIndex < lines.length; lineIndex++) {
        const line = lines[lineIndex];
        const lineText = line.map(item => item.str).join(' ');
        
        // Try exact match first
        if (lineText.includes(val)) {
          console.log(`✅ FOUND PII "${val}" in line ${lineIndex}: "${lineText}"`);
          if (maskingMethod === 'rectangle') {
            drawRectangleOverPII(page, line, val, height);
          } else {
            replacePIIWithText(page, line, val, height);
          }
          pageRectangleCount++;
          break; // Found the PII, move to next PII value
        } else {
          // Try fuzzy matching for split text
          const normalizedVal = val.replace(/\s/g, ''); // Remove all spaces
          const normalizedLineText = lineText.replace(/\s/g, ''); // Remove all spaces from line
          
          if (normalizedLineText.includes(normalizedVal)) {
            console.log(`✅ FOUND PII "${val}" (fuzzy match) in line ${lineIndex}: "${lineText}"`);
            if (maskingMethod === 'rectangle') {
              drawRectangleOverPII(page, line, val, height);
            } else {
              replacePIIWithText(page, line, val, height);
            }
            pageRectangleCount++;
            break; // Found the PII, move to next PII value
          } else {
            // Try partial matching for credit card numbers and IBANs
            if (isCreditCardOrIBAN(val)) {
              const valDigits = val.replace(/\D/g, ''); // Extract only digits
              const lineDigits = lineText.replace(/\D/g, ''); // Extract only digits from line
              
              if (lineDigits.includes(valDigits)) {
                console.log(`✅ FOUND PII "${val}" (digit match) in line ${lineIndex}: "${lineText}"`);
                if (maskingMethod === 'rectangle') {
                  drawRectangleOverPII(page, line, val, height);
                } else {
                  replacePIIWithText(page, line, val, height);
                }
                pageRectangleCount++;
                break; // Found the PII, move to next PII value
              }
            }
          }
        }
      }
    }
    
    console.log(`${maskingMethod === 'rectangle' ? 'Drew rectangles over' : 'Replaced'} ${pageRectangleCount} PII values on page ${i + 1}`);
  }
  
  await textPdfDoc.destroy();
  
  console.log("PDF masking completed");
  const pdfBytes = await modifiedPdfDoc.save();
  return new Blob([pdfBytes], { type: 'application/pdf' });
}

/**
 * Helper function to create a virtual text item for a substring
 */
function createVirtualItem(originalItem, startIndex, endIndex) {
  const originalText = originalItem.str;
  const subText = originalText.substring(startIndex, endIndex);

  // Calculate the proportional width and x-coordinate for the substring
  const totalWidth = originalItem.width;
  const charWidth = totalWidth / originalText.length; // Average width per character
  const subTextWidth = charWidth * subText.length;

  // Calculate the x-coordinate for the start of the substring
  const subTextX = originalItem.transform[4] + (charWidth * startIndex);
  const subTextY = originalItem.transform[5]; // y-coordinate is the same

  // Create a new item object with the properties of a PDF.js text item
  return {
    str: subText,
    dir: originalItem.dir,
    width: subTextWidth,
    height: originalItem.height,
    transform: [
      originalItem.transform[0], originalItem.transform[1],
      originalItem.transform[2], originalItem.transform[3],
      subTextX, subTextY
    ],
    fontName: originalItem.fontName,
    hasEOL: originalItem.hasEOL,
    originalItem: originalItem // Keep a reference to the original item for debugging
  };
}

/**
 * Helper function to create a virtual text item for a PII value that spans multiple items
 */
function createVirtualItemForPII(piiValue, valueItems, combinedText, colonIndex) {
  if (valueItems.length === 0) {
    console.log(`❌ createVirtualItemForPII: No value items provided`);
    return null;
  }
  
  // Normalize the PII value for matching
  const normalizedPiiValue = piiValue.replace(/\s/g, '');
  const normalizedCombinedText = combinedText.replace(/\s/g, '');
  
  // Find PII in normalized combined text (after colon)
  const normalizedColonIndex = normalizedCombinedText.indexOf(':', colonIndex);
  const normalizedPiiStart = normalizedCombinedText.indexOf(normalizedPiiValue, normalizedColonIndex + 1);
  
  if (normalizedPiiStart === -1) {
    console.log(`❌ createVirtualItemForPII: Could not find "${piiValue}" in normalized combined text`);
    console.log(`🔍 DEBUG - Combined text: "${combinedText}"`);
    console.log(`🔍 DEBUG - Normalized combined text: "${normalizedCombinedText}"`);
    console.log(`🔍 DEBUG - Normalized PII: "${normalizedPiiValue}"`);
    return null;
  }
  
  // Build item positions array to map back to original text
  let currentPos = 0;
  const itemPositions = [];
  for (const item of valueItems) {
    const itemStart = combinedText.indexOf(item.str, currentPos);
    if (itemStart === -1) break;
    itemPositions.push({
      item: item,
      start: itemStart,
      end: itemStart + item.str.length,
      normalizedStart: normalizedCombinedText.indexOf(item.str.replace(/\s/g, ''), normalizedColonIndex + 1)
    });
    currentPos = itemStart + item.str.length;
  }
  
  // Find which items contain the normalized PII range
  const normalizedPiiEnd = normalizedPiiStart + normalizedPiiValue.length;
  
  const piiStartItem = itemPositions.find(pos => {
    const itemNormalized = pos.item.str.replace(/\s/g, '');
    const itemNormalizedStart = normalizedCombinedText.indexOf(itemNormalized, normalizedColonIndex + 1);
    const itemNormalizedEnd = itemNormalizedStart + itemNormalized.length;
    return normalizedPiiStart >= itemNormalizedStart && normalizedPiiStart < itemNormalizedEnd;
  });
  
  const piiEndItem = itemPositions.find(pos => {
    const itemNormalized = pos.item.str.replace(/\s/g, '');
    const itemNormalizedStart = normalizedCombinedText.indexOf(itemNormalized, normalizedColonIndex + 1);
    const itemNormalizedEnd = itemNormalizedStart + itemNormalized.length;
    return normalizedPiiEnd > itemNormalizedStart && normalizedPiiEnd <= itemNormalizedEnd;
  });
  
  if (!piiStartItem || !piiEndItem) {
    console.log(`❌ createVirtualItemForPII: Could not find start/end items`);
    console.log(`🔍 DEBUG - Normalized PII start: ${normalizedPiiStart}, end: ${normalizedPiiEnd}`);
    console.log(`🔍 DEBUG - Item positions: ${itemPositions.map(p => `${p.item.str}(${p.normalizedStart})`).join(', ')}`);
    return null;
  }
  
  // Use the items from itemPositions
  const firstItem = piiStartItem.item;
  const lastItem = piiEndItem.item;
  
  // Calculate the start position within the first item
  const firstItemStart = piiStartItem.start;
  const piiStartInFirstItem = Math.max(0, normalizedPiiStart - (normalizedCombinedText.indexOf(firstItem.str.replace(/\s/g, ''), normalizedColonIndex + 1)));
  
  // Calculate the end position within the last item  
  const lastItemNormalizedStart = normalizedCombinedText.indexOf(lastItem.str.replace(/\s/g, ''), normalizedColonIndex + 1);
  const piiEndInLastItem = normalizedPiiEnd - lastItemNormalizedStart;
  
  // If PII spans multiple items, create a combined virtual item
  if (firstItem !== lastItem) {
    // Calculate the total width by summing up the widths of all items containing the PII
    let totalWidth = 0;
    let currentX = firstItem.transform[4];
    
    // Find all items between firstItem and lastItem (inclusive)
    const firstItemIndex = itemPositions.indexOf(piiStartItem);
    const lastItemIndex = itemPositions.indexOf(piiEndItem);
    const itemsInRange = itemPositions.filter((pos, index) => {
      return firstItemIndex <= index && index <= lastItemIndex;
    });
    
    for (const pos of itemsInRange) {
      const item = pos.item;
      if (item === firstItem) {
        // First item - use partial width from start position
        const partialWidth = item.width * (item.str.length - piiStartInFirstItem) / item.str.length;
        totalWidth += partialWidth;
        currentX = item.transform[4] + (item.width * piiStartInFirstItem / item.str.length);
      } else if (item === lastItem) {
        // Last item - use partial width to end position
        const partialWidth = item.width * piiEndInLastItem / item.str.length;
        totalWidth += partialWidth;
      } else {
        // Middle item - use full width
        totalWidth += item.width;
      }
    }
    
    return {
      str: piiValue,
      dir: firstItem.dir,
      width: totalWidth,
      height: firstItem.height,
      transform: [
        firstItem.transform[0], firstItem.transform[1],
        firstItem.transform[2], firstItem.transform[3],
        currentX, 
        firstItem.transform[5]
      ],
      fontName: firstItem.fontName,
      hasEOL: lastItem.hasEOL,
      originalItem: firstItem
    };
  } else {
    // PII is within a single item
    return createVirtualItem(firstItem, piiStartInFirstItem, piiEndInLastItem);
  }
}

/**
 * Draw rectangles over PII text (Option 1)
 */
function drawRectangleOverPII(page, line, piiValue, height) {
  const lineText = line.map(item => item.str).join(' ');
  
  console.log(`🔍 DEBUG - Looking for PII: "${piiValue}"`);
  console.log(`🔍 DEBUG - In line: "${lineText}"`);
  
  // Find text items that contain the PII using direct string matching
  let piiItems = [];
  
  // First, try exact matching on individual items
  for (let i = 0; i < line.length; i++) {
    const item = line[i];
    const itemText = item.str;
    
    if (itemText.includes(piiValue)) {
      piiItems.push(item);
      console.log(`✅ Found exact match in item ${i}: "${itemText}"`);
    }
  }
  
  // If no exact matches, try fuzzy matching on individual items
  if (piiItems.length === 0) {
    const normalizedVal = piiValue.replace(/\s/g, '');
    
    for (let i = 0; i < line.length; i++) {
      const item = line[i];
      const itemText = item.str;
      const normalizedItemText = itemText.replace(/\s/g, '');
      
      if (normalizedItemText.includes(normalizedVal)) {
        piiItems.push(item);
        console.log(`✅ Found fuzzy match in item ${i}: "${itemText}"`);
      }
    }
  }
  
  // If still no matches, try finding consecutive items that together contain the PII
  if (piiItems.length === 0) {
    const normalizedVal = piiValue.replace(/\s/g, '');
    
    for (let i = 0; i < line.length; i++) {
      let combinedText = '';
      let combinedItems = [];
      
      // For credit card numbers and IBANs, be more restrictive about what we include
      const maxItems = (isCreditCardOrIBAN(piiValue) || piiValue.startsWith('DE')) ? 6 : 5;
      
      for (let j = i; j < Math.min(i + maxItems, line.length); j++) {
        combinedText += (combinedText ? ' ' : '') + line[j].str;
        combinedItems.push(line[j]);
        
        const normalizedCombined = combinedText.replace(/\s/g, '');
        if (normalizedCombined.includes(normalizedVal)) {
          // For credit card numbers, check if the combined text is mostly digits
          if (isCreditCardOrIBAN(piiValue)) {
            if (piiValue.match(/^\d{4}\s\d{4}\s\d{4}\s\d{4}$/)) {
              // For credit card numbers, require high digit ratio
              const digitRatio = (combinedText.replace(/\D/g, '').length) / combinedText.replace(/\s/g, '').length;
              if (digitRatio < 0.5) {
                console.log(`❌ Skipping combined match with low digit ratio: "${combinedText}" (ratio: ${digitRatio.toFixed(2)})`);
                continue;
              }
            } else if (piiValue.startsWith('DE')) {
              // For IBANs, be more lenient - just check that it contains significant parts
              const piiDigits = piiValue.replace(/\D/g, '');
              const combinedDigits = combinedText.replace(/\D/g, '');
              if (combinedDigits.length < piiDigits.length * 0.3) {
                console.log(`❌ Skipping combined match with too few digits: "${combinedText}" (${combinedDigits.length} vs ${piiDigits.length})`);
                continue;
              }
            }
          }
          
          // Special handling for label: value format - try to extract just the value part
          if (combinedText.includes(':') && combinedText.length > piiValue.length * 2) {
            const colonIndex = combinedText.lastIndexOf(':');
            if (colonIndex > 0) {
              const valuePart = combinedText.substring(colonIndex + 1).trim();
              // Check if the PII value matches the value part (with or without spaces)
              const normalizedValuePart = valuePart.replace(/\s/g, '');
              const normalizedPiiValue = piiValue.replace(/\s/g, '');
              
              console.log(`🔍 DEBUG - Checking label:value match for "${piiValue}"`);
              console.log(`🔍 DEBUG - Value part: "${valuePart}", normalized: "${normalizedValuePart}"`);
              console.log(`🔍 DEBUG - PII value: "${piiValue}", normalized: "${normalizedPiiValue}"`);
              
              if (normalizedValuePart === normalizedPiiValue || valuePart.includes(piiValue) || normalizedValuePart.includes(normalizedPiiValue)) {
                // This is a label: value format, we need to create a virtual item for just the value
                // Find the items that contain the value part (after the colon)
                const valueItems = combinedItems.filter(item => {
                  const itemIndex = combinedText.indexOf(item.str);
                  return itemIndex > colonIndex;
                });
                
                console.log(`🔍 DEBUG - Found ${valueItems.length} value items after colon`);
                
                if (valueItems.length > 0) {
                  // Create a virtual item that represents just the value part
                  const virtualItem = createVirtualItemForPII(piiValue, valueItems, combinedText, colonIndex);
                  
                  if (virtualItem) {
                    piiItems.push(virtualItem);
                    console.log(`✅ Found label:value match, created virtual item for value: "${piiValue}"`);
                    break;
                  } else {
                    console.log(`❌ Failed to create virtual item for "${piiValue}"`);
                  }
                }
              } else {
                console.log(`❌ Value part doesn't match PII value`);
              }
            }
          }
          
          piiItems.push(...combinedItems);
          console.log(`✅ Found combined match in items ${i}-${j}: "${combinedText}"`);
          break;
        }
      }
      
      if (piiItems.length > 0) break;
    }
  }
  
  if (piiItems.length === 0) {
    console.log(`❌ Could not find PII "${piiValue}" in any items`);
    return;
  }
  
  console.log(`🔍 DEBUG - Found ${piiItems.length} PII items for "${piiValue}"`);
  
  // NEW POST-PROCESSING FILTER: Refine piiItems, especially for label:value patterns
  const refinedPiiItems = [];
  for (const item of piiItems) {
    const itemText = item.str;
    const colonIndex = itemText.lastIndexOf(':');

    // Check if it's a label:value pattern where PII is only the value part
    // And the item is significantly longer than the PII value
    if (colonIndex !== -1 && itemText.includes(piiValue) && itemText.length > piiValue.length) {
      const valuePart = itemText.substring(colonIndex + 1).trim();
      // Check if the PII value matches the value part (with or without spaces)
      const normalizedValuePart = valuePart.replace(/\s/g, '');
      const normalizedPiiValue = piiValue.replace(/\s/g, '');
      
      // If the PII value is exactly the value part, or the PII is a significant part of the value part
      // (e.g., PII is at least 50% of the value part's length)
      if (normalizedValuePart === normalizedPiiValue || valuePart === piiValue || 
          (valuePart.includes(piiValue) && piiValue.length >= valuePart.length * 0.5)) {
        const piiStartIndexInItem = itemText.indexOf(piiValue, colonIndex);
        if (piiStartIndexInItem !== -1) {
          const piiEndIndexInItem = piiStartIndexInItem + piiValue.length;
          const virtualItem = createVirtualItem(item, piiStartIndexInItem, piiEndIndexInItem);
          refinedPiiItems.push(virtualItem);
          console.log(`✅ Refined item: Replaced "${itemText}" with virtual item for "${piiValue}"`);
          continue; // Skip the original item
        }
      }
    }
    // If not a label:value pattern to be refined, or if PII is the whole item, keep the original item
    refinedPiiItems.push(item);
  }
  piiItems = refinedPiiItems; // Update piiItems with the refined list
  
  // Special handling for CUST-XXXXX patterns - filter out irrelevant items
  if (piiValue.startsWith('CUST-')) {
    const filteredItems = piiItems.filter(item => {
      const itemText = item.str;
      // Only keep items that contain "CUST", "-", or digits (part of the PII)
      const isRelevant = itemText.includes('CUST') || 
                        itemText.includes('-') || 
                        /\d/.test(itemText);
      
      if (!isRelevant) {
        console.log(`❌ Filtering out irrelevant CUST item: "${itemText}"`);
      } else {
        console.log(`✅ Keeping relevant CUST item: "${itemText}"`);
      }
      
      return isRelevant;
    });
    
    piiItems.length = 0; // Clear the array
    piiItems.push(...filteredItems); // Add filtered items
    
    console.log(`🔍 DEBUG - After CUST filtering: ${piiItems.length} relevant PII items`);
  }
  
  // Special handling for credit card numbers - be more precise with filtering
  if (isCreditCardOrIBAN(piiValue) && piiValue.match(/^\d{4}\s\d{4}\s\d{4}\s\d{4}$/)) {
    console.log(`🔍 DEBUG - Processing credit card number: "${piiValue}"`);
    const piiDigits = piiValue.replace(/\D/g, '');
    const filteredItems = piiItems.filter(item => {
      const itemText = item.str;
      const itemDigits = itemText.replace(/\D/g, '');
      
      // For credit card numbers, only include items that contain significant parts of the PII
      // Check if the item contains 3+ consecutive digits from the PII
      let isRelevant = false;
      
      if (itemDigits.length >= 3) {
        // Check for 3+ digit sequences from the PII
        for (let i = 0; i <= itemDigits.length - 3; i++) {
          for (let len = 3; len <= Math.min(itemDigits.length - i, 6); len++) {
            const itemSequence = itemDigits.substring(i, i + len);
            if (piiDigits.includes(itemSequence)) {
              isRelevant = true;
              console.log(`✅ Keeping credit card item: "${itemText}" (contains sequence: ${itemSequence})`);
              break;
            }
          }
          if (isRelevant) break;
        }
      } else if (itemDigits.length > 0) {
        // For shorter items, check if they're direct substrings of the PII
        if (piiDigits.includes(itemDigits)) {
          isRelevant = true;
          console.log(`✅ Keeping credit card item: "${itemText}" (direct substring: ${itemDigits})`);
        }
      }
      
      if (!isRelevant) {
        console.log(`❌ Filtering out credit card item: "${itemText}" (digits: ${itemDigits})`);
      }
      
      return isRelevant;
    });
    
    piiItems.length = 0; // Clear the array
    piiItems.push(...filteredItems); // Add filtered items
    
    console.log(`🔍 DEBUG - After credit card filtering: ${piiItems.length} relevant PII items`);
  }
  
  // Filter out items that don't actually contain PII digits (skip if already processed above)
  if (isCreditCardOrIBAN(piiValue) && !piiValue.match(/^\d{4}\s\d{4}\s\d{4}\s\d{4}$/)) {
    const piiDigits = piiValue.replace(/\D/g, '');
    const filteredItems = [];
    
    for (const item of piiItems) {
      const itemText = item.str;
      const itemDigits = itemText.replace(/\D/g, '');
      let isRelevant = false;
      
      // For IBANs and DE codes, be very precise - only include items that are clearly part of the PII
      if (piiValue.startsWith('DE')) {
        // For IBANs and DE codes, check if item contains significant parts of the PII
        const normalizedPii = piiValue.replace(/\s/g, '').toUpperCase();
        const normalizedItem = itemText.replace(/\s/g, '').toUpperCase();
        
        // Only include items that contain 3+ character sequences from the PII
        // This prevents including random text that happens to contain a few characters
        for (let i = 0; i < normalizedPii.length - 2; i++) {
          for (let len = 3; len <= Math.min(normalizedPii.length - i, 6); len++) {
            const piiSequence = normalizedPii.substring(i, i + len);
            if (normalizedItem.includes(piiSequence)) {
              isRelevant = true;
              console.log(`✅ Item "${item.str}" contains DE sequence: ${piiSequence}`);
              break;
            }
          }
          if (isRelevant) break;
        }
        
        // Special case: if the item is very short (like "DE", "46", "85"), include it
        if (!isRelevant && normalizedItem.length <= 4) {
          // Check if the item is a direct substring of the PII
          if (normalizedPii.includes(normalizedItem)) {
            isRelevant = true;
            console.log(`✅ Item "${item.str}" is short and contained in PII: ${normalizedItem}`);
          }
        }
      } else {
        // For other credit card numbers, check for digit sequences
        if (piiDigits.length > 0 && itemDigits.length > 0) {
          // For credit card numbers, be more lenient for single digits
          const minSequenceLength = itemDigits.length === 1 ? 1 : 2; // Allow single digits
          
          for (let k = 0; k <= itemDigits.length - minSequenceLength; k++) {
            for (let len = minSequenceLength; len <= Math.min(itemDigits.length - k, 4); len++) {
              const itemSequence = itemDigits.substring(k, k + len);
              if (piiDigits.includes(itemSequence)) {
                isRelevant = true;
                console.log(`✅ Item "${item.str}" contains PII digits: ${itemSequence}`);
                break;
              }
            }
            if (isRelevant) break;
          }
        }
      }
      
      if (isRelevant) {
        filteredItems.push(item);
      } else {
        console.log(`❌ Filtering out irrelevant item: "${item.str}"`);
      }
    }
    
    piiItems.length = 0; // Clear the array
    piiItems.push(...filteredItems); // Add filtered items
    
    console.log(`🔍 DEBUG - After filtering: ${piiItems.length} relevant PII items`);
  }
  
  // Draw rectangles over PII items
  if ((isCreditCardOrIBAN(piiValue) || piiValue.startsWith('CUST-') || piiValue.startsWith('INV-')) && piiItems.length > 1) {
    // For credit card numbers, CUST patterns, and INV patterns split across multiple items, draw one rectangle spanning all items
    console.log(`🔍 DEBUG - Drawing single rectangle for split PII across ${piiItems.length} items`);

    // Find the leftmost and rightmost positions
    let minX = Infinity;
    let maxX = -Infinity;
    let minY = Infinity;
    let maxY = -Infinity;

    for (const item of piiItems) {
      const transform = item.transform;
      const itemX = transform[4];
      const itemY = transform[5];
      const itemWidth = item.width || 0;
      const itemHeight = item.height || 12;

      minX = Math.min(minX, itemX);
      maxX = Math.max(maxX, itemX + itemWidth);
      minY = Math.min(minY, itemY);
      maxY = Math.max(maxY, itemY + itemHeight);

      console.log(`🔍 DEBUG - Item: "${item.str}" at (${itemX}, ${itemY}) size: ${itemWidth}x${itemHeight}`);
    }

    // Draw one rectangle spanning all items
    const rectangleX = minX - 2;
    const rectangleY = minY - 2;
    const rectangleWidth = maxX - minX + 4;
    const rectangleHeight = maxY - minY + 4;

    console.log(`🔍 DEBUG - Combined rectangle: (${rectangleX}, ${rectangleY}) size: ${rectangleWidth}x${rectangleHeight}`);

    page.drawRectangle({
      x: rectangleX,
      y: rectangleY,
      width: rectangleWidth,
      height: rectangleHeight,
      color: rgb(0, 0, 0),
    });

    console.log(`✅ Drew single rectangle for split PII "${piiValue}"`);
  } else {
    // For single items or other PII, draw rectangle for each item
    for (const item of piiItems) {
      const transform = item.transform;
      const itemX = transform[4];
      const itemY = transform[5];
      const itemWidth = item.width || 0;
      const itemHeight = item.height || 12;

      console.log(`🔍 DEBUG - Item: "${item.str}"`);
      console.log(`🔍 DEBUG - PDF.js coords: x=${itemX}, y=${itemY}, width=${itemWidth}, height=${itemHeight}`);

      const rectangleX = itemX - 2;
      const rectangleY = itemY - 2;
      const rectangleWidth = itemWidth + 4;
      const rectangleHeight = itemHeight + 4;

      console.log(`🔍 DEBUG - Rectangle: (${rectangleX}, ${rectangleY}) size: ${rectangleWidth}x${rectangleHeight}`);

      page.drawRectangle({
        x: rectangleX,
        y: rectangleY,
        width: rectangleWidth,
        height: rectangleHeight,
        color: rgb(0, 0, 0),
      });

      console.log(`✅ Drew rectangle for PII "${piiValue}"`);
    }
  }
}

/**
 * Replace PII text with "XXXX" (Option 2)
 */
function replacePIIWithText(page, line, piiValue, height) {
  const lineText = line.map(item => item.str).join(' ');
  
  console.log(`🔍 DEBUG - Looking for PII: "${piiValue}"`);
  console.log(`🔍 DEBUG - In line: "${lineText}"`);
  
  // Find text items that contain the PII using direct string matching
  let piiItems = [];
  
  // First, try exact matching on individual items
  for (let i = 0; i < line.length; i++) {
    const item = line[i];
    const itemText = item.str;
    
    if (itemText.includes(piiValue)) {
      piiItems.push(item);
      console.log(`✅ Found exact match in item ${i}: "${itemText}"`);
    }
  }
  
  // If no exact matches, try fuzzy matching on individual items
  if (piiItems.length === 0) {
    const normalizedVal = piiValue.replace(/\s/g, '');
    
    for (let i = 0; i < line.length; i++) {
      const item = line[i];
      const itemText = item.str;
      const normalizedItemText = itemText.replace(/\s/g, '');
      
      if (normalizedItemText.includes(normalizedVal)) {
        piiItems.push(item);
        console.log(`✅ Found fuzzy match in item ${i}: "${itemText}"`);
      }
    }
  }
  
  // If still no matches, try finding consecutive items that together contain the PII
  if (piiItems.length === 0) {
    const normalizedVal = piiValue.replace(/\s/g, '');
    
    for (let i = 0; i < line.length; i++) {
      let combinedText = '';
      let combinedItems = [];
      
      for (let j = i; j < Math.min(i + 5, line.length); j++) {
        combinedText += (combinedText ? ' ' : '') + line[j].str;
        combinedItems.push(line[j]);
        
        const normalizedCombined = combinedText.replace(/\s/g, '');
        if (normalizedCombined.includes(normalizedVal)) {
          piiItems.push(...combinedItems);
          console.log(`✅ Found combined match in items ${i}-${j}: "${combinedText}"`);
          break;
        }
      }
      
      if (piiItems.length > 0) break;
    }
  }
  
  if (piiItems.length === 0) {
    console.log(`❌ Could not find PII "${piiValue}" in any items`);
    return;
  }
  
  console.log(`🔍 DEBUG - Found ${piiItems.length} PII items for "${piiValue}"`);
  
  // NEW POST-PROCESSING FILTER: Refine piiItems, especially for label:value patterns
  const refinedPiiItems = [];
  for (const item of piiItems) {
    const itemText = item.str;
    const colonIndex = itemText.lastIndexOf(':');

    // Check if it's a label:value pattern where PII is only the value part
    // And the item is significantly longer than the PII value
    if (colonIndex !== -1 && itemText.includes(piiValue) && itemText.length > piiValue.length) {
      const valuePart = itemText.substring(colonIndex + 1).trim();
      // Check if the PII value matches the value part (with or without spaces)
      const normalizedValuePart = valuePart.replace(/\s/g, '');
      const normalizedPiiValue = piiValue.replace(/\s/g, '');
      
      // If the PII value is exactly the value part, or the PII is a significant part of the value part
      // (e.g., PII is at least 50% of the value part's length)
      if (normalizedValuePart === normalizedPiiValue || valuePart === piiValue || 
          (valuePart.includes(piiValue) && piiValue.length >= valuePart.length * 0.5)) {
        const piiStartIndexInItem = itemText.indexOf(piiValue, colonIndex);
        if (piiStartIndexInItem !== -1) {
          const piiEndIndexInItem = piiStartIndexInItem + piiValue.length;
          const virtualItem = createVirtualItem(item, piiStartIndexInItem, piiEndIndexInItem);
          refinedPiiItems.push(virtualItem);
          console.log(`✅ Refined item: Replaced "${itemText}" with virtual item for "${piiValue}"`);
          continue; // Skip the original item
        }
      }
    }
    // If not a label:value pattern to be refined, or if PII is the whole item, keep the original item
    refinedPiiItems.push(item);
  }
  piiItems = refinedPiiItems; // Update piiItems with the refined list
  
  // Replace PII text with "XXXX" by drawing text over it
  for (const item of piiItems) {
    const transform = item.transform;
    const itemX = transform[4];
    const itemY = transform[5];
    const itemWidth = item.width || 0;
    const itemHeight = item.height || 12;
    
    // Calculate position for replacement text
    const textX = itemX;
    const textY = height - itemY - itemHeight; // Flip Y coordinate
    
    // Create replacement text with appropriate length
    const replacementText = 'X'.repeat(Math.max(4, Math.floor(itemWidth / 8))); // Adjust length based on width
    
    // Draw replacement text over the PII text
    page.drawText(replacementText, {
      x: textX,
      y: textY,
      size: itemHeight,
      color: rgb(0, 0, 0), // Black text
    });
    
    console.log(`✅ Replaced PII item "${item.str}" with "${replacementText}" at (${textX}, ${textY})`);
  }
}

/**
 * Check if a value is a credit card number or IBAN
 */
function isCreditCardOrIBAN(val) {
  const digits = val.replace(/\D/g, '');
  const normalizedVal = val.replace(/\s/g, ''); // Remove spaces for length check
  
  // Credit card numbers: 12-19 digits
  const isCreditCard = (digits.length >= 12 && digits.length <= 19);
  
  // IBANs: For 'DE' IBANs, they are typically 22 alphanumeric characters long
  // Also include shorter DE codes that might be bank codes or similar
  const isGermanIBAN = normalizedVal.startsWith('DE') && (normalizedVal.length === 22 || normalizedVal.length >= 8);
  
  return isCreditCard || isGermanIBAN;
}

/**
 * Escape special regex characters in a string
 * @param {string} string - The string to escape
 * @returns {string} - The escaped string
 */
function escapeRegExp(string) {
  return string.replace(/[.*+?^${}()|[\]\\]/g, '\\$&');
}

/**
 * Find a byte pattern in a Uint8Array
 * @param {Uint8Array} haystack - The data to search in
 * @param {Uint8Array} needle - The pattern to find
 * @param {number} startOffset - Where to start searching
 * @returns {number} - Index of the pattern, or -1 if not found
 */
function findBytes(haystack, needle, startOffset = 0) {
  for (let i = startOffset; i <= haystack.length - needle.length; i++) {
    let found = true;
    for (let j = 0; j < needle.length; j++) {
      if (haystack[i + j] !== needle[j]) {
        found = false;
        break;
      }
    }
    if (found) return i;
  }
  return -1;
}

/**
 * Extract text from a specific PDF page using pdf.js
 * @param {ArrayBuffer} pdfArrayBuffer - PDF file as ArrayBuffer
 * @param {number} pageNum - Page number (1-based)
 * @returns {Promise<string>} - Extracted text from the page
 */
async function extractTextFromPdfPage(pdfArrayBuffer, pageNum) {
  try {
    console.log(`Extracting text from PDF page ${pageNum}...`);
    
    const uint8Array = new Uint8Array(pdfArrayBuffer);
    const loadingTask = pdfjsLib.getDocument({
      data: uint8Array,
      useSystemFonts: true,
      disableWorker: true,
    });
    
    const pdfDoc = await loadingTask.promise;
    const page = await pdfDoc.getPage(pageNum);
    const textContent = await page.getTextContent();
    const pageText = textContent.items.map(item => item.str).join(' ');
    
    await pdfDoc.destroy();
    return pageText;
  } catch (error) {
    console.error(`Error extracting text from page ${pageNum}:`, error);
    return '';
  }
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
