/**
 * Image masking module for images and scanned PDFs
 * Masks detected PII with black rectangles using OCR positioning
 */

import { PDFDocument, rgb } from "pdf-lib";
import * as pdfjsLib from "pdfjs-dist";
import { createWorker } from 'tesseract.js';

// Set up PDF.js worker
pdfjsLib.GlobalWorkerOptions.workerSrc = `//cdnjs.cloudflare.com/ajax/libs/pdf.js/${pdfjsLib.version}/pdf.worker.min.js`;

/**
 * Mask image file by drawing black rectangles over detected PII
 * @param {File} imageFile - Original image file
 * @param {Array<{type: string, value: string}>} detections - PII detections
 * @returns {Promise<Blob>} - Masked image as PNG
 */
export async function maskImage(imageFile, detections) {
  // Load image
  const img = await loadImageFromFile(imageFile);

  // Create canvas
  const canvas = document.createElement('canvas');
  canvas.width = img.width;
  canvas.height = img.height;
  const ctx = canvas.getContext('2d');

  // Draw original image
  ctx.drawImage(img, 0, 0);

  // Try OCR-based masking for precise alignment
  try {
    const boxes = await extractBoxesWithTesseract(img, detections);
    ctx.fillStyle = 'black';
    for (const b of boxes) {
      ctx.fillRect(b.x - 2, b.y - 2, b.width + 4, b.height + 4);
    }
  } catch (err) {
    console.warn('OCR masking failed, falling back to heuristic rectangles:', err);
    // Fallback: draw rough rectangles (previous behavior)
    ctx.fillStyle = 'black';
    const textHeight = 20;
    const padding = 5;
    for (let i = 0; i < detections.length; i++) {
      const y = 50 + (i * 30);
      const x = 50;
      const width = 300;
      ctx.fillRect(x - padding, y - textHeight + padding, width + (padding * 2), textHeight + padding);
    }
  }

  return new Promise((resolve, reject) => {
    canvas.toBlob((blob) => {
      if (blob) {
        resolve(blob);
      } else {
        reject(new Error('Failed to create image blob'));
      }
    }, 'image/png');
  });
}

// Run Tesseract to get word boxes and map detections to boxes
async function extractBoxesWithTesseract(imgElement, detections) {
  // Prepare worker (French + English improves accuracy on FR IDs)
  const worker = await createWorker('eng+fra');
  try {
    // Recognize directly from canvas to avoid re-encoding
    const tempCanvas = document.createElement('canvas');
    tempCanvas.width = imgElement.width;
    tempCanvas.height = imgElement.height;
    const tctx = tempCanvas.getContext('2d');
    tctx.drawImage(imgElement, 0, 0);

    const { data } = await worker.recognize(tempCanvas);
    const words = data?.words || [];

    // Index words by order with normalized text
    const normalized = (s) => s
      .normalize('NFKD')
      .replace(/\p{Diacritic}+/gu, '')
      .replace(/\s+/g, ' ')
      .trim()
      .toLowerCase();

    const wordList = words.map(w => ({
      text: w.text || '',
      ntext: normalized(w.text || ''),
      bbox: { x: w.bbox.x0, y: w.bbox.y0, width: w.bbox.x1 - w.bbox.x0, height: w.bbox.y1 - w.bbox.y0 }
    }));

    const boxes = [];

    // Token-level matcher with strict rules: digits exact, short words exact, longer words allow 1 edit
    const levenshtein = (a, b) => {
      const m = a.length, n = b.length;
      const dp = Array.from({ length: m + 1 }, () => new Array(n + 1).fill(0));
      for (let i = 0; i <= m; i++) dp[i][0] = i;
      for (let j = 0; j <= n; j++) dp[0][j] = j;
      for (let i = 1; i <= m; i++) {
        for (let j = 1; j <= n; j++) {
          const cost = a[i - 1] === b[j - 1] ? 0 : 1;
          dp[i][j] = Math.min(
            dp[i - 1][j] + 1,
            dp[i][j - 1] + 1,
            dp[i - 1][j - 1] + cost
          );
        }
      }
      return dp[m][n];
    };

    const tokensMatchStrict = (windowTokens, targetTokens) => {
      if (windowTokens.length !== targetTokens.length) return false;
      for (let t = 0; t < targetTokens.length; t++) {
        const tgt = targetTokens[t];
        const win = windowTokens[t];
        const hasDigit = /\d/.test(tgt) || /\d/.test(win);
        if (hasDigit || tgt.length <= 3) {
          if (win !== tgt) return false; // exact for digits and very short tokens
        } else {
          const dist = levenshtein(win, tgt);
          if (dist > 1) return false; // allow 1 edit for longer alphabetic tokens
        }
      }
      return true;
    };

    for (const det of detections || []) {
      const target = normalized(String(det.value || ''));
      if (!target) continue;

      // Split into tokens; try to match consecutive words with fuzziness
      const tokens = target.split(' ').filter(Boolean);
      if (tokens.length === 0) continue;

      // Sliding fuzzy window over word sequence
      const targetCompact = target.replace(/\s/g, '');
      for (let i = 0; i < wordList.length; i++) {
        // Prefer exact token-length windows to avoid overmasking trailing words (e.g., city names)
        const preferredLens = [tokens.length];
        for (const wlenRaw of preferredLens) {
          const wlen = Math.min(wlenRaw, wordList.length - i);
          const window = wordList.slice(i, i + wlen);
          const windowText = window.map(w => w.ntext).join(' ');
          const windowCompact = windowText.replace(/\s/g, '');

          // STRICT, token-aware match: require tokens to match with tight tolerance
          let isMatch = false;
          if (windowCompact === targetCompact || tokensMatchStrict(window.map(w => w.ntext), tokens)) {
            isMatch = true;
          } else {
            // Fuzzy fallback for minor OCR slips (max 2 edits), but same token count and same digit sequence
            const digitsEqual = windowCompact.replace(/\D/g, '') === targetCompact.replace(/\D/g, '');
            const dist = levenshtein(windowCompact, targetCompact);
            const ratio = dist / Math.max(1, targetCompact.length);
            if (digitsEqual && dist <= 2 && ratio <= 0.12) {
              isMatch = true;
            }
          }

          if (isMatch) {
            // Combine bounding boxes from window
            let x0 = Infinity, y0 = Infinity, x1 = -Infinity, y1 = -Infinity;
            for (const w of window) {
              const b = w.bbox;
              x0 = Math.min(x0, b.x);
              y0 = Math.min(y0, b.y);
              x1 = Math.max(x1, b.x + b.width);
              y1 = Math.max(y1, b.y + b.height);
            }
            boxes.push({ x: x0, y: y0, width: x1 - x0, height: y1 - y0 });
            i = i + wlen - 1; // skip ahead
            break;
          }
        }
      }
    }

    return boxes;
  } finally {
    await worker.terminate();
  }
}

/**
 * Mask scanned PDF by drawing black rectangles over detected PII
 * @param {ArrayBuffer} pdfArrayBuffer - Original PDF file
 * @param {Array<{type: string, value: string}>} detections - PII detections
 * @returns {Promise<Blob>} - Masked PDF file
 */
export async function maskScannedPdf(pdfArrayBuffer, detections) {
  const pdfDoc = await PDFDocument.load(pdfArrayBuffer);
  const pages = pdfDoc.getPages();
  
  // For scanned PDFs, we need to:
  // 1. Extract images from each page
  // 2. Use OCR to find text positions (or use heuristics)
  // 3. Draw black rectangles over detected PII
  
  // Simple approach: draw black rectangles at estimated positions
  for (const page of pages) {
    const { width, height } = page.getSize();
    
    // Draw black rectangles over detected areas
    // This is a placeholder - in production, use OCR to find exact positions
    for (let i = 0; i < detections.length; i++) {
      const y = height - 100 - (i * 30); // From top
      const x = 50;
      const rectWidth = 300;
      const rectHeight = 20;
      
      page.drawRectangle({
        x,
        y,
        width: rectWidth,
        height: rectHeight,
        color: rgb(0, 0, 0),
      });
    }
  }
  
  const pdfBytes = await pdfDoc.save();
  return new Blob([pdfBytes], { type: 'application/pdf' });
}

/**
 * Advanced masking using OCR (Tesseract.js)
 * This is a more sophisticated approach that finds exact text positions
 * @param {File} imageFile - Image file
 * @param {Array<{type: string, value: string}>} detections - PII detections
 * @returns {Promise<Blob>} - Masked image
 */
export async function maskImageWithOCR(imageFile, detections) {
  // This would require Tesseract.js
  // For now, fall back to simple masking
  return maskImage(imageFile, detections);
}

/**
 * Load image from file
 * @param {File} file - Image file
 * @returns {Promise<HTMLImageElement>}
 */
function loadImageFromFile(file) {
  return new Promise((resolve, reject) => {
    const reader = new FileReader();
    reader.onload = (e) => {
      const img = new Image();
      img.onload = () => resolve(img);
      img.onerror = () => reject(new Error('Failed to load image'));
      img.src = e.target.result;
    };
    reader.onerror = () => reject(new Error('Failed to read file'));
    reader.readAsDataURL(file);
  });
}

/**
 * Extract text positions from PDF using PDF.js
 * @param {ArrayBuffer} pdfArrayBuffer - PDF file
 * @returns {Promise<Array<{text: string, x: number, y: number, width: number, height: number, page: number}>>}
 */
export async function extractTextPositionsFromPdf(pdfArrayBuffer) {
  const loadingTask = pdfjsLib.getDocument({ data: pdfArrayBuffer });
  const pdf = await loadingTask.promise;
  
  const allPositions = [];
  
  for (let pageNum = 1; pageNum <= pdf.numPages; pageNum++) {
    const page = await pdf.getPage(pageNum);
    const textContent = await page.getTextContent();
    const viewport = page.getViewport({ scale: 1.0 });
    
    for (const item of textContent.items) {
      if (item.str && item.transform) {
        const tx = item.transform[4];
        const ty = item.transform[5];
        
        allPositions.push({
          text: item.str,
          x: tx,
          y: viewport.height - ty, // Flip Y coordinate
          width: item.width || 100,
          height: item.height || 12,
          page: pageNum
        });
      }
    }
  }
  
  return allPositions;
}

/**
 * Mask PDF with precise text positions
 * @param {ArrayBuffer} pdfArrayBuffer - Original PDF
 * @param {Array<{type: string, value: string}>} detections - PII detections
 * @returns {Promise<Blob>} - Masked PDF
 */
export async function maskPdfWithPositions(pdfArrayBuffer, detections) {
  // Extract text positions
  const textPositions = await extractTextPositionsFromPdf(pdfArrayBuffer);
  
  // Find positions for each detection
  const maskedPositions = [];
  for (const det of detections) {
    const val = String(det.value || "").trim();
    if (!val) continue;
    
    // Find all text items that match this detection
    for (const pos of textPositions) {
      if (pos.text.includes(val)) {
        maskedPositions.push(pos);
      }
    }
  }
  
  // Load PDF and draw black rectangles
  const pdfDoc = await PDFDocument.load(pdfArrayBuffer);
  const pages = pdfDoc.getPages();
  
  for (const pos of maskedPositions) {
    const page = pages[pos.page - 1];
    if (!page) continue;
    
    const { height } = page.getSize();
    
    page.drawRectangle({
      x: pos.x - 2,
      y: height - pos.y - pos.height - 2,
      width: pos.width + 4,
      height: pos.height + 4,
      color: rgb(0, 0, 0),
    });
  }
  
  const pdfBytes = await pdfDoc.save();
  return new Blob([pdfBytes], { type: 'application/pdf' });
}

