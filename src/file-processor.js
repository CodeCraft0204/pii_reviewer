/**
 * File processor for detecting file types and extracting text
 */

import mammoth from "mammoth";
import * as pdfjsLib from "pdfjs-dist";

// Set up PDF.js worker
pdfjsLib.GlobalWorkerOptions.workerSrc = `//cdnjs.cloudflare.com/ajax/libs/pdf.js/${pdfjsLib.version}/pdf.worker.min.js`;

/**
 * Detect file type from file object
 * @param {File} file - The file to detect
 * @returns {Object} - { type: 'docx'|'doc'|'pdf'|'image', mimeType: string }
 */
export function detectFileType(file) {
  const name = file.name.toLowerCase();
  const mimeType = file.type.toLowerCase();

  // Check by extension first
  if (name.endsWith('.docx')) {
    return { type: 'docx', mimeType: file.type || 'application/vnd.openxmlformats-officedocument.wordprocessingml.document' };
  }
  if (name.endsWith('.doc')) {
    return { type: 'doc', mimeType: file.type || 'application/msword' };
  }
  if (name.endsWith('.pdf')) {
    return { type: 'pdf', mimeType: file.type || 'application/pdf' };
  }
  
  // Check for image types
  if (name.match(/\.(jpg|jpeg|png|gif|bmp|tiff|webp)$/i) || mimeType.startsWith('image/')) {
    return { type: 'image', mimeType: file.type };
  }

  // Fallback to mime type
  if (mimeType.includes('wordprocessingml')) return { type: 'docx', mimeType };
  if (mimeType.includes('msword')) return { type: 'doc', mimeType };
  if (mimeType.includes('pdf')) return { type: 'pdf', mimeType };
  if (mimeType.startsWith('image/')) return { type: 'image', mimeType };

  throw new Error(`Unsupported file type: ${name} (${mimeType})`);
}

/**
 * Extract text from DOCX file
 * @param {ArrayBuffer} arrayBuffer - The DOCX file as ArrayBuffer
 * @returns {Promise<string>} - Extracted text
 */
export async function extractTextFromDocx(arrayBuffer) {
  const { value } = await mammoth.extractRawText({ arrayBuffer });
  return (value || "").replace(/\r/g, "");
}

/**
 * Extract text from DOC file (older Word format)
 * @param {ArrayBuffer} arrayBuffer - The DOC file as ArrayBuffer
 * @returns {Promise<string>} - Extracted text
 */
export async function extractTextFromDoc(arrayBuffer) {
  // mammoth also supports .doc files
  try {
    const { value } = await mammoth.extractRawText({ arrayBuffer });
    return (value || "").replace(/\r/g, "");
  } catch (error) {
    throw new Error(`Failed to extract text from DOC file: ${error.message}`);
  }
}

/**
 * Extract text from PDF file
 * @param {ArrayBuffer} arrayBuffer - The PDF file as ArrayBuffer
 * @returns {Promise<{text: string, isScanned: boolean}>} - Extracted text and whether it's a scanned PDF
 */
export async function extractTextFromPdf(arrayBuffer) {
  const loadingTask = pdfjsLib.getDocument({ data: arrayBuffer });
  const pdf = await loadingTask.promise;
  
  let fullText = "";
  let totalChars = 0;
  
  for (let pageNum = 1; pageNum <= pdf.numPages; pageNum++) {
    const page = await pdf.getPage(pageNum);
    const textContent = await page.getTextContent();
    const pageText = textContent.items.map(item => item.str).join(' ');
    fullText += pageText + '\n';
    totalChars += pageText.length;
  }
  
  // Heuristic: if very little text extracted, it's likely a scanned PDF
  const isScanned = totalChars < 100 && pdf.numPages > 0;
  
  return {
    text: fullText.trim(),
    isScanned
  };
}

/**
 * Check if file is an image
 * @param {File} file - The file to check
 * @returns {boolean}
 */
export function isImageFile(file) {
  const { type } = detectFileType(file);
  return type === 'image';
}

/**
 * Load image from file
 * @param {File} file - The image file
 * @returns {Promise<HTMLImageElement>}
 */
export function loadImage(file) {
  return new Promise((resolve, reject) => {
    const reader = new FileReader();
    reader.onload = (e) => {
      const img = new Image();
      img.onload = () => resolve(img);
      img.onerror = reject;
      img.src = e.target.result;
    };
    reader.onerror = reject;
    reader.readAsDataURL(file);
  });
}

/**
 * Extract text and metadata from any supported file
 * @param {File} file - The file to process
 * @returns {Promise<{text: string, fileType: string, isScanned: boolean}>}
 */
export async function extractTextFromFile(file) {
  const { type } = detectFileType(file);
  const arrayBuffer = await file.arrayBuffer();
  
  switch (type) {
    case 'docx':
      return {
        text: await extractTextFromDocx(arrayBuffer),
        fileType: 'docx',
        isScanned: false
      };
    
    case 'doc':
      return {
        text: await extractTextFromDoc(arrayBuffer),
        fileType: 'doc',
        isScanned: false
      };
    
    case 'pdf': {
      const { text, isScanned } = await extractTextFromPdf(arrayBuffer);
      return {
        text,
        fileType: 'pdf',
        isScanned
      };
    }
    
    case 'image':
      return {
        text: '',
        fileType: 'image',
        isScanned: true // Images are treated as scanned documents
      };
    
    default:
      throw new Error(`Unsupported file type: ${type}`);
  }
}

