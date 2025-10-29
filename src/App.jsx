import React, { useState } from "react";
import { saveAs } from "file-saver";
import { readFileAsText } from "./eval.js";
import { detectFileType, extractTextFromFile, isImageFile } from "./file-processor.js";
import { maskDocx, maskDoc, maskPdf } from "./text-masker.js";
import { maskImage, maskScannedPdf } from "./image-masker.js";

export default function App() {
  const [originalFile, setOriginalFile] = useState(null);
  const [jsonFile, setJsonFile] = useState(null);
  const [busy, setBusy] = useState(false);
  const [log, setLog] = useState("");

  const logLine = (m) => setLog((s) => (s ? s + "\n" : "") + m);

  const onGenerate = async () => {
    if (!originalFile || !jsonFile) {
      alert("Please upload an original file (DOC/DOCX/PDF/Image) and a JSON file with PII detections.");
      return;
    }
    setBusy(true);
    setLog("");

    // Add timeout to prevent infinite hanging
    const timeoutId = setTimeout(() => {
      logLine("❌ Operation timed out after 30 seconds");
      setBusy(false);
    }, 30000);

    try {
      // Detect file type
      logLine("Detecting file type…");
      const fileType = detectFileType(originalFile);
      logLine(`File type detected: ${fileType.type}`);

      // Read JSON detections
      logLine("Reading JSON detections…");
      const jsonText = await readFileAsText(jsonFile);
      const parsed = JSON.parse(jsonText);
      const detections = Array.isArray(parsed?.pii)
        ? parsed.pii.map((x) => ({ type: x.type, value: String(x.value ?? "") }))
        : [];
      logLine(`Found ${detections.length} PII detections`);

      // Extract text and determine masking approach
      logLine("Extracting text from file…");
      const { text, fileType: detectedType, isScanned } = await extractTextFromFile(originalFile);
      logLine(`Extracted text length: ${text.length}`);
      logLine(`Is scanned document: ${isScanned}`);

      // Generate masked file based on file type
      let maskedBlob;
      const originalArrayBuffer = await originalFile.arrayBuffer();

      if (isScanned || isImageFile(originalFile)) {
        // Use image masking (black rectangles)
        logLine("Applying image masking (black rectangles)…");
        if (fileType.type === 'pdf') {
          maskedBlob = await maskScannedPdf(originalArrayBuffer, detections);
        } else {
          maskedBlob = await maskImage(originalFile, detections);
        }
      } else {
        // Use text masking (replace with 'x' characters)
        logLine("Applying text masking (replace with 'x' characters)…");
        switch (fileType.type) {
          case 'docx':
            maskedBlob = await maskDocx(originalArrayBuffer, detections);
            break;
          case 'doc':
            maskedBlob = await maskDoc(originalArrayBuffer, detections);
            break;
          case 'pdf':
            maskedBlob = await maskPdf(originalArrayBuffer, detections, text, 'rectangle'); // Use 'rectangle' or 'text' for masking method
            break;
          default:
            throw new Error(`Unsupported file type for text masking: ${fileType.type}`);
        }
      }

      // Generate output filename
      const originalName = originalFile.name;
      const nameWithoutExt = originalName.replace(/\.[^/.]+$/, "");
      const extension = originalName.split('.').pop().toLowerCase();
      
      let outputExtension;
      if (fileType.type === 'image') {
        outputExtension = 'png'; // Images are converted to PNG
      } else {
        outputExtension = extension; // Keep original extension for documents
      }

      const outputFileName = `${nameWithoutExt}_MASKED.${outputExtension}`;

      logLine("Downloading masked file…");
      saveAs(maskedBlob, outputFileName);

      logLine("✅ Done — masked file downloaded.");
    } catch (e) {
      console.error(e);
      logLine("❌ " + (e?.message || e));
      alert("Failed: " + (e?.message || e));
    } finally {
      clearTimeout(timeoutId);
      setBusy(false);
    }
  };

  return (
    <div style={{ minHeight: "100vh", background: "#f7fafc", padding: 24 }}>
      <div style={{ maxWidth: 900, margin: "0 auto", background: "#fff", borderRadius: 16, boxShadow: "0 6px 20px rgba(0,0,0,0.06)", padding: 24 }}>
        <h1 style={{ fontSize: 22, marginBottom: 6 }}>PII Masking Tool</h1>
        <p style={{ fontSize: 14, color: "#718096", marginBottom: 20 }}>
          Upload a document (DOC/DOCX/PDF) or image, and a JSON file with PII detections to create a masked version.
        </p>
        <div style={{ display: "grid", gap: 12 }}>
          <div>
            <div style={{ fontSize: 13, fontWeight: 600, marginBottom: 6 }}>Original File</div>
            <input 
              type="file" 
              accept=".doc,.docx,.pdf,.jpg,.jpeg,.png,.gif,.bmp,.tiff,.webp" 
              onChange={(e) => setOriginalFile(e.target.files?.[0] || null)} 
            />
            {originalFile && (
              <div style={{ fontSize: 12, color: "#718096", marginTop: 4 }}>
                Selected: {originalFile.name} ({detectFileType(originalFile)?.type || 'unknown'})
              </div>
            )}
          </div>

          <div>
            <div style={{ fontSize: 13, fontWeight: 600, marginBottom: 6 }}>PII Detections JSON</div>
            <input type="file" accept=".json" onChange={(e) => setJsonFile(e.target.files?.[0] || null)} />
            {jsonFile && <div style={{ fontSize: 12, color: "#718096", marginTop: 4 }}>Selected: {jsonFile.name}</div>}
          </div>

          <button
            disabled={busy}
            onClick={onGenerate}
            style={{
              padding: "10px 14px",
              borderRadius: 10,
              background: "#111",
              color: "#fff",
              fontWeight: 600,
              opacity: busy ? 0.6 : 1,
              cursor: busy ? "not-allowed" : "pointer"
            }}
          >
            {busy ? "Masking…" : "Generate Masked File"}
          </button>
        </div>

        <pre style={{ background: "#f1f5f9", color: "#1f2937", fontSize: 12, padding: 12, borderRadius: 10, marginTop: 16, maxHeight: 260, overflow: "auto", whiteSpace: "pre-wrap" }}>
{log || "Logs will appear here…"}
        </pre>
      </div>
    </div>
  );
}
