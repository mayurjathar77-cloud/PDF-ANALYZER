import { useRef, useState } from "react";
import pdfWorkerUrl from "pdfjs-dist/build/pdf.worker.min.mjs?url";

const stopWords = new Set(
  "a an and are as at be been but by for from had has have he her hers him his i if in into is it its may more most no not of on or our she so than that the their them then there these they this to was we were will with would you your".split(" ")
);

function formatFileSize(bytes) {
  if (bytes < 1024 * 1024) return `${Math.max(1, Math.round(bytes / 1024))} KB`;
  return `${(bytes / (1024 * 1024)).toFixed(1)} MB`;
}

function splitSentences(text) {
  if (typeof Intl.Segmenter === "function") {
    const segmenter = new Intl.Segmenter("en", { granularity: "sentence" });
    return Array.from(segmenter.segment(text), (part) => part.segment.trim()).filter(Boolean);
  }
  return text.match(/[^.!?]+[.!?]+|[^.!?]+$/g)?.map((sentence) => sentence.trim()).filter(Boolean) || [];
}

function summarize(text, length) {
  const sentences = splitSentences(text.replace(/\s+/g, " "));
  const usable = sentences.filter((sentence) => sentence.split(/\s+/).length >= 5);
  if (usable.length === 0) return "";

  const termFrequency = new Map();
  const wordLists = usable.map((sentence) =>
    (sentence.toLowerCase().match(/[a-z0-9][a-z0-9'-]*/g) || []).filter(
      (word) => word.length > 2 && !stopWords.has(word)
    )
  );

  for (const words of wordLists) {
    for (const word of new Set(words)) termFrequency.set(word, (termFrequency.get(word) || 0) + 1);
  }

  let maxFrequency = 1;
  for (const frequency of termFrequency.values()) {
    if (frequency > maxFrequency) maxFrequency = frequency;
  }
  const scored = usable.map((sentence, index) => {
    const words = wordLists[index];
    const contentWords = words.length || 1;
    const frequencyScore = words.reduce((score, word) => score + (termFrequency.get(word) || 0) / maxFrequency, 0) / contentWords;
    const positionBonus = index === 0 ? 0.18 : index < Math.ceil(usable.length * 0.15) ? 0.08 : 0;
    return { sentence, index, score: frequencyScore + positionBonus - (words.length < 5 ? 0.35 : 0) };
  });

  const fraction = { concise: 0.18, balanced: 0.3, detailed: 0.45 }[length] || 0.3;
  const count = Math.min(usable.length, Math.max(1, Math.ceil(usable.length * fraction)));
  return scored
    .sort((a, b) => b.score - a.score || a.index - b.index)
    .slice(0, count)
    .sort((a, b) => a.index - b.index)
    .map((item) => item.sentence)
    .join("\n\n");
}

async function extractPdfText(file) {
  const pdfjsLib = await import("pdfjs-dist");
  pdfjsLib.GlobalWorkerOptions.workerSrc = pdfWorkerUrl;
  const pdf = await pdfjsLib.getDocument({ data: new Uint8Array(await file.arrayBuffer()) }).promise;
  const pages = [];
  for (let pageNumber = 1; pageNumber <= pdf.numPages; pageNumber += 1) {
    const page = await pdf.getPage(pageNumber);
    const content = await page.getTextContent();
    pages.push(content.items.map((item) => item.str).filter(Boolean).join(" "));
  }
  return { text: pages.join("\n\n"), pageCount: pdf.numPages };
}

async function makeWordDocument(file, summary, pageCount) {
  const { Document, HeadingLevel, Packer, Paragraph, TextRun } = await import("docx");
  const paragraphs = summary.split(/\n\n+/).map(
    (text) => new Paragraph({ children: [new TextRun({ text, size: 22 })], spacing: { after: 180, line: 320 } })
  );
  const document = new Document({
    creator: "MYPDF",
    title: `Summary of ${file.name}`,
    description: "A locally generated summary of a PDF document",
    sections: [{
      properties: {},
      children: [
        new Paragraph({ text: "Document Summary", heading: HeadingLevel.TITLE, spacing: { after: 100 } }),
        new Paragraph({
          children: [new TextRun({ text: file.name, italics: true, color: "737B8C", size: 20 })],
          spacing: { after: 360 }
        }),
        ...paragraphs,
        new Paragraph({
          children: [new TextRun({ text: `Generated with MYPDF · ${pageCount} page${pageCount === 1 ? "" : "s"}`, color: "9299A8", size: 18 })],
          spacing: { before: 300 }
        })
      ]
    }]
  });
  return Packer.toBlob(document);
}

function BrandMark() {
  return (
    <span className="brand-mark" aria-hidden="true">
      <svg viewBox="0 0 28 28" fill="none">
        <path d="M7.5 4.5h8l5 5v14h-13v-19Z" stroke="currentColor" strokeWidth="1.8" strokeLinejoin="round" />
        <path d="M15.5 4.8v5h5M11 14h6M11 18h6" stroke="currentColor" strokeWidth="1.8" strokeLinecap="round" strokeLinejoin="round" />
        <path d="m18.7 16.8 1.1 1.1 2.2-2.5" stroke="#fff" strokeWidth="1.7" strokeLinecap="round" strokeLinejoin="round" />
      </svg>
    </span>
  );
}

function App() {
  const fileInputRef = useRef(null);
  const [file, setFile] = useState(null);
  const [summaryLength, setSummaryLength] = useState("balanced");
  const [summary, setSummary] = useState("");
  const [pageCount, setPageCount] = useState(0);
  const [status, setStatus] = useState("");
  const [isLoading, setIsLoading] = useState(false);
  const [isDownloading, setIsDownloading] = useState(false);
  const [isDragging, setIsDragging] = useState(false);

  function chooseFile(nextFile) {
    setStatus("");
    setSummary("");
    if (!nextFile) return;
    if (nextFile.type !== "application/pdf" && !nextFile.name.toLowerCase().endsWith(".pdf")) {
      setStatus("Please choose a PDF file.");
      return;
    }
    if (nextFile.size > 50 * 1024 * 1024) {
      setStatus("This file is larger than 50 MB. Choose a smaller PDF.");
      return;
    }
    if (nextFile.size === 0) {
      setStatus("This file is empty. Choose another PDF.");
      return;
    }
    setFile(nextFile);
    setPageCount(0);
  }

  function clearFile() {
    setFile(null);
    setSummary("");
    setPageCount(0);
    setStatus("");
    if (fileInputRef.current) fileInputRef.current.value = "";
  }

  async function handleGenerate() {
    if (!file) return;
    setIsLoading(true);
    setStatus("");
    setSummary("");
    try {
      const result = await extractPdfText(file);
      if (result.text.trim().split(/\s+/).length < 20) {
        throw new Error("We couldn't find enough selectable text in this PDF. Scanned documents need OCR before they can be summarized.");
      }
      const nextSummary = summarize(result.text, summaryLength);
      if (!nextSummary) throw new Error("We couldn't find enough readable sentences to summarize in this PDF.");
      setPageCount(result.pageCount);
      setSummary(nextSummary);
    } catch (error) {
      setStatus(error instanceof Error ? error.message : "Something went wrong while reading this PDF.");
    } finally {
      setIsLoading(false);
    }
  }

  async function handleDownload() {
    if (!file || !summary) return;
    setIsDownloading(true);
    setStatus("");
    try {
      const blob = await makeWordDocument(file, summary, pageCount);
      const url = URL.createObjectURL(blob);
      const link = document.createElement("a");
      link.href = url;
      link.download = `${file.name.replace(/\.pdf$/i, "").replace(/[<>:"/\\|?*]/g, "_")}-summary.docx`;
      document.body.append(link);
      link.click();
      link.remove();
      window.setTimeout(() => URL.revokeObjectURL(url), 1000);
    } catch (error) {
      setStatus(error instanceof Error ? error.message : "The Word document could not be created.");
    } finally {
      setIsDownloading(false);
    }
  }

  function handleDrop(event) {
    event.preventDefault();
    setIsDragging(false);
    chooseFile(event.dataTransfer.files?.[0]);
  }

  return (
    <>
      <header className="topbar">
        <a className="brand" href="#" aria-label="MYPDF home">
          <BrandMark />
          <span>MYPDF<span className="brand-period">.</span></span>
        </a>
        <div className="topbar-note"><span className="lock-icon" aria-hidden="true">⌑</span> Private by design</div>
      </header>

      <main>
        <section className="intro">
          <div className="eyebrow"><span className="eyebrow-dot" /> YOUR READING, MADE LIGHTER</div>
          <h1>Big document.<br /><span>Clear takeaways.</span></h1>
          <p className="intro-copy">Turn any PDF into a thoughtful summary, ready to take with you in a Word document.</p>
          <div className="trust-row">
            <span><svg viewBox="0 0 20 20" aria-hidden="true"><path d="m5 10 3.2 3.2L15.5 6" fill="none" stroke="currentColor" strokeWidth="1.7" strokeLinecap="round" strokeLinejoin="round" /></svg> No account needed</span>
            <span><svg viewBox="0 0 20 20" aria-hidden="true"><path d="m5 10 3.2 3.2L15.5 6" fill="none" stroke="currentColor" strokeWidth="1.7" strokeLinecap="round" strokeLinejoin="round" /></svg> Your files stay private</span>
          </div>
        </section>

        <section className="workspace" aria-label="PDF summarizer">
          <div className="main-card">
            <div className="card-heading">
              <div className="step-badge">01</div>
              <div><h2>Upload your PDF</h2><p>Choose a document to get started</p></div>
            </div>

            {!file ? (
              <div
                className={`dropzone${isDragging ? " drag-over" : ""}`}
                role="button"
                tabIndex={0}
                aria-label="Choose a PDF file or drop it here"
                onClick={() => fileInputRef.current?.click()}
                onKeyDown={(event) => {
                  if (event.key === "Enter" || event.key === " ") {
                    event.preventDefault();
                    fileInputRef.current?.click();
                  }
                }}
                onDragEnter={(event) => { event.preventDefault(); setIsDragging(true); }}
                onDragOver={(event) => event.preventDefault()}
                onDragLeave={() => setIsDragging(false)}
                onDrop={handleDrop}
              >
                <input
                  id="file-input"
                  ref={fileInputRef}
                  type="file"
                  accept="application/pdf,.pdf"
                  hidden
                  onClick={(event) => event.stopPropagation()}
                  onChange={(event) => chooseFile(event.target.files?.[0])}
                />
                <div className="upload-icon" aria-hidden="true">
                  <svg viewBox="0 0 48 48" fill="none"><path d="M15 5.5h13l9 9v25H15a4 4 0 0 1-4-4v-26a4 4 0 0 1 4-4Z" stroke="currentColor" strokeWidth="1.8" strokeLinejoin="round" /><path d="M28 6v9h9M24 32V20m0 0-5 5m5-5 5 5" stroke="currentColor" strokeWidth="1.8" strokeLinecap="round" strokeLinejoin="round" /></svg>
                </div>
                <p className="drop-title">Drop your PDF here</p>
                <p className="drop-subtitle">or <span className="browse-button">browse files</span> on your device</p>
                <div className="file-limit"><span className="pdf-label">PDF</span> Up to 50 MB</div>
              </div>
            ) : (
              <div className="selected-file">
                <div className="file-icon" aria-hidden="true"><svg viewBox="0 0 24 24" fill="none"><path d="M6.5 3.5h7l4 4v13h-11v-17Z" stroke="currentColor" strokeWidth="1.5" strokeLinejoin="round" /><path d="M13.5 3.8v4h4M9 12h6M9 15h6" stroke="currentColor" strokeWidth="1.5" strokeLinecap="round" /></svg></div>
                <div className="file-details"><strong>{file.name}</strong><span>{formatFileSize(file.size)}</span></div>
                <button type="button" className="remove-file" aria-label="Remove selected PDF" onClick={clearFile}>×</button>
              </div>
            )}

            <div className="options-row">
              <label htmlFor="summary-length">Summary length</label>
              <div className="select-wrap">
                <select id="summary-length" value={summaryLength} onChange={(event) => setSummaryLength(event.target.value)} aria-label="Summary length">
                  <option value="concise">Concise</option>
                  <option value="balanced">Balanced</option>
                  <option value="detailed">Detailed</option>
                </select>
                <svg viewBox="0 0 16 16" aria-hidden="true"><path d="m4 6 4 4 4-4" fill="none" stroke="currentColor" strokeWidth="1.5" strokeLinecap="round" strokeLinejoin="round" /></svg>
              </div>
            </div>

            <button className={`primary-button${isLoading ? " is-loading" : ""}`} type="button" disabled={!file || isLoading} onClick={handleGenerate}>
              <span>{isLoading ? "Reading your document…" : "Generate my summary"}</span>
              <svg viewBox="0 0 20 20" aria-hidden="true"><path d="M4 10h11m-4-4 4 4-4 4" fill="none" stroke="currentColor" strokeWidth="1.7" strokeLinecap="round" strokeLinejoin="round" /></svg>
            </button>
            <div className="status-message" role="status" aria-live="polite">{status}</div>
            {isLoading && <div className="progress-track"><span /></div>}

            {summary && (
              <div className="result-card">
                <div className="result-heading">
                  <div className="result-check" aria-hidden="true"><svg viewBox="0 0 20 20"><path d="m5 10 3.2 3.2L15.5 6" fill="none" stroke="currentColor" strokeWidth="1.8" strokeLinecap="round" strokeLinejoin="round" /></svg></div>
                  <div><strong>Your summary is ready</strong><span>{pageCount} page{pageCount === 1 ? "" : "s"} · {summary.split(/\s+/).length} words</span></div>
                </div>
                <div className="summary-preview">{summary}</div>
                <button className="download-button" type="button" disabled={isDownloading} onClick={handleDownload}>
                  <svg viewBox="0 0 20 20" aria-hidden="true"><path d="M10 3v9m0 0 3.5-3.5M10 12 6.5 8.5M4 13v3h12v-3" fill="none" stroke="currentColor" strokeWidth="1.5" strokeLinecap="round" strokeLinejoin="round" /></svg>
                  {isDownloading ? "Preparing Word document…" : "Download Word document"}
                </button>
              </div>
            )}
          </div>

          <aside className="side-column">
            <div className="how-card">
              <div className="side-heading"><span className="sparkle" aria-hidden="true">✳</span><h2>How it works</h2></div>
              <ol className="steps-list">
                <li><span className="list-number">1</span><div><strong>Add your PDF</strong><p>Drop in a report, paper, or any PDF.</p></div></li>
                <li><span className="list-number">2</span><div><strong>We find the key ideas</strong><p>A focused summary, made in moments.</p></div></li>
                <li><span className="list-number">3</span><div><strong>Take it with you</strong><p>Download and edit in Microsoft Word.</p></div></li>
              </ol>
            </div>
            <div className="privacy-card">
              <div className="privacy-icon" aria-hidden="true"><svg viewBox="0 0 24 24" fill="none"><path d="M12 3 5 6v5c0 4.5 2.9 7.7 7 10 4.1-2.3 7-5.5 7-10V6l-7-3Z" stroke="currentColor" strokeWidth="1.6" strokeLinejoin="round" /><path d="m9 12 2 2 4-4" stroke="currentColor" strokeWidth="1.6" strokeLinecap="round" strokeLinejoin="round" /></svg></div>
              <div><strong>Your documents stay yours</strong><p>Everything happens in your browser. Your files are never uploaded or stored.</p></div>
            </div>
            <div className="note-card"><span aria-hidden="true">✦</span><p>Works best with PDFs that contain selectable text.</p></div>
          </aside>
        </section>
      </main>

      <footer><span>MYPDF<span className="brand-period">.</span></span><span>© Design and Developed By Team MJ💗 All Rights are Reserved</span></footer>
    </>
  );
}

export default App;
