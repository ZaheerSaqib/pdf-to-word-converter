import io
import os
from flask import Flask, request, send_file, jsonify, render_template_string
from pypdf import PdfReader
from docx import Document
from docx.shared import Pt

app = Flask(__name__)

GITHUB_URL = "https://github.com/ZaheerSaqib/pdf-to-word-converter"

HTML_PAGE = """<!DOCTYPE html>
<html lang="en">
<head>
  <meta charset="UTF-8" />
  <meta name="viewport" content="width=device-width, initial-scale=1.0" />
  <title>PDF to Word Converter — CortexLabs by Zaheer</title>
  <style>
    :root {
      --bg: #f4f7f5;
      --surface: #ffffff;
      --surface-alt: #f9fbfa;
      --ink: #0f172a;
      --body: #334155;
      --muted: #64748b;
      --border: #e2e8f0;
      --accent: #059669;
      --accent-hover: #047857;
      --accent-soft: #ecfdf5;
      --accent-ring: rgba(5, 150, 105, 0.18);
      --drop-bg: linear-gradient(160deg, #10b981 0%, #059669 100%);
      --drop-border: rgba(255, 255, 255, 0.72);
      --shadow-sm: 0 2px 10px rgba(15, 23, 42, 0.05);
      --shadow-md: 0 10px 25px -8px rgba(15, 23, 42, 0.08);
      --shadow-lg: 0 22px 44px -14px rgba(5, 150, 105, 0.38);
    }
    body.dark {
      --bg: #090d14;
      --surface: #111827;
      --surface-alt: #0f1623;
      --ink: #f8fafc;
      --body: #cbd5e1;
      --muted: #94a3b8;
      --border: #1e293b;
      --accent: #10b981;
      --accent-hover: #34d399;
      --accent-soft: rgba(16, 185, 129, 0.14);
      --accent-ring: rgba(16, 185, 129, 0.24);
      --drop-bg: linear-gradient(160deg, #047857 0%, #064e3b 100%);
      --drop-border: rgba(167, 243, 208, 0.55);
      --shadow-sm: 0 2px 10px rgba(0, 0, 0, 0.35);
      --shadow-md: 0 12px 28px -8px rgba(0, 0, 0, 0.55);
      --shadow-lg: 0 24px 48px -16px rgba(0, 0, 0, 0.8);
    }

    * { box-sizing: border-box; margin: 0; padding: 0; }

    @keyframes fadeUp {
      from { opacity: 0; transform: translateY(12px); }
      to { opacity: 1; transform: translateY(0); }
    }

    @keyframes pulseSoft {
      0%, 100% { opacity: 1; }
      50% { opacity: 0.5; }
    }

    body {
      font-family: -apple-system, BlinkMacSystemFont, "Segoe UI", Roboto, Inter, sans-serif;
      background: var(--bg);
      color: var(--ink);
      line-height: 1.55;
      min-height: 100vh;
      display: flex;
      flex-direction: column;
      transition: background-color 0.25s ease, color 0.25s ease;
    }

    header {
      height: 64px;
      background: var(--surface);
      border-bottom: 1px solid var(--border);
      padding: 0 32px;
      display: flex;
      align-items: center;
      justify-content: space-between;
      box-shadow: var(--shadow-sm);
      position: sticky;
      top: 0;
      z-index: 10;
    }

    .brand {
      display: flex;
      align-items: baseline;
      gap: 8px;
      font-weight: 800;
      font-size: 1.18rem;
      letter-spacing: -0.025em;
      color: var(--ink);
    }
    .brand-dot {
      width: 8px;
      height: 8px;
      border-radius: 50%;
      background: var(--accent);
      display: inline-block;
      margin-right: 2px;
      transform: translateY(-1px);
    }
    .brand .by {
      font-weight: 600;
      font-size: 0.84rem;
      color: var(--muted);
      letter-spacing: 0;
    }

    .header-actions {
      display: flex;
      align-items: center;
      gap: 10px;
    }

    .btn-github {
      display: inline-flex;
      align-items: center;
      gap: 7px;
      text-decoration: none;
      background: var(--surface-alt);
      color: var(--ink);
      border: 1px solid var(--border);
      padding: 7px 12px;
      border-radius: 8px;
      font-size: 0.82rem;
      font-weight: 600;
      transition: transform 0.15s ease, border-color 0.15s ease, box-shadow 0.15s ease;
    }
    .btn-github:hover {
      border-color: var(--accent);
      transform: translateY(-1px);
      box-shadow: 0 4px 12px var(--accent-ring);
    }
    .star-icon {
      color: #eab308;
      font-size: 0.9rem;
      line-height: 1;
    }

    .btn-ghost {
      background: var(--surface);
      color: var(--ink);
      border: 1px solid var(--border);
      padding: 7px 13px;
      border-radius: 8px;
      font-size: 0.82rem;
      font-weight: 600;
      cursor: pointer;
      transition: transform 0.15s ease, border-color 0.15s ease, background-color 0.15s ease;
    }
    .btn-ghost:hover {
      border-color: var(--accent);
      transform: translateY(-1px);
    }
    .btn-ghost:active { transform: translateY(0); }

    .btn-primary {
      background: var(--accent);
      color: #ffffff;
      border: none;
      padding: 10px 18px;
      border-radius: 8px;
      font-size: 0.88rem;
      font-weight: 700;
      cursor: pointer;
      text-decoration: none;
      display: inline-flex;
      align-items: center;
      gap: 8px;
      transition: transform 0.15s ease, background-color 0.15s ease, box-shadow 0.15s ease;
    }
    .btn-primary:hover {
      background: var(--accent-hover);
      transform: translateY(-1px);
      box-shadow: 0 8px 18px -6px rgba(5, 150, 105, 0.45);
    }
    .btn-primary:active { transform: translateY(0); }

    main {
      flex: 1;
      width: 100%;
      max-width: 920px;
      margin: 0 auto;
      padding: 44px 24px 64px;
    }

    .hero-block {
      text-align: center;
      margin-bottom: 28px;
      animation: fadeUp 0.32s ease both;
    }
    .hero-title {
      font-size: 2.25rem;
      font-weight: 800;
      letter-spacing: -0.035em;
      line-height: 1.18;
      color: var(--ink);
    }
    .hero-sub {
      margin-top: 8px;
      font-size: 0.96rem;
      color: var(--muted);
      font-weight: 500;
    }

    .dropzone {
      background: var(--drop-bg);
      border-radius: 16px;
      padding: 42px 24px;
      text-align: center;
      color: #ffffff;
      box-shadow: var(--shadow-lg);
      animation: fadeUp 0.42s ease both;
      transition: transform 0.2s ease, box-shadow 0.2s ease, filter 0.2s ease;
    }
    .dropzone:hover { transform: translateY(-2px); }
    .dropzone.dragover {
      filter: brightness(1.06);
      transform: scale(1.01);
    }

    .dropzone-inner {
      border: 1.5px dashed var(--drop-border);
      border-radius: 12px;
      padding: 44px 24px;
      max-width: 720px;
      margin: 0 auto;
      background: rgba(255, 255, 255, 0.04);
      transition: border-color 0.2s ease, background-color 0.2s ease;
    }
    .dropzone.dragover .dropzone-inner {
      border-color: #ffffff;
      background: rgba(255, 255, 255, 0.12);
    }

    .choose-btn {
      background: var(--surface);
      color: var(--ink);
      border: none;
      padding: 14px 28px;
      font-size: 0.95rem;
      font-weight: 700;
      letter-spacing: -0.01em;
      border-radius: 9px;
      cursor: pointer;
      display: inline-flex;
      align-items: center;
      gap: 10px;
      box-shadow: 0 8px 20px rgba(15, 23, 42, 0.2);
      transition: transform 0.16s ease, box-shadow 0.16s ease;
    }
    .choose-btn:hover {
      transform: translateY(-2px);
      box-shadow: 0 12px 24px rgba(15, 23, 42, 0.28);
    }
    .choose-btn:active { transform: translateY(0); }

    .drop-hint {
      margin-top: 14px;
      font-size: 0.9rem;
      font-weight: 500;
      color: rgba(255, 255, 255, 0.92);
    }

    .work-card {
      background: var(--surface);
      border: 1px solid var(--border);
      border-radius: 16px;
      padding: 44px 24px;
      text-align: center;
      box-shadow: var(--shadow-md);
      display: none;
      animation: fadeUp 0.3s ease both;
    }

    .doc-badge {
      width: 88px;
      margin: 0 auto 14px;
      border-radius: 10px;
      overflow: hidden;
      border: 1px solid var(--border);
      box-shadow: var(--shadow-sm);
    }
    .doc-top {
      background: var(--accent-soft);
      padding: 16px 12px;
    }
    .doc-lines div {
      height: 6px;
      background: var(--accent);
      opacity: 0.4;
      border-radius: 3px;
      margin-bottom: 6px;
    }
    .doc-lines div:last-child {
      width: 65%;
      margin-bottom: 0;
    }
    .doc-tag {
      background: #dc2626;
      color: #ffffff;
      font-size: 0.74rem;
      font-weight: 800;
      padding: 5px 0;
      letter-spacing: 0.08em;
      transition: background-color 0.2s ease;
    }
    .doc-tag.docx { background: #059669; }

    .file-meta {
      font-size: 0.88rem;
      font-weight: 600;
      color: var(--muted);
      margin-bottom: 6px;
    }
    .doc-stats {
      font-size: 0.82rem;
      font-weight: 600;
      color: var(--accent);
      margin-bottom: 12px;
      min-height: 18px;
    }
    .status-label {
      font-size: 1.12rem;
      font-weight: 700;
      color: var(--ink);
      margin-bottom: 18px;
    }
    .status-label.working { animation: pulseSoft 1.3s ease-in-out infinite; }
    .error-msg {
      color: #dc2626;
      font-size: 0.9rem;
      font-weight: 600;
      margin-bottom: 16px;
      display: none;
    }

    .info-panel {
      background: var(--surface);
      border: 1px solid var(--border);
      border-radius: 14px;
      padding: 28px 32px;
      margin-top: 28px;
      box-shadow: var(--shadow-sm);
      animation: fadeUp 0.52s ease both;
    }

    .info-grid {
      display: grid;
      grid-template-columns: 1.05fr 1fr;
      gap: 36px;
      align-items: start;
    }
    @media (max-width: 720px) {
      .info-grid { grid-template-columns: 1fr; gap: 24px; }
      .info-panel { padding: 22px 20px; }
    }

    .info-copy {
      border-left: 3px solid var(--accent);
      padding-left: 16px;
    }
    .info-heading {
      font-size: 0.98rem;
      font-weight: 700;
      color: var(--ink);
      margin-bottom: 6px;
      letter-spacing: -0.01em;
    }
    .info-copy p {
      color: var(--body);
      font-size: 0.92rem;
      font-weight: 400;
    }

    .checklist {
      list-style: none;
      display: flex;
      flex-direction: column;
      gap: 12px;
    }
    .checklist li {
      display: flex;
      align-items: flex-start;
      gap: 10px;
      font-size: 0.89rem;
      font-weight: 500;
      color: var(--body);
      line-height: 1.45;
      padding: 6px 8px;
      border-radius: 8px;
      transition: transform 0.15s ease, background-color 0.15s ease;
    }
    .checklist li:hover {
      transform: translateX(3px);
      background: var(--surface-alt);
    }
    .checklist strong {
      color: var(--ink);
      font-weight: 800;
      text-transform: uppercase;
      letter-spacing: 0.03em;
      font-size: 0.82rem;
    }
    .check {
      color: var(--accent);
      font-weight: 800;
      font-size: 0.92rem;
      line-height: 1.4;
      flex-shrink: 0;
    }

    footer {
      text-align: center;
      padding: 20px 24px;
      font-size: 0.82rem;
      font-weight: 500;
      color: var(--muted);
      border-top: 1px solid var(--border);
      background: var(--surface);
    }
  </style>
</head>
<body>
  <header>
    <div class="brand">
      <span class="brand-dot"></span>
      <span>CortexLabs</span>
      <span class="by">by Zaheer</span>
    </div>
    <div class="header-actions">
      <a class="btn-github" href="__GITHUB_URL__" target="_blank" rel="noopener noreferrer" title="Open GitHub & Star">
        <svg width="15" height="15" viewBox="0 0 24 24" fill="currentColor" aria-hidden="true">
          <path d="M12 0C5.37 0 0 5.37 0 12c0 5.31 3.435 9.795 8.205 11.385.6.105.825-.255.825-.57 0-.285-.015-1.23-.015-2.235-3.015.555-3.795-.735-4.035-1.41-.135-.345-.72-1.41-1.23-1.695-.42-.225-1.02-.78-.015-.795.945-.015 1.62.87 1.845 1.23 1.08 1.815 2.805 1.305 3.495.99.105-.78.42-1.305.765-1.605-2.67-.3-5.46-1.335-5.46-5.925 0-1.305.465-2.385 1.23-3.225-.12-.3-.54-1.53.12-3.18 0 0 1.005-.315 3.3 1.23.96-.27 1.98-.405 3-.405s2.04.135 3 .405c2.295-1.56 3.3-1.23 3.3-1.23.66 1.65.24 2.88.12 3.18.765.84 1.23 1.905 1.23 3.225 0 4.605-2.805 5.625-5.475 5.925.435.375.81 1.095.81 2.22 0 1.605-.015 2.895-.015 3.3 0 .315.225.69.825.57A12.02 12.02 0 0 0 24 12c0-6.63-5.37-12-12-12z"/>
        </svg>
        <span>GitHub</span>
        <span class="star-icon">★</span>
      </a>
      <button class="btn-ghost" id="themeToggle" type="button">Dark Theme</button>
    </div>
  </header>

  <main>
    <div class="hero-block">
      <h1 class="hero-title">PDF to Word Converter</h1>
      <p class="hero-sub">Turn PDF files into clean, editable Word documents in seconds.</p>
    </div>

    <div class="dropzone" id="dropzone">
      <div class="dropzone-inner">
        <input type="file" id="fileInput" accept=".pdf,application/pdf" hidden />
        <button class="choose-btn" id="chooseBtn" type="button">+ Choose File</button>
        <p class="drop-hint">or drop a PDF file here</p>
      </div>
    </div>

    <div class="work-card" id="workCard">
      <div class="doc-badge">
        <div class="doc-top">
          <div class="doc-lines"><div></div><div></div><div></div></div>
        </div>
        <div class="doc-tag" id="docTag">PDF</div>
      </div>
      <div class="file-meta" id="fileMeta"></div>
      <div class="doc-stats" id="docStats"></div>
      <div class="status-label working" id="statusLabel">Working...</div>
      <div class="error-msg" id="errorMsg"></div>
      <div id="downloadArea" style="display:none; justify-content:center; gap:12px; flex-wrap:wrap;">
        <a class="btn-primary" id="downloadLink" href="#">Download Word (.docx)</a>
        <button class="btn-ghost" id="resetBtn" type="button">Convert Another</button>
      </div>
    </div>

    <div class="info-panel">
      <div class="info-grid">
        <div class="info-copy">
          <div class="info-heading">Editable DOCX Output</div>
          <p>
            Convert PDF documents into editable Word (.docx) files quickly. Extract page text cleanly
            and open the result in Microsoft Word or Google Docs without extra setup.
          </p>
        </div>
        <ul class="checklist">
          <li>
            <span class="check">✓</span>
            <span><strong>Fast local conversion</strong> with clean page-by-page text extraction</span>
          </li>
          <li>
            <span class="check">✓</span>
            <span><strong>Standard .docx format</strong> ready for editing, comments, and sharing</span>
          </li>
          <li>
            <span class="check">✓</span>
            <span><strong>Zero account required</strong> with instant browser-to-file export</span>
          </li>
        </ul>
      </div>
    </div>
  </main>

  <footer>CortexLabs by Zaheer — PDF to Word Utility</footer>

  <script>
    const themeToggle = document.getElementById("themeToggle");
    themeToggle.addEventListener("click", () => {
      document.body.classList.toggle("dark");
      themeToggle.textContent = document.body.classList.contains("dark") ? "Light Theme" : "Dark Theme";
    });

    const dropzone = document.getElementById("dropzone");
    const fileInput = document.getElementById("fileInput");
    const chooseBtn = document.getElementById("chooseBtn");
    const workCard = document.getElementById("workCard");
    const docTag = document.getElementById("docTag");
    const fileMeta = document.getElementById("fileMeta");
    const docStats = document.getElementById("docStats");
    const statusLabel = document.getElementById("statusLabel");
    const errorMsg = document.getElementById("errorMsg");
    const downloadArea = document.getElementById("downloadArea");
    const downloadLink = document.getElementById("downloadLink");
    const resetBtn = document.getElementById("resetBtn");

    let activeBlobUrl = null;

    chooseBtn.addEventListener("click", () => fileInput.click());

    ["dragenter", "dragover"].forEach(evt => {
      dropzone.addEventListener(evt, e => {
        e.preventDefault();
        dropzone.classList.add("dragover");
      });
    });

    ["dragleave", "drop"].forEach(evt => {
      dropzone.addEventListener(evt, e => {
        e.preventDefault();
        dropzone.classList.remove("dragover");
      });
    });

    dropzone.addEventListener("drop", e => {
      const file = e.dataTransfer.files[0];
      if (file) handleFile(file);
    });

    fileInput.addEventListener("change", () => {
      if (fileInput.files[0]) handleFile(fileInput.files[0]);
    });

    resetBtn.addEventListener("click", () => {
      if (activeBlobUrl) {
        URL.revokeObjectURL(activeBlobUrl);
        activeBlobUrl = null;
      }
      fileInput.value = "";
      workCard.style.display = "none";
      dropzone.style.display = "block";
    });

    async function handleFile(file) {
      if (!file.name.toLowerCase().endsWith(".pdf")) {
        alert("Please select a valid .pdf file.");
        return;
      }

      if (activeBlobUrl) {
        URL.revokeObjectURL(activeBlobUrl);
        activeBlobUrl = null;
      }

      dropzone.style.display = "none";
      workCard.style.display = "block";
      downloadArea.style.display = "none";
      errorMsg.style.display = "none";
      docStats.textContent = "";
      docTag.textContent = "PDF";
      docTag.classList.remove("docx");

      const sizeMb = (file.size / (1024 * 1024)).toFixed(2);
      fileMeta.textContent = `${file.name} (${sizeMb} MB)`;
      statusLabel.textContent = "Working...";
      statusLabel.classList.add("working");

      const formData = new FormData();
      formData.append("file", file);

      try {
        const res = await fetch("/convert", { method: "POST", body: formData });
        if (!res.ok) {
          const errData = await res.json().catch(() => ({}));
          throw new Error(errData.error || "Conversion failed.");
        }

        const pages = Number(res.headers.get("X-Pages") || 0);
        const words = Number(res.headers.get("X-Words") || 0);
        const outName = file.name.replace(/\.pdf$/i, "") + ".docx";
        const blob = await res.blob();
        activeBlobUrl = URL.createObjectURL(blob);

        docTag.textContent = "DOCX";
        docTag.classList.add("docx");
        docStats.textContent = `${pages} page(s) • ${words.toLocaleString()} words extracted`;
        statusLabel.classList.remove("working");
        statusLabel.textContent = "Ready";
        downloadLink.href = activeBlobUrl;
        downloadLink.setAttribute("download", outName);
        downloadLink.style.display = "inline-flex";
        downloadArea.style.display = "flex";
      } catch (err) {
        statusLabel.classList.remove("working");
        statusLabel.textContent = "Could not convert file";
        errorMsg.textContent = err.message;
        errorMsg.style.display = "block";
        downloadArea.style.display = "flex";
        downloadLink.style.display = "none";
      }
    }
  </script>
</body>
</html>
""".replace("__GITHUB_URL__", GITHUB_URL)


def pdf_to_docx_bytes(pdf_bytes: bytes) -> tuple[bytes, int, int]:
    reader = PdfReader(io.BytesIO(pdf_bytes))
    doc = Document()

    style = doc.styles["Normal"]
    style.font.name = "Calibri"
    style.font.size = Pt(11)

    page_count = len(reader.pages)
    word_count = 0

    for idx, page in enumerate(reader.pages, start=1):
        if idx > 1:
            doc.add_page_break()
        text = page.extract_text() or ""
        for line in text.splitlines():
            clean = line.strip()
            if clean:
                word_count += len(clean.split())
                doc.add_paragraph(clean)

    out = io.BytesIO()
    doc.save(out)
    out.seek(0)
    return out.read(), page_count, word_count


@app.route("/")
def index():
    return render_template_string(HTML_PAGE)


@app.route("/convert", methods=["POST"])
def convert():
    uploaded = request.files.get("file")
    if not uploaded or not uploaded.filename:
        return jsonify({"error": "No file uploaded."}), 400

    pdf_bytes = uploaded.read()
    if not pdf_bytes:
        return jsonify({"error": "Uploaded file is empty."}), 400

    try:
        docx_bytes, page_count, word_count = pdf_to_docx_bytes(pdf_bytes)
    except Exception as exc:
        return jsonify({"error": f"Failed to parse PDF: {exc}"}), 400

    base_name = os.path.splitext(os.path.basename(uploaded.filename))[0]
    out_name = f"{base_name}.docx"
    response = send_file(
        io.BytesIO(docx_bytes),
        as_attachment=True,
        download_name=out_name,
        mimetype="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
    )
    response.headers["X-Pages"] = str(page_count)
    response.headers["X-Words"] = str(word_count)
    return response


if __name__ == "__main__":
    app.run(host="127.0.0.1", port=5000, debug=True)