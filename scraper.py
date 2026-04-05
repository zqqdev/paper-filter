#!/usr/bin/env python3
"""
IEEE Document Scraper — Local web tool for searching and downloading
documents by keyword match inside document content.

Usage:
    pip install -r requirements.txt
    python scraper.py

Then open http://localhost:5000 in your browser.
"""

import os
import re
import shutil
import tempfile
import threading
import time
from pathlib import Path
from urllib.parse import urljoin, urlparse, unquote

import requests
import urllib3
from bs4 import BeautifulSoup
from flask import Flask, render_template_string, request, jsonify

# Text extraction
import pdfplumber
from pptx import Presentation
from docx import Document as DocxDocument

# Suppress InsecureRequestWarning when SSL verification is disabled
urllib3.disable_warnings(urllib3.exceptions.InsecureRequestWarning)

app = Flask(__name__)

# ---------------------------------------------------------------------------
# Configuration
# ---------------------------------------------------------------------------
DOWNLOAD_DIR = Path.home() / "ieee_downloads"
DOWNLOAD_DIR.mkdir(exist_ok=True)

DOCUMENT_EXTENSIONS = {
    ".pdf", ".doc", ".docx", ".ppt", ".pptx",
    ".xls", ".xlsx", ".csv", ".txt", ".rtf",
}

# Extensions we can actually extract text from
EXTRACTABLE_EXTENSIONS = {".pdf", ".pptx", ".docx", ".txt", ".csv", ".rtf"}

HEADERS = {
    "User-Agent": (
        "Mozilla/5.0 (Windows NT 10.0; Win64; x64) "
        "AppleWebKit/537.36 (KHTML, like Gecko) "
        "Chrome/120.0.0.0 Safari/537.36"
    )
}

# Shared state for background tasks
task_status = {}

# ---------------------------------------------------------------------------
# Text extraction from documents
# ---------------------------------------------------------------------------

def extract_text(filepath: Path) -> str:
    """Extract text content from a document file."""
    ext = filepath.suffix.lower()
    try:
        if ext == ".pdf":
            return _extract_pdf(filepath)
        elif ext == ".pptx":
            return _extract_pptx(filepath)
        elif ext == ".docx":
            return _extract_docx(filepath)
        elif ext in (".txt", ".csv", ".rtf"):
            return filepath.read_text(errors="ignore")
        else:
            return ""
    except Exception as e:
        print(f"  [WARN] Could not extract text from {filepath.name}: {e}")
        return ""


def _extract_pdf(filepath: Path) -> str:
    parts = []
    with pdfplumber.open(filepath) as pdf:
        for page in pdf.pages:
            text = page.extract_text()
            if text:
                parts.append(text)
    return "\n".join(parts)


def _extract_pptx(filepath: Path) -> str:
    parts = []
    prs = Presentation(str(filepath))
    for slide in prs.slides:
        for shape in slide.shapes:
            if shape.has_text_frame:
                for para in shape.text_frame.paragraphs:
                    text = para.text.strip()
                    if text:
                        parts.append(text)
            if shape.has_table:
                for row in shape.table.rows:
                    for cell in row.cells:
                        text = cell.text.strip()
                        if text:
                            parts.append(text)
        if slide.has_notes_slide and slide.notes_slide.notes_text_frame:
            notes = slide.notes_slide.notes_text_frame.text.strip()
            if notes:
                parts.append(notes)
    return "\n".join(parts)


def _extract_docx(filepath: Path) -> str:
    doc = DocxDocument(str(filepath))
    parts = [p.text for p in doc.paragraphs if p.text.strip()]
    for table in doc.tables:
        for row in table.rows:
            for cell in row.cells:
                if cell.text.strip():
                    parts.append(cell.text.strip())
    return "\n".join(parts)


def find_snippets(text: str, keyword: str, case_sensitive: bool,
                  max_snippets: int = 3, context_chars: int = 80) -> list:
    """Find keyword occurrences and return surrounding context snippets."""
    if not text:
        return []
    flags = 0 if case_sensitive else re.IGNORECASE
    pattern = re.compile(re.escape(keyword), flags)
    snippets = []
    for m in pattern.finditer(text):
        start = max(0, m.start() - context_chars)
        end = min(len(text), m.end() + context_chars)
        snippet = text[start:end].replace("\n", " ").strip()
        if start > 0:
            snippet = "..." + snippet
        if end < len(text):
            snippet = snippet + "..."
        snippets.append(snippet)
        if len(snippets) >= max_snippets:
            break
    return snippets


# ---------------------------------------------------------------------------
# Scraping logic
# ---------------------------------------------------------------------------

def get_all_doc_links(url: str, verify_ssl: bool = False) -> list:
    """Fetch page and return ALL document links (no keyword filter)."""
    resp = requests.get(url, headers=HEADERS, timeout=30, verify=verify_ssl)
    resp.raise_for_status()
    soup = BeautifulSoup(resp.text, "html.parser")
    results = []
    seen = set()

    for a_tag in soup.find_all("a", href=True):
        href = a_tag["href"]
        parsed = urlparse(href)
        path_lower = unquote(parsed.path).lower()
        if not any(path_lower.endswith(ext) for ext in DOCUMENT_EXTENSIONS):
            continue

        full_url = urljoin(url, href)
        if full_url in seen:
            continue
        seen.add(full_url)

        link_text = a_tag.get_text(strip=True)
        filename = unquote(urlparse(full_url).path.split("/")[-1])
        ext = Path(filename).suffix.lower()

        results.append({
            "url": full_url,
            "filename": filename,
            "link_text": link_text or filename,
            "extension": ext,
        })

    return results


def content_search_worker(task_id: str, url: str, keyword: str,
                          case_sensitive: bool, search_mode: str,
                          verify_ssl: bool):
    """
    Background worker:
      1. Fetch page, get all doc links
      2. If mode is 'filename': filter by filename/link text only
      3. If mode is 'content': download each doc, extract text, search
      4. Update task_status with progress and results
    """
    status = task_status[task_id]
    status["status"] = "fetching_page"

    try:
        all_links = get_all_doc_links(url, verify_ssl)
    except Exception as e:
        status["status"] = "error"
        status["error"] = f"Failed to fetch page: {e}"
        return

    status["total"] = len(all_links)

    if search_mode == "filename":
        status["status"] = "searching"
        matches = []
        for i, doc in enumerate(all_links):
            searchable = f"{doc['link_text']} {doc['filename']}"
            if case_sensitive:
                found = keyword in searchable
            else:
                found = keyword.lower() in searchable.lower()
            if found:
                doc["match_type"] = "filename"
                doc["snippets"] = []
                matches.append(doc)
            status["done"] = i + 1
        status["results"] = matches
        status["status"] = "complete"
        return

    # Content search mode
    status["status"] = "scanning"
    tmpdir = tempfile.mkdtemp(prefix="ieee_scraper_")
    matches = []

    try:
        for i, doc in enumerate(all_links):
            status["done"] = i + 1
            status["current_file"] = doc["filename"]

            # Quick check: filename/link text match (no download needed)
            searchable = f"{doc['link_text']} {doc['filename']}"
            if case_sensitive:
                name_match = keyword in searchable
            else:
                name_match = keyword.lower() in searchable.lower()

            if name_match:
                doc["match_type"] = "filename"
                doc["snippets"] = ["(Matched in filename / link text)"]
                matches.append(doc)
                continue

            # Skip files we can't extract text from
            ext = doc["extension"].lower()
            if ext not in EXTRACTABLE_EXTENSIONS:
                continue

            # Download to temp, extract text, search
            try:
                r = requests.get(doc["url"], headers=HEADERS, timeout=60,
                                 stream=True, verify=verify_ssl)
                r.raise_for_status()
                tmp_path = Path(tmpdir) / doc["filename"]
                with open(tmp_path, "wb") as f:
                    for chunk in r.iter_content(chunk_size=8192):
                        f.write(chunk)

                text = extract_text(tmp_path)
                if not text:
                    continue

                if case_sensitive:
                    found = keyword in text
                else:
                    found = keyword.lower() in text.lower()

                if found:
                    snippets = find_snippets(text, keyword, case_sensitive)
                    doc["match_type"] = "content"
                    doc["snippets"] = snippets
                    matches.append(doc)

                tmp_path.unlink(missing_ok=True)

            except Exception as e:
                status.setdefault("warnings", []).append(
                    f"{doc['filename']}: {e}"
                )

        status["results"] = matches
        status["status"] = "complete"
    finally:
        shutil.rmtree(tmpdir, ignore_errors=True)


def download_worker(task_id: str, files: list, folder: str, verify_ssl: bool):
    """Background worker to download selected files."""
    dest = DOWNLOAD_DIR / folder
    dest.mkdir(parents=True, exist_ok=True)
    status = task_status[task_id]
    status["status"] = "running"

    for i, f in enumerate(files):
        try:
            r = requests.get(f["url"], headers=HEADERS, timeout=60,
                             stream=True, verify=verify_ssl)
            r.raise_for_status()
            filepath = dest / f["filename"]
            counter = 1
            while filepath.exists():
                stem = Path(f["filename"]).stem
                ext = Path(f["filename"]).suffix
                filepath = dest / f"{stem}_{counter}{ext}"
                counter += 1
            with open(filepath, "wb") as out:
                for chunk in r.iter_content(chunk_size=8192):
                    out.write(chunk)
            status["files"].append({
                "filename": filepath.name,
                "path": str(filepath),
                "size": filepath.stat().st_size,
            })
        except Exception as e:
            status["errors"].append({"filename": f["filename"], "error": str(e)})
        status["done"] = i + 1

    status["status"] = "complete"


# ---------------------------------------------------------------------------
# HTML Template
# ---------------------------------------------------------------------------

HTML_TEMPLATE = r"""
<!DOCTYPE html>
<html lang="en">
<head>
<meta charset="UTF-8">
<meta name="viewport" content="width=device-width, initial-scale=1.0">
<title>Document Scraper</title>
<link rel="preconnect" href="https://fonts.googleapis.com">
<link href="https://fonts.googleapis.com/css2?family=JetBrains+Mono:wght@400;600&family=DM+Sans:ital,wght@0,400;0,500;0,600;0,700;1,400&display=swap" rel="stylesheet">
<style>
  :root {
    --bg: #0e1117;
    --surface: #161b22;
    --surface2: #1c2333;
    --border: #2a3140;
    --text: #e6edf3;
    --text-dim: #8b949e;
    --accent: #58a6ff;
    --accent-glow: rgba(88, 166, 255, 0.15);
    --green: #3fb950;
    --red: #f85149;
    --orange: #d29922;
    --yellow-glow: rgba(210,153,34,0.15);
    --green-glow: rgba(63,185,80,0.12);
    --radius: 8px;
  }
  * { box-sizing: border-box; margin: 0; padding: 0; }
  body {
    font-family: 'DM Sans', system-ui, sans-serif;
    background: var(--bg);
    color: var(--text);
    min-height: 100vh;
  }
  .header {
    padding: 2rem 2rem 1rem;
    border-bottom: 1px solid var(--border);
    background: var(--surface);
  }
  .header h1 {
    font-size: 1.4rem; font-weight: 700; letter-spacing: -0.02em;
    display: flex; align-items: center; gap: 0.6rem;
  }
  .header h1 .icon {
    width: 28px; height: 28px; background: var(--accent); border-radius: 6px;
    display: flex; align-items: center; justify-content: center;
    font-size: 14px; color: var(--bg); font-weight: 700;
  }
  .header p { color: var(--text-dim); font-size: 0.85rem; margin-top: 0.4rem; padding-left: 2.4rem; }
  .main { padding: 1.5rem 2rem 3rem; max-width: 1100px; }

  .form-grid { display: grid; grid-template-columns: 1fr 1fr auto; gap: 0.75rem; align-items: end; }
  .field label {
    display: block; font-size: 0.75rem; font-weight: 600;
    text-transform: uppercase; letter-spacing: 0.06em;
    color: var(--text-dim); margin-bottom: 0.35rem;
  }
  .field input {
    width: 100%; padding: 0.6rem 0.8rem; font-size: 0.9rem;
    font-family: 'JetBrains Mono', monospace;
    background: var(--surface); border: 1px solid var(--border);
    border-radius: var(--radius); color: var(--text); outline: none;
    transition: border-color 0.2s;
  }
  .field input:focus { border-color: var(--accent); box-shadow: 0 0 0 3px var(--accent-glow); }
  .field input::placeholder { color: var(--text-dim); opacity: 0.5; }

  .btn {
    padding: 0.6rem 1.4rem; font-size: 0.85rem; font-weight: 600;
    font-family: 'DM Sans', sans-serif; border: none;
    border-radius: var(--radius); cursor: pointer;
    transition: all 0.15s; white-space: nowrap;
  }
  .btn-primary { background: var(--accent); color: var(--bg); }
  .btn-primary:hover { filter: brightness(1.15); }
  .btn-primary:disabled { opacity: 0.4; cursor: not-allowed; }
  .btn-success { background: var(--green); color: var(--bg); }
  .btn-success:hover { filter: brightness(1.1); }
  .btn-success:disabled { opacity: 0.4; cursor: not-allowed; }

  .options-row {
    margin-top: 0.6rem; display: flex; align-items: center;
    gap: 1.2rem; flex-wrap: wrap;
  }
  .checkbox-label {
    font-size: 0.8rem; color: var(--text-dim);
    display: flex; align-items: center; gap: 0.4rem; cursor: pointer;
  }
  .checkbox-label input[type="checkbox"] { accent-color: var(--accent); }

  .mode-toggle {
    display: inline-flex; border: 1px solid var(--border);
    border-radius: var(--radius); overflow: hidden;
  }
  .mode-toggle button {
    padding: 0.35rem 0.9rem; font-size: 0.78rem; font-weight: 600;
    font-family: 'DM Sans', sans-serif; background: transparent;
    color: var(--text-dim); border: none; cursor: pointer; transition: all 0.15s;
  }
  .mode-toggle button.active { background: var(--accent); color: var(--bg); }
  .mode-toggle button:not(.active):hover { background: var(--surface2); color: var(--text); }

  .status-bar {
    margin-top: 1.2rem; padding: 0.7rem 1rem; background: var(--surface2);
    border-radius: var(--radius); font-size: 0.82rem;
    font-family: 'JetBrains Mono', monospace; color: var(--text-dim);
    display: none; flex-direction: column; gap: 0.5rem;
  }
  .status-bar.visible { display: flex; }
  .status-row { display: flex; align-items: center; gap: 0.6rem; }
  .spinner {
    width: 14px; height: 14px; border: 2px solid var(--border);
    border-top-color: var(--accent); border-radius: 50%;
    animation: spin 0.7s linear infinite; flex-shrink: 0;
  }
  @keyframes spin { to { transform: rotate(360deg); } }
  .progress-bar-outer {
    width: 100%; height: 4px; background: var(--surface);
    border-radius: 2px; overflow: hidden;
  }
  .progress-bar-inner {
    height: 100%; background: var(--accent); border-radius: 2px;
    width: 0%; transition: width 0.3s;
  }

  .results-section { margin-top: 1.5rem; display: none; }
  .results-section.visible { display: block; }
  .results-header {
    display: flex; align-items: center; justify-content: space-between;
    margin-bottom: 0.75rem;
  }
  .results-header h2 { font-size: 0.95rem; font-weight: 600; }
  .badge {
    font-size: 0.7rem; font-weight: 600; padding: 0.2rem 0.55rem;
    border-radius: 99px; font-family: 'JetBrains Mono', monospace;
    margin-left: 0.5rem;
  }
  .badge-blue { background: var(--accent-glow); color: var(--accent); }

  .results-table { width: 100%; border-collapse: collapse; }
  .results-table thead th {
    font-size: 0.7rem; font-weight: 600; text-transform: uppercase;
    letter-spacing: 0.06em; color: var(--text-dim); text-align: left;
    padding: 0.5rem 0.75rem; border-bottom: 1px solid var(--border);
  }
  .results-table tbody tr {
    border-bottom: 1px solid var(--border); transition: background 0.1s;
  }
  .results-table tbody tr:hover { background: var(--surface2); }
  .results-table td { padding: 0.6rem 0.75rem; font-size: 0.82rem; vertical-align: top; }
  .results-table td:first-child { width: 30px; vertical-align: middle; }
  .results-table input[type="checkbox"] { accent-color: var(--accent); }

  .ext-badge {
    display: inline-block; font-size: 0.65rem; font-weight: 600;
    font-family: 'JetBrains Mono', monospace; text-transform: uppercase;
    padding: 0.15rem 0.45rem; border-radius: 4px;
    background: var(--surface); border: 1px solid var(--border); color: var(--text-dim);
  }
  .match-badge {
    display: inline-block; font-size: 0.6rem; font-weight: 600;
    padding: 0.12rem 0.4rem; border-radius: 4px;
  }
  .match-content { background: var(--green-glow); color: var(--green); }
  .match-filename { background: var(--yellow-glow); color: var(--orange); }

  .link-text { color: var(--text); display: inline; }
  .file-url {
    font-size: 0.7rem; color: var(--text-dim);
    font-family: 'JetBrains Mono', monospace;
    max-width: 500px; overflow: hidden; text-overflow: ellipsis;
    white-space: nowrap; display: block; margin-top: 0.15rem;
  }
  .file-url a { color: var(--text-dim); text-decoration: none; }
  .file-url a:hover { color: var(--accent); }

  .snippet {
    font-size: 0.75rem; color: var(--text-dim); font-style: italic;
    margin-top: 0.3rem; line-height: 1.5; padding: 0.35rem 0.5rem;
    background: var(--surface); border-radius: 4px; border-left: 2px solid var(--green);
  }
  .snippet mark {
    background: rgba(63,185,80,0.25); color: var(--green);
    border-radius: 2px; padding: 0 2px; font-style: normal; font-weight: 600;
  }

  .download-bar { margin-top: 1rem; display: flex; align-items: center; gap: 1rem; }
  .dl-progress-wrap { flex: 1; display: none; }
  .dl-progress-wrap.visible { display: block; }
  .dl-progress-text {
    font-size: 0.75rem; color: var(--text-dim);
    font-family: 'JetBrains Mono', monospace; margin-top: 0.25rem;
  }

  .select-actions { display: flex; gap: 0.6rem; }
  .select-actions button {
    background: none; border: none; color: var(--accent); cursor: pointer;
    font-family: 'DM Sans', sans-serif; font-size: 0.75rem;
  }
  .select-actions button:hover { text-decoration: underline; }

  .empty-state {
    text-align: center; padding: 3rem 1rem; color: var(--text-dim);
    font-size: 0.9rem; display: none;
  }
  .warnings {
    margin-top: 0.75rem; font-size: 0.75rem; color: var(--orange);
    font-family: 'JetBrains Mono', monospace;
  }

  @media (max-width: 700px) {
    .form-grid { grid-template-columns: 1fr; }
    .main { padding: 1rem; }
  }
</style>
</head>
<body>

<div class="header">
  <h1><span class="icon">S</span> Document Scraper</h1>
  <p>Search inside document content — downloads and scans PDFs, PPTX, DOCX and more.</p>
</div>

<div class="main">
  <div class="form-grid">
    <div class="field">
      <label>Page URL</label>
      <input type="url" id="urlInput" placeholder="https://www.ieee802.org/11/Reports/wng_update.htm"
             value="https://www.ieee802.org/11/Reports/wng_update.htm">
    </div>
    <div class="field">
      <label>Keyword</label>
      <input type="text" id="keywordInput" placeholder="e.g. sensing, WLAN, 6GHz">
    </div>
    <button class="btn btn-primary" id="searchBtn" onclick="doSearch()">Search</button>
  </div>

  <div class="options-row">
    <span style="font-size:0.75rem; color:var(--text-dim); font-weight:600;">SEARCH MODE</span>
    <div class="mode-toggle">
      <button id="modeContent" class="active" onclick="setMode('content')">Inside content</button>
      <button id="modeFilename" onclick="setMode('filename')">Filename only</button>
    </div>
    <label class="checkbox-label">
      <input type="checkbox" id="caseSensitive"> Case sensitive
    </label>
    <label class="checkbox-label">
      <input type="checkbox" id="skipSSL" checked> Skip SSL verification
    </label>
  </div>

  <div class="status-bar" id="statusBar">
    <div class="status-row">
      <div class="spinner" id="statusSpinner"></div>
      <span id="statusText">Searching...</span>
    </div>
    <div class="progress-bar-outer" id="scanProgress" style="display:none;">
      <div class="progress-bar-inner" id="scanProgressBar"></div>
    </div>
  </div>

  <div class="results-section" id="resultsSection">
    <div class="results-header">
      <h2>Results <span class="badge badge-blue" id="countBadge">0</span></h2>
      <div class="select-actions">
        <button onclick="selectAll()">Select all</button>
        <button onclick="selectNone()">Select none</button>
      </div>
    </div>
    <table class="results-table">
      <thead>
        <tr><th></th><th>Document</th><th>Type</th><th>Match</th></tr>
      </thead>
      <tbody id="resultsBody"></tbody>
    </table>
    <div class="download-bar">
      <button class="btn btn-success" id="downloadBtn" onclick="doDownload()">Download selected</button>
      <div class="dl-progress-wrap" id="dlProgressWrap">
        <div class="progress-bar-outer"><div class="progress-bar-inner" id="dlProgressBar"></div></div>
        <div class="dl-progress-text" id="dlProgressText"></div>
      </div>
    </div>
    <div class="warnings" id="warnings"></div>
  </div>

  <div class="empty-state" id="emptyState">No matching documents found. Try a different keyword.</div>
</div>

<script>
let searchResults = [];
let currentMode = 'content';

function setMode(mode) {
  currentMode = mode;
  document.getElementById('modeContent').classList.toggle('active', mode === 'content');
  document.getElementById('modeFilename').classList.toggle('active', mode === 'filename');
}

function esc(s) {
  const d = document.createElement('div');
  d.textContent = s;
  return d.innerHTML;
}

function highlightKeyword(text, keyword, caseSensitive) {
  const escaped = keyword.replace(/[.*+?^${}()|[\]\\]/g, '\\$&');
  const re = new RegExp(`(${escaped})`, caseSensitive ? 'g' : 'gi');
  return esc(text).replace(re, '<mark>$1</mark>');
}

async function doSearch() {
  const url = document.getElementById('urlInput').value.trim();
  const keyword = document.getElementById('keywordInput').value.trim();
  if (!url || !keyword) return alert('Please enter both a URL and a keyword.');

  const btn = document.getElementById('searchBtn');
  btn.disabled = true;
  btn.textContent = 'Searching...';
  hideResults();

  try {
    const resp = await fetch('/api/search', {
      method: 'POST',
      headers: {'Content-Type': 'application/json'},
      body: JSON.stringify({
        url, keyword,
        case_sensitive: document.getElementById('caseSensitive').checked,
        skip_ssl: document.getElementById('skipSSL').checked,
        mode: currentMode,
      }),
    });
    const data = await resp.json();
    if (data.error) {
      showStatus('Error: ' + data.error, false);
      btn.disabled = false; btn.textContent = 'Search';
      return;
    }
    pollSearch(data.task_id, keyword, document.getElementById('caseSensitive').checked);
  } catch (e) {
    showStatus('Network error: ' + e.message, false);
    btn.disabled = false; btn.textContent = 'Search';
  }
}

async function pollSearch(taskId, keyword, caseSensitive) {
  const btn = document.getElementById('searchBtn');
  const progressOuter = document.getElementById('scanProgress');
  const progressBar = document.getElementById('scanProgressBar');

  const interval = setInterval(async () => {
    try {
      const resp = await fetch('/api/status/' + taskId);
      const s = await resp.json();

      if (s.status === 'error') {
        clearInterval(interval);
        showStatus('Error: ' + (s.error || 'Unknown error'), false);
        btn.disabled = false; btn.textContent = 'Search';
        return;
      }
      if (s.status === 'fetching_page') {
        showStatus('Fetching page and finding document links...', true);
      } else if (s.status === 'searching' || s.status === 'scanning') {
        const total = s.total || 0, done = s.done || 0;
        const pct = total > 0 ? Math.round((done / total) * 100) : 0;
        const cf = s.current_file ? ` \u2014 ${s.current_file}` : '';
        const label = currentMode === 'content'
          ? `Scanning document content: ${done}/${total}${cf}`
          : `Checking filenames: ${done}/${total}`;
        showStatus(label, true);
        progressOuter.style.display = 'block';
        progressBar.style.width = pct + '%';
      }
      if (s.status === 'complete') {
        clearInterval(interval);
        hideStatus();
        const results = s.results || [];
        if (results.length === 0) {
          document.getElementById('emptyState').style.display = 'block';
        } else {
          searchResults = results;
          renderResults(results, keyword, caseSensitive);
        }
        if (s.warnings && s.warnings.length > 0) {
          document.getElementById('warnings').textContent =
            '\u26A0 ' + s.warnings.length + ' file(s) could not be processed.';
        }
        btn.disabled = false; btn.textContent = 'Search';
      }
    } catch (e) {
      clearInterval(interval);
      showStatus('Error polling status.', false);
      btn.disabled = false; btn.textContent = 'Search';
    }
  }, 600);
}

function renderResults(results, keyword, caseSensitive) {
  const tbody = document.getElementById('resultsBody');
  tbody.innerHTML = '';
  results.forEach((r, i) => {
    const tr = document.createElement('tr');
    const matchLabel = r.match_type === 'content' ? 'content' : 'filename';
    const matchClass = r.match_type === 'content' ? 'match-content' : 'match-filename';

    let snippetHtml = '';
    if (r.snippets && r.snippets.length > 0 && r.match_type === 'content') {
      snippetHtml = r.snippets.map(s =>
        '<div class="snippet">' + highlightKeyword(s, keyword, caseSensitive) + '</div>'
      ).join('');
    }

    tr.innerHTML =
      '<td><input type="checkbox" data-idx="' + i + '" checked></td>' +
      '<td>' +
        '<span class="link-text">' + esc(r.link_text) + '</span>' +
        '<span class="file-url"><a href="' + esc(r.url) + '" target="_blank">' + esc(r.filename) + '</a></span>' +
        snippetHtml +
      '</td>' +
      '<td><span class="ext-badge">' + esc(r.extension.replace('.','')) + '</span></td>' +
      '<td><span class="match-badge ' + matchClass + '">' + matchLabel + '</span></td>';
    tbody.appendChild(tr);
  });
  document.getElementById('countBadge').textContent = results.length;
  document.getElementById('resultsSection').classList.add('visible');
  document.getElementById('emptyState').style.display = 'none';
}

function selectAll() {
  document.querySelectorAll('#resultsBody input[type=checkbox]').forEach(c => c.checked = true);
}
function selectNone() {
  document.querySelectorAll('#resultsBody input[type=checkbox]').forEach(c => c.checked = false);
}

async function doDownload() {
  const checked = [...document.querySelectorAll('#resultsBody input[type=checkbox]:checked')];
  if (!checked.length) return alert('Select at least one document.');

  const files = checked.map(c => searchResults[parseInt(c.dataset.idx)]);
  const keyword = document.getElementById('keywordInput').value.trim();
  const folder = keyword.replace(/[^a-zA-Z0-9_-]/g, '_');

  const resp = await fetch('/api/download', {
    method: 'POST',
    headers: {'Content-Type': 'application/json'},
    body: JSON.stringify({
      files, folder,
      skip_ssl: document.getElementById('skipSSL').checked,
    }),
  });
  const data = await resp.json();
  if (data.error) return alert(data.error);

  document.getElementById('downloadBtn').disabled = true;
  document.getElementById('dlProgressWrap').classList.add('visible');
  pollDownload(data.task_id, files.length);
}

async function pollDownload(taskId, total) {
  const bar = document.getElementById('dlProgressBar');
  const text = document.getElementById('dlProgressText');
  const interval = setInterval(async () => {
    try {
      const resp = await fetch('/api/status/' + taskId);
      const s = await resp.json();
      const pct = Math.round((s.done / total) * 100);
      bar.style.width = pct + '%';
      text.textContent = s.done + ' / ' + total + ' files';
      if (s.status === 'complete') {
        clearInterval(interval);
        document.getElementById('downloadBtn').disabled = false;
        const errors = s.errors && s.errors.length ? ' (' + s.errors.length + ' failed)' : '';
        text.textContent = 'Done! ' + s.files.length + ' files saved to ' + s.folder + errors;
      }
    } catch (e) {
      clearInterval(interval);
      text.textContent = 'Error polling download status.';
    }
  }, 500);
}

function showStatus(msg, spinning) {
  const bar = document.getElementById('statusBar');
  bar.classList.add('visible');
  document.getElementById('statusText').textContent = msg;
  document.getElementById('statusSpinner').style.display = spinning ? 'block' : 'none';
}
function hideStatus() {
  document.getElementById('statusBar').classList.remove('visible');
  document.getElementById('scanProgress').style.display = 'none';
  document.getElementById('scanProgressBar').style.width = '0%';
}
function hideResults() {
  document.getElementById('resultsSection').classList.remove('visible');
  document.getElementById('emptyState').style.display = 'none';
  document.getElementById('dlProgressWrap').classList.remove('visible');
  document.getElementById('warnings').textContent = '';
}

document.getElementById('keywordInput').addEventListener('keydown', e => {
  if (e.key === 'Enter') doSearch();
});
</script>
</body>
</html>
"""

# ---------------------------------------------------------------------------
# Flask routes
# ---------------------------------------------------------------------------

@app.route("/")
def index():
    return render_template_string(HTML_TEMPLATE)


@app.route("/api/search", methods=["POST"])
def api_search():
    data = request.get_json()
    url = data.get("url", "").strip()
    keyword = data.get("keyword", "").strip()
    case_sensitive = data.get("case_sensitive", False)
    skip_ssl = data.get("skip_ssl", True)
    mode = data.get("mode", "content")

    if not url or not keyword:
        return jsonify({"error": "URL and keyword are required."}), 400

    task_id = f"search_{int(time.time()*1000)}"
    task_status[task_id] = {
        "status": "starting",
        "total": 0,
        "done": 0,
        "results": [],
        "current_file": "",
    }

    t = threading.Thread(
        target=content_search_worker,
        args=(task_id, url, keyword, case_sensitive, mode, not skip_ssl),
    )
    t.daemon = True
    t.start()

    return jsonify({"task_id": task_id})


@app.route("/api/download", methods=["POST"])
def api_download():
    data = request.get_json()
    files = data.get("files", [])
    folder = data.get("folder", "downloads")
    skip_ssl = data.get("skip_ssl", True)

    if not files:
        return jsonify({"error": "No files provided."}), 400

    task_id = f"dl_{int(time.time()*1000)}"
    task_status[task_id] = {
        "status": "starting",
        "total": len(files),
        "done": 0,
        "files": [],
        "errors": [],
        "folder": str(DOWNLOAD_DIR / folder),
    }

    t = threading.Thread(
        target=download_worker,
        args=(task_id, files, folder, not skip_ssl),
    )
    t.daemon = True
    t.start()

    return jsonify({"task_id": task_id})


@app.route("/api/status/<task_id>")
def api_status(task_id):
    status = task_status.get(task_id)
    if not status:
        return jsonify({"error": "Unknown task."}), 404
    return jsonify(status)


# ---------------------------------------------------------------------------
# Main
# ---------------------------------------------------------------------------

if __name__ == "__main__":
    print(f"\n📂 Downloads will be saved to: {DOWNLOAD_DIR}")

    # Quick SSL connectivity test
    test_url = "https://www.ieee802.org/11/Reports/wng_update.htm"
    print(f"🔒 Testing connection to {test_url}...")
    try:
        r = requests.get(test_url, verify=False, timeout=10, headers=HEADERS)
        print(f"   ✓ Connection OK (status {r.status_code}, {len(r.text)} bytes)")
    except Exception as e:
        print(f"   ✗ Connection failed: {e}")
        print("   This may be a network issue unrelated to SSL.")

    print(f"🌐 Open http://localhost:5000 in your browser\n")
    app.run(debug=False, port=5000)
