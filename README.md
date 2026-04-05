# Document Scraper

A local web tool for searching and downloading documents from web pages by keyword. Built for scraping IEEE 802.11 WNG update pages and similar static HTML sites with document links.

## Setup

```bash
# Requires Python 3.8+
pip install -r requirements.txt
python scraper.py
```

Then open **http://localhost:5000** in your browser.

## How it works

1. Enter a page URL (e.g. `https://www.ieee802.org/11/Reports/wng_update.htm`)
2. Enter a keyword to filter by
3. Click **Search** — the tool fetches the page, finds all downloadable file links, and filters by keyword match in the link text or filename
4. Select which files to download and click **Download selected**
5. Files are saved to `~/ieee_downloads/<keyword>/`

## Supported file types

PDF, DOC, DOCX, PPT, PPTX, XLS, XLSX, CSV, TXT, ZIP, GZ, TAR, RTF, ODT, ODP.

## Limitations

- Only scrapes a single page (no recursive crawling). If documents are spread across multiple pages, run searches on each URL.
- Keyword matching is against link text and filenames only — it does not search inside document content.
- The target site must serve static HTML. JavaScript-rendered content won't be found.
