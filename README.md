MetaCleaner — Windows Metadata Remover (GUI)

Overview

- Drag‑and‑drop or add files/folders, preview what will be removed, and clean in place with optional backups.
- Ships as a single portable EXE for Windows 10/11 x64 — no installs required.
- Bundles ExifTool, pikepdf and pywin32 for broad, safe coverage.

Download

- Go to GitHub Releases for this repo and download `MetaCLeaner_Portable.exe`.
- Windows SmartScreen may warn on first run (unsigned); click “More info” → “Run anyway”.

Quick Start

- Add Files or Add Folder (duplicates ignored). You can also drag‑and‑drop.
- Click Scan to see Type and Will Clean columns.
- Keep “Backup originals (.bak)” enabled for safety.
- Click Clean. Status shows “Cleaned (verified)”, “Cleaned”, “No metadata”, or “Unsupported”.
- Logs appear in the bottom pane and are saved per‑session to `logs/MetaCleaner_YYYYMMDD_HHMMSS.log` next to the EXE.

What gets cleaned

- JPEG: EXIF, XMP (APP1), IPTC (APP13) — no recompression.
- PNG: tEXt/iTXt/zTXt/tIME chunks.
- GIF: comment extensions.
- PDF: Info dictionary + XMP metadata.
- Office (DOCX/XLSX/PPTX): `docProps/*` (core/app/custom) + `docProps/thumbnail.*`, with references cleaned.
- Legacy Office (DOC/XLS/PPT): OLE property sets (SummaryInformation, DocumentSummaryInformation) removed in place.
- RTF: removes the entire `{\info ...}` group (author/company/title/etc.).
- Mislabelled Office files: `.doc/.xls/.ppt` that are actually OOXML/RTF/Word2003XML are detected and cleaned appropriately without changing the extension.

Verification (content‑only hashing)

- Enable “Verify content unchanged” to compare only the content portions, not metadata bytes:
  - JPEG: compressed scan data
  - PNG: concatenated `IDAT`
  - GIF: all data excluding comments
  - PDF: page content streams
  - DOCX/XLSX/PPTX: ZIP content parts (excludes docProps/ and rels)
  - DOC/XLS/PPT: core OLE streams (WordDocument + 0/1Table, Workbook/Book, PowerPoint Document)
  - RTF: document with `{\info ...}` stripped
- Tolerant verification: if sensitive tags detected by ExifTool (Author/Title/Company/…) were present before and are gone after, or OLE property sets were removed, the app accepts the result as “Cleaned (verified)” even if ancillary bytes changed.

Backups

- If enabled, a `.bak` (or `.bak.N`) copy of the original is created next to the file before replacement.
- For ExifTool paths, the app creates `.bak` itself and removes it if nothing changed.

Logs

- Every session writes a timestamped log to `logs/` next to the EXE as well as to the on‑screen Log pane.

Notes & limits

- Encrypted/password‑protected PDFs and certain proprietary formats may be “Unsupported”.
- Media containers (MP3/MP4/etc.) are handled via ExifTool; content‑only verification is not yet defined for all.
- The EXE is not code‑signed; SmartScreen may show a one‑time warning.

Build from source (optional)

- Python 3.12 on Windows:
  - `python -m venv .venv`
  - `.\.venv\Scripts\python -m pip install -U pip`
  - `.\.venv\Scripts\pip install -r requirements.txt`
  - Place `vendor\exiftool\ExifTool.exe` (optional; CI fetches it automatically)
  - `pyinstaller --noconsole --onefile --name MetaCLeaner_Portable --add-binary "vendor\exiftool\ExifTool.exe;exiftool" meta_cleaner.pyw`

Privacy & safety

- Cleaning is done to a temporary file and then atomically replaced.
- Backups contain the original metadata; keep them private or delete after validation.
