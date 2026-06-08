"""
DocToExcel - FastAPI Backend
Converts PDF, DOCX, Images, CSV, TXT to Excel (.xlsx)
"""

import os
import re as _re
import uuid
import logging
import traceback
from pathlib import Path
from typing import Optional

from fastapi import FastAPI, File, Form, UploadFile, HTTPException
from fastapi.middleware.cors import CORSMiddleware
from fastapi.responses import FileResponse, JSONResponse
import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter

# ── Logging ──────────────────────────────────────────────────────────────────
logging.basicConfig(level=logging.INFO, format="%(levelname)s: %(message)s")
log = logging.getLogger(__name__)

# ── App Setup ─────────────────────────────────────────────────────────────────
app = FastAPI(title="DocToExcel API", version="1.0.0")

app.add_middleware(
    CORSMiddleware,
    allow_origins=["http://localhost:5173", "http://127.0.0.1:5173"],
    allow_credentials=True,
    allow_methods=["*"],
    allow_headers=["*"],
)

OUTPUTS_DIR = Path(__file__).parent / "outputs"
OUTPUTS_DIR.mkdir(exist_ok=True)

ALLOWED_EXTENSIONS = {".pdf", ".docx", ".doc", ".png", ".jpg", ".jpeg", ".csv", ".txt"}

# ── Excel Helpers ─────────────────────────────────────────────────────────────

def style_header_row(ws, row_num: int, num_cols: int):
    """Apply teal header styling to a row."""
    header_fill = PatternFill(start_color="009E76", end_color="009E76", fill_type="solid")
    header_font = Font(bold=True, color="FFFFFF", size=11)
    for col in range(1, num_cols + 1):
        cell = ws.cell(row=row_num, column=col)
        cell.fill = header_fill
        cell.font = header_font
        cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)


def apply_table_borders(ws, min_row: int, max_row: int, min_col: int, max_col: int):
    """Apply thin borders to a cell range."""
    thin = Side(style="thin", color="CCCCCC")
    border = Border(left=thin, right=thin, top=thin, bottom=thin)
    for row in ws.iter_rows(min_row=min_row, max_row=max_row, min_col=min_col, max_col=max_col):
        for cell in row:
            cell.border = border


_INVALID_SHEET_CHARS = str.maketrans({c: "" for c in r"\/*?:[]\'"})

def sanitize_sheet_name(name: str) -> str:
    """Remove Excel-invalid characters and truncate to 31 chars."""
    return name.translate(_INVALID_SHEET_CHARS).strip()[:31] or "Sheet"




# ── Stage 4: Merge continuation rows ──────────────────────────────────────────

def _merge_continuation_rows(rows: list[list[str]], col_boundaries: list[int] | None = None) -> list[list[str]]:
    """
    Bank statement descriptions often wrap across multiple visual lines, producing
    extra rows with content only in the Description column.

    A row is a *continuation* of the preceding row when:
    - No anchor column (columns to the left of the widest/description column)
      contains a pure date  →  it is not a new transaction
    - No financial-column cell (columns to the right of description) contains an
      amount  →  it carries no new financial data

    The description column is inferred as the widest column by pixel width.
    """
    if not rows:
        return rows

    if col_boundaries and len(col_boundaries) > 2:
        widths   = [col_boundaries[i + 1] - col_boundaries[i]
                    for i in range(len(col_boundaries) - 1)]
        desc_col = widths.index(max(widths))
    else:
        desc_col = min(1, len(rows[0]) - 1)

    anchor_cols = list(range(desc_col))
    ncols_hint  = max(len(rows[0]), desc_col + 2)

    merged: list[list[str]] = [list(rows[0])]

    for row in rows[1:]:
        # Skip empty rows
        if not any(c.strip() for c in row):
            continue

        # A complete header row is always a new row — never a continuation of
        # whatever came before (e.g. a MICR/address metadata row).
        if _is_header_row(row):
            merged.append(list(row))
            continue

        # Skip stray header-fragment rows — lines from multi-line column headers
        # (e.g. lone "Date" sitting below "Value" in a "Value Date" cell) that
        # contain only header keywords but do NOT form a complete header row.
        row_words = [w for c in row for w in c.strip().split() if c.strip()]
        if row_words and all(w.lower() in _HEADER_KEYWORDS for w in row_words):
            continue

        anchor_has_date = any(
            bool(_PURE_DATE_RE.match(row[i].strip()))
            for i in anchor_cols
            if i < len(row) and row[i].strip()
        )
        has_financial = any(
            _NUMERIC_RE.match(row[i].strip())
            for i in range(desc_col + 1, max(len(row), ncols_hint))
            if i < len(row) and row[i].strip()
        )

        if not anchor_has_date and not has_financial:
            prev = merged[-1]
            while len(prev) <= desc_col:
                prev.append("")
            text = " ".join(c.strip() for c in row if c.strip())
            if text:
                prev[desc_col] = (prev[desc_col] + " " + text).strip() if prev[desc_col] else text
        else:
            merged.append(list(row))

    return merged


# ── Stage 5: Build single-sheet output ────────────────────────────────────────

def _is_header_row(row: list[str]) -> bool:
    """
    Return True if this row is a table column-header line.

    Strict criteria:
    - Every non-empty cell must be ≤ 3 words (real headers are concise labels)
    - Word set must hit both required keyword groups
    """
    non_empty = [c.strip() for c in row if c.strip()]
    if not non_empty:
        return False
    if any(len(c.split()) > 3 for c in non_empty):
        return False        # long cell → address / metadata, not a header
    texts = {w.lower() for c in non_empty for w in c.split()}
    return all(texts & grp for grp in _HEADER_REQUIRED_GROUPS)


def _build_single_sheet_rows(all_sections: list[tuple[str, list[list[str]]]]) -> list[list[str]]:
    """
    Combine all sections into one sheet:

    - For sections that contain a header row: skip everything before the header
      (address text, MICR lines, metadata), emit the header exactly once globally,
      then emit all subsequent data rows.
    - For sections without a header row: include only if they contain financial
      amounts (skips pure-text metadata blocks that contain dates like "16 Jul 2019").
    - Duplicate header rows from subsequent pages are silently dropped.
    """
    combined: list[list[str]] = []
    header_emitted = False

    for _name, rows in all_sections:
        if not rows:
            continue

        header_idx = next((i for i, row in enumerate(rows) if _is_header_row(row)), None)

        if header_idx is not None:
            if not header_emitted:
                header_row = list(rows[header_idx])
                # Absorb any immediately-following header-keyword-only rows into
                # the header.  This handles multi-line column headers in the PDF
                # (e.g. "Value" on line 1, "Date" on line 2 of the same cell).
                start_data = header_idx + 1
                while start_data < len(rows):
                    frag = rows[start_data]
                    frag_words = [w for c in frag for w in c.strip().split() if c.strip()]
                    if frag_words and all(w.lower() in _HEADER_KEYWORDS for w in frag_words):
                        # Merge fragment words into the header cell that covers
                        # the same column position.
                        for ci, cell in enumerate(frag):
                            if cell.strip() and ci < len(header_row):
                                header_row[ci] = (
                                    (header_row[ci] + " " + cell.strip()).strip()
                                    if header_row[ci] else cell.strip()
                                )
                        start_data += 1
                    else:
                        break
                combined.append(header_row)
                header_emitted = True
            else:
                start_data = header_idx + 1
            for row in rows[start_data:]:
                if not _is_header_row(row):
                    # Skip stray header-fragment rows (incomplete keyword-only lines)
                    row_words = [w for c in row for w in c.strip().split() if c.strip()]
                    if (row_words
                            and all(w.lower() in _HEADER_KEYWORDS for w in row_words)
                            and not _is_header_row(row)):
                        continue
                    combined.append(row)
        else:
            has_financial = any(
                _NUMERIC_RE.match(cell.strip())
                for row in rows for cell in row if cell.strip()
            )
            if not has_financial:
                continue
            for row in rows:
                # Skip pure metadata rows (no date, no financial amount)
                row_has_amount = any(
                    _NUMERIC_RE.match(c.strip()) for c in row if c.strip()
                )
                row_has_date = any(
                    _PURE_DATE_RE.match(c.strip()) for c in row if c.strip()
                )
                if row_has_amount or row_has_date:
                    combined.append(row)

    return combined


def auto_fit_columns(ws):
    """Set reasonable column widths based on content."""
    for col in ws.columns:
        max_len = 0
        col_letter = get_column_letter(col[0].column)
        for cell in col:
            if cell.value:
                max_len = max(max_len, len(str(cell.value)))
        adjusted = min(max(max_len + 4, 10), 60)
        ws.column_dimensions[col_letter].width = adjusted


def write_data_to_sheet(ws, rows: list[list], sheet_title: str = "Converted Data"):
    """Write a list-of-lists to a worksheet with nice formatting."""
    ws.title = sheet_title[:31]  # Excel sheet name limit

    if not rows:
        ws["A1"] = "No data found in the document."
        return

    for r_idx, row in enumerate(rows, start=1):
        for c_idx, value in enumerate(row, start=1):
            cell = ws.cell(row=r_idx, column=c_idx, value=value)
            cell.alignment = Alignment(vertical="center", wrap_text=False)

    # Style first row as header if it looks like one
    if len(rows) > 1:
        style_header_row(ws, 1, len(rows[0]))

    apply_table_borders(ws, 1, len(rows), 1, max(len(r) for r in rows))
    auto_fit_columns(ws)

    # Freeze the header row
    ws.freeze_panes = "A2"


def save_workbook(wb: openpyxl.Workbook, stem: str) -> str:
    """Save workbook to outputs dir and return just the filename."""
    unique_id = uuid.uuid4().hex[:8]
    filename = f"{stem}_{unique_id}.xlsx"
    out_path = OUTPUTS_DIR / filename
    wb.save(out_path)
    log.info("Saved: %s", out_path)
    return filename


# ── Converters ────────────────────────────────────────────────────────────────

def convert_csv(content: bytes, stem: str) -> str:
    """CSV → Excel. Detects delimiter automatically."""
    import csv
    import io

    text = content.decode("utf-8-sig", errors="replace")
    # Sniff delimiter
    try:
        dialect = csv.Sniffer().sniff(text[:4096], delimiters=",;\t|")
    except csv.Error:
        dialect = csv.excel

    reader = csv.reader(io.StringIO(text), dialect)
    rows = [row for row in reader if any(cell.strip() for cell in row)]

    wb = openpyxl.Workbook()
    ws = wb.active
    write_data_to_sheet(ws, rows, "CSV Data")
    return save_workbook(wb, stem)


def convert_txt(content: bytes, stem: str) -> str:
    """TXT → Excel. Tries tab-separated first, then splits into lines."""
    import csv
    import io

    text = content.decode("utf-8-sig", errors="replace")
    lines = [l for l in text.splitlines() if l.strip()]

    # Detect if tab-separated
    if lines and "\t" in lines[0]:
        reader = csv.reader(io.StringIO(text), delimiter="\t")
        rows = [row for row in reader if any(cell.strip() for cell in row)]
    else:
        # Put each line as a row with a single column
        rows = [["Line", "Content"]]
        for i, line in enumerate(lines, start=1):
            rows.append([i, line])

    wb = openpyxl.Workbook()
    ws = wb.active
    write_data_to_sheet(ws, rows, "Text Data")
    return save_workbook(wb, stem)


def _setup_tesseract():
    """Ensure pytesseract can find the Tesseract binary."""
    import shutil
    import pytesseract
    if not shutil.which("tesseract"):
        pytesseract.pytesseract.tesseract_cmd = r"C:\Program Files\Tesseract-OCR\tesseract.exe"


# ── PDF Table Extraction Pipeline ────────────────────────────────────────────
#
# Architecture overview:
#   1. GROUP words into visual lines using a y-band approach (not a fixed tolerance)
#   2. SPLIT lines into sections at large vertical gaps
#   3. DETECT header row by keyword density — only accept lines where the majority
#      of words are header keywords (rejects address/MICR lines that happen to end
#      with "Deposit Withdrawal Balance")
#   4. CLUSTER multi-word column headers (e.g. "Value Date") by gap size, then place
#      column boundaries at cluster midpoints — not between every word
#   5. ASSIGN words to columns; pull stray year tokens back to adjacent date column
#   6. MERGE continuation rows (multiline descriptions) into their anchor row
#   7. FILTER metadata sections from single-sheet output

_HEADER_KEYWORDS = {
    "date", "value", "description", "debit", "credit", "balance", "amount",
    "withdrawal", "deposit", "cheque", "chq", "narration", "particulars",
    "transaction", "reference", "remarks",
}
_HEADER_REQUIRED_GROUPS = [
    {"debit", "credit", "amount", "withdrawal", "deposit"},
    {"date", "description", "balance", "narration", "particulars"},
]
# Known multi-word bank column headers — consecutive words matching one of
# these ordered bigrams are kept in the same cluster (single column).
_HEADER_BIGRAMS = {
    ("value", "date"),
    ("debit", "amount"),
    ("credit", "amount"),
    ("cheque", "no"),
    ("cheque", "no."),
    ("cheque", "number"),
    ("chq", "no"),
    ("chq", "no."),
    ("chq", "number"),
    ("reference", "no"),
    ("reference", "no."),
    ("reference", "number"),
    ("transaction", "date"),
    ("posting", "date"),
    ("sl", "no"),
    ("sl", "no."),
    ("s", "no"),
    ("s.", "no"),
}

_NUMERIC_RE      = _re.compile(r"^[\$\-]?\d{1,10}(,\d{3})*\.\d{2}$")
_PURE_DATE_RE    = _re.compile(
    r'^(\d{1,2}[/\-]\d{1,2}[/\-]\d{2,4}'
    r'|\d{1,2}\s+[A-Za-z]{3,9}\.?\s+\d{2,4}'
    r'|[A-Za-z]{3,9}\.?\s+\d{1,2},?\s+\d{2,4}'
    r')$'
)
_PARTIAL_DATE_RE = _re.compile(r'\d{1,2}\s+[A-Za-z]{3,9}\.?\s*$')
_LONE_YEAR_RE    = _re.compile(r'^\d{2}$|^\d{4}$')


# ── Stage 1: Group words into visual lines ─────────────────────────────────────

def _group_into_lines(words: list[dict]) -> list[list[dict]]:
    """
    Group words into visual lines using bounding-box overlap, NOT y-tolerance.

    Algorithm:
    - Sort words by top coordinate.
    - Maintain the bottom edge of the current line (max bottom seen so far).
    - A new word joins the current line when its top edge is BELOW the current
      line's bottom edge minus a small allowance (25 % of avg height). This
      means words must genuinely OVERLAP the current line's vertical extent.
    - Otherwise a new line starts.

    Why this is better than tolerance-based approaches:
    - Two consecutive table rows whose bounding boxes DON'T overlap are always
      placed in separate line-groups, regardless of how close they are.
    - Words on the same visual line that have slightly different top values
      (common in pdfplumber output) are still grouped together because their
      boxes overlap.
    - The MICR line and the table header line are NEVER merged even when they
      are only 1-2 pt apart, because the MICR box ends before the header box
      begins.
    """
    if not words:
        return []

    words = sorted(words, key=lambda w: (w["top"], w["left"]))

    lines:      list[list[dict]] = []
    current     = [words[0]]
    line_bottom = words[0]["top"] + max(words[0]["height"], 1)

    for word in words[1:]:
        avg_h     = sum(w["height"] for w in current) / len(current)
        # Allow a small downward overlap tolerance (25 % of avg height) to
        # handle words on the same line whose tops differ slightly.
        overlap_threshold = line_bottom - avg_h * 0.25

        if word["top"] < overlap_threshold:
            # Word's top is above the current line's adjusted bottom → same line
            current.append(word)
            line_bottom = max(line_bottom, word["top"] + max(word["height"], 1))
        else:
            lines.append(sorted(current, key=lambda w: w["left"]))
            current     = [word]
            line_bottom = word["top"] + max(word["height"], 1)

    lines.append(sorted(current, key=lambda w: w["left"]))
    return lines


# ── Stage 2: Detect column boundaries from the header line ────────────────────

def _find_col_boundaries(line_groups: list[list[dict]], img_width: int) -> list[int] | None:
    """
    Scan line_groups for the table header row and return column-boundary x-positions.

    A genuine header line:
    - contains words from BOTH required keyword groups
    - has at least 3 words
    - has ≥ 50 % of its words as header keywords (rejects MICR/address lines that
      accidentally end with "… Deposit Withdrawal Balance Date")

    Multi-word column headers ("Value Date", "Chq No") are detected by matching
    consecutive word pairs against a known-bigrams list.  Boundaries are placed at
    the midpoint between adjacent clusters, not between individual words.
    """
    for line in line_groups:
        texts = {w["text"].lower() for w in line}
        if not (all(texts & grp for grp in _HEADER_REQUIRED_GROUPS) and len(line) >= 3):
            continue

        kw_count = sum(1 for w in line if w["text"].lower() in _HEADER_KEYWORDS)
        if kw_count < len(line) / 2:
            continue          # mostly non-keyword words → address/metadata line

        header_words = sorted(line, key=lambda w: w["left"])

        # Cluster consecutive words that form known multi-word column headers
        # (e.g. "Value" + "Date" → "Value Date").  Any other consecutive pair
        # starts a new cluster — gap size is NOT used (too unreliable across PDFs).
        clusters: list[list[dict]] = [[header_words[0]]]
        for word in header_words[1:]:
            bigram = (clusters[-1][-1]["text"].lower(), word["text"].lower())
            if bigram in _HEADER_BIGRAMS:
                clusters[-1].append(word)
            else:
                clusters.append([word])

        # Place boundaries at the midpoint between adjacent clusters
        boundaries = [0]          # start from page left edge (not header word edge)
        for i in range(len(clusters) - 1):
            right_cur  = max(w["right"] for w in clusters[i])
            left_next  = clusters[i + 1][0]["left"]
            boundaries.append((right_cur + left_next) // 2)
        boundaries.append(img_width)
        return boundaries

    return None


# ── Stage 3: Assign words to columns ──────────────────────────────────────────

def _line_to_cells_fixed(line: list[dict], col_boundaries: list[int]) -> list[str]:
    """
    Map each word to a column bucket using col_boundaries.

    Special rules:
    - The rightmost column (Balance) only accepts numbers.  Non-numeric words that
      drift right are folded back into the Description column (second-to-last).
    - A lone 2- or 4-digit year (e.g. "19", "2019") that drifts one column to the
      right of a partial date ("16 Jun") is pulled back into the date's column.
    """
    num_cols = len(col_boundaries) - 1
    if num_cols < 1:
        return [" ".join(w["text"] for w in line)]

    cells    = [""] * num_cols
    last     = num_cols - 1
    prev_col: int | None = None

    for word in line:
        # Find which column bucket this word's left edge falls into
        col_idx = 0
        for j in range(1, num_cols):
            if word["left"] >= col_boundaries[j]:
                col_idx = j

        # Rightmost column guard: only numerics and header keywords allowed
        if (col_idx == last and num_cols > 1
                and not _NUMERIC_RE.match(word["text"])
                and word["text"].lower() not in _HEADER_KEYWORDS):
            col_idx = last - 1

        # Lone-year pull-back: "16 Jun" | "19" → "16 Jun 19" | ""
        if (prev_col is not None
                and col_idx != prev_col
                and _LONE_YEAR_RE.match(word["text"])
                and _PARTIAL_DATE_RE.search(cells[prev_col])):
            col_idx = prev_col

        cells[col_idx] = (cells[col_idx] + " " + word["text"]).strip()
        prev_col = col_idx

    return cells


def _line_to_cells_gap(line: list[dict]) -> list[str]:
    """Split a line into cells by detecting large horizontal gaps (fallback)."""
    cells: list[str] = []
    bucket = [line[0]]
    for word in line[1:]:
        prev         = bucket[-1]
        gap          = word["left"] - prev["right"]
        avg_char_w   = prev["width"] / max(len(prev["text"]), 1)
        if gap > avg_char_w * 2.5:
            cells.append(" ".join(w["text"] for w in bucket))
            bucket = [word]
        else:
            bucket.append(word)
    cells.append(" ".join(w["text"] for w in bucket))
    return cells


def _ocr_image_to_sections(img) -> list[tuple[str, list[list[str]]]]:
    """
    OCR an image and split into sections by large vertical gaps between line groups.
    Each section gets its own column boundaries (derived from its header row).
    Returns [(section_title, rows), ...].
    """
    import pytesseract

    data = pytesseract.image_to_data(img, output_type=pytesseract.Output.DICT)

    words = []
    for i in range(len(data["text"])):
        text = data["text"][i].strip()
        if text and int(data["conf"][i]) > 0:
            words.append({
                "text": text,
                "left": data["left"][i],
                "top": data["top"][i],
                "width": data["width"][i],
                "height": data["height"][i],
                "right": data["left"][i] + data["width"][i],
            })

    if not words:
        return []

    line_groups = _group_into_lines(words)
    img_width = getattr(img, "width", 9999)

    # Determine threshold for a "section break" gap (2.5× the median line spacing)
    if len(line_groups) > 1:
        tops = [line[0]["top"] for line in line_groups]
        spacings = sorted([tops[i + 1] - tops[i] for i in range(len(tops) - 1)])
        median_spacing = spacings[len(spacings) // 2]
        gap_threshold = median_spacing * 2.5
    else:
        gap_threshold = 9999

    # Split line_groups into section buckets at large vertical gaps
    section_buckets: list[list[list[dict]]] = [[]]
    for i, line in enumerate(line_groups):
        section_buckets[-1].append(line)
        if i + 1 < len(line_groups):
            gap = line_groups[i + 1][0]["top"] - line[0]["top"]
            if gap > gap_threshold:
                section_buckets.append([])

    # Convert each bucket to rows with per-section column detection
    sections: list[tuple[str, list[list[str]]]] = []
    for bucket in section_buckets:
        if not bucket:
            continue
        col_boundaries = _find_col_boundaries(bucket, img_width)
        rows: list[list[str]] = []
        for line in bucket:
            cells = _line_to_cells_fixed(line, col_boundaries) if col_boundaries else _line_to_cells_gap(line)
            if any(c.strip() for c in cells):
                rows.append(cells)
        if rows:
            rows = _merge_continuation_rows(rows, col_boundaries)
            title = sanitize_sheet_name(" ".join(c for c in rows[0] if c.strip())) or "Section"
            sections.append((title, rows))

    return sections


def _ocr_pdf_pages(content: bytes) -> list[tuple[str, list[list[str]]]]:
    """
    Render each PDF page to an image, OCR it, and split by vertical gaps.
    Returns [(sheet_name, rows), ...] — one entry per section across all pages.
    """
    try:
        import pytesseract
    except ImportError:
        raise HTTPException(
            status_code=500,
            detail="pytesseract not installed. Run: pip install pytesseract  (and install Tesseract OCR)"
        )
    _setup_tesseract()

    try:
        import pdfplumber
    except ImportError:
        raise HTTPException(status_code=500, detail="pdfplumber is not installed.")

    import io

    results: list[tuple[str, list[list[str]]]] = []
    used_names: dict[str, int] = {}

    with pdfplumber.open(io.BytesIO(content)) as pdf:
        for page_num, page in enumerate(pdf.pages, start=1):
            img = page.to_image(resolution=200).original
            for title, rows in _ocr_image_to_sections(img):
                base = f"P{page_num} {title}"[:31]
                n = used_names.get(base, 0)
                used_names[base] = n + 1
                sheet_name = base if n == 0 else f"{base[:28]} {n}"
                results.append((sheet_name[:31], rows))

    return results


def convert_pdf(content: bytes, stem: str, mode: str = "single", password: str = "") -> str:
    """PDF → Excel. mode='separate' puts each table/section on its own sheet;
    mode='single' stacks everything into one sheet.
    Raises HTTP 423 if the PDF is password-protected and no/wrong password given."""
    try:
        import pdfplumber
    except ImportError:
        raise HTTPException(status_code=500, detail="pdfplumber is not installed. Run: pip install pdfplumber")

    import io
    from pdfminer.pdfdocument import PDFPasswordIncorrect
    from pdfminer.pdfparser import PDFSyntaxError

    wb = openpyxl.Workbook()
    wb.remove(wb.active)

    # Collect all (sheet_name, rows) sections first
    all_sections: list[tuple[str, list[list[str]]]] = []

    open_kwargs = {"password": password} if password else {}
    try:
        pdf_file = pdfplumber.open(io.BytesIO(content), **open_kwargs)
    except (PDFPasswordIncorrect, PDFSyntaxError, Exception) as e:
        err = str(e).lower()
        if "password" in err or "encrypt" in err or "incorrect" in err or isinstance(e, PDFPasswordIncorrect):
            raise HTTPException(status_code=423, detail="password_required")
        raise

    # ── Two-pass word-based extraction ────────────────────────────────────────
    #
    # Pass 1 — collect every page's words and line_groups; find the BEST column
    #          boundaries across the entire document (the cleanest header row may
    #          be on page 2 even though we need it for page 1).
    #
    # Pass 2 — convert each page's line_groups into rows, using per-section
    #          boundaries when available, falling back to the document-level
    #          best boundaries otherwise.  This solves the MICR-merges-with-header
    #          problem: page 1's header may be obscured but page 2's is clean.

    def _normalise_words(raw_words):
        return [
            {
                "text":   w["text"].strip(),
                "left":   int(w["x0"]),
                "top":    int(w["top"]),
                "width":  int(w["x1"] - w["x0"]),
                "height": int(w["bottom"] - w["top"]),
                "right":  int(w["x1"]),
            }
            for w in raw_words if w["text"].strip()
        ]

    def _section_buckets(line_groups):
        """Split line_groups into section buckets at large vertical gaps."""
        if len(line_groups) > 1:
            tops = [ln[0]["top"] for ln in line_groups]
            spacings = sorted([tops[i + 1] - tops[i] for i in range(len(tops) - 1)])
            gap_threshold = spacings[len(spacings) // 2] * 2.5
        else:
            gap_threshold = 9999
        buckets: list[list[list[dict]]] = [[]]
        for i, line in enumerate(line_groups):
            buckets[-1].append(line)
            if i + 1 < len(line_groups):
                if line_groups[i + 1][0]["top"] - line[0]["top"] > gap_threshold:
                    buckets.append([])
        return buckets

    def _lines_to_rows(line_groups, col_boundaries):
        """Convert a list of line_groups to rows using col_boundaries."""
        rows: list[list[str]] = []
        for line in line_groups:
            if len(line) <= 2:
                merged = " ".join(w["text"] for w in line).strip()
                if merged:
                    rows.append([merged])
            else:
                cells = (
                    _line_to_cells_fixed(line, col_boundaries)
                    if col_boundaries else _line_to_cells_gap(line)
                )
                if any(c.strip() for c in cells):
                    rows.append(cells)
        return rows

    # ── Single pass: read words + detect boundaries + build sections ─────────
    # We need one pass because pdfplumber closes the PDF on context exit.
    # Strategy:
    #   1. Extract all page word data.
    #   2. Find the best (most-column) boundaries from any page.
    #   3. Process every section bucket, using per-section boundaries when
    #      available, falling back to the document-best when not (e.g. page 1
    #      where MICR text may prevent local header detection).

    best_boundaries: list[int] | None = None
    page_data: list[tuple[int, int, list]] = []  # (page_num, img_width, line_groups)

    with pdf_file as pdf:
        for page_num, page in enumerate(pdf.pages, start=1):
            raw = page.extract_words(keep_blank_chars=False, use_text_flow=False)
            if not raw:
                page_data.append((page_num, int(page.width), []))
                continue
            words       = _normalise_words(raw)
            img_width   = int(page.width)
            line_groups = _group_into_lines(words)
            page_data.append((page_num, img_width, line_groups))

            # Track the cleanest header found so far (most columns wins)
            boundaries = _find_col_boundaries(line_groups, img_width)
            if boundaries and (
                best_boundaries is None or len(boundaries) > len(best_boundaries)
            ):
                best_boundaries = boundaries

    # Build sections now that best_boundaries is known
    for page_num, img_width, line_groups in page_data:
        if not line_groups:
            continue
        for bucket in _section_buckets(line_groups):
            if not bucket:
                continue
            # Per-section boundaries preferred; fall back to document-best
            col_boundaries = _find_col_boundaries(bucket, img_width) or best_boundaries
            rows = _lines_to_rows(bucket, col_boundaries)
            if rows:
                rows = _merge_continuation_rows(rows, col_boundaries)
                title = sanitize_sheet_name(
                    " ".join(c for c in rows[0] if c.strip()) or f"Page {page_num}"
                )
                all_sections.append((title, rows))

    if not all_sections:
        log.info("No text layer found in PDF; attempting OCR fallback.")
        all_sections = _ocr_pdf_pages(content)

    if not all_sections:
        ws = wb.create_sheet(title="Empty PDF")
        ws["A1"] = "No extractable content found in this PDF."
        return save_workbook(wb, stem)

    if mode == "single":
        combined_rows = _build_single_sheet_rows(all_sections)
        ws = wb.create_sheet(title="All Data")
        write_data_to_sheet(ws, combined_rows, "All Data")
    else:
        for sheet_name, rows in all_sections:
            safe = sanitize_sheet_name(sheet_name)
            ws = wb.create_sheet(title=safe)
            write_data_to_sheet(ws, rows, safe)

    return save_workbook(wb, stem)


def convert_docx(content: bytes, stem: str) -> str:
    """DOCX → Excel. Extracts all tables; falls back to paragraph text."""
    try:
        from docx import Document
    except ImportError:
        raise HTTPException(status_code=500, detail="python-docx is not installed. Run: pip install python-docx")

    import io

    doc = Document(io.BytesIO(content))
    wb = openpyxl.Workbook()
    wb.remove(wb.active)
    sheets_created = 0

    # Extract tables
    for t_idx, table in enumerate(doc.tables, start=1):
        rows = []
        for row in table.rows:
            cells = [cell.text.strip() for cell in row.cells]
            if any(cells):
                rows.append(cells)
        if rows:
            sheet_name = f"Table {t_idx}"
            ws = wb.create_sheet(title=sheet_name)
            write_data_to_sheet(ws, rows, sheet_name)
            sheets_created += 1

    # If no tables, extract paragraphs
    if sheets_created == 0:
        paragraphs = [p.text.strip() for p in doc.paragraphs if p.text.strip()]
        if paragraphs:
            rows = [["#", "Paragraph"]] + [[i, p] for i, p in enumerate(paragraphs, 1)]
            ws = wb.create_sheet(title="Document Text")
            write_data_to_sheet(ws, rows, "Document Text")
            sheets_created += 1

    if sheets_created == 0:
        ws = wb.create_sheet(title="Empty Document")
        ws["A1"] = "No content found in this document."

    return save_workbook(wb, stem)


def _detect_columns_from_gaps(
    line_groups: list[list[dict]], img_width: int
) -> list[int] | None:
    """
    When no header row is detected, infer column boundaries from the distribution
    of large inter-word gaps across all lines.  Gaps that appear consistently in
    the same x-range across multiple lines mark column separators.
    """
    gap_midpoints: list[float] = []

    for line in line_groups:
        if len(line) < 2:
            continue
        words = sorted(line, key=lambda w: w["left"])
        gaps = [
            (words[i + 1]["left"] - words[i]["right"],
             (words[i]["right"] + words[i + 1]["left"]) / 2)
            for i in range(len(words) - 1)
        ]
        sizes = [g for g, _ in gaps]
        if not sizes:
            continue
        median_gap = sorted(sizes)[len(sizes) // 2]
        threshold = max(median_gap * 1.5, 15)
        for gap, mid in gaps:
            if gap >= threshold:
                gap_midpoints.append(mid)

    if not gap_midpoints:
        return None

    gap_midpoints.sort()
    clusters: list[list[float]] = [[gap_midpoints[0]]]
    for mid in gap_midpoints[1:]:
        if mid - clusters[-1][-1] < 40:
            clusters[-1].append(mid)
        else:
            clusters.append([mid])

    # Only trust boundaries that appear in at least 25 % of lines
    min_support = max(2, len(line_groups) // 4)
    boundaries = [0]
    for cluster in clusters:
        if len(cluster) >= min_support:
            boundaries.append(int(sum(cluster) / len(cluster)))
    boundaries.append(img_width)

    return boundaries if len(boundaries) > 2 else None


def _merge_repeating_cols(rows: list[list[str]]) -> list[list[str]]:
    """
    When an image table has repeated column groups (e.g. Name|Phone|Name|Phone|Name|Phone),
    normalize to the base group by stacking each repetition as new rows.

    Detection: find the smallest period P (≥ 2) that divides the column count evenly
    and has at least 2 repetitions.  The period with the most non-empty data wins.
    """
    if not rows:
        return rows

    ncols = max(len(r) for r in rows)
    if ncols < 4:
        return rows

    best_period = None
    best_score  = -1

    for period in range(2, ncols // 2 + 1):
        if ncols % period != 0:
            continue
        reps = ncols // period
        # Score = total non-empty cells when stacked (reward smaller period)
        score = 0
        for row in rows:
            padded = (row + [""] * ncols)[:ncols]
            for r in range(reps):
                chunk = padded[r * period:(r + 1) * period]
                score += sum(1 for c in chunk if c.strip())
        # Prefer the smallest period (penalise larger periods slightly)
        adjusted = score - period * 0.1
        if adjusted > best_score:
            best_score  = adjusted
            best_period = period

    if best_period is None or best_period == ncols:
        return rows

    reps    = ncols // best_period
    stacked = []
    for row in rows:
        padded = (row + [""] * ncols)[:ncols]
        for r in range(reps):
            chunk = padded[r * best_period:(r + 1) * best_period]
            if any(c.strip() for c in chunk):
                stacked.append(chunk)
    return stacked


def convert_image(content: bytes, stem: str, filename: str, merge_cols: bool = False) -> str:
    """Image (PNG/JPG) → Excel via OCR (pytesseract) with column detection."""
    try:
        import pytesseract
        from PIL import Image
        from pytesseract import Output
    except ImportError:
        raise HTTPException(
            status_code=500,
            detail=(
                "pytesseract or Pillow is not installed. "
                "Run: pip install pytesseract Pillow  "
                "Also install Tesseract OCR engine from https://github.com/UB-Mannheim/tesseract/wiki"
            ),
        )

    import io, os, shutil

    # Auto-locate Tesseract on Windows when it is not in the process PATH
    if not shutil.which("tesseract"):
        _win_paths = [
            r"C:\Program Files\Tesseract-OCR\tesseract.exe",
            r"C:\Program Files (x86)\Tesseract-OCR\tesseract.exe",
            os.path.expandvars(r"%LOCALAPPDATA%\Programs\Tesseract-OCR\tesseract.exe"),
        ]
        for _p in _win_paths:
            if os.path.isfile(_p):
                pytesseract.pytesseract.tesseract_cmd = _p
                break

    try:
        img = Image.open(io.BytesIO(content))
        if img.mode not in ("RGB", "L"):
            img = img.convert("RGB")

        data = pytesseract.image_to_data(img, output_type=Output.DICT)
    except Exception as e:
        raise HTTPException(status_code=500, detail=f"OCR failed: {str(e)}")

    # ── Build word list with y-band line grouping ────────────────────────────
    # For photos, pytesseract's block/line numbering can misalign words on the
    # same visual row.  Instead, cluster words by y-centre within ±50% of the
    # average character height.
    from collections import defaultdict

    all_words: list[dict] = []
    for i in range(len(data["text"])):
        text = data["text"][i].strip()
        conf = int(data["conf"][i])
        if not text or conf <= 20:
            continue
        w = data["width"][i]
        h = data["height"][i]
        all_words.append({
            "text":   text,
            "left":   data["left"][i],
            "top":    data["top"][i],
            "right":  data["left"][i] + w,
            "width":  w,
            "height": max(h, 1),
        })

    wb = openpyxl.Workbook()
    ws = wb.active

    if not all_words:
        write_data_to_sheet(ws, [["Result"], ["No text could be extracted via OCR."]], "OCR Result")
        return save_workbook(wb, stem)

    # Y-band clustering: group words whose vertical centres are within
    # tolerance of the current running line centre.
    avg_h   = sum(w["height"] for w in all_words) / len(all_words)
    y_tol   = avg_h * 0.6
    sorted_w = sorted(all_words, key=lambda w: (w["top"] + w["height"] / 2, w["left"]))

    line_groups: list[list[dict]] = []
    cur_line:    list[dict]       = [sorted_w[0]]
    cur_cy = sorted_w[0]["top"] + sorted_w[0]["height"] / 2

    for wd in sorted_w[1:]:
        wd_cy = wd["top"] + wd["height"] / 2
        if abs(wd_cy - cur_cy) <= y_tol:
            cur_line.append(wd)
            cur_cy = sum(w["top"] + w["height"] / 2 for w in cur_line) / len(cur_line)
        else:
            line_groups.append(sorted(cur_line, key=lambda w: w["left"]))
            cur_line = [wd]
            cur_cy   = wd_cy
    line_groups.append(sorted(cur_line, key=lambda w: w["left"]))

    img_width = img.width

    # ── Column detection ─────────────────────────────────────────────────────
    # 1. Header-keyword based (bank statements etc.)
    col_boundaries = _find_col_boundaries(line_groups, img_width)

    # 2. Zone histogram: divide width into 50 coarse zones.
    #    Only runs of ≥ 2 consecutive empty zones count as column separators
    #    (~img_width/25 px minimum gap).  This avoids splitting on the small
    #    spaces between individual words within a column.
    if col_boundaries is None:
        n_zones  = 50
        zone_w   = max(1.0, img_width / n_zones)
        zone_cov = [False] * n_zones
        for wd in all_words:
            z0 = int(wd["left"]  / zone_w)
            z1 = int(wd["right"] / zone_w)
            for z in range(max(0, z0), min(n_zones, z1 + 1)):
                zone_cov[z] = True

        min_gap_zones = 2          # at least 2 empty zones = real column gap
        histo_bounds  = [0]
        z = 0
        while z < n_zones:
            if not zone_cov[z]:
                gap_start = z
                while z < n_zones and not zone_cov[z]:
                    z += 1
                gap_end = z
                if gap_end - gap_start >= min_gap_zones:
                    mid_px = int(((gap_start + gap_end) / 2) * zone_w)
                    histo_bounds.append(mid_px)
            else:
                z += 1
        histo_bounds.append(img_width)
        if len(histo_bounds) > 2:
            col_boundaries = histo_bounds

    # 3. Gap-distribution fallback
    if col_boundaries is None:
        col_boundaries = _detect_columns_from_gaps(line_groups, img_width)

    # ── Render rows ──────────────────────────────────────────────────────────
    if col_boundaries and len(col_boundaries) > 2:
        rows = []
        for line in line_groups:
            cells = _line_to_cells_fixed(line, col_boundaries)
            if any(c.strip() for c in cells):
                rows.append(cells)
        if not rows:
            rows = [["Result"], ["No structured content detected."]]
    else:
        rows = [["Line #", "Extracted Text"]]
        for i, line in enumerate(line_groups, start=1):
            text = " ".join(wd["text"] for wd in line)
            rows.append([i, text])

    if merge_cols:
        rows = _merge_repeating_cols(rows)

    write_data_to_sheet(ws, rows, "OCR Result")
    return save_workbook(wb, stem)


# ── Routes ────────────────────────────────────────────────────────────────────

@app.get("/")
async def root():
    return {"message": "DocToExcel API is running. POST /api/convert to convert files."}


@app.post("/api/convert")
async def convert_file(file: UploadFile = File(...), mode: str = Form("single"), password: str = Form(""), merge_cols: str = Form("false")):
    """
    Accept an uploaded file and convert it to Excel.
    mode: 'separate' (default) = one sheet per table/section; 'single' = all in one sheet.
    Returns JSON: { output_filename, message }
    """
    original_name = file.filename or "document"
    path = Path(original_name)
    ext = path.suffix.lower()
    stem = path.stem[:40]  # cap length for filename

    log.info("Received file: %s (ext=%s, content_type=%s)", original_name, ext, file.content_type)

    if ext not in ALLOWED_EXTENSIONS:
        raise HTTPException(
            status_code=400,
            detail=f'File type "{ext}" is not supported. Please upload PDF, DOCX, PNG, JPG, CSV, or TXT.',
        )

    content = await file.read()
    if len(content) == 0:
        raise HTTPException(status_code=400, detail="The uploaded file is empty.")

    try:
        if ext == ".csv":
            output_filename = convert_csv(content, stem)
            msg = f'"{original_name}" has been converted successfully! Your Excel file is ready.'
        elif ext == ".txt":
            output_filename = convert_txt(content, stem)
            msg = f'"{original_name}" has been converted successfully! Your Excel file is ready.'
        elif ext == ".pdf":
            output_filename = convert_pdf(content, stem, mode=mode, password=password)
            if mode == "single":
                msg = f'"{original_name}" converted! All tables are combined in one sheet.'
            else:
                msg = f'"{original_name}" converted! Each table/section is on a separate sheet.'
        elif ext in (".docx", ".doc"):
            output_filename = convert_docx(content, stem)
            msg = f'"{original_name}" converted! Tables and text have been extracted to Excel.'
        elif ext in (".png", ".jpg", ".jpeg"):
            do_merge = merge_cols.lower() == "true"
            output_filename = convert_image(content, stem, original_name, merge_cols=do_merge)
            msg = f'"{original_name}" processed via OCR. Extracted text has been placed in the Excel file.'
        else:
            raise HTTPException(status_code=400, detail=f"Unsupported file type: {ext}")

    except HTTPException:
        raise
    except Exception as e:
        log.error("Conversion error:\n%s", traceback.format_exc())
        raise HTTPException(status_code=500, detail=f"Conversion failed: {str(e)}")

    return JSONResponse({
        "output_filename": output_filename,
        "message": msg,
    })


@app.get("/api/download/{filename}")
async def download_file(filename: str):
    """Serve a previously converted Excel file."""
    # Security: ensure no path traversal
    safe_name = Path(filename).name
    if safe_name != filename:
        raise HTTPException(status_code=400, detail="Invalid filename.")

    file_path = OUTPUTS_DIR / safe_name
    if not file_path.exists():
        raise HTTPException(status_code=404, detail=f'File "{safe_name}" not found.')

    return FileResponse(
        path=file_path,
        filename=safe_name,
        media_type="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
    )


@app.get("/api/files")
async def list_files():
    """List all converted files in the outputs directory."""
    files = []
    for f in sorted(OUTPUTS_DIR.glob("*.xlsx"), key=lambda x: x.stat().st_mtime, reverse=True):
        files.append({
            "filename": f.name,
            "size_kb": round(f.stat().st_size / 1024, 1),
        })
    return {"files": files}
