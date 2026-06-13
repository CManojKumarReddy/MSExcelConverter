"""
End-to-end tests for PDF → Excel conversion (main.convert_pdf).

Guards the structured-vs-word-based routing documented in CONVERSION_NOTES.md §4:
  * Bordered PDFs with well-segmented tables use page.extract_tables() (sample-tables).
  * Bank statements whose tables have NO inter-row lines (pdfplumber crams a whole
    page into one newline-stuffed row) must be detected as degenerate and fall back
    to the word-based pipeline that segments transactions correctly (SCB statement).

These don't need Tesseract — they use the PDF text layer.

Run:  cd backend && python -m pytest tests/test_pdf_conversion.py -v
"""


def _all_text(sheets) -> str:
    return " ".join(c for rows in sheets.values() for r in rows for c in r if c).lower()


def _max_cols(sheets) -> int:
    return max(
        (sum(1 for c in r if c.strip()) for rows in sheets.values() for r in rows),
        default=0,
    )


def _max_newlines_in_any_cell(sheets) -> int:
    """A high value means a whole table got crammed into one cell (the bug)."""
    return max(
        (c.count("\n") for rows in sheets.values() for r in rows for c in r),
        default=0,
    )


class TestScbBankStatement:
    """Standard Chartered statement: column lines but NO inter-row lines.

    Regression target: extract_tables() crammed each page's transactions into a
    single newline-stuffed row. Must fall back to the word-based pipeline.
    """

    FIXTURE = "scb_statement.pdf"

    def test_transactions_are_row_segmented(self, convert_pdf_to_sheets):
        sheets = convert_pdf_to_sheets(self.FIXTURE, mode="single")
        rows = next(iter(sheets.values()))
        # Word-based path yields one transaction per row → many rows, not ~3.
        assert len(rows) >= 30, f"only {len(rows)} rows — looks crammed"

    def test_no_crammed_multiline_cells(self, convert_pdf_to_sheets):
        sheets = convert_pdf_to_sheets(self.FIXTURE, mode="single")
        # The degenerate extraction stuffed 10-30 values into one cell; the fix
        # must keep cells essentially single-line.
        assert _max_newlines_in_any_cell(sheets) <= 1, "cells still crammed with newlines"

    def test_has_seven_column_table(self, convert_pdf_to_sheets):
        sheets = convert_pdf_to_sheets(self.FIXTURE, mode="single")
        # Date | Value Date | Description | Cheque | Deposit | Withdrawal | Balance
        assert _max_cols(sheets) >= 6, f"max cols too low ({_max_cols(sheets)})"

    def test_known_values_present(self, convert_pdf_to_sheets):
        sheets = convert_pdf_to_sheets(self.FIXTURE, mode="single")
        text = _all_text(sheets)
        for token in ["balance forward", "114,453.65", "atm withdrawal", "16 jun 19"]:
            assert token in text, f"missing {token!r}"


class TestSampleTablesPdf:
    """Bordered PDF with properly row-segmented tables — must keep using the
    structured extract_tables() path (NOT regress to word-based)."""

    FIXTURE = "sample_tables.pdf"

    def test_many_tables_extracted(self, convert_pdf_to_sheets):
        sheets = convert_pdf_to_sheets(self.FIXTURE, mode="separate")
        # The doc has ~29 tables; structured extraction should yield many sheets.
        assert len(sheets) >= 20, f"only {len(sheets)} sheets — structured path may have regressed"

    def test_cells_not_crammed(self, convert_pdf_to_sheets):
        sheets = convert_pdf_to_sheets(self.FIXTURE, mode="separate")
        assert _max_newlines_in_any_cell(sheets) <= 1

    def test_known_content_present(self, convert_pdf_to_sheets):
        sheets = convert_pdf_to_sheets(self.FIXTURE, mode="separate")
        text = _all_text(sheets)
        # Distinct content from the structured tables.
        hits = sum(t in text for t in ["policy functions", "daniel radcliffe", "data cell"])
        assert hits >= 2, f"expected structured-table content, got hits={hits}"
