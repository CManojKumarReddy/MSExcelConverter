"""
End-to-end tests for image → Excel conversion (main.convert_image).

These guard the hard-won behaviour documented in backend/CONVERSION_NOTES.md:
  * the best-of-N strategy selector picks a sane candidate, and
  * known-tricky images (dark-background two-table, coloured-header bank statement)
    come out as real multi-column tables, not OCR garbage.

Assertions are deliberately tolerant of minor OCR noise: we check for a MAJORITY
of expected tokens and for structural properties (multi-column, row count, chosen
strategy) rather than exact text — Tesseract output varies slightly across versions.

Run:  cd backend && venv\\Scripts\\python -m pytest -v
"""
from conftest import requires_tesseract, requires_rapidocr, all_text, max_nonempty_cols


def _tokens_present(text: str, tokens: list[str]) -> int:
    """Count how many of the expected tokens appear in the flattened text."""
    return sum(1 for t in tokens if t.lower() in text)


@requires_tesseract
class TestTwoTableDarkImage:
    """Dark-mode screenshot, light text, two stacked tables (Employee / Project).

    Regression target: this used to come out as 16 columns of OCR garbage because
    the grid detector mis-fired on the dark background and returned early.
    """

    FIXTURE = "two_table_dark.png"

    def test_does_not_pick_grid_garbage(self, convert_image_to_rows, caplog):
        rows, chosen = convert_image_to_rows(self.FIXTURE, caplog)
        # The bogus 16-column grid candidate must lose the scoring.
        assert chosen != "grid", f"grid garbage was chosen (rows={rows[:2]})"

    def test_extracts_employee_data(self, convert_image_to_rows, caplog):
        rows, _ = convert_image_to_rows(self.FIXTURE, caplog)
        text = all_text(rows)
        expected = ["emp-101", "sarah", "jenkins", "engineering",
                    "95,000", "david", "marketing", "elena", "sales"]
        hits = _tokens_present(text, expected)
        assert hits >= len(expected) * 0.7, f"only {hits}/{len(expected)} tokens in: {text[:300]}"

    def test_extracts_project_data(self, convert_image_to_rows, caplog):
        rows, _ = convert_image_to_rows(self.FIXTURE, caplog)
        text = all_text(rows)
        expected = ["project", "alpha", "design", "completed", "beta", "gamma", "planning"]
        hits = _tokens_present(text, expected)
        assert hits >= len(expected) * 0.7, f"only {hits}/{len(expected)} tokens in: {text[:300]}"

    def test_is_multi_column(self, convert_image_to_rows, caplog):
        rows, _ = convert_image_to_rows(self.FIXTURE, caplog)
        # A real table split — not everything crammed into one cell.
        assert max_nonempty_cols(rows) >= 4, f"max cols too low; rows={rows[:3]}"


@requires_tesseract
class TestFirstBankImage:
    """Coloured (purple) header band, white labels, key-value metadata + table."""

    FIXTURE = "sample2.png"

    def test_chooses_header_band_strategy(self, convert_image_to_rows, caplog):
        rows, chosen = convert_image_to_rows(self.FIXTURE, caplog)
        assert chosen == "header_band", f"expected header_band, got {chosen!r}"

    def test_has_five_column_table(self, convert_image_to_rows, caplog):
        rows, _ = convert_image_to_rows(self.FIXTURE, caplog)
        # DATE | DESCRIPTION | WITHDRAWAL | DEPOSIT | BALANCE
        assert max_nonempty_cols(rows) >= 5, f"max cols too low; rows={rows[:8]}"

    def test_extracts_transaction_rows(self, convert_image_to_rows, caplog):
        rows, _ = convert_image_to_rows(self.FIXTURE, caplog)
        text = all_text(rows)
        expected = ["internet", "electric", "payroll", "balance",
                    "27,508", "253.68", "deposit", "withdrawal"]
        hits = _tokens_present(text, expected)
        assert hits >= len(expected) * 0.6, f"only {hits}/{len(expected)} tokens in: {text[:400]}"

    def test_metadata_present(self, convert_image_to_rows, caplog):
        rows, _ = convert_image_to_rows(self.FIXTURE, caplog)
        text = all_text(rows)
        # Key-value header section above the table is preserved as text rows.
        assert "first bank" in text or "account summary" in text, f"metadata missing: {text[:200]}"


@requires_tesseract
class TestAccountTransactionsImage:
    """Light-background bank statement (Date|Description|Debit|Credit|Balance).

    Regression target: a `conf > 20` OCR gate silently dropped valid description
    cells that Tesseract emitted with conf 0-15 (e.g. '10/08 POS PURCHASE',
    '11/02 CHECK 1249', '11/09 SERVICE CHARGE') — leaving rows with no description.
    """

    FIXTURE = "account_transactions.png"

    def test_no_rows_missing_description(self, convert_image_to_rows, caplog):
        rows, _ = convert_image_to_rows(self.FIXTURE, caplog)
        # Every row that has a date (col 0) and amounts must also have a non-empty
        # description (col 1).  These are the rows that previously came out blank.
        import re
        date_re = re.compile(r"^\d{2}/\d{2}$")
        offenders = []
        for r in rows:
            if len(r) >= 2 and date_re.match(r[0].strip()):
                if not r[1].strip():
                    offenders.append(r)
        assert not offenders, f"rows with a date but no description: {offenders}"

    def test_recovered_descriptions_present(self, convert_image_to_rows, caplog):
        rows, _ = convert_image_to_rows(self.FIXTURE, caplog)
        text = all_text(rows)
        # The specific descriptions that the conf gate used to drop.
        for token in ["check 1249", "service charge"]:
            assert token in text, f"missing recovered description {token!r}"
        assert text.count("pos purchase") >= 1  # 10/08 row was previously blank

    def test_five_columns(self, convert_image_to_rows, caplog):
        rows, _ = convert_image_to_rows(self.FIXTURE, caplog)
        # Date | Description | Debit | Credit | Balance
        assert max_nonempty_cols(rows) >= 4


@requires_tesseract
class TestSbiRuledStatement:
    """SBI statement screenshot: key-value metadata on top + a bordered table with
    FAINT grey rules and multi-line cells (Date / Narration / Ref / Debit / Credit
    / Balance).  Otsu can't see the rules; a fixed-threshold vertical-line pass
    (the `ruled_columns` strategy) recovers the 6-column structure.
    """

    FIXTURE = "sbi_statement.png"

    def test_chooses_ruled_columns(self, convert_image_to_rows, caplog):
        rows, chosen = convert_image_to_rows(self.FIXTURE, caplog)
        assert chosen == "ruled_columns", f"expected ruled_columns, got {chosen!r}"

    def test_six_column_structure(self, convert_image_to_rows, caplog):
        rows, _ = convert_image_to_rows(self.FIXTURE, caplog)
        assert max_nonempty_cols(rows) >= 5, f"too few columns; rows={rows[:3]}"

    def test_credit_column_separated(self, convert_image_to_rows, caplog):
        """The lone Credit value (80.00) must not be merged with Balance."""
        rows, _ = convert_image_to_rows(self.FIXTURE, caplog)
        # Find the CREDIT INTEREST row; 80.00 and the balance must be in separate cells.
        credit_row = next((r for r in rows if any("CREDIT INTEREST" in c for c in r)), None)
        assert credit_row is not None, "CREDIT INTEREST row not found"
        cells_with_8000 = [c for c in credit_row if "80.00" in c]
        assert cells_with_8000 == ["80.00"], f"80.00 not isolated: {credit_row}"

    def test_amounts_and_dates_present(self, convert_image_to_rows, caplog):
        rows, _ = convert_image_to_rows(self.FIXTURE, caplog)
        text = all_text(rows)
        for token in ["05-dec-19", "21,700.00", "19,088.46", "withdrawal", "80.00"]:
            assert token in text, f"missing {token!r}"

    def test_header_row_recovered(self, convert_image_to_rows, caplog):
        """The table header (skipped by full-page OCR) is recovered via the
        header-band pass and placed above the transactions."""
        rows, _ = convert_image_to_rows(self.FIXTURE, caplog)
        header_row = next(
            (r for r in rows if any("narration" in c.lower() for c in r)), None
        )
        assert header_row is not None, "header row with 'Narration' not found"
        joined = " ".join(header_row).lower()
        # Header spans the key columns.
        for kw in ["date", "narration", "debit", "credit", "balance"]:
            assert kw in joined, f"header missing {kw!r}: {header_row}"


@requires_tesseract
class TestContactsPhoto:
    """Camera PHOTO of a screen (not a screenshot): a contacts list with repeating
    Name/Phone column-pairs.  Low-contrast + screen noise made the local adaptive
    threshold output pure OCR garbage; a denoise + Otsu fallback (triggered by low
    mean confidence) recovers most names and numbers.
    """

    FIXTURE = "contacts_photo.jpg"

    def test_recovers_real_data_not_garbage(self, convert_image_to_rows, caplog):
        rows, _ = convert_image_to_rows(self.FIXTURE, caplog)
        text = all_text(rows)
        # A spread of names + phone numbers that the denoise fallback recovers.
        names = ["kumar", "naveen", "darmaraju", "mahesh", "praven", "venu"]
        phones = ["7902550932", "9353763052", "9884144696", "9662373609"]
        name_hits = sum(1 for n in names if n in text)
        phone_hits = sum(1 for p in phones if p in text)
        assert name_hits >= 4, f"only {name_hits}/6 names recovered: {text[:300]}"
        assert phone_hits >= 3, f"only {phone_hits}/4 phone numbers recovered"

    def test_has_phone_number_column(self, convert_image_to_rows, caplog):
        rows, _ = convert_image_to_rows(self.FIXTURE, caplog)
        import re
        # At least several cells should be clean 10-digit phone numbers.
        ten_digit = sum(
            1 for r in rows for c in r if re.fullmatch(r"\d{10}", c.strip())
        )
        assert ten_digit >= 8, f"only {ten_digit} clean 10-digit numbers"

    @requires_rapidocr
    def test_rapidocr_recovers_full_list(self, convert_image_to_rows, caplog):
        """With the local neural OCR (RapidOCR) installed, the low-confidence photo
        path uses it and recovers far more than Tesseract — including the top-left
        block Tesseract dropped (viswa/jagadeesh/anjayya) and most phone numbers."""
        rows, _ = convert_image_to_rows(self.FIXTURE, caplog)
        text = all_text(rows)
        topleft = ["viswa", "jagadeesh", "anjayya", "suryanarayana", "srinivs", "sunil"]
        hits = sum(1 for n in topleft if n in text)
        assert hits >= 4, f"top-left names not recovered ({hits}/6): {text[:300]}"
        import re
        ten_digit = sum(1 for r in rows for c in r if re.fullmatch(r"\d{10}", c.strip()))
        assert ten_digit >= 30, f"expected many phone numbers via RapidOCR, got {ten_digit}"

    @requires_rapidocr
    def test_excel_chrome_stripped(self, convert_image_to_rows, caplog):
        """The source is a photo of an Excel window, so OCR also captures the app's
        chrome: the Name Box ('A26'), the column-letter band ('B'/'C D E'/'F'/'G'),
        and the row-number gutter ('2 viswa').  The end-of-pipeline validation must
        drop all of it so phone numbers aren't shoved down a row by a bogus header."""
        import re
        rows, _ = convert_image_to_rows(self.FIXTURE, caplog)
        first_col = [r[0].strip() for r in rows if r]
        # No Name Box token survives ("A26"-style) anywhere.
        flat = [c.strip() for r in rows for c in r]
        assert not any(re.fullmatch(r"[A-Z]{1,3}\d{1,7}", c) for c in flat), \
            f"Excel Name Box leaked into output: {flat[:8]}"
        # No lone column-letter cell ("B", "C D E") survives.
        assert not any(re.fullmatch(r"[A-Z]( [A-Z])*", c) for c in flat), \
            f"column-letter band leaked into output: {flat[:8]}"
        # Row-number gutter stripped: the first real name is 'viswa', not '2 viswa'.
        assert "viswa" in first_col, f"row-number gutter not stripped: {first_col[:5]}"
        assert not any(re.match(r"\d{1,3}\s+\w", c) for c in first_col), \
            f"a '<rownum> name' gutter prefix survived: {first_col[:5]}"

    @requires_rapidocr
    def test_tilt_aligns_top_row_phones(self, convert_image_to_rows, caplog):
        """The photo is tilted ~2°, which smeared the y-band row grouping: the top
        row's phones (viswa's, shaki's) dropped into the next row and merged with it
        ('9980831997 9963393260').  The tilt refinement must re-group so each name
        sits with its OWN phone — eliminating two-numbers-in-one-cell collisions and
        putting viswa's number on viswa's row."""
        import re
        rows, _ = convert_image_to_rows(self.FIXTURE, caplog)
        # No cell glues two phone numbers together any more.
        collisions = [c.strip() for r in rows for c in r if re.search(r"\d{9,}\s+\d{9,}", c.strip())]
        assert not collisions, f"two phones still share a cell: {collisions[:3]}"
        # viswa is on the FIRST row and carries its own number (9963393260), not blank
        # and not merged onto Unknown's row.
        assert rows[0][0].strip().lower().endswith("viswa"), f"viswa not first: {rows[0]}"
        assert "9963393260" in " ".join(rows[0]), f"viswa's phone not on viswa's row: {rows[0]}"
        # Unknown (second row) carries its own number, not viswa's.
        assert "9980831997" in " ".join(rows[1]) and "9963393260" not in " ".join(rows[1]), \
            f"Unknown's row is wrong: {rows[1]}"

    @requires_rapidocr
    def test_last_name_phone_column_is_split(self, convert_image_to_rows, caplog):
        """The rightmost Name/Phone pair was glued into one column by OCR
        ('pargunan 7981233678', 'shyam 8978973222', …).  The validation pass must
        split it so names and phone numbers land in separate columns."""
        import re
        rows, _ = convert_image_to_rows(self.FIXTURE, caplog)
        text = all_text(rows)
        # These names live in the previously-glued last column — they must survive…
        for name in ("pargunan", "shyam", "ramana reddy", "lingswamy"):
            assert name in text, f"{name!r} missing after split: {text[:300]}"
        # …with NO cell still gluing a name onto a trailing phone number.
        glued = [c.strip() for r in rows for c in r
                 if re.match(r"^.*[A-Za-z].*\s+\d{9,}$", c.strip())]
        assert not glued, f"name+phone cells still glued: {glued[:5]}"


def test_strip_spreadsheet_chrome_unit():
    """Deterministic unit check (no OCR): the chrome stripper removes the Name Box
    row, blanks the column-letter band, strips the row-number gutter, and peels a
    column letter glued onto a name — while keeping all real names and phones."""
    import main
    rows = [
        ["A26", "", "", "", ""],                                  # Name Box
        ["2 viswa", "B", "C D E", "F", "G"],                       # gutter + letter band
        ["3 Unknown", "9980831997", "vijay", "9019519925", "shaki H"],
        ["10", "6235118029", "venu", "9944325281", "Sriram 9841507278"],
        ["16 siva prasad", "9052191414", "vijaykumar", "9895702096", ""],
    ]
    out = main._strip_spreadsheet_chrome(rows)
    flat = [c.strip() for r in out for c in r]
    assert ["A26"] not in [[c for c in r if c.strip()] for r in out]   # Name Box row gone
    assert "viswa" in flat and "Unknown" in flat                       # gutter stripped
    assert "siva prasad" in flat
    assert "B" not in flat and "C D E" not in flat and "G" not in flat  # band blanked
    assert "shaki" in flat and "shaki H" not in flat                    # trailing letter peeled
    assert "9980831997" in flat and "6235118029" in flat                # phones preserved
    # The lone row-number "10" became an empty name cell, but kept its phone.
    assert any(r[0].strip() == "" and "6235118029" in r for r in out)


def test_strip_spreadsheet_chrome_noop_on_plain_table():
    """A normal photo/scan (no Excel chrome) must pass through untouched — the
    cleanup is gated so it never mangles ordinary tabular data."""
    import main
    rows = [
        ["Name", "Phone"],
        ["Alice", "9963393260"],
        ["Bob", "9980831997"],
    ]
    assert main._strip_spreadsheet_chrome(rows) == rows


def test_split_glued_name_phone_columns_unit():
    """Deterministic unit check (no OCR): a column that glued 'name 9876543210'
    splits into a name column + a phone column; pure-phone and pure-name columns
    are left alone, and a contact's two numbers stay together in the phone cell."""
    import main
    rows = [
        ["Unknown", "9980831997 9963393260", "vijay", "9019519925", "shaki"],
        ["Suryanarayana", "8921251270", "bharath", "9951762752", "pargunan 7981233678 8129608945"],
        ["jagadeesh", "9941302401", "shantharam", "9894932671", "shyam 8978973222"],
        ["anjayya", "9493964986", "kumar", "7902550932", "ramana reddy 8848436887"],
    ]
    out = main._split_glued_name_phone_columns(rows)
    # The glued last column became two columns: names then phones.
    assert out[1][4] == "pargunan" and out[1][5] == "7981233678 8129608945"
    assert out[2][4] == "shyam" and out[2][5] == "8978973222"
    assert out[3][4] == "ramana reddy" and out[3][5] == "8848436887"  # multi-word name kept whole
    assert out[0][4] == "shaki" and out[0][5] == ""                   # name-only cell, no phone
    # Pure-name and pure-phone columns are untouched (col B's two-number cell stays put).
    assert [r[0] for r in out] == ["Unknown", "Suryanarayana", "jagadeesh", "anjayya"]
    assert out[0][1] == "9980831997 9963393260"   # NOT split — no letter in front


def test_split_glued_name_phone_columns_noop():
    """A clean table whose columns are already separated must pass through
    unchanged — no spurious column splitting."""
    import main
    rows = [
        ["Alice", "9963393260", "Bob", "9980831997"],
        ["Carol", "8921251270", "Dave", "9951762752"],
        ["Eve", "9941302401", "Frank", "9894932671"],
    ]
    assert main._split_glued_name_phone_columns(rows) == rows


def _word(text, left, top, w=80, h=22):
    return {"text": text, "left": left, "top": top, "right": left + w, "width": w, "height": h}


def test_count_number_collisions():
    """The collision counter flags only cells with two long (phone-like) numbers."""
    import main
    rows = [
        ["viswa", "9980831997 9963393260"],   # collision
        ["Unknown", "9980831997"],            # single number, fine
        ["x 12 34", "ab 9876543210"],         # short ints / name+phone, not a collision
    ]
    assert main._count_number_collisions(rows) == 1


def test_group_lines_tilted_regroups_a_tilted_grid():
    """Deterministic (no OCR): on a 3x3 grid tilted so y grows with x, the untilted
    grouping mixes rows, but grouping at the matching slope recovers exactly 3 clean
    rows of 3 words each."""
    import main
    words = []
    for i in range(3):                       # rows
        base = 60 + 26 * i
        for x in (100, 500, 900):            # columns
            words.append(_word(f"r{i}c{x}", x, base + int(x * 0.04)))  # tilt 0.04
    y_tol = 22 * 0.6
    tilted = main._group_lines_tilted(words, 0.04, y_tol)
    assert [len(ln) for ln in tilted] == [3, 3, 3], \
        f"tilt-correct grouping should give 3 full rows: {[[w['text'] for w in ln] for ln in tilted]}"
    # Untilted grouping does NOT cleanly recover the three rows.
    raw = main._group_lines_tilted(words, 0.0, y_tol)
    assert [len(ln) for ln in raw] != [3, 3, 3]


def test_refine_tilted_rows_noop_when_clean():
    """No two-number collisions => refinement must return the rows unchanged
    (straight images are never re-grouped)."""
    import main
    rows = [
        ["viswa", "9963393260", "vijay", "9019519925"],
        ["Unknown", "9980831997", "bharath", "9951762752"],
    ]
    assert main._refine_tilted_rows(rows, [], 1000, 13.0) is rows


def test_desplit_glued_words():
    """Unit check: OCR-glued transaction keywords are split; normal words aren't."""
    import main
    assert main._desplit_glued("POSPURCHASE") == "POS PURCHASE"
    assert main._desplit_glued("PREAUTHORIZEDCREDIT") == "PREAUTHORIZED CREDIT"
    assert main._desplit_glued("ATMWITHDRAWAL") == "ATM WITHDRAWAL"
    assert main._desplit_glued("INTERESTCREDIT") == "INTEREST CREDIT"
    # Must NOT touch standalone keywords or words that merely contain them.
    for safe in ("PURCHASE", "CHECK", "ACCREDITED", "CREDITOR", "DEPOSIT", "Engineering"):
        assert main._desplit_glued(safe) == safe


@requires_tesseract
def test_scorer_rejects_garbage_rows():
    """Unit check on the scorer: clean tables must outrank OCR-noise tables."""
    import main
    clean = [
        ["DATE", "DESC", "AMOUNT"],
        ["03/02", "Internet Bill", "75.99"],
        ["03/05", "Electric Bill", "253.68"],
        ["03/06", "Phone Bill", "44.10"],
    ]
    garbage = [
        ["abl", "e 1:", "7)", "lrlite", "ard", "Em", "iploye", "e", "Dat", "a", "_", "�", "-", "x", "y", "z"],
        ["Telle", "yaa", "Tey", "a", "Nama", "ae", "'", "D", "99", "rt", "Ss", "1", "f", "_", "ee", "z"],
    ]
    assert main._score_table_rows(clean) > main._score_table_rows(garbage)
