# Conversion Notes & Learnings

Hard-won lessons from building and debugging the document→Excel converters in `main.py`.
Read this **before** changing any image/PDF detection logic — it will save hours.

> **Regression tests:** `backend/tests/test_image_conversion.py` locks in the behaviour below
> against real fixture images (`backend/tests/fixtures/`). Run `python -m pytest` after any change
> to the image pipeline. Add a fixture + assertions when you fix a new problem image.

---

## 1. Debugging Methodology (the #1 efficiency lesson)

> **Never iterate blind on screenshots. A screenshot ≠ the file.**

We once burned ~10 rounds tweaking image detection based on screenshots. The breakthrough came
only after capturing the *actual uploaded file* and printing raw pixel values — which revealed
the "white" document was actually a **dark purple-gray background (luminance ~98)**. Every
absolute-brightness heuristic had been failing silently for that reason.

**The fast path to diagnose a bad conversion:**

1. **Capture the real input.** Temporarily save the uploaded bytes in the `/api/convert` route
   (`convert_file`, ~line 1567) before conversion:
   ```python
   if ext in (".png", ".jpg", ".jpeg"):
       (OUTPUTS_DIR.parent / "debug_input").mkdir(exist_ok=True)
       (OUTPUTS_DIR.parent / "debug_input" / f"last_upload{ext}").write_bytes(content)
   ```
   Upload once through the UI, then remove the snippet.

2. **Run a throwaway script** that imports `main` and calls the detection functions directly.
   Set the Tesseract path manually — the auto-locate logic only runs inside `convert_image`:
   ```python
   import pytesseract
   pytesseract.pytesseract.tesseract_cmd = r"C:\Program Files\Tesseract-OCR\tesseract.exe"
   ```

3. **Inspect, don't assume.** Print per-row luminance / dark-fraction profiles to *see* where
   bands and backgrounds actually are:
   ```python
   import numpy as np
   gray = np.array(Image.open(path).convert("L"))
   row_lum = gray[:, gray.shape[1]//5 : 4*gray.shape[1]//5].mean(axis=1)
   for r in range(0, len(row_lum), 10): print(r, round(row_lum[r]))
   ```

---

## 2. Image Pipeline Architecture — Best-Of-N Strategy Selection

`convert_image` does NOT use first-match-wins (an early strategy that mis-fires used to override
everything — that caused the dark-image 16-column garbage regression). Instead it **runs every
applicable strategy as a candidate, scores each result, and keeps the best.** A strategy that
mis-fires simply loses on score.

| Candidate | Function(s) | Best for |
|---|---|---|
| `grid` | `_detect_table_grid` → `_extract_cells_from_grid` | Real drawn gridlines: morphological H/V line detection → intersections → crop each cell → OCR with **PSM 7** (single line). |
| `header_band` | `_detect_header_band` → `_columns_from_header_band` | Solid colored header bar (e.g. purple) with white labels. The header's word x-positions **are** the column boundaries. *Solved FIRST BANK.* |
| `word_based` | `_find_col_boundaries` → histogram → right-edge → `_detect_columns_from_gaps` | Safe baseline, always present. |
| `left_aligned` | `_columns_from_left_edges` | Left-aligned multi-column tables (employee/project tables) — clusters word LEFT edges across rows. *Solved the dark two-table image.* |
| `ruled_columns` | `_columns_from_ruled_lines` | Bordered light-bg statements with FAINT grey rules (e.g. SBI). Detects vertical rules via a fixed high threshold (Otsu misses them), then renders the good whole-image OCR into those column x-boundaries. *Solved the SBI ruled statement — separated the lone Credit value.* |

**Scorer — `_score_table_rows(rows)`** rates how table-like a candidate is (pure function over the
rows): + column consistency, + real-token ratio, + real cells per row (genuine column structure),
+ fill ratio; − garbage ratio (×5, sinks OCR noise), − implausible column counts (>8). This is what
demotes the bogus 16-column grid (score ~3) below a clean candidate (~10).

**Confidence prior:** `header_band` gets a +2.0 bonus because a detected colored header is a strong
image-type signal AND it legitimately prepends metadata text rows above the table (which would
otherwise look "inconsistent" to the scorer and lose to a wrongly-merged `left_aligned`).

> **Adding a strategy is safe** — `max()` over candidates means a new strategy can only win when it
> scores higher; it never makes existing cases worse. Tune the scorer, not the detectors.

### Scorer-tuning pitfalls (learned the hard way)
- **A "real token" regex (`[A-Za-z0-9]{2,}`) does NOT separate good from garbage.** OCR noise like
  `Telle`, `abl`, `9923`, `Nama` all pass it. Rewarding *real-cells-per-row* alone let the bogus
  16-column grid (full of plausible-looking fragments) win. Counter it with a strong **garbage-ratio
  penalty (×5)** and an **implausible-column-count penalty (>8 cols)** — the grid had 16 cols and a
  0.23 garbage ratio vs a clean candidate's 7 cols / 0.0.
- **Correct sparse tables look "inconsistent" to a naive scorer.** A right answer with metadata text
  rows above a 5-col table (FIRST BANK) has mixed column counts → low consistency → it lost to a
  *wrongly* merged but uniform `left_aligned`. Fix: the `header_band` confidence bonus, because the
  band is a high-confidence structural detection, not an inferred guess.
- **Calibrate against real candidate component dumps, not intuition.** Print each candidate's
  `(n, max_cols, fill_ratio, real_ratio, garbage_ratio, consistency, real_per_row)` and pick weights
  that separate the known-good from known-bad with margin. Thin margins (e.g. 7.79 vs 7.61) are
  fragile — aim for a clear gap.

---

## 3. Pixel-Level Gotchas

- **Use RELATIVE brightness, not absolute thresholds.** `_detect_header_band` compares each row
  to the **page-median luminance** (band = rows ≥35% darker than median). An absolute `< 110`
  test failed because the whole background sat at ~98 luminance.

- **Many "white" documents aren't.** Dark-mode screenshots and tinted exports have mid-gray
  backgrounds. Always verify against real pixel values before trusting a brightness assumption.

- **Detect text POLARITY before OCR, not brightness.** Tesseract reads black-on-white only. A
  light-text-on-dark image (dark-mode screenshot, mean luminance ~45) reads as garbage unless the
  grayscale is inverted first (`gray = 255 - gray`). But do NOT decide by mean brightness — a
  dark-text-on-mid-gray bank statement (mean ~98) must NOT be inverted. Use Otsu pixel mass: the
  minority pixel class is the "ink"; if dark pixels are the MAJORITY the background is dark →
  invert. This one fix took the dark two-table image from 3 garbled lines to all rows readable.

- **Camera PHOTOS of screens need a denoise fallback, gated on OCR confidence.** A low-contrast
  phone photo of a monitor (glare, blur, screen moiré) makes the 31×31 adaptive threshold produce
  speckle → pure OCR garbage. Classic noise metrics (residual std, Laplacian variance) do NOT flag
  it — it reads as *low* high-frequency content because it's blurry, not noisy. The reliable signal
  is **mean OCR confidence**: the garbage pass scored ~20 while every clean screenshot scores 70-95.
  So: after the primary adaptive pass, if mean confidence < 55, retry with `fastNlMeansDenoising` +
  global Otsu and keep whichever reads more confidently. Clean images never cross the threshold, so
  they're untouched (no second OCR pass, no regression). Took the contacts photo from garbage to
  most names/numbers recovered. Phone digits still imperfect — photos can't match a screenshot.

- **For degraded photos, swap the OCR ENGINE, not the preprocessing — RapidOCR (local neural OCR).**
  When Tesseract's mean confidence is low (degraded camera photo of a screen), `convert_image` calls
  `_ocr_words_rapidocr` — **RapidOCR** (`rapidocr-onnxruntime`, PP-OCR models via ONNXRuntime): free,
  local, CPU-only, models bundled (no card / quota / internet / download). It returns `(box, text,
  score)` per region, mapped into the same `all_words` shape, so all the downstream column-detection
  / best-of-N / merge logic is reused unchanged. On the contacts photo it took recovery from ~0-2
  top-left names + ~8 phones (Tesseract) to **10/10 names + ~46 phone numbers**. **Gating:** triggered
  at Tesseract mean-conf **< 65** (contacts is ~57; clean docs are 70-95, so they keep their tuned
  Tesseract path untouched — zero regression, no added latency). RapidOCR adds a few seconds (ONNX
  init + inference) and does text+boxes only (no native table structure — PP-Structure/PaddleOCR is
  the heavier option for that). This is the realistic "our own AI": run a pretrained neural model
  locally, don't train one. Tests: `requires_rapidocr` marker; `TestContactsPhoto`.

- **Don't keep chasing preprocessing for a degraded photo region — it's a Tesseract ceiling.**
  Measured on the contacts photo's unreadable top-left block (glare/blur): `denoise+Otsu` (the
  current fallback) already recovered the MOST text; CLAHE (any clip/tile), bilateral filter,
  unsharp mask, lighter denoise, and 3× upscale all did **equal or worse**. Two traps proven here:
  (1) **CLAHE does not fix a localized glare/blur patch** (only a smooth global illumination
  gradient, which this wasn't); (2) **mean confidence is NOT a reliable proxy for actual recovery** —
  the highest-confidence variant (raw denoised grayscale, conf 75) recovered *fewer* real names than
  denoise+Otsu (conf ~59). So a confidence-based "best-of-N preprocessing" selector can pick a
  worse-reading variant. Conclusion: when a photo *region* is genuinely degraded, no thresholding
  trick recovers it — the fix is a better capture (flat, evenly-lit, or a screenshot) or a cloud
  vision model (Gemini/Azure), not more OpenCV. Don't add speculative preprocessing that doesn't
  beat denoise+Otsu on a held-out check.

- **Invert colored header bands before OCR.** White-on-purple is invisible to Tesseract +
  adaptive threshold. Crop the band → `255 - gray` → Otsu threshold → black-on-white → readable.
  (See `_columns_from_header_band`.)

- **Faint grey table rules need a FIXED threshold, not Otsu.** Light-grey gridlines on a near-white
  page sit just below white; Otsu's global split puts them on the white side, so morphology finds
  nothing (0 intersections). A fixed `gray < 240` catches them; long-kernel morphology isolates the
  rules from text. Use the rules for **column x-boundaries only**, then render the good whole-image
  OCR into them — per-cell re-OCR of small/multi-line cells (`_extract_cells_from_grid`) garbles
  dates and drops wrapped narration. Gate ruled detection to light backgrounds (mean > 200) so a
  dark/inverted image's foreground can't be mistaken for a grid. (See `_columns_from_ruled_lines`.)

- **Full-page OCR often SKIPS a table's header row** when it sits between horizontal rules — the
  page-level PSM-6 segmentation drops it, so the columns come out unlabelled. Fix: detect the
  horizontal rules too, OCR each inter-rule band on its own, and take the first band containing ≥4
  table-header keywords (`date/narration/debit/credit/balance/...`). Map those words into the ruled
  columns and **insert** the header at the table boundary (`_header_row_for_ruled`). Note: keep the
  metadata rows rendered IN columns rather than splitting them to single-cell text — the split
  tanks the candidate's score (consistency/real-per-row drop) and it loses to `word_based`. Insert
  the header inline instead of via the metadata-split path.

- **Per-region OCR beats whole-image OCR.** PSM 7 (single text line) on a cropped cell or band is
  far more accurate than PSM 6 on the full page.

- **Split metadata from the table.** Key-value header text above the table (address, account
  summary) pollutes column detection — especially right-edge numeric clustering. `convert_image`
  splits `line_groups` at the header-band y-position so detection runs only on table rows.

- **Fuzzy-match OCR'd headers to known names** with `difflib.SequenceMatcher` (≥0.6 ratio) to fix
  glitches like `O�SCRIPTION` → `DESCRIPTION`. Note: a positional char-by-char comparison FAILS
  on shifted text (a dropped leading char misaligns everything); sequence matching handles it.

- **Do NOT gate OCR words on confidence (`conf > 20` was wrong).** Tesseract confidence is
  unreliable — it assigns conf **0–15 to perfectly-correct words** (`CHECK 1249`, `SERVICE CHARGE`,
  a `POS PURCHASE` line). A `conf > 20` filter silently dropped those cells, leaving table rows with
  a date and amounts but a blank description. Keep every emitted word (`conf >= 0`; drop only
  Tesseract's `-1` structural markers) and let the **best-of-N scorer** handle any noise. On the
  Account-Transactions image this recovered 3 dropped descriptions AND let `word_based` win with a
  proper 5-column split (Credit and Balance had previously merged because the Credit column was too
  sparse to detect once its values were filtered out). Lesson: an upstream filter that drops data is
  far more dangerous than downstream noise the scorer can rank away.

- **Re-split OCR-glued keywords, conservatively (`_desplit_glued`).** All-caps statements sometimes
  lose spaces (`POSPURCHASE`, `ATMWITHDRAWAL`, `PREAUTHORIZEDCREDIT`). Split a known transaction
  keyword off the END of an all-caps `<prefix><keyword>` token only — never mid-word — so
  `ACCREDITED`/`CREDITOR`/standalone `DEPOSIT` are untouched. **Do this for text, never for
  numbers:** we deliberately do NOT "correct" OCR'd amounts (e.g. `.26`→`-26`) — silently altering a
  financial figure risks corrupting a correct one. Text cleanup is safe; number cleanup is not.

---

## 4. PDF Conversion Notes

`convert_pdf` (line 684):

- **Try `page.extract_tables()` first** for bordered PDFs, filtered by `_is_useful_table`
  (≥2 columns, ≥2 non-empty rows). Structured tables found this way bypass the word-based path.

- **But reject DEGENERATE extracted tables (`_is_degenerate_table`).** A bank statement with column
  lines but NO horizontal lines between transaction rows makes pdfplumber cram an entire page into
  ONE row, each cell holding many newline-separated values (e.g. SCB statement → a 3-row table
  whose Withdrawal cell is `"1,500.00\n1,500.00\n966.00\n…"`). Signal: some cell holds ≥4 text
  lines AND more lines than the table has rows. When ANY table on a page is degenerate, abandon the
  structured path for the **whole page** and fall back to the word-based pipeline (which segments
  one transaction per row correctly). This does NOT trip `sample-tables.pdf` (well-segmented,
  ≤1 line/cell). *This was a real regression: adding extract_tables-first for sample-tables broke
  the SCB statement until this guard was added.*

- **`_merge_continuation_rows` (line 74) is BANKING-SPECIFIC.** It merges rows that lack a date
  and an amount into the previous row (for wrapped transaction descriptions). It must **not** run
  on general tables — it once mashed all 29 tables of `sample-tables.pdf` together. It now only
  runs on the word-based fallback path, never on tables from `extract_tables()`.

- **Lesson: structured extraction is not always better than word-based.** For bordered, fully-ruled
  tables it wins; for ruled-columns-but-unruled-rows statements it fails. Validate the extracted
  result before trusting it — same "don't blindly prefer one strategy" principle as the image
  best-of-N selector.

---

## 5. The Tesseract Ceiling (set expectations)

- The column **structure** can be made correct with the techniques above.
- Residual **text** glitches (`Stotement`, `0306`, `= Payroll Run`) are Tesseract's accuracy
  limit on low-contrast / dark images — **not a logic bug**. Don't chase them with more
  preprocessing; it tends to break other rows.
- Matching commercial tools (jpgtoexcel.com, imagetotext.info) on raw text accuracy requires a
  **cloud OCR engine** (Google Cloud Vision or Azure Computer Vision). That's the known path for
  a future step-change in accuracy — a single API call would replace the entire Tesseract
  preprocessing pipeline.

---

## 6. Verified Result — FIRST BANK image

**Before:** everything mashed into 2 columns (description + a blob of amounts in column B).

**After** (tier-2 colored-header-band detection):

| DATE | DESCRIPTION | WITHDRAWAL | DEPOSIT | BALANCE |
|---|---|---|---|---|
| | Previous balance | | | 27,584.38 |
| 03/02 | Internet Bill | 75.99 | | 27,508.39 |
| 03/05 | Electric Bill | 253.68 | | 27,254.71 |
| 0306 | Check No. 4598 | | 456.84 | 27,711.85 |
| | Deposit from Credit Card Processor | | 5,891.26 | 33,602.81 |
| 03/12 | Payroll Run | 3,893.75 | | 29,708.06 |
| 03/16 | Debit Transaction | 243.46 | | 29,464.60 |
| 03/21 | Rent Bill | 750.00 | | 28,714.60 |
| 03/21 | Check No. 234 | | 263.84 | 28,983.44 |

Withdrawals, deposits, and balances all land in the correct columns.

---

## 7. Testing (regression guard)

`backend/tests/` is a pytest suite that locks in everything above against real fixture images.

```bash
cd backend
pip install -r requirements-dev.txt   # pytest, first time only
python -m pytest                       # ~26s
```

**How it's built (and why):**
- **Real fixtures, not synthetic.** `tests/fixtures/two_table_dark.png` (dark two-table) and
  `sample2.png` (FIRST BANK). Synthetic images using PIL's default font OCR differently from real
  renders and gave false confidence — always test the actual file.
- **Tolerant assertions.** Check for a *majority* of expected tokens (≥60-70%) and **structural
  properties** (chosen strategy, max non-empty columns, row count) — NOT exact text. Tesseract
  output varies (`Satary`, `27,25471`); exact-match tests would be flaky.
- **Assert the chosen strategy.** Tests capture the `chose '<strategy>'` log line, so they verify
  *which* candidate won (e.g. dark image must NOT pick `grid`; FIRST BANK must pick `header_band`).
  This is the strongest regression guard — it catches a scorer regression even if output looks ok.
- **Memoize per fixture.** OCR is ~10s/image; `conftest.py` converts each fixture once per session
  (8 tests → 2 OCR runs). Cut the suite from 107s to 26s.
- **Auto-skip without Tesseract** via `requires_tesseract` so CI on a machine without the engine
  doesn't fail spuriously.
- **One pure-function unit test** (`test_scorer_rejects_garbage_rows`) checks `_score_table_rows`
  directly — fast, no OCR, pins the clean-beats-garbage invariant.

**When you fix a new problem image:** drop it in `tests/fixtures/`, add a class with a couple of
tolerant assertions (tokens + structure + strategy), and confirm `pytest` stays green.

### Tilt: rotate the GROUPING math, never the image (`_refine_tilted_rows`)

**The defect.** The contacts photo is tilted ~2° (NOT the ~0.5° a global projection-profile
estimator reports — that's confounded by the laptop bezel + mild perspective). Measured directly
from word boxes: within the top row, the y-centre drifts ~37px from the left edge to the right edge
(slope ≈ 0.044). Since the y-band grouping clusters by y-centre within ±~14px, that 37px spread
**smears whole rows** — the top row's phones drop into the next row and merge ("9980831997
9963393260"), and the middle/right column-pairs end up shifted a full row against the names.

**What does NOT work — rotating the image.** Tested twice: rotating the bitmap and re-running OCR
produces brand-new word boxes and *scrambles* the result into a worse cascade. Both rotation signs
confirmed. RapidOCR is already robust to the small angle; re-interpolation only hurts.

**What DOES work — correcting the grouping coordinate.** Group by a tilt-adjusted y,
`cy - left*slope`, with NO image rotation and NO re-OCR. At slope ≈ 0.035 the contacts grid is
recovered essentially perfectly (every name with its own phone, top to bottom).

**Estimating the slope is the hard part — so don't.** Five global estimators (least-squares,
median-pair, centroid/box projection profile, line-count, clean-row count) all either
*underestimate* (perspective makes a single global angle ambiguous) or *false-trigger on a straight
image* (one picked +2° on `sbi_statement`). The pipeline's own `_score_table_rows` can't help
either: it scores column *consistency*, which is identical whether a row is correctly aligned or
shifted by one — margins were noise (11.34 vs 11.38).

**The robust trigger = the defect itself.** Instead of estimating tilt, detect the *symptom*: a cell
holding **two phone-length numbers** (`\d{9,}\s+\d{9,}`). Only when that exists, re-group at a few
fixed slopes (0.02/0.035/0.05), re-detect columns with the existing detectors, and keep the version
with the **fewest** such collisions (tie-break: higher table score). Because it fires only on the
defect and must strictly reduce it, every collision-free image — i.e. all the straight screenshots —
is a guaranteed no-op (`test_refine_tilted_rows_noop_when_clean`). This fixed the whole contacts
table, not just the two cells the user spotted.

**Lesson:** when a global parameter is unidentifiable, optimise the concrete defect signal directly
and gate on its presence — far more robust than estimating the parameter. (See also the CLAHE and
"don't rotate the image" findings — preprocessing that perturbs the bitmap rarely helps OCR here.)

### Spreadsheet-screenshot chrome (`_strip_spreadsheet_chrome`)

When the uploaded image is a **photo/screenshot of an Excel/Sheets window**, OCR also reads the
app's own UI, which is not data:
- the **Name Box** ("A26") → lands as a lone first row;
- the **column-letter band** ("A B C D …") → OCR folds it into the first data row as separate
  cells ("B", "C D E", "F", "G"), and that spurious row **shoves the first contacts' phone numbers
  down a row** (the user-reported symptom);
- the **row-number gutter** ("2", "3", …) → glues onto the leftmost name ("2 viswa").

A final validation pass in `convert_image` (right before `write_data_to_sheet`) strips all three —
but **only when the rows actually look like a spreadsheet screenshot** (a Name-Box token, or ≥2
column-letter cells in the top rows). Ordinary photos/scans are a guaranteed no-op
(`test_strip_spreadsheet_chrome_noop_on_plain_table`). Edits are bounded to the **top 3 rows** for
the letter band / glued-letter cases and to **column 0** for the gutter, so a legitimate
single-letter or leading-number cell deeper in the table is never touched.

### Glued name+phone column (`_split_glued_name_phone_columns`)

When OCR misses the gap between a name column and its phone column, every cell in
that column reads as `"name 9876543210"` (e.g. the rightmost pair on the contacts
photo: `pargunan 7981233678`, `shyam 8978973222`, `ramana reddy 8848436887`). The
validation pass splits such a column into a name column + a phone column.

Gating that keeps it safe:
- A column is split **only** when a strong majority of its filled cells (≥60%, and
  ≥3 cells) match `"<name> <phone>"`.
- The name part **must contain a letter** — so a pure two-phone cell like
  `9980831997 9963393260` is *not* mistaken for name+phone and the phone column is
  left intact.
- The phone part requires runs of **≥9 digits**, so short references/amounts/dates
  (`CHECK 1249`, `253.68`, `03/02`) never trigger a split.
- A contact's two numbers stay together in one phone cell (`7981233678 8129608945`).

Clean tables whose columns are already separate are a no-op
(`test_split_glued_name_phone_columns_noop`); the bank-statement/employee fixtures
confirm no false splits.

**Honest limit:** the chrome removal does *not* repair the one OCR artifact it exposes — viswa's
own phone got vertically mis-grouped into the next row's cell ("9980831997 9963393260") at the very
top edge. The two numbers' order is ambiguous, so re-splitting/re-shifting would risk corrupting
good data; we strip the chrome and leave that single merged cell rather than guess. Everything from
the second contact down aligns correctly.

---

*Line numbers refer to `backend/main.py` as of this writing. If the file shifts, re-grep the
function names rather than trusting the numbers.*
