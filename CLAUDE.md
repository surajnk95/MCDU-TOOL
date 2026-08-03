# CLAUDE.md

This file provides guidance to Claude Code (claude.ai/code) when working with code in this repository.

## What This Project Is

A local web tool that extracts text from phone photos of a Boeing 777-9 MCDU display and places it into a 13-row × 40-column reference grid. The grid always has 40 physical columns; physical columns 0 and 39 are always blank; labeled/usable columns are 1–38.

## Running the App

```bash
python app.py           # starts server at http://127.0.0.1:8766
./run.sh                # same via shell wrapper
PORT=9000 python app.py # custom port
```

Install dependencies first:

```bash
pip install -r requirements.txt           # standard (Tesseract only)
pip install -r requirements-hybrid.txt    # adds PaddleOCR (Python 3.11–3.13)
```

Tesseract must be installed separately. The app auto-discovers it at common paths or reads the `TESSERACT_CMD` environment variable.

## Running Tests

```bash
python -m pytest tests/test_app.py -v
python -m pytest tests/test_app.py -v -k "TestClassName"   # single class
python -m pytest tests/test_app.py -v -k "test_method_name" # single test
```

Tests live in `tests/test_app.py` and `tests/test_regression.py` (a synthetic end-to-end harness) and import `app` directly. No test database or fixtures needed — most tests use synthetic `PIL` images or call pure functions.

## Architecture

The entire backend is a single file: `app.py`. It runs a stdlib `ThreadingHTTPServer` (port 8766) — no framework. All API routes are handled in `McmduHandler.do_POST`. The frontend is `static/{index.html,app.js,styles.css}`.

Locate functions with `grep -n "def <name>" app.py` rather than trusting line numbers in docs — `app.py` is ~4,300 lines and any number written down here goes stale within a few commits.

**API endpoints** (all POST, JSON in/out):
- `/api/detect-display` — finds MCDU screen corners from a raw photo
- `/api/flatten-display` — perspective-warps the screen to a rectangle
- `/api/analyze` — runs OCR and returns a populated grid
- `/api/refine-grid` — second OCR pass with re-focused crop
- `/api/review-requirements` — validates specific cells against expected values
- `/api/remember-grid` — saves corrections from a complete grid
- `/api/remember-templates` — learns character templates from changed cells
- `/api/fuse-grids` — combines 2–3 grids from separate photos of the same page
- `/api/export-docx` — writes the grid to a `.docx` file in `data/exports/`

**Image processing pipeline** (inside `analyze`):
1. Load image → detect display corners → perspective warp to 1600×N px
2. Estimate grid origin (`estimate_grid_origin`) to align cell boundaries with character gaps
3. Run Tesseract (TSV + box passes) → place words/chars into grid
4. Optionally run PaddleOCR (lazy-init in background thread) → fuse results (`fuse_engine_grids`)
5. Apply learned corrections (`apply_corrections`) and character templates (`apply_templates`)
6. Disambiguate O/0 (`disambiguate_o_zero`) and recover dash lines (`recover_dash_lines`)

**Learning / persistence** (`data/` folder, created at startup):
- `corrections.json` — maps `"row:N|raw_row_text"` → corrected row text; a correction is only reused on an exact raw-text match
- `templates.json` — maps character label → list of pixel feature vectors; learned only from cells the user explicitly changes

Two baselines drive learning, and they are deliberately different. `/api/analyze` returns both `grid` (after `apply_corrections`) and `rawGrid` (before it). Row corrections must key on `rawGrid`, or a second correction to the same row is stored under the first correction's text and can never be reached again — `apply_corrections` is a single pass. Template learning keys on the *displayed* grid, so only cells the user actually retyped contribute samples. The frontend keeps these as `state.rawGrid` and `state.sourceGrid`.

Anything that learns from or classifies against the warped image must use the **cleaned** warp (`clean_warped_image`): learning on the raw warp while inferring on the cleaned one bakes the cursor, glare and entry outlines into stored feature vectors.

**OCR fusion** (`fuse_engine_grids`): Tesseract and PaddleOCR grids are compared cell by cell. Agreement boosts confidence; disagreement marks the cell with a warning colour for manual review.

**Display detection** (`detect_display` → `refine_display_corners`): finds the largest dark rectangular region via connected components, then refines the four corner lines using robust line fits on edge pixels.

**Concurrency**: `analyze` fans OCR passes out across `ThreadPoolExecutor`s while `tesseract` subprocesses read `data/mcdu_user_words.txt` and `mcdu_user_patterns.txt`. Never write those (or the JSON stores) with a plain `write_text` — use `write_text_atomic`, and keep `ensure_dirs()` out of any per-OCR-call path.

## Key Constants

```python
ROWS = 13; COLS = 40; FIRST_DATA_COL = 1; LAST_DATA_COL = 38
SCREEN_W = 1600          # warp target width
OCR_WHITELIST            # characters Tesseract is restricted to
BLANK_TEMPLATE_KEY       # sentinel for a "this cell should be empty" template
```

## Offline constraint

This tool runs on a restricted office machine. It must make **no outbound network connections** — no cloud OCR, no web APIs, no telemetry, no CDN assets. `tests/test_app.py::OfflineGuaranteeTests` enforces this and will fail if network-capable calls appear in `app.py`.

PaddleOCR is the one component that reaches the network (it downloads model weights on first construction). It is gated behind `PADDLE_ENABLED` (`MCDU_ENABLE_PADDLE=1`) and the UI checkbox defaults to off. Do not re-enable either by default.

## OCR pass selection — measured, do not revert casually

Defaults here were set by measuring against the real office photos in `test_photos/` (transcribed ground truth, whitespace-collapsed row similarity) *and* the synthetic harness. Changing them without re-measuring will regress accuracy.

- **Per-row strip OCR (`per_row_strip_ocr`) is off by default.** PSM 7 assumes its input is a line of text, so it invents characters inside an empty entry box, and the inflated character count out-scores a better whole-image read in `score_candidate_row`. Disabling it moved row similarity 68.3% → 72.1% and was faster. Synthetic clean/blur-1.2/blur-2.0 stayed at 100%. Opt in per request with `{"rowStripOcr": true}`.
- `_PER_ROW_SCORE_WEIGHT` discounts that candidate when it *is* enabled, for the same reason.
- **All eight preprocessing variants are kept by default** — each wins rows outright. `fast=True` (`{"fastMode": true}`) drops to `FAST_VARIANTS`, ~25% faster at no measured accuracy cost.
- The orientation probe scores only 0/90/270. A 180 pass was previously computed and discarded unread.
- **`repair_confusable_tokens` runs on each candidate grid *before* row selection**, not on the final grid. A label bleeding into the row below arrives as `CR2 ALT` over `CRZ ALT` — two different strings, so `deduplicate_adjacent_rows` cannot see the repeat and the boxed value underneath stays lost. Normalising first is worth ~1.5 points on its own. It never rewrites a token without letters, so altitudes and speeds (`250`, `12000`, `2000A`) are safe.

`test_photos/` holds the reference photos and is gitignored — never commit them. `MCDU_VOCABULARY` feeds three things at once (the Tesseract user-words file, `mcdu_row_score` ranking, and token repair), so adding a real page word is usually a cheap win.

## Disabled subsystems

`ATLAS_ENABLED = False` gates the glyph-atlas classifier (`_build_glyph_atlas`, `classify_cell_atlas`, `fuse_atlas_grid`). It caps at ~75% cross-font accuracy on available system fonts, below the 85% threshold. Leave it off unless you are working on that specifically.

`analyze` calls `calibrate_grid([], screen_size)` with an empty box list on purpose — the UI grid is exactly 40×13 and OCR must use the identical pitch, so the fitted-scale path in `calibrate_axis` is intentionally never taken in production.
