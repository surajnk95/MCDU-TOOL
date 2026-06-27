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

Tests live in `tests/test_app.py` and import `app` directly. No test database or fixtures needed — most tests use synthetic `PIL` images or call pure functions.

## Architecture

The entire backend is a single file: `app.py`. It runs a stdlib `ThreadingHTTPServer` (port 8766) — no framework. All API routes are handled in `McmduHandler.do_POST` (line 2480). The frontend is `static/{index.html,app.js,styles.css}`.

**API endpoints** (all POST, JSON in/out):
- `/api/detect-display` — finds MCDU screen corners from a raw photo
- `/api/flatten-display` — perspective-warps the screen to a rectangle
- `/api/analyze` — runs OCR and returns a populated grid
- `/api/refine-grid` — second OCR pass with re-focused crop
- `/api/review-requirements` — validates specific cells against expected values
- `/api/remember` — saves corrected cells as row-level corrections
- `/api/remember-grid` — saves corrections from a complete grid
- `/api/remember-templates` — learns character templates from changed cells
- `/api/export-docx` — writes the grid to a `.docx` file in `data/exports/`

**Image processing pipeline** (inside `analyze`):
1. Load image → detect display corners → perspective warp to 1600×N px
2. Estimate grid origin (`estimate_grid_origin`) to align cell boundaries with character gaps
3. Run Tesseract (TSV + box passes) → place words/chars into grid
4. Optionally run PaddleOCR (lazy-init in background thread) → fuse results (`fuse_engine_grids`)
5. Apply learned corrections (`apply_corrections`) and character templates (`apply_templates`)
6. Disambiguate O/0 (`disambiguate_o_zero`) and recover dash lines (`recover_dash_lines`)

**Learning / persistence** (`data/` folder, created at startup):
- `corrections.json` — maps `"row:raw_row_text"` → corrected row text; a correction is only reused on an exact raw-text match
- `templates.json` — maps character label → list of pixel feature vectors; learned only from cells the user explicitly changes

**OCR fusion** (`fuse_engine_grids`, line 1157): Tesseract and PaddleOCR grids are compared cell by cell. Agreement boosts confidence; disagreement marks the cell with a warning colour for manual review.

**Display detection** (`detect_display` → `refine_display_corners`, line 1372): finds the largest dark rectangular region via connected components, then refines the four corner lines using robust line fits on edge pixels.

## Key Constants

```python
ROWS = 13; COLS = 40; FIRST_DATA_COL = 1; LAST_DATA_COL = 38
SCREEN_W = 1600          # warp target width
OCR_WHITELIST            # characters Tesseract is restricted to
BLANK_TEMPLATE_KEY       # sentinel for a "this cell should be empty" template
```

## `single_python_version/`

Contains `mcdu_tool_single.py` — a self-contained single-file version. Kept in sync manually when significant changes land in `app.py`.
