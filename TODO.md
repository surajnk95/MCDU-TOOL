# MCDU Tool — Accuracy & Robustness Roadmap

This is an implementation backlog for improving OCR accuracy and handling real-world
office photos (blur, glare, angled shots, rotated images, colored text, message boxes,
and **multiple MCDU screens in one photo**).

## Hard constraints (do not violate)

- **100% local / offline.** No cloud OCR, no web APIs, no telemetry. Everything runs on
  the user's laptop and a locked-down office laptop.
- **No new system-level installs beyond what already exists.** Allowed dependencies are
  the pure-pip wheels already in `requirements.txt` (`numpy`, `opencv-python-headless`,
  `Pillow`, `python-docx`) plus the existing system `tesseract` binary. PaddleOCR stays
  **optional** (`requirements-hybrid.txt`) and must never become required.
- **OpenCV is already installed and is the main lever** — it is currently almost unused.
  Prefer `cv2` for all new image processing.
- Keep `app.py` runnable with the standard requirements only; degrade gracefully if an
  optional feature can't run.

## Key reference points in `app.py`

| Area | Function(s) |
| --- | --- |
| Image load | `load_image` (~205) |
| Display detection | `detect_display` (~1557), `refine_display_corners` (~1372), `component_corners` (~1469), `connected_components` (~1238) |
| Warp / grid align | `warp_screen` (~352), `align_warp_to_grid` (~323), `estimate_grid_origin` (~275), `calibrate_grid` (~991) |
| Preprocessing | `preprocess_for_ocr` (~357), `preprocessing_variants` (~365) |
| Tesseract | `run_tesseract_tsv` (~742), `run_tesseract_boxes` (~808) |
| Templates | `cell_feature` (~374), `feature_distance` (~393), `classify_from_templates` (~424), `apply_templates` (~1994) |
| Post-processing | `clean_ocr_text` (~883), `normalize_mcdu_phrase` (~892), `disambiguate_o_zero` (~2060), `recover_dash_lines` (~2107) |
| Orchestration | `analyze` (~2252), `refine_grid` (~2357) |
| Constants | `MCDU_VOCABULARY`, `MCDU_WORD_HINTS`, `OCR_WHITELIST`, `ROWS`, `COLS` |

---

## IMPORTANT — color-channel correction

An earlier suggestion was to use the **green channel** for preprocessing. **Do not do that.**
The MCDU draws magenta (active/modified), cyan (boxes), amber (cautions), and green text —
magenta and amber are nearly invisible in the green channel.

**Use the max-of-RGB channel (equivalently the HSV "Value" channel):** `V = max(R, G, B)`.
All colored text and white text are bright in at least one channel, while the screen
background stays dark. This becomes the base grayscale for OCR and for `cell_feature`.

---

## Priority 1 — Multi-screen detection & selection (headline feature)

**Problem:** Photos often contain a second full screen or a partial neighbor screen at the
edge. `detect_display` only keeps the single largest dark rectangle and can latch onto the
wrong one or straddle two.

1. **Detect every candidate screen.** In `detect_display` / `connected_components`, collect
   *all* dark rectangular components, not just the largest.
2. **Score and filter candidates** by: area, aspect ratio close to the MCDU screen ratio,
   rectangularity (fill of its bounding box), and **completeness** — reject screens whose
   bounding box touches the image border (these are cut-off neighbors).
3. **Return all valid candidates** from `/api/detect-display` (list of corner sets +
   thumbnails / bounding boxes), not a single quad.
4. **Auto-pick the best candidate by default** (most complete + most central + largest),
   but let the user click a different detected screen in the UI to override.
5. **Crop to the chosen screen before warp & OCR** so the other display can never pollute
   the grid. Everything downstream (`warp_screen`, `analyze`) operates on the single chosen
   screen only.
6. UI: when 2+ screens are found, draw selectable outlines over the photo; clicking one
   sets the active corners. Keep the existing manual corner-drag as the final fallback.

**Acceptance:** photos 1/2/5 (one main + partial neighbor) lock onto the main screen; photo
3 (two screens) lets the user choose either; no grid ever spans two displays.

## Priority 2 — Orientation handling (rotated photos)

**Problem:** Some photos are rotated 90° (e.g. a sideways two-screen shot). No orientation
handling exists today.

7. **Honor EXIF orientation** in `load_image` via `PIL.ImageOps.exif_transpose(img)`.
8. **Auto-detect rotation** when EXIF is absent/wrong: run a quick OCR confidence probe at
   0/90/180/270 and keep the highest-confidence orientation (cheap because it can run on a
   downscaled image).
9. **Manual rotate buttons** (⟲ / ⟳ 90°) in the UI as a guaranteed fallback.

**Acceptance:** the rotated two-screen photo is auto-corrected to upright before detection.

## Priority 3 — Display-aware preprocessing (largest accuracy gain on these photos)

10. **Max-channel (HSV Value) base image** for OCR and `cell_feature` — replaces
    `ImageOps.grayscale` in `preprocess_for_ocr` and `cell_feature`.
11. **2× upscale** the warped screen (`cv2.resize`, `INTER_CUBIC`/`LANCZOS`) before OCR so
    Tesseract sees ~30–40px cap height.
12. **Denoise + deblur** variants: `cv2.fastNlMeansDenoising` for phone grain and an
    unsharp mask (`image − GaussianBlur`) for soft blur. Add as new entries in
    `preprocessing_variants` so the existing best-of-N scorer can choose them.
13. **Adaptive / Otsu binarization** (`cv2.adaptiveThreshold` or Otsu) to produce a clean
    black/white image — natural fit for monochrome MCDU output.
14. **Remove the cursor crosshair.** Detect the magenta `+` cursor (color + thin-cross
    shape) and inpaint it (`cv2.inpaint`) before OCR. This is what turns `EXEC` into
    `EXLC` and clips `SEL SPD`.
15. **Glare suppression.** Detect saturated white blooms (very high V, low saturation,
    large blob) and inpaint them before OCR.
16. **Handle colored message boxes.** Detect filled bright rectangles (blue/cyan/amber
    background with white text) and OCR them **without inverting** (correct polarity),
    since the global invert path mangles white-on-color regions.
17. **Ignore boxed entry-field outlines.** Suppress thin straight rectangle borders so they
    aren't misread as characters or fed into `recover_dash_lines` as dashes.

## Priority 4 — Recognition core (exploit the fixed grid)

18. **Per-cell recognition.** The MCDU is a fixed-pitch font in a known 13×40 lattice. Crop
    each cell at the known `calibrate_grid` coordinates and recognize it directly, via
    Tesseract `--psm 10` (single char) and/or template matching. This is the biggest
    blur-resistance win and resolves most O/0, B/8, S/5, I/1 confusions that the current
    regex-only `disambiguate_o_zero` cannot.
19. **Feed Tesseract the vocabulary.** Write `MCDU_VOCABULARY` to a temp file and pass
    `--user-words` (and `--user-patterns` for formats like `FL\d{2,3}`, `\.\d{3}`,
    `\d+/\d+`). Add `-c user_defined_dpi=300`.
20. **Force `--oem 1` (LSTM) everywhere.** `run_tesseract_tsv` currently omits `--oem`;
    make it consistent with `run_tesseract_boxes`.
21. **Upgrade the template matcher.** Replace the 16×24 binary-mask + mean-abs-distance in
    `cell_feature`/`feature_distance` with normalized cross-correlation
    (`cv2.matchTemplate`, `TM_CCOEFF_NORMED`); center/deskew the glyph before comparing.
    Optionally **pre-seed templates** by rendering the MCDU font glyph set so recognition
    works before the user has made any corrections (today `templates.json` starts empty).

## Priority 5 — Domain-knowledge validation (very high value, cheap)

22. **Per-field format validator.** These pages are highly structured. Use the row label to
    constrain and auto-correct the value, e.g.:
    - `CRZ ALT` → `FLxxx` or 4–5 digit altitude
    - `*SPD` (`ECON/LRC/SEL/RTA SPD`) → `.xxx`
    - top-left time → `HH:MM:SSz`; top-right date → `DD MON YY`
    - `OPT / MAX / RECMD` → `FLxxx` or altitude
    - `*/FUEL`, `ETA/FUEL` → `HHMMz/ nnn.n`
    When OCR output is one edit away from the field's expected format, snap to it.
23. **Capture color semantics (optional export field).** Record whether a value was magenta
    (modified/active) vs white (normal); optionally surface/export it, since that's
    operationally meaningful on these pages.

## Priority 6 — Workflow & quality

24. **Multi-photo fusion.** Allow loading 2–3 shots of the same page and majority-vote per
    cell — blur in one frame is sharp in another. Pure local logic over existing grids.
25. **Blur / quality gate.** Compute variance-of-Laplacian (`cv2.Laplacian`) on the warped
    screen and warn "image too blurry — retake" before spending an OCR pass.

---

## Suggested implementation order

1. P3 #10–13 (max-channel + upscale + denoise + binarize) — small, compounding, immediate.
2. P1 (multi-screen detect/select) — the headline feature.
3. P2 (orientation) — unblocks rotated photos.
4. P3 #14–17 (cursor / glare / message boxes / outlines) — fixes the specific corruptions.
5. P4 #18 (per-cell recognition) — the big one; then #19–21.
6. P5 #22 (field validators) — cleans up the long tail.
7. P6 (#24–25) — workflow polish.

## Testing

A set of real office photos exercises every failure mode (single screen + partial neighbor,
two full screens, 90°-rotated two-screen, glare, magenta cursor over text, colored message
boxes). After each priority lands, re-run these through the tool and confirm:
- correct screen is isolated (no two-display bleed),
- rotation auto-corrected,
- `EXEC` no longer reads `EXLC`,
- magenta values (`FL204`, `.860`, `12000`) and message-box text read correctly,
- numeric fields match their expected formats.

Keep `python -m pytest tests/test_app.py -v` green, and add tests for new pure functions
(screen scoring/selection, orientation probe, field validators, cursor/glare masks).

---

# PHASE 2 — Post-photo-testing audit backlog

Findings from running the real pipeline end-to-end on five worst-case phone photos
(June 2026 audit). Phase 1 above is fully implemented. Same hard constraints apply:
100% local, no new dependencies, OpenCV-first, PaddleOCR stays optional.

## P2-A. Bugs observed on real photos

**A1. Label duplication eats the value row (critical).** `analyze()` picks the best OCR
variant PER ROW independently; the same physical word (e.g. "CRZ ALT") can land in row N
in one variant and row N+1 in another, so it appears twice and the real value row
(magenta `12000` / `FL204`) is displaced. `mcdu_row_score` rewards the vocabulary label
on both rows. Fix: (1) after per-row selection, dedupe identical text in adjacent rows —
keep the row whose source word y-centre is closest to that row's band; (2) in
`row_candidate_score`, penalize a candidate that duplicates the chosen row above, and
add a bonus for digit-bearing tokens in the row directly under a label row.

**A2. Column drift splits tokens (e.g. `FL344 FL383 FL344` read as `FL 344 FL38 3FL3 44`).**
`estimate_grid_origin` corrects phase only, never pitch/scale, and `analyze` passes
`calibrate_grid([], ...)`. Fix: extend the warp-alignment step (`align_warp_to_grid`) to
also estimate a small per-axis SCALE from projected ink columns and bake it into the
warp itself, so backend and browser grid stay consistent.

**A3. Boxed values fuse with their frame (`12000` in an entry box unreadable).** The box
bottom edge merges with digit bases; the outline eraser skips filled boxes. Fix attempt:
morphological line removal — binarize, extract long horizontal/vertical strokes with
`cv2.getStructuringElement` + `MORPH_OPEN`, subtract them, then `MORPH_CLOSE` to re-seal
digit gaps; apply locally per detected box. Accept that worst-case blur may stay
unreadable.

**A4. `fuse_grids` tie-break is nondeterministic.** `max(set(votes), key=votes.count)`
on a 1-1 conflict depends on set order. Make it deterministic (prefer first grid) or
mark the cell as a conflict; optionally accept confidence grids in `/api/fuse-grids`
and weight votes.

**A5. Title row misreads (`MOD M.850 D/D` -> `MA Y4.890D/0`).** The inverse-video MOD
chip (black text on white) breaks polarity. Fix: detect small inverse-video chips and
OCR them un-inverted (reuse the message-box machinery), and add a page-title fuzzy
validator — row 0 comes from a small closed vocabulary (ACT/MOD + page names), so
edit-distance snapping on row 0 is safe.

## P2-B. Accuracy upgrades

**B1. Pre-seeded glyph atlas as a recognition engine (top recommendation).** The font is
fixed and the grid known. Render the MCDU-style font (B612 Mono is a close open match)
for the whitelist charset at several blur/thickness levels into a glyph atlas; classify
each cell by NCC (`cv2.matchTemplate`) against it. `classify_from_templates` exists but
ships empty. Fuse atlas results as a third engine beside Tesseract/Paddle.

**B2. (Optional, heavy) Fine-tune Tesseract LSTM on the MCDU font** with synthetic
renders + photo augmentation (tesstrain). Train on the dev machine; office laptop only
needs the resulting .traineddata file.

**B3. Fuzzy phrase normalization.** `MCDU_PHRASE_REPLACEMENTS` is exact-match, so
`ENGEOUT>`/`<COSPD`/`LRCSPO` survive. Match with edit distance 1 against the phrase and
vocabulary lists (difflib already imported).

**B4. Space preservation via char boxes.** Word-level placement compacts spaces
(`ENG OUT` -> `ENGEOUT`). When char boxes are available for a word, place characters by
their individual x positions instead of packing from the word's left edge.

## P2-C. Efficiency

**C1. Parallelize Tesseract subprocess calls** in `analyze` (variants x PSMs, box passes,
row strips) with a ThreadPoolExecutor — currently all serial; expect 3-5x wall-time win.

**C2. Prune the variant x PSM matrix.** Early-exit losing variants after 2 PSMs; add
per-stage timing logs to learn which variant/PSM combinations actually win.

**C3. Skip redundant passes** (e.g. skip per-row strip OCR when the word grid is already
high-confidence).

## P2-D. Housekeeping

**D1.** `cleanup_exports` only deletes `.png` — exported `.docx` accumulate forever.
**D2.** `whole_grid_focused_recheck` output skips `validate_field_formats`.
**D3.** Blur threshold 80 is loosely calibrated; add a "marginal" band (80-150).
**D4.** (Optional) split `app.py` (~3,400 lines) into modules.

## P2-E. Testing gap

No end-to-end regression coverage. Build a synthetic MCDU screen renderer in the test
suite: render a known 13x40 grid in a monospace font on a black screen (optionally
warp/blur), run the pipeline, assert the grid reads back. Optionally keep 2-3 heavily
downscaled real photos (~200 KB) as fixtures. This would have caught both the
orientation bug and A1.

## P2-F. Finding from the regression harness (Phase 3 candidate)

Building the P2-E synthetic harness surfaced a real weakness: the per-row
variant selection in analyze() confuses content between rows when rows have low
mcdu_row_score (short numerics, single words) — high-scoring content from an
adjacent row can win the wrong row. The harness grid was set to compound
high-score strings to get a stable 100% baseline, so it does NOT guard against
this. Worth a Phase 3 item: make row selection position-aware (a candidate's
word y-centres must fall within the row's band) so it cannot borrow content
from neighbouring rows. This is the same class of issue as A1 (label
duplication) and likely explains several garbled rows on the real photos.

## P3 note — pre-existing double-letter limitation (found during Prompt B)

merge_ocr_grids' same_nearby guard drops a char-source character when the word
grid has the same character within ±2 columns. This correctly suppresses
cross-source shift duplicates, but also drops a LEGITIMATE double letter when
the word source read it as single and the char source supplied the missing one
(e.g. word "AL" + char "L" -> "AL", not "ALL"). Pre-existing (not introduced by
Prompt B). Low frequency. A position-aware fix (compare source word x-extents)
would resolve it; fold into P2-F if that round happens.
