"""P2-E: Synthetic MCDU regression harness.

Renders known grids in a monospace font, drives the full analyze() pipeline,
and asserts character-readback accuracy above documented thresholds.

Each test calls real Tesseract OCR and takes 10-40 seconds.  Run selectively:
  python -m pytest tests/test_regression.py -v -s
"""

import base64
import io
import unittest
from pathlib import Path

import numpy as np
from PIL import Image, ImageDraw, ImageFont, ImageFilter

import app

# ---------------------------------------------------------------------------
# Geometry — matches production pipeline constants
# ---------------------------------------------------------------------------
_SW = app.SCREEN_W      # 1600 px
_SH = 1300              # within [MIN_SCREEN_H, MAX_SCREEN_H]; cell_h = 100 px
_CW = _SW / app.COLS    # 40.0 px wide per column
_CH = _SH / app.ROWS    # 100.0 px tall per row
_FIRST = app.FIRST_DATA_COL   # 1
_LAST  = app.LAST_DATA_COL    # 38
_DATA_COLS = _LAST - _FIRST + 1  # 38

# ---------------------------------------------------------------------------
# Fixed test-grid content.
# 38-char strings (index 0 → col 1 … index 37 → col 38).
# Space = empty cell.  Only characters from OCR_WHITELIST are used.
# Phrases are written in their normalised form so that pipeline corrections
# produce a byte-identical result.
# ---------------------------------------------------------------------------
_GRID_ROWS: list[str] = [
    "ACT RTA CRZ",  # row 0  — title; snap_title_row corrects OCR errors
    "RECMDSPD",     # row 1  — RECMD+SPD (two high-score WORD_HINTS, score ≈ 43)
    "LEGSFUEL",     # row 2  — LEGS+FUEL
    "MAXSTEP",      # row 3  — MAX+STEP
    "OPTALT",       # row 4  — OPT+ALT
    "ECONRTE",      # row 5  — ECON+RTE
    "LRCOUT",       # row 6  — LRC+OUT
    "SELSTEP",      # row 7  — SEL+STEP
    "ENGMAX",       # row 8  — ENG+MAX
    "TOFUEL",       # row 9  — TO+FUEL
    "STEPSPD",      # row 10 — STEP+SPD
    "OPTLEGS",      # row 11 — OPT+LEGS
    "ECONALT",      # row 12 — ECON+ALT
]

# ---------------------------------------------------------------------------
# Helpers
# ---------------------------------------------------------------------------

def _find_font(paths: list[str]) -> str | None:
    for p in paths:
        if Path(p).exists():
            return p
    return None


def _atlas_font() -> str | None:
    """Return the first available font from the atlas build list."""
    return _find_font(app._ATLAS_FONT_PATHS)


def _alt_font() -> str | None:
    """Return a monospace font NOT in the atlas build list (cross-font test)."""
    candidates = [
        # SF Mono: common macOS terminal font; not in atlas paths
        "/System/Library/Fonts/SFNSMono.ttf",
        "/System/Library/Fonts/SFNSMonoItalic.ttf",
        # Ubuntu Mono Bold: common Linux font (regular is in atlas paths on some distros)
        "/usr/share/fonts/truetype/ubuntu/UbuntuMono-B.ttf",
    ]
    atlas_set = set(app._ATLAS_FONT_PATHS)
    return _find_font([p for p in candidates if p not in atlas_set])


def _make_image(
    rows: list[str],
    font_path: str | None,
    font_size: int = 55,
    sw: int = _SW,
    sh: int = _SH,
) -> Image.Image:
    """Render row strings onto a black MCDU-sized canvas.

    Character i in row_text r occupies data column (FIRST + i), row r.
    Space characters produce blank (empty) cells.
    """
    img = Image.new("RGB", (sw, sh), (0, 0, 0))
    draw = ImageDraw.Draw(img)
    cw = sw / app.COLS
    ch = sh / app.ROWS

    try:
        font: ImageFont.ImageFont | ImageFont.FreeTypeFont = (
            ImageFont.truetype(font_path, font_size)
            if font_path and Path(font_path).exists()
            else ImageFont.load_default()
        )
    except Exception:
        font = ImageFont.load_default()

    for r, row_text in enumerate(rows[: app.ROWS]):
        for i, glyph in enumerate(row_text[:_DATA_COLS]):
            col = _FIRST + i
            if col > _LAST or glyph in (" ", ""):
                continue
            cx = col * cw + cw / 2
            cy = r * ch + ch / 2
            try:
                bb = draw.textbbox((0, 0), glyph, font=font)
                tx = cx - (bb[2] - bb[0]) / 2 - bb[0]
                ty = cy - (bb[3] - bb[1]) / 2 - bb[1]
            except Exception:
                tx, ty = cx, cy
            draw.text((tx, ty), glyph, fill=(255, 255, 255), font=font)

    return img


def _expected_grid(rows: list[str]) -> list[list[str]]:
    """Convert row strings to the canonical expected grid (empty cells for spaces)."""
    grid = app.empty_grid()
    for r, row_text in enumerate(rows[: app.ROWS]):
        for i, ch in enumerate(row_text[:_DATA_COLS]):
            col = _FIRST + i
            if col <= _LAST and ch and ch != " ":
                grid[r][col] = ch
    return app.normalize_grid_guards(grid)


def _to_data_url(img: Image.Image) -> str:
    buf = io.BytesIO()
    img.save(buf, format="PNG")
    return f"data:image/png;base64,{base64.b64encode(buf.getvalue()).decode()}"


def _full_corners(w: int, h: int) -> list[dict]:
    return [
        {"x": 0.0, "y": 0.0},
        {"x": float(w), "y": 0.0},
        {"x": float(w), "y": float(h)},
        {"x": 0.0, "y": float(h)},
    ]


def _run_analyze(img: Image.Image, **extra) -> dict:
    w, h = img.size
    return app.analyze(
        {
            "image": _to_data_url(img),
            "corners": _full_corners(w, h),
            "hybridOcr": False,
            **extra,
        }
    )


def _accuracy(
    result: list[list[str]], expected: list[list[str]]
) -> tuple[float, int, int]:
    """Return (accuracy, correct, total) counting non-empty expected cells only."""
    total = correct = 0
    for r in range(app.ROWS):
        for c in range(_FIRST, _LAST + 1):
            ex = expected[r][c]
            if not ex:
                continue
            total += 1
            if result[r][c] == ex:
                correct += 1
    return correct / max(1, total), correct, total


def _render_char_cell(
    ch: str,
    font: ImageFont.FreeTypeFont,
    cell_w: int = int(_CW),
    cell_h: int = int(_CH),
) -> np.ndarray:
    """Render one character bright-on-dark and return as uint8 ndarray."""
    canvas = Image.new("L", (cell_w, cell_h), 0)
    draw = ImageDraw.Draw(canvas)
    cx, cy = cell_w / 2, cell_h / 2
    try:
        bb = draw.textbbox((0, 0), ch, font=font)
        tx = cx - (bb[2] - bb[0]) / 2 - bb[0]
        ty = cy - (bb[3] - bb[1]) / 2 - bb[1]
    except Exception:
        tx, ty = 0.0, 0.0
    draw.text((tx, ty), ch, fill=255, font=font)
    return np.asarray(canvas, dtype=np.uint8)


# ===========================================================================
class SyntheticRegressionTests(unittest.TestCase):
    """End-to-end accuracy tests driven from synthetic MCDU screen renders.

    Slow: each test invokes Tesseract (10–40 s).
    """

    @classmethod
    def setUpClass(cls) -> None:
        # Ensure the exports directory exists so analyze() can save previews.
        app.ensure_dirs()

    # ------------------------------------------------------------------
    # Test 1 — Clean render
    # ------------------------------------------------------------------
    def test_clean_render_high_accuracy(self) -> None:
        """Clean render (no blur, no distortion) must achieve >= 90% readback."""
        font = _atlas_font()
        if font is None:
            self.skipTest("No monospace atlas font found on this system")

        img = _make_image(_GRID_ROWS, font)
        expected = _expected_grid(_GRID_ROWS)
        result = _run_analyze(img)
        grid: list[list[str]] = result["grid"]

        acc, correct, total = _accuracy(grid, expected)
        print(f"\n[P2-E test1] clean-render accuracy: {correct}/{total} = {acc:.1%}")
        self.assertGreaterEqual(
            acc,
            0.90,
            f"Clean-render accuracy {acc:.1%} ({correct}/{total}) is below the 90% threshold",
        )

    # ------------------------------------------------------------------
    # Test 2 — Blurred + perspective-warped render
    # ------------------------------------------------------------------
    def test_blurred_warped_render_passing_accuracy(self) -> None:
        """Mild blur + perspective warp must still achieve >= 70% readback."""
        font = _atlas_font()
        if font is None:
            self.skipTest("No monospace atlas font found on this system")

        clean = _make_image(_GRID_ROWS, font)
        blurred = clean.filter(ImageFilter.GaussianBlur(radius=2))

        # Slight keystone: shift each corner ~20–25 px inward/outward
        w, h = blurred.size
        warp_corners = [
            {"x": 25.0,      "y": 15.0      },   # TL
            {"x": w - 20.0,  "y": 22.0      },   # TR
            {"x": w - 15.0,  "y": h - 18.0  },   # BR
            {"x": 18.0,      "y": h - 12.0  },   # BL
        ]
        expected = _expected_grid(_GRID_ROWS)
        result = app.analyze(
            {
                "image": _to_data_url(blurred),
                "corners": warp_corners,
                "hybridOcr": False,
            }
        )
        grid: list[list[str]] = result["grid"]

        acc, correct, total = _accuracy(grid, expected)
        print(f"\n[P2-E test2] blurred+warped accuracy: {correct}/{total} = {acc:.1%}")
        self.assertGreaterEqual(
            acc,
            0.70,
            f"Blurred+warped accuracy {acc:.1%} ({correct}/{total}) is below the 70% threshold",
        )

    # ------------------------------------------------------------------
    # Test 3 — 90° rotated render
    # ------------------------------------------------------------------
    def test_rotated_90_probe_and_recovery(self) -> None:
        """probe_orientation must return 90 for a 90°-CW-rotated render; post-correction >= 70%."""
        font = _atlas_font()
        if font is None:
            self.skipTest("No monospace atlas font found on this system")

        clean = _make_image(_GRID_ROWS, font)

        # Simulate phone held 90° CW (top of screen pointing left in the photo).
        # PIL rotate(-90) = 90° CW → portrait image 1300 × 1600.
        rotated = clean.rotate(-90, expand=True)

        correction = app.probe_orientation(rotated)
        print(f"\n[P2-E test3] probe_orientation returned: {correction}°")
        self.assertEqual(
            correction,
            90,
            f"Expected probe_orientation=90 for a 90°-CW render, got {correction}",
        )

        # Apply the orientation correction and run the full pipeline
        corrected = rotated.rotate(correction, expand=True)  # → 1600 × 1300
        expected = _expected_grid(_GRID_ROWS)
        result = _run_analyze(corrected)
        grid: list[list[str]] = result["grid"]

        acc, correct, total = _accuracy(grid, expected)
        print(f"[P2-E test3] post-correction accuracy: {correct}/{total} = {acc:.1%}")
        self.assertGreaterEqual(
            acc,
            0.70,
            f"Post-correction accuracy {acc:.1%} ({correct}/{total}) is below the 70% threshold",
        )

    # ------------------------------------------------------------------
    # Test 4 — Separator dash row
    # ------------------------------------------------------------------
    def test_separator_dash_row_recovered(self) -> None:
        """A thin bright horizontal line in row 6 must be recovered as dashes (>= 95% cols)."""
        font = _atlas_font()
        if font is None:
            self.skipTest("No monospace atlas font found on this system")

        # Render normal content except row 6, which gets no characters
        rows_no6 = list(_GRID_ROWS)
        rows_no6[6] = ""
        img = _make_image(rows_no6, font)

        # Draw a thin (3 px) bright horizontal line through the vertical midpoint
        # of row 6, spanning exactly the data columns (1..38).
        draw = ImageDraw.Draw(img)
        cy = int(6 * _CH + _CH / 2)
        x0 = int(_FIRST * _CW)
        x1 = int((_LAST + 1) * _CW)
        draw.line([(x0, cy), (x1, cy)], fill=(255, 255, 255), width=3)

        result = _run_analyze(img)
        grid: list[list[str]] = result["grid"]

        dash_cols = sum(1 for c in range(_FIRST, _LAST + 1) if grid[6][c] == "-")
        total_cols = _DATA_COLS
        ratio = dash_cols / total_cols
        print(f"\n[P2-E test4] dash-row coverage: {dash_cols}/{total_cols} = {ratio:.1%}")
        self.assertGreaterEqual(
            ratio,
            0.95,
            f"Dash-row coverage {ratio:.1%} ({dash_cols}/{total_cols}) is below 95%",
        )

    # ------------------------------------------------------------------
    # Test 5 — Glyph-atlas accuracy guard (cross-font)
    # ------------------------------------------------------------------
    def test_atlas_alnum_accuracy_on_alt_font(self) -> None:
        """Atlas must classify >= 85% of alnum chars rendered in a non-atlas font.

        Always measures and reports accuracy.  Skips (not fails) when
        ATLAS_ENABLED=False so the CI stays green while the atlas is disabled.
        """
        atlas = app._get_glyph_atlas()
        if not atlas:
            self.skipTest("No atlas fonts available on this system — cannot build atlas")

        alt_path = _alt_font()
        if alt_path is None:
            self.skipTest("No non-atlas monospace font found for cross-font test")

        try:
            font = ImageFont.truetype(alt_path, 64)
        except Exception as exc:
            self.skipTest(f"Could not load alt font {alt_path}: {exc}")

        alnum_chars = [c for c in app.OCR_WHITELIST if c.isalnum()]
        correct = 0
        wrong: list[str] = []

        for ch in alnum_chars:
            cell_arr = _render_char_cell(ch, font)
            res = app._atlas_classify_patch(cell_arr, atlas)
            if res is not None and res[0] == ch:
                correct += 1
            else:
                got = res[0] if res is not None else "—"
                wrong.append(f"{ch}→{got}")

        total = len(alnum_chars)
        acc = correct / max(1, total)

        print(
            f"\n[P2-E test5] atlas cross-font ({Path(alt_path).name}):"
            f" {correct}/{total} = {acc:.1%}"
        )
        if wrong:
            print(f"[P2-E test5] misclassified: {', '.join(wrong[:24])}")

        if not app.ATLAS_ENABLED:
            self.skipTest(
                f"ATLAS_ENABLED=False; measured cross-font accuracy="
                f"{acc:.1%} ({correct}/{total}) — informational only"
            )

        self.assertGreaterEqual(
            acc,
            0.85,
            f"Atlas cross-font accuracy {acc:.1%} ({correct}/{total}) is below 85% threshold",
        )
