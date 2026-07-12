import base64
import io
import math
import os
import tempfile
import time
import unittest
from pathlib import Path
from unittest.mock import patch

import app
import numpy as np
from PIL import Image, ImageDraw, ImageFilter, ImageOps


def grid_with_row(text: str, row: int = 0) -> list[list[str]]:
    grid = app.empty_grid()
    padded = text[: app.COLS].ljust(app.COLS)
    grid[row] = ["" if char == " " else char for char in padded]
    return app.normalize_grid_guards(grid)


class OcrNormalizationTests(unittest.TestCase):
    def test_speed_altitude_slash_is_not_changed_to_degree(self) -> None:
        self.assertEqual(app.clean_ocr_text("180/4530"), "180/4530")
        self.assertEqual(app.clean_ocr_text("100/280"), "100/280")

    def test_heading_position_format_recovers_degree(self) -> None:
        self.assertEqual(app.clean_ocr_text("000/0.0NM"), "000°/0.0NM")
        self.assertEqual(app.clean_ocr_text("000O/0.0NM"), "000°/0.0NM")

    def test_track_heading_recovers_degree(self) -> None:
        # LEGS/RTE page: "217°TRK" — degree dropped or misread before TRK.
        self.assertEqual(app.clean_ocr_text("217TRK"), "217°TRK")
        self.assertEqual(app.clean_ocr_text("217oTRK"), "217°TRK")
        self.assertEqual(app.clean_ocr_text("217°TRK"), "217°TRK")

    def test_standalone_course_recovers_degree(self) -> None:
        # LEGS page course column: "268°" with the degree read as the letter o/O.
        self.assertEqual(app.clean_ocr_text("268o"), "268°")
        self.assertEqual(app.clean_ocr_text("268O"), "268°")
        self.assertEqual(app.clean_ocr_text("268°"), "268°")

    def test_glidepath_angle_recovers_degree(self) -> None:
        # Descent legs: "GP 3.00°" — degree misread as letter o after the angle.
        self.assertEqual(app.clean_ocr_text("3.00o"), "3.00°")
        self.assertEqual(app.clean_ocr_text("3.00°"), "3.00°")

    def test_plain_numbers_are_not_given_a_degree(self) -> None:
        # Speeds/altitudes/distances must never gain a spurious degree.
        self.assertEqual(app.clean_ocr_text("180"), "180")
        self.assertEqual(app.clean_ocr_text("2680"), "2680")
        self.assertEqual(app.clean_ocr_text("FL204"), "FL204")
        self.assertEqual(app.clean_ocr_text("60NM"), "60NM")

    def test_common_mcdu_compounds_recover_spacing(self) -> None:
        self.assertEqual(app.normalize_mcdu_phrase("TOFL204"), "TO FL204")
        self.assertEqual(app.normalize_mcdu_phrase("KBFIETA/FUEL"), "KBFI ETA/FUEL")
        self.assertEqual(app.normalize_mcdu_phrase("TOT/D"), "TO T/D")


class DisplayGeometryTests(unittest.TestCase):
    def test_grid_origin_moves_boundaries_into_character_gaps(self) -> None:
        image = app.Image.new("RGB", (800, 260), "black")
        draw = ImageDraw.Draw(image)
        target_x = -6.0
        target_y = 3.0
        cell_w = image.width / app.COLS
        cell_h = image.height / app.ROWS
        for row in range(app.ROWS):
            for col in range(1, app.COLS - 1, 3):
                center_x = target_x + (col + 0.5) * cell_w
                center_y = target_y + (row + 0.5) * cell_h
                draw.rectangle(
                    (center_x - 8, center_y - 8, center_x + 8, center_y + 8),
                    fill="white",
                )
        origin_x, origin_y, scale_x, scale_y = app.estimate_grid_origin(image)
        self.assertAlmostEqual(origin_x, target_x, delta=1.5)
        self.assertAlmostEqual(origin_y, target_y, delta=1.5)
        self.assertAlmostEqual(scale_x, 1.0, places=6)
        self.assertAlmostEqual(scale_y, 1.0, places=6)

    def test_border_lines_refine_a_skewed_phone_photo(self) -> None:
        try:
            import cv2  # noqa: F401
        except ImportError:
            self.skipTest("OpenCV is optional at runtime")
        image = app.Image.new("RGB", (800, 600), (130, 130, 130))
        draw = ImageDraw.Draw(image)
        expected = [(130, 100), (690, 125), (650, 500), (100, 470)]
        draw.polygon(expected, fill=(5, 5, 5))
        draw.line(expected + [expected[0]], fill=(235, 235, 235), width=7, joint="curve")
        rough = [
            {"x": 100.0, "y": 100.0},
            {"x": 690.0, "y": 100.0},
            {"x": 690.0, "y": 500.0},
            {"x": 100.0, "y": 500.0},
        ]
        refined = app.refine_display_corners(image, rough)
        for actual, target in zip(refined, expected):
            self.assertLess(math.dist((actual["x"], actual["y"]), target), 7.0)

    def test_collapsed_bottom_left_corner_is_reconstructed(self) -> None:
        points = [(316.0, 317.0), (2426.0, 316.0), (2404.0, 1945.0), (637.0, 1944.0)]
        corrected = app.regularize_quadrilateral(points)
        self.assertEqual(corrected[:3], points[:3])
        self.assertAlmostEqual(corrected[3][0], 294.0)
        self.assertAlmostEqual(corrected[3][1], 1946.0)

    def test_normal_perspective_is_preserved(self) -> None:
        points = [(692.0, 486.0), (2428.0, 520.0), (2393.0, 1887.0), (659.0, 1860.0)]
        self.assertEqual(app.regularize_quadrilateral(points), points)

    def test_ocr_geometry_uses_the_exact_visible_grid_pitch(self) -> None:
        boxes = [
            {"left": col * 40 + 10, "top": row * 100 + 20, "width": 20, "height": 50}
            for row in range(6)
            for col in range(6)
        ]
        geometry = app.calibrate_grid(boxes, (1600, 1300))
        self.assertEqual(geometry["cell_w"], 40.0)
        self.assertEqual(geometry["cell_h"], 100.0)

    def test_aligned_grid_geometry_has_no_second_ocr_offset(self) -> None:
        geometry = app.calibrate_grid([], (1600, 1300))
        self.assertEqual(geometry["origin_x"], 0.0)
        self.assertEqual(geometry["origin_y"], 0.0)


class PitchCorrectionTests(unittest.TestCase):
    """Tests for A2: per-axis scale / pitch correction in estimate_grid_origin."""

    def _make_grid_image(
        self,
        width: int,
        height: int,
        pitch_x: float,
        pitch_y: float,
        char_fill_ratio: float = 0.55,
    ) -> "app.Image.Image":
        """Render solid rectangles at every data cell at the given pitches."""
        image = app.Image.new("RGB", (width, height), "black")
        draw = ImageDraw.Draw(image)
        char_w = pitch_x * char_fill_ratio
        char_h = pitch_y * char_fill_ratio
        for row in range(app.ROWS):
            for col in range(app.FIRST_DATA_COL, app.LAST_DATA_COL + 1):
                cx = (col + 0.5) * pitch_x
                cy = (row + 0.5) * pitch_y
                draw.rectangle(
                    (cx - char_w / 2, cy - char_h / 2, cx + char_w / 2, cy + char_h / 2),
                    fill="white",
                )
        return image

    def test_well_aligned_image_returns_unit_scale(self) -> None:
        """A correctly-pitched image must produce scale == 1.0 on both axes."""
        width, height = app.SCREEN_W, 520
        pitch_x = width / app.COLS   # exactly nominal
        pitch_y = height / app.ROWS  # exactly nominal
        image = self._make_grid_image(width, height, pitch_x, pitch_y)
        _, _, scale_x, scale_y = app.estimate_grid_origin(image)
        self.assertAlmostEqual(scale_x, 1.0, places=6,
                               msg=f"Expected scale_x=1.0, got {scale_x}")
        self.assertAlmostEqual(scale_y, 1.0, places=6,
                               msg=f"Expected scale_y=1.0, got {scale_y}")

    def test_well_aligned_image_is_not_transformed(self) -> None:
        """align_warp_to_grid must return the original image object for a well-aligned image."""
        width, height = app.SCREEN_W, 520
        pitch_x = width / app.COLS
        pitch_y = height / app.ROWS
        image = self._make_grid_image(width, height, pitch_x, pitch_y)
        result, offsets = app.align_warp_to_grid(image)
        self.assertIs(result, image,
                      "No transform should be applied when pitch and phase are both correct")
        self.assertEqual(offsets, (0.0, 0.0))

    def test_wrong_pitch_image_recovers_scale(self) -> None:
        """estimate_grid_origin must detect and return a pitch correction for a 1.5% pitch error."""
        drift = 0.015  # 1.5% — one of the linspace(0.97, 1.03, 13) grid points
        width, height = app.SCREEN_W, 520
        nominal_pitch_x = width / app.COLS    # 40.0
        pitch_x = nominal_pitch_x * (1 + drift)  # 40.6
        pitch_y = height / app.ROWS           # 40.0 (no y-axis drift)
        # fill_ratio=0.88 → gap width ≈ 4.9 px; a 0.5-step scale error (1.01 vs 1.015)
        # shifts boundary ~3.8 px at the image edges, pushing it outside the gap and
        # producing a noticeably higher ink score, making 1.015 the clear winner.
        image = self._make_grid_image(width, height, pitch_x, pitch_y, char_fill_ratio=0.88)
        _, _, scale_x, scale_y = app.estimate_grid_origin(image)
        self.assertAlmostEqual(
            scale_x, 1.0 + drift, delta=0.003,
            msg=f"Expected scale_x≈{1+drift:.3f}, got {scale_x:.4f}",
        )
        # Y axis has no drift — must stay at 1.0
        self.assertAlmostEqual(scale_y, 1.0, places=4,
                               msg=f"Expected scale_y=1.0, got {scale_y}")


class CorrectionSafetyTests(unittest.TestCase):
    def test_similar_dynamic_row_is_not_replaced(self) -> None:
        source = grid_with_row(" FL204  0458Z/ 60NM")
        corrected = grid_with_row(" FL204  0458Z/ 60NM")
        new_grid = grid_with_row(" FL205  0459Z/ 60NM")
        key = app.correction_key(0, app.grid_row_from_payload(source, 0))
        with patch.object(app, "load_corrections", return_value={key: app.grid_row_from_payload(corrected, 0)}):
            result = app.apply_corrections(new_grid)
        self.assertEqual(result[0], new_grid[0])

    def test_exact_row_correction_still_applies(self) -> None:
        source = grid_with_row(" POS REE")
        corrected = grid_with_row(" POS REF")
        key = app.correction_key(0, app.grid_row_from_payload(source, 0))
        with patch.object(app, "load_corrections", return_value={key: app.grid_row_from_payload(corrected, 0)}):
            result = app.apply_corrections(source)
        self.assertEqual(app.grid_row_from_payload(result, 0), app.grid_row_from_payload(corrected, 0))


class OcrMergeTests(unittest.TestCase):
    def test_character_boxes_fill_blank_cells_without_overwriting_words(self) -> None:
        words = app.empty_grid()
        word_confidence = app.empty_confidence_grid()
        chars = app.empty_grid()
        char_confidence = app.empty_confidence_grid()
        words[0][2] = "R"
        words[0][4] = "F"
        word_confidence[0][2] = 0.91
        word_confidence[0][4] = 0.91
        chars[0][2] = "P"
        char_confidence[0][2] = 0.75
        chars[0][3] = "E"
        char_confidence[0][3] = 0.75
        merged, confidence = app.merge_ocr_grids(words, word_confidence, chars, char_confidence)
        self.assertEqual(merged[0][2], "R")
        self.assertEqual(merged[0][3], "E")
        self.assertLess(confidence[0][2], 0.62)

    def test_character_merge_does_not_duplicate_nearby_arrow(self) -> None:
        words = app.empty_grid()
        word_confidence = app.empty_confidence_grid()
        chars = app.empty_grid()
        char_confidence = app.empty_confidence_grid()
        words[0][37], word_confidence[0][37] = ">", 0.95
        chars[0][38], char_confidence[0][38] = ">", 0.75
        merged, _ = app.merge_ocr_grids(words, word_confidence, chars, char_confidence)
        self.assertEqual(merged[0][37], ">")
        self.assertEqual(merged[0][38], "")

    def test_character_merge_rejects_standalone_dash_and_arrow(self) -> None:
        words = app.empty_grid()
        word_confidence = app.empty_confidence_grid()
        chars = app.empty_grid()
        char_confidence = app.empty_confidence_grid()
        chars[0][4], char_confidence[0][4] = "-", 0.8
        chars[1][10], char_confidence[1][10] = ">", 0.8
        merged, _ = app.merge_ocr_grids(words, word_confidence, chars, char_confidence)
        self.assertEqual(merged[0][4], "")
        self.assertEqual(merged[1][10], "")

    def test_hybrid_fusion_rewards_agreement_and_marks_conflicts(self) -> None:
        primary = app.empty_grid()
        primary_confidence = app.empty_confidence_grid()
        secondary = app.empty_grid()
        secondary_confidence = app.empty_confidence_grid()
        primary[0][1], primary_confidence[0][1] = "A", 0.7
        secondary[0][1], secondary_confidence[0][1] = "A", 0.9
        secondary[0][2], secondary_confidence[0][2] = "B", 0.9
        primary[0][3], primary_confidence[0][3] = "O", 0.9
        secondary[0][3], secondary_confidence[0][3] = "0", 0.9

        fused, confidence, summary = app.fuse_engine_grids(
            primary,
            primary_confidence,
            secondary,
            secondary_confidence,
        )

        self.assertEqual(fused[0][1], "A")
        self.assertGreater(confidence[0][1], 0.8)
        self.assertEqual(fused[0][2], "B")
        self.assertEqual(fused[0][3], "0")
        self.assertLessEqual(confidence[0][3], 0.45)
        self.assertEqual(
            summary,
            {"rowsSelected": 1, "agreements": 1, "blanksFilled": 0, "disagreements": 1},
        )

    def test_paddle_array_output_is_converted_to_word_boxes(self) -> None:
        words = app.paddle_words_from_payload(
            {
                "rec_texts": ["ENG OUT"],
                "rec_scores": np.asarray([0.92], dtype=np.float32),
                "dt_polys": np.asarray([[[10, 20], [90, 20], [90, 50], [10, 50]]], dtype=np.float32),
            }
        )
        self.assertEqual(len(words), 1)
        self.assertEqual(words[0]["text"], "ENG OUT")
        self.assertAlmostEqual(words[0]["conf"], 92.0, places=3)
        self.assertEqual((words[0]["left"], words[0]["top"], words[0]["width"], words[0]["height"]), (10, 20, 80, 30))

    def test_edge_pass_adds_unique_bottom_text_without_replacing_top_text(self) -> None:
        primary = [{"text": "MOD LRC D/D", "left": 500, "top": 0, "width": 460, "height": 70}]
        edge = [
            {"text": "MUD LRC D/U", "left": 530, "top": 22, "width": 450, "height": 70},
            {"text": "<ERASE", "left": 110, "top": 1205, "width": 230, "height": 60},
        ]
        merged = app.merge_edge_words(primary, edge, (1600, 1236), 30)
        self.assertEqual([word["text"] for word in merged], ["MOD LRC D/D", "<ERASE"])

    def test_corner_words_are_mapped_back_to_the_full_display(self) -> None:
        mapped = app.remap_crop_words(
            [{"text": "LRC>", "left": 440, "top": 70, "width": 320, "height": 100}],
            (1000, 1080),
            2.0,
            40,
        )
        self.assertEqual((mapped[0]["left"], mapped[0]["top"]), (1200, 1095))
        self.assertEqual((mapped[0]["width"], mapped[0]["height"]), (160, 50))

    def test_unique_focused_word_is_added_away_from_display_edges(self) -> None:
        primary = [{"text": "STEP", "left": 80, "top": 690, "width": 160, "height": 70}]
        focused = [{"text": "0", "left": 60, "top": 790, "width": 45, "height": 60}]
        merged = app.merge_unique_words(primary, focused)
        self.assertEqual([word["text"] for word in merged], ["STEP", "0"])

    def test_separator_normalization_removes_isolated_ocr_symbol(self) -> None:
        grid = app.empty_grid()
        for col in range(1, 35):
            grid[9][col] = "-"
        grid[9][21] = ">"
        image = app.Image.new("RGB", (1600, 1300), "black")
        recovered = app.recover_dash_lines(grid, image)
        self.assertTrue(all(recovered[9][col] == "-" for col in range(1, 35)))

    def test_full_width_separator_line_overrides_ocr_garbage(self) -> None:
        # A thin full-width bright line is the MCDU divider; OCR often misreads it
        # as random characters. The image-based pass must force the whole data row
        # to dashes, clearing the garbage.
        grid = app.empty_grid()
        for col, ch in zip(range(1, 12), "PSTEGCNGING"):
            grid[9][col] = ch
        image = app.Image.new("RGB", (1600, 1300), "black")
        draw = ImageDraw.Draw(image)
        cell_h = 1300 / app.ROWS
        y = int((9 + 0.5) * cell_h)
        draw.line((1 * 40, y, 39 * 40, y), fill="white", width=4)
        recovered = app.recover_dash_lines(grid, image)
        self.assertTrue(
            all(recovered[9][col] == "-" for col in range(app.FIRST_DATA_COL, app.LAST_DATA_COL + 1)),
            "full-width separator row should become all dashes",
        )

    def test_dash_recovery_does_not_fill_between_text_blocks(self) -> None:
        grid = app.empty_grid()
        grid[0][2] = "A"
        grid[0][35] = "B"
        image = app.Image.new("RGB", (1600, 1300), "black")
        draw = ImageDraw.Draw(image)
        draw.line((10 * 40, 50, 31 * 40, 50), fill="white", width=5)
        recovered = app.recover_dash_lines(grid, image)
        self.assertFalse(any(recovered[0][col] == "-" for col in range(3, 35)))

    def test_balanced_recheck_does_not_fill_intentional_word_spacing(self) -> None:
        grid = app.empty_grid()
        grid[0][5] = "A"
        grid[0][8] = "B"
        image = app.Image.new("RGB", (1600, 1300), "black")
        with patch.object(app, "focused_grid_read", return_value="     AX B".ljust(38)):
            refined, _ = app.whole_grid_focused_recheck(image, grid, "balanced")
        self.assertEqual(refined[0][6], "")


class TemplateLearningTests(unittest.TestCase):
    def test_only_changed_cells_are_learned(self) -> None:
        source = grid_with_row(" ABC")
        corrected = grid_with_row(" ADC")
        with tempfile.TemporaryDirectory() as temp_dir:
            template_path = Path(temp_dir) / "templates.json"
            with (
                patch.object(app, "TEMPLATES", template_path),
                patch.object(app, "load_image"),
                patch.object(app, "warp_screen"),
                patch.object(app, "cell_feature", return_value=[1.0]),
                patch.object(app, "load_templates", return_value={}),
                patch.object(app, "save_templates") as save,
            ):
                result = app.remember_templates(
                    {
                        "image": "unused",
                        "corners": [{}, {}, {}, {}],
                        "sourceGrid": source,
                        "grid": corrected,
                    }
                )
        self.assertEqual(result["learned"], 1)
        saved_templates = save.call_args.args[0]
        self.assertEqual(set(saved_templates), {"D"})

    def test_deleted_dash_learns_a_blank_visual_template(self) -> None:
        source = grid_with_row(" -")
        corrected = grid_with_row("")
        with (
            patch.object(app, "load_image"),
            patch.object(app, "warp_screen"),
            patch.object(app, "cell_feature", return_value=[1.0]),
            patch.object(app, "load_templates", return_value={}),
            patch.object(app, "save_templates") as save,
        ):
            result = app.remember_templates(
                {"image": "unused", "corners": [{}, {}, {}, {}], "sourceGrid": source, "grid": corrected}
            )
        self.assertEqual(result["learned"], 1)
        self.assertEqual(set(save.call_args.args[0]), {app.BLANK_TEMPLATE_KEY})

    def test_blank_template_removes_only_a_predicted_dash(self) -> None:
        grid = app.empty_grid()
        grid[0][1] = "-"
        grid[0][2] = "A"
        image = app.Image.new("RGB", (1600, 1300), "black")
        with (
            patch.object(app, "load_templates", return_value={app.BLANK_TEMPLATE_KEY: [[1.0]]}),
            patch.object(app, "classify_from_templates", return_value=(app.BLANK_TEMPLATE_KEY, 0.05)),
            patch.object(app, "cell_feature", return_value=[1.0]),
        ):
            result = app.apply_templates(grid, image)
        self.assertEqual(result[0][1], "")
        self.assertEqual(result[0][2], "A")

    def test_dash_template_cannot_generate_content_in_a_blank_cell(self) -> None:
        grid = app.empty_grid()
        image = app.Image.new("RGB", (1600, 1300), "white")
        with (
            patch.object(app, "load_templates", return_value={"-": [[1.0]]}),
            patch.object(app, "classify_from_templates", return_value=("-", 0.01)),
            patch.object(app, "cell_feature", return_value=[1.0]),
        ):
            result = app.apply_templates(grid, image)
        self.assertEqual(result[0][1], "")

    def test_blank_template_key_survives_reload(self) -> None:
        with tempfile.TemporaryDirectory() as temp_dir:
            path = Path(temp_dir) / "templates.json"
            path.write_text('{"__BLANK__": [[1.0]], "A": [[2.0]]}', encoding="utf-8")
            with patch.object(app, "TEMPLATES", path):
                templates = app.load_templates()
        self.assertEqual(set(templates), {app.BLANK_TEMPLATE_KEY, "A"})


class BlankCorrectionTests(unittest.TestCase):
    def test_completely_blank_corrected_row_is_saved(self) -> None:
        source = grid_with_row(" ----------", row=11)
        corrected = app.empty_grid()
        with (
            patch.object(app, "load_corrections", return_value={}),
            patch.object(app, "save_corrections") as save,
        ):
            result = app.remember_grid({"sourceGrid": source, "grid": corrected})
        self.assertEqual(result["saved"], 1)
        saved = save.call_args.args[0]
        self.assertEqual(len(saved), 1)
        self.assertFalse(next(iter(saved.values())).strip())


class PreprocessingTests(unittest.TestCase):
    def _make_mcdu_image(self) -> Image.Image:
        """80×60 RGB image: dark background with white, magenta, and cyan text pixels."""
        img = Image.new("RGB", (80, 60), (10, 10, 10))
        draw = ImageDraw.Draw(img)
        draw.rectangle((5, 5, 15, 15), fill=(230, 230, 230))   # white text
        draw.rectangle((20, 5, 30, 15), fill=(220, 0, 220))     # magenta text
        draw.rectangle((35, 5, 45, 15), fill=(0, 210, 210))     # cyan text
        return img

    def test_max_channel_gray_is_brighter_than_luminance_for_magenta(self) -> None:
        # A pure magenta pixel (255, 0, 255) has luminance ≈ 73 but max-channel = 255.
        img = Image.new("RGB", (4, 4), (255, 0, 255))
        mc = app.max_channel_gray(img)
        lum = ImageOps.grayscale(img)
        mc_val = np.asarray(mc).mean()
        lum_val = np.asarray(lum).mean()
        self.assertGreater(mc_val, lum_val + 50)

    def test_max_channel_gray_is_brighter_than_luminance_for_amber(self) -> None:
        # Amber (255, 191, 0) has luminance ≈ 178 but max-channel = 255.
        img = Image.new("RGB", (4, 4), (255, 191, 0))
        mc = app.max_channel_gray(img)
        lum = ImageOps.grayscale(img)
        mc_val = np.asarray(mc).mean()
        lum_val = np.asarray(lum).mean()
        self.assertGreater(mc_val, lum_val + 50)

    def test_max_channel_gray_keeps_dark_background_dark(self) -> None:
        img = Image.new("RGB", (4, 4), (8, 8, 8))
        mc = app.max_channel_gray(img)
        self.assertLessEqual(np.asarray(mc).mean(), 10)

    def test_preprocessing_variants_returns_at_least_three(self) -> None:
        img = self._make_mcdu_image()
        variants = app.preprocessing_variants(img)
        self.assertGreaterEqual(len(variants), 3)
        names = [name for name, _ in variants]
        self.assertIn("contrast", names)
        self.assertIn("inverted", names)
        self.assertIn("grayscale", names)

    def test_preprocessing_variants_returns_pil_images(self) -> None:
        img = self._make_mcdu_image()
        for name, variant in app.preprocessing_variants(img):
            self.assertIsInstance(variant, Image.Image), f"{name!r} variant is not a PIL Image"

    def test_cv2_variants_present_when_opencv_available(self) -> None:
        try:
            import cv2  # noqa: F401
        except ImportError:
            self.skipTest("OpenCV not installed")
        img = self._make_mcdu_image()
        names = [name for name, _ in app.preprocessing_variants(img)]
        self.assertIn("upscale2x", names)
        self.assertIn("denoised", names)
        self.assertIn("unsharp", names)
        self.assertIn("otsu", names)
        self.assertIn("adaptive", names)

    def test_upscale2x_variant_is_twice_the_input_size(self) -> None:
        try:
            import cv2  # noqa: F401
        except ImportError:
            self.skipTest("OpenCV not installed")
        img = self._make_mcdu_image()  # 80×60
        variants = dict(app.preprocessing_variants(img))
        upscaled = variants["upscale2x"]
        self.assertEqual(upscaled.size, (160, 120))

    def test_otsu_variant_is_binary(self) -> None:
        try:
            import cv2  # noqa: F401
        except ImportError:
            self.skipTest("OpenCV not installed")
        img = self._make_mcdu_image()
        variants = dict(app.preprocessing_variants(img))
        otsu_arr = np.asarray(variants["otsu"])
        unique_vals = set(otsu_arr.flatten().tolist())
        self.assertTrue(unique_vals.issubset({0, 255}), f"Otsu output has non-binary values: {unique_vals}")

    def test_adaptive_variant_is_binary(self) -> None:
        try:
            import cv2  # noqa: F401
        except ImportError:
            self.skipTest("OpenCV not installed")
        img = self._make_mcdu_image()
        variants = dict(app.preprocessing_variants(img))
        adaptive_arr = np.asarray(variants["adaptive"])
        unique_vals = set(adaptive_arr.flatten().tolist())
        self.assertTrue(unique_vals.issubset({0, 255}), f"Adaptive output has non-binary values: {unique_vals}")

    def test_cell_feature_uses_max_channel(self) -> None:
        # Image: dark background with a magenta rectangle inside cell (6, 20).
        # Magenta (220, 0, 220) has max-channel=220 but luminance≈73.
        # The threshold is max(90, mean+std*0.65); with a dark background the mean is
        # very low, so threshold lands at the 90 floor.  max-channel pixels at 220
        # clear it; luminance pixels at 73 do not.
        img = Image.new("RGB", (1600, 1300), (5, 5, 5))
        draw = ImageDraw.Draw(img)
        cell_w = 1600 / app.COLS
        cell_h = 1300 / app.ROWS
        cx1 = int(20 * cell_w) + 5
        cy1 = int(6 * cell_h) + 5
        cx2 = int(21 * cell_w) - 5
        cy2 = int(7 * cell_h) - 5
        draw.rectangle((cx1, cy1, cx2, cy2), fill=(220, 0, 220))

        feature = app.cell_feature(img, row=6, col=20)
        self.assertEqual(len(feature), 16 * 24)
        # After centering + zero-mean normalization the bright pixels are positive.
        # sum ≈ 0 (zero mean), but max > 0 confirms the glyph was detected.
        self.assertGreater(max(feature), 0)


def image_to_data_url(image: Image.Image) -> str:
    buf = io.BytesIO()
    image.save(buf, format="PNG")
    return "data:image/png;base64," + base64.b64encode(buf.getvalue()).decode()


def make_two_screen_image() -> Image.Image:
    """400×300 image: two side-by-side dark MCDU-shaped rectangles on a gray background."""
    img = Image.new("RGB", (400, 300), (200, 200, 200))
    draw = ImageDraw.Draw(img)
    draw.rectangle((10, 30, 185, 145), fill=(0, 0, 0))   # screen 1: 175×115, aspect ~1.52
    draw.rectangle((215, 30, 390, 145), fill=(0, 0, 0))  # screen 2: 175×115, aspect ~1.52
    return img


def make_partial_screen_image() -> Image.Image:
    """400×300 image: one complete screen + one screen cut off at the left edge."""
    img = Image.new("RGB", (400, 300), (200, 200, 200))
    draw = ImageDraw.Draw(img)
    draw.rectangle((0, 30, 130, 145), fill=(0, 0, 0))    # cut-off: touches left edge
    draw.rectangle((200, 30, 390, 145), fill=(0, 0, 0))  # complete: 190×115, aspect ~1.65
    return img


class OrientationTests(unittest.TestCase):
    def test_probe_orientation_returns_valid_angle(self) -> None:
        img = Image.new("RGB", (120, 90), (200, 200, 200))
        draw = ImageDraw.Draw(img)
        draw.rectangle((5, 5, 115, 85), fill=(0, 0, 0))
        result = app.probe_orientation(img)
        self.assertIn(result, (0, 90, 180, 270))

    def test_probe_orientation_returns_zero_when_all_scores_equal(self) -> None:
        # Uniform dark image — all orientations score 0; default 0 should be returned.
        img = Image.new("RGB", (100, 100), (0, 0, 0))
        result = app.probe_orientation(img)
        self.assertEqual(result, 0)

    def test_detect_display_includes_rotation_field(self) -> None:
        img = Image.new("RGB", (400, 300), (200, 200, 200))
        draw = ImageDraw.Draw(img)
        draw.rectangle((50, 50, 350, 250), fill=(0, 0, 0))
        result = app.detect_display({"image": image_to_data_url(img)})
        self.assertIn("rotation", result)
        self.assertIn(result["rotation"], (0, 90, 180, 270))

    def test_skip_orientation_probe_returns_rotation_zero(self) -> None:
        img = Image.new("RGB", (400, 300), (200, 200, 200))
        draw = ImageDraw.Draw(img)
        draw.rectangle((50, 50, 350, 250), fill=(0, 0, 0))
        result = app.detect_display(
            {"image": image_to_data_url(img), "skipOrientationProbe": True}
        )
        self.assertEqual(result["rotation"], 0)

    def test_exif_transpose_already_applied_in_load_image(self) -> None:
        # load_image wraps exif_transpose, so a pure RGB image (no EXIF) must come
        # back as RGB without raising.
        img = Image.new("RGB", (80, 60), (10, 20, 30))
        result = app.load_image(image_to_data_url(img))
        self.assertEqual(result.mode, "RGB")
        self.assertEqual(result.size, (80, 60))


class MultiScreenDetectionTests(unittest.TestCase):
    def test_two_screens_returns_two_candidates(self) -> None:
        payload = {"image": image_to_data_url(make_two_screen_image())}
        result = app.detect_display(payload)
        self.assertIn("candidates", result)
        self.assertGreaterEqual(len(result["candidates"]), 2)

    def test_response_always_includes_corners_and_best_index(self) -> None:
        payload = {"image": image_to_data_url(make_two_screen_image())}
        result = app.detect_display(payload)
        self.assertIn("corners", result)
        self.assertIn("bestIndex", result)
        self.assertEqual(result["bestIndex"], 0)
        self.assertEqual(len(result["corners"]), 4)

    def test_each_candidate_has_required_fields(self) -> None:
        payload = {"image": image_to_data_url(make_two_screen_image())}
        result = app.detect_display(payload)
        for candidate in result["candidates"]:
            self.assertIn("corners", candidate)
            self.assertIn("confidence", candidate)
            self.assertIn("boundingBox", candidate)
            self.assertIn("touchesEdge", candidate)
            self.assertEqual(len(candidate["corners"]), 4)

    def test_edge_touching_screen_scores_lower_than_complete_screen(self) -> None:
        payload = {"image": image_to_data_url(make_partial_screen_image())}
        result = app.detect_display(payload)
        self.assertGreaterEqual(len(result["candidates"]), 2)
        edge_candidates = [c for c in result["candidates"] if c["touchesEdge"]]
        complete_candidates = [c for c in result["candidates"] if not c["touchesEdge"]]
        self.assertTrue(edge_candidates, "expected at least one edge-touching candidate")
        self.assertTrue(complete_candidates, "expected at least one complete candidate")
        best_complete = max(c["score"] for c in complete_candidates)
        best_edge = max(c["score"] for c in edge_candidates)
        self.assertGreater(best_complete, best_edge)

    def test_auto_pick_avoids_edge_touching_screen(self) -> None:
        payload = {"image": image_to_data_url(make_partial_screen_image())}
        result = app.detect_display(payload)
        best = result["candidates"][result["bestIndex"]]
        self.assertFalse(best["touchesEdge"], "auto-pick should prefer the complete screen")

    def test_single_screen_returns_one_candidate_with_best_index_zero(self) -> None:
        img = Image.new("RGB", (400, 300), (200, 200, 200))
        draw = ImageDraw.Draw(img)
        draw.rectangle((50, 50, 350, 250), fill=(0, 0, 0))  # one large complete screen
        payload = {"image": image_to_data_url(img)}
        result = app.detect_display(payload)
        self.assertEqual(result["bestIndex"], 0)
        self.assertEqual(result["candidates"][0]["touchesEdge"], False)


# ---------------------------------------------------------------------------
# Helper builders for #14–#17 tests
# ---------------------------------------------------------------------------

def make_magenta_cross_image(
    width: int = 120, height: int = 100, arm: int = 4
) -> np.ndarray:
    """Dark RGB image with a centred magenta + cross (R=220, G=0, B=220)."""
    arr = np.zeros((height, width, 3), dtype=np.uint8)
    cx, cy = width // 2, height // 2
    arr[cy - arm : cy + arm + 1, :, 0] = 220  # H-arm: R channel
    arr[cy - arm : cy + arm + 1, :, 2] = 220  # H-arm: B channel
    arr[:, cx - arm : cx + arm + 1, 0] = 220  # V-arm: R channel
    arr[:, cx - arm : cx + arm + 1, 2] = 220  # V-arm: B channel
    return arr


def make_magenta_L_image(width: int = 60, height: int = 100, thick: int = 5) -> np.ndarray:
    """Dark RGB image with an L-shaped magenta blob (not a cross)."""
    arr = np.zeros((height, width, 3), dtype=np.uint8)
    arr[:, :thick, 0] = 220             # vertical bar on the left
    arr[:, :thick, 2] = 220
    arr[-thick:, :, 0] = 220            # horizontal bar at the bottom
    arr[-thick:, :, 2] = 220
    return arr


def make_magenta_block_image(width: int = 60, height: int = 60) -> np.ndarray:
    """Dark RGB image with a centred filled magenta block (a stand-in for a
    centred magenta glyph such as a digit). Fills its corners, so it must NOT be
    mistaken for a crosshair and erased."""
    arr = np.zeros((height, width, 3), dtype=np.uint8)
    cx, cy = width // 2, height // 2
    half = min(width, height) // 4
    arr[cy - half : cy + half, cx - half : cx + half, 0] = 220
    arr[cy - half : cy + half, cx - half : cx + half, 2] = 220
    return arr


def make_glare_image(
    width: int = 200, height: int = 150, blob_size: int = 60
) -> np.ndarray:
    """Dark RGB image with one large bright-white glare blob."""
    arr = np.full((height, width, 3), 20, dtype=np.uint8)
    cx, cy = width // 2, height // 2
    r = blob_size // 2
    arr[cy - r : cy + r, cx - r : cx + r] = 255
    return arr


def make_thin_rect_image(width: int = 400, height: int = 260) -> np.ndarray:
    """Dark RGB image with a thin bright rectangle outline (~4 cols × 1 row).

    Image is sized 400×260 so cell_w=10, cell_h=20.  The rectangle at
    (20,10)-(60,30) is 40×20 px = 4 cols × 1 row — within the 9-col limit.
    """
    arr = np.zeros((height, width, 3), dtype=np.uint8)
    x0, y0, x1, y1 = 20, 10, 60, 30  # not touching the image edge
    arr[y0 : y0 + 2, x0 : x1] = 220  # top border
    arr[y1 - 2 : y1, x0 : x1] = 220  # bottom border
    arr[y0 : y1, x0 : x0 + 2] = 220  # left border
    arr[y0 : y1, x1 - 2 : x1] = 220  # right border
    return arr


class CleanWarpedTests(unittest.TestCase):
    """Tests for #14 (cursor erasure), #15 (glare), #16 (message boxes), #17 (outlines)."""

    def _skip_no_cv2(self) -> None:
        try:
            import cv2  # noqa: F401
        except ImportError:
            self.skipTest("OpenCV not installed")

    # --- #14 cursor erasure ------------------------------------------------

    def test_erase_magenta_cursor_removes_cross(self) -> None:
        self._skip_no_cv2()
        rgb = make_magenta_cross_image()
        cx, cy = rgb.shape[1] // 2, rgb.shape[0] // 2
        # Centre pixel should be magenta before erasure
        self.assertEqual(rgb[cy, cx, 0], 220)
        self.assertEqual(rgb[cy, cx, 1], 0)
        result = app.erase_magenta_cursor(rgb)
        # After inpaint, the cross centre should no longer be pure magenta
        centre = result[cy, cx]
        is_magenta = centre[0] > 180 and centre[1] < 40 and centre[2] > 180
        self.assertFalse(is_magenta, "Centre pixel should be inpainted, not still magenta")

    def test_erase_magenta_cursor_preserves_L_shape(self) -> None:
        self._skip_no_cv2()
        # An L-shaped magenta blob is NOT a cross — it should NOT be inpainted
        rgb = make_magenta_L_image()
        before_sum = int(np.sum(rgb[:, :, 0] > 100))  # count R-high pixels
        result = app.erase_magenta_cursor(rgb)
        after_sum = int(np.sum(result[:, :, 0] > 100))
        # Most magenta pixels should survive (the L is not a cross)
        self.assertGreater(after_sum, before_sum * 0.7, "L-shaped magenta was wrongly erased")

    def test_erase_magenta_cursor_noop_on_dark_image(self) -> None:
        self._skip_no_cv2()
        rgb = np.zeros((80, 100, 3), dtype=np.uint8)
        result = app.erase_magenta_cursor(rgb)
        np.testing.assert_array_equal(result, rgb)

    def test_erase_magenta_cursor_preserves_centred_glyph(self) -> None:
        self._skip_no_cv2()
        # A centred filled magenta block (like a magenta digit/value) fills its
        # corners and must survive — it is not a crosshair.
        rgb = make_magenta_block_image()
        before = int(np.sum(rgb[:, :, 0] > 100))
        result = app.erase_magenta_cursor(rgb)
        after = int(np.sum(result[:, :, 0] > 100))
        self.assertGreater(
            after, before * 0.9, "Centred magenta glyph was wrongly erased as a cursor"
        )

    # --- #15 glare suppression ---------------------------------------------

    def test_suppress_glare_removes_large_bright_blob(self) -> None:
        self._skip_no_cv2()
        rgb = make_glare_image(blob_size=80)
        h, w = rgb.shape[:2]
        cell_w = w / app.COLS
        cell_h = h / app.ROWS
        # Blob is large enough to exceed the min-area threshold
        result = app.suppress_glare(rgb, cell_w, cell_h)
        cx, cy = w // 2, h // 2
        # The centre of the blob should no longer be pure white after inpainting
        centre = result[cy, cx]
        self.assertLess(int(centre.min()), 240, "Glare centre should be reduced by inpainting")

    def test_suppress_glare_preserves_small_bright_pixels(self) -> None:
        self._skip_no_cv2()
        # A single bright pixel (text-sized) must survive — it is below the area threshold
        rgb = np.zeros((80, 100, 3), dtype=np.uint8)
        rgb[40, 50] = 255
        cell_w = 100 / app.COLS
        cell_h = 80 / app.ROWS
        result = app.suppress_glare(rgb, cell_w, cell_h)
        np.testing.assert_array_equal(result[40, 50], [255, 255, 255])

    # --- #17 entry-outline erasure -----------------------------------------

    def test_erase_entry_outlines_removes_thin_rectangle(self) -> None:
        self._skip_no_cv2()
        rgb = make_thin_rect_image()           # 400×260
        h, w = rgb.shape[:2]
        cell_w = w / app.COLS                  # 400/40 = 10
        cell_h = h / app.ROWS                  # 260/13 = 20
        # Top border at y=10 should be bright before erasure
        self.assertGreater(int(rgb[10, 40, 0]), 100)
        result = app.erase_entry_outlines(rgb, cell_w, cell_h)
        # After erasure, the top border pixel should be dark
        self.assertLess(int(result[10, 40, 0]), 50, "Top border pixel should be erased")

    # --- composite + #16 ---------------------------------------------------

    def test_clean_warped_image_returns_same_size(self) -> None:
        self._skip_no_cv2()
        warped = Image.new("RGB", (1600, 1040), (0, 0, 0))
        result = app.clean_warped_image(warped)
        self.assertIsInstance(result, Image.Image)
        self.assertEqual(result.size, (1600, 1040))

    def test_clean_warped_image_noop_on_dark_screen(self) -> None:
        self._skip_no_cv2()
        warped = Image.new("RGB", (400, 260), (5, 5, 5))
        result = app.clean_warped_image(warped)
        # No cursor, no glare, no outlines → image should be virtually identical
        arr_in = np.asarray(warped)
        arr_out = np.asarray(result)
        self.assertEqual(arr_in.shape, arr_out.shape)

    def test_message_box_preprocessing_makes_white_text_dark(self) -> None:
        self._skip_no_cv2()
        # Blue background (R=0, G=0, B=200) with a white text patch (R=G=B=255)
        crop = np.zeros((40, 80, 3), dtype=np.uint8)
        crop[:] = [0, 0, 200]                     # blue background
        crop[10:30, 20:60] = [255, 255, 255]      # white text area
        result_arr = np.asarray(app._preprocess_message_box_region(crop))
        # White text → (255−0)×255/255 = 255, inverted → 0 (dark)
        white_region = result_arr[10:30, 20:60]
        # Blue bg → (255−200ish)×200ish/255 ≈ 43, inverted → ~212 (light)
        bg_region = result_arr[:, :10]
        self.assertLess(float(white_region.mean()), 60, "White text should become dark")
        self.assertGreater(float(bg_region.mean()), 100, "Blue bg should become light")


class PerRowStripOcrTests(unittest.TestCase):
    """Tests for per_row_strip_ocr (#18) — per-row Tesseract --psm 7 pass."""

    def _make_mcdu_warped(self, width: int = 1600, height: int = 1040) -> Image.Image:
        return Image.new("RGB", (width, height), (0, 0, 0))

    def test_per_row_strip_ocr_returns_correct_grid_dimensions(self) -> None:
        warped = self._make_mcdu_warped()
        geometry = app.calibrate_grid([], warped.size)
        grid, conf = app.per_row_strip_ocr(warped, geometry)
        self.assertEqual(len(grid), app.ROWS)
        self.assertEqual(len(conf), app.ROWS)
        for row in range(app.ROWS):
            self.assertEqual(len(grid[row]), app.COLS)
            self.assertEqual(len(conf[row]), app.COLS)

    def test_per_row_strip_ocr_only_writes_to_data_columns(self) -> None:
        # Guard cells (0 and 39) must never be written, regardless of OCR output
        warped = self._make_mcdu_warped()
        geometry = app.calibrate_grid([], warped.size)
        grid, _ = app.per_row_strip_ocr(warped, geometry)
        for row in range(app.ROWS):
            self.assertEqual(grid[row][0], "", f"Row {row} col 0 guard was written")
            self.assertEqual(grid[row][app.COLS - 1], "", f"Row {row} col 39 guard was written")
        # All characters written must be from the OCR whitelist
        for row in range(app.ROWS):
            for col in range(app.FIRST_DATA_COL, app.LAST_DATA_COL + 1):
                char = grid[row][col]
                if char:
                    self.assertIn(char, app.OCR_WHITELIST, f"Cell [{row}][{col}] = {char!r} not in whitelist")

    def test_per_row_strip_ocr_guard_cells_remain_empty(self) -> None:
        warped = self._make_mcdu_warped()
        geometry = app.calibrate_grid([], warped.size)
        grid, _ = app.per_row_strip_ocr(warped, geometry)
        for row in range(app.ROWS):
            self.assertEqual(grid[row][0], "", "Column 0 guard must remain empty")
            self.assertEqual(grid[row][app.COLS - 1], "", "Column 39 guard must remain empty")

    def test_per_row_strip_ocr_does_not_hallucinate_on_blank_image(self) -> None:
        # A blank (all-black) display has no ink, so the ink gate must skip every
        # strip — no characters may be written, otherwise PSM 7 would inject junk
        # into empty rows that row_candidate_score could pick over a blank row.
        warped = self._make_mcdu_warped()
        geometry = app.calibrate_grid([], warped.size)
        grid, _ = app.per_row_strip_ocr(warped, geometry)
        written = sum(bool(cell) for row in grid for cell in row)
        self.assertEqual(written, 0, "Blank display produced hallucinated characters")

    def test_per_row_strip_ocr_does_not_crash_on_tiny_image(self) -> None:
        warped = self._make_mcdu_warped(width=80, height=52)
        geometry = app.calibrate_grid([], warped.size)
        try:
            grid, conf = app.per_row_strip_ocr(warped, geometry)
        except Exception as exc:  # noqa: BLE001
            self.fail(f"per_row_strip_ocr raised unexpectedly: {exc}")
        self.assertEqual(len(grid), app.ROWS)


class FieldValidatorTests(unittest.TestCase):
    """Tests for _snap_fl_token, _snap_decimal_token, and validate_field_formats (#22)."""

    @staticmethod
    def _label_value(label: str, value: str) -> list[list[str]]:
        grid = app.empty_grid()
        for text, r in ((label, 0), (value, 1)):
            padded = text[: app.COLS].ljust(app.COLS)
            grid[r] = ["" if ch == " " else ch for ch in padded]
        return app.normalize_grid_guards(grid)

    def test_mach_speed_leading_dot_restored_in_speed_row(self) -> None:
        out = app.validate_field_formats(self._label_value(" ECON SPD", "  733"))
        self.assertEqual("".join(c or " " for c in out[1]).strip(), ".733")
        out = app.validate_field_formats(self._label_value(" RTA SPD", "  482"))
        self.assertEqual("".join(c or " " for c in out[1]).strip(), ".482")

    def test_three_digits_not_given_dot_outside_speed_context(self) -> None:
        # Altitudes / other 3-digit values must never gain a leading dot.
        out = app.validate_field_formats(self._label_value(" CRZ ALT", "  340"))
        self.assertEqual("".join(c or " " for c in out[1]).strip(), "340")
        out = app.validate_field_formats(self._label_value(" OPT", "  344"))
        self.assertEqual("".join(c or " " for c in out[1]).strip(), "344")

    # --- _snap_fl_token ---

    def test_snap_fl_correct_token_returns_none(self) -> None:
        self.assertIsNone(app._snap_fl_token("FL204"))
        self.assertIsNone(app._snap_fl_token("FL89"))
        self.assertIsNone(app._snap_fl_token("FL350"))

    def test_snap_fl_fixes_second_char_substitution(self) -> None:
        # L → I (very common on some fonts)
        self.assertEqual(app._snap_fl_token("FI204"), "FL204")
        self.assertEqual(app._snap_fl_token("F1350"), "FL350")

    def test_snap_fl_fixes_first_char_substitution(self) -> None:
        # F → E (serifs can cause E↔F confusion)
        self.assertEqual(app._snap_fl_token("EL204"), "FL204")
        self.assertEqual(app._snap_fl_token("PL350"), "FL350")

    def test_snap_fl_rejects_unrelated_tokens(self) -> None:
        self.assertIsNone(app._snap_fl_token("ECON"))
        self.assertIsNone(app._snap_fl_token(".860"))
        self.assertIsNone(app._snap_fl_token("FL"))       # too short
        self.assertIsNone(app._snap_fl_token("EL2"))      # only 1 digit (too short)
        self.assertIsNone(app._snap_fl_token("ABCDE"))    # not FL-like

    def test_snap_fl_preserves_length(self) -> None:
        result = app._snap_fl_token("EL204")
        self.assertIsNotNone(result)
        self.assertEqual(len(result), len("EL204"))  # type: ignore[arg-type]

    # --- _snap_decimal_token ---

    def test_snap_decimal_correct_token_returns_none(self) -> None:
        self.assertIsNone(app._snap_decimal_token(".860"))
        self.assertIsNone(app._snap_decimal_token(".840"))

    def test_snap_decimal_fixes_comma(self) -> None:
        self.assertEqual(app._snap_decimal_token(",860"), ".860")
        self.assertEqual(app._snap_decimal_token(",840"), ".840")

    def test_snap_decimal_rejects_unrelated(self) -> None:
        self.assertIsNone(app._snap_decimal_token("ECON"))
        self.assertIsNone(app._snap_decimal_token("FL350"))
        self.assertIsNone(app._snap_decimal_token(".86"))    # only 2 digits
        self.assertIsNone(app._snap_decimal_token("860"))   # missing dot (length change)

    # --- validate_field_formats ---

    def _make_grid_with_tokens(self, row_tokens: dict[int, list[tuple[int, str]]]) -> list[list[str]]:
        """Build a 13×40 grid with specific tokens at (row, start_col)."""
        grid = app.empty_grid()
        for row, tokens in row_tokens.items():
            for start_col, text in tokens:
                for i, ch in enumerate(text):
                    if app.FIRST_DATA_COL <= start_col + i <= app.LAST_DATA_COL:
                        grid[row][start_col + i] = ch
        return grid

    def test_validate_standalone_fl_snapping(self) -> None:
        grid = self._make_grid_with_tokens({2: [(5, "FI204")]})
        result = app.validate_field_formats(grid)
        self.assertEqual("".join(result[2][5:10]).rstrip(), "FL204")

    def test_validate_standalone_decimal_snapping(self) -> None:
        grid = self._make_grid_with_tokens({4: [(20, ",860")]})
        result = app.validate_field_formats(grid)
        self.assertEqual("".join(result[4][20:24]), ".860")

    def test_validate_label_driven_fl_from_crz_alt(self) -> None:
        # Label row 1: "CRZ ALT", data row 2: "EL350"
        grid = self._make_grid_with_tokens({
            1: [(2, "CRZ"), (6, "ALT")],
            2: [(2, "EL350")],
        })
        result = app.validate_field_formats(grid)
        self.assertEqual("".join(result[2][2:7]), "FL350")

    def test_validate_label_driven_spd_from_econ_spd(self) -> None:
        # Label row 3: "ECON SPD", data row 4: ",840"
        grid = self._make_grid_with_tokens({
            3: [(2, "ECON"), (7, "SPD")],
            4: [(2, ",840")],
        })
        result = app.validate_field_formats(grid)
        self.assertEqual("".join(result[4][2:6]), ".840")

    def test_validate_noop_on_already_correct(self) -> None:
        grid = self._make_grid_with_tokens({
            1: [(2, "CRZ"), (6, "ALT")],
            2: [(2, "FL350"), (10, ".840")],
        })
        result = app.validate_field_formats(grid)
        self.assertEqual("".join(result[2][2:7]), "FL350")
        self.assertEqual("".join(result[2][10:14]), ".840")

    def test_validate_does_not_touch_guard_columns(self) -> None:
        grid = app.empty_grid()
        result = app.validate_field_formats(grid)
        for row in range(app.ROWS):
            self.assertEqual(result[row][0], "")
            self.assertEqual(result[row][app.COLS - 1], "")


class ColorSemanticsTests(unittest.TestCase):
    """Tests for _classify_cell_color and extract_color_semantics (#23)."""

    def test_classify_magenta(self) -> None:
        # OpenCV HSV magenta: H=150, S=255, V=255
        crop = np.zeros((8, 6, 3), dtype=np.uint8)
        crop[:, :, 0] = 150
        crop[:, :, 1] = 255
        crop[:, :, 2] = 255
        self.assertEqual(app._classify_cell_color(crop), "magenta")

    def test_classify_white(self) -> None:
        # White: H=0, S=0, V=255
        crop = np.zeros((8, 6, 3), dtype=np.uint8)
        crop[:, :, 1] = 0
        crop[:, :, 2] = 255
        self.assertEqual(app._classify_cell_color(crop), "white")

    def test_classify_cyan(self) -> None:
        # Cyan: H=90, S=255, V=255
        crop = np.zeros((8, 6, 3), dtype=np.uint8)
        crop[:, :, 0] = 90
        crop[:, :, 1] = 255
        crop[:, :, 2] = 255
        self.assertEqual(app._classify_cell_color(crop), "cyan")

    def test_classify_amber(self) -> None:
        # Amber/yellow: H=25, S=255, V=255
        crop = np.zeros((8, 6, 3), dtype=np.uint8)
        crop[:, :, 0] = 25
        crop[:, :, 1] = 255
        crop[:, :, 2] = 255
        self.assertEqual(app._classify_cell_color(crop), "amber")

    def test_classify_dark_cell_returns_empty(self) -> None:
        # All-dark: V=50 (below 120 threshold)
        crop = np.zeros((8, 6, 3), dtype=np.uint8)
        crop[:, :, 2] = 50
        self.assertEqual(app._classify_cell_color(crop), "")

    def test_extract_color_semantics_dimensions(self) -> None:
        warped = Image.new("RGB", (1600, 1040), (0, 0, 0))
        grid = app.empty_grid()
        result = app.extract_color_semantics(warped, grid)
        self.assertEqual(len(result), app.ROWS)
        for row in result:
            self.assertEqual(len(row), app.COLS)

    def test_extract_color_semantics_empty_cells_are_blank(self) -> None:
        warped = Image.new("RGB", (1600, 1040), (0, 0, 0))
        grid = app.empty_grid()  # all cells empty
        result = app.extract_color_semantics(warped, grid)
        for row in result:
            for cell in row:
                self.assertEqual(cell, "")

    def test_extract_color_semantics_identifies_magenta_cell(self) -> None:
        # Draw a magenta rectangle in cell (2, 5)
        img = Image.new("RGB", (1600, 1040), (0, 0, 0))
        draw = ImageDraw.Draw(img)
        cell_w = 1600 / app.COLS
        cell_h = 1040 / app.ROWS
        x1 = int(5 * cell_w) + 2
        y1 = int(2 * cell_h) + 2
        x2 = int(6 * cell_w) - 2
        y2 = int(3 * cell_h) - 2
        draw.rectangle((x1, y1, x2, y2), fill=(220, 0, 220))  # magenta
        grid = app.empty_grid()
        grid[2][5] = "A"  # cell must be non-empty to get a color label
        result = app.extract_color_semantics(img, grid)
        self.assertEqual(result[2][5], "magenta")

    def test_extract_color_semantics_guard_columns_are_blank(self) -> None:
        warped = Image.new("RGB", (1600, 1040), (255, 255, 255))
        grid = app.empty_grid()
        result = app.extract_color_semantics(warped, grid)
        for row in range(app.ROWS):
            self.assertEqual(result[row][0], "")
            self.assertEqual(result[row][app.COLS - 1], "")


class TesseractConfigTests(unittest.TestCase):
    """Tests for shared Tesseract flags added by _tesseract_extra_args (#19/#20)."""

    def test_extra_args_contains_oem1(self) -> None:
        args = app._tesseract_extra_args()
        oem_idx = next((i for i, v in enumerate(args) if v == "--oem"), None)
        self.assertIsNotNone(oem_idx, "--oem missing from _tesseract_extra_args")
        self.assertEqual(args[oem_idx + 1], "1")  # type: ignore[index]

    def test_extra_args_contains_dpi300(self) -> None:
        args = app._tesseract_extra_args()
        self.assertTrue(
            any("user_defined_dpi=300" in arg for arg in args),
            "user_defined_dpi=300 missing from _tesseract_extra_args",
        )

    def test_vocab_words_file_written_by_ensure_dirs(self) -> None:
        with tempfile.TemporaryDirectory() as temp_dir:
            with (
                patch.object(app, "DATA", Path(temp_dir)),
                patch.object(app, "EXPORTS", Path(temp_dir) / "exports"),
                patch.object(app, "CORRECTIONS", Path(temp_dir) / "corrections.json"),
                patch.object(app, "TEMPLATES", Path(temp_dir) / "templates.json"),
                patch.object(app, "VOCAB_WORDS", Path(temp_dir) / "mcdu_user_words.txt"),
                patch.object(app, "VOCAB_PATTERNS", Path(temp_dir) / "mcdu_user_patterns.txt"),
                patch.object(app, "cleanup_exports"),
            ):
                app.ensure_dirs()
                words_file = Path(temp_dir) / "mcdu_user_words.txt"
                patterns_file = Path(temp_dir) / "mcdu_user_patterns.txt"
                self.assertTrue(words_file.exists(), "mcdu_user_words.txt not created")
                self.assertTrue(patterns_file.exists(), "mcdu_user_patterns.txt not created")
                words = words_file.read_text(encoding="utf-8").splitlines()
                self.assertIn("LEGS", words)
                self.assertIn("CRZ", words)
                patterns = patterns_file.read_text(encoding="utf-8").splitlines()
                # Tesseract user-pattern syntax: \n is a digit (regex \d is rejected).
                self.assertTrue(any(r"FL\n\n" in p for p in patterns))
                self.assertFalse(
                    any(r"\d" in p for p in patterns),
                    "patterns must not use regex \\d — Tesseract rejects it",
                )


class TemplateMatcherTests(unittest.TestCase):
    """Tests for the upgraded NCC template matcher (#21)."""

    def test_feature_distance_identical_is_zero(self) -> None:
        a = [1.5, -0.7, 0.3, 0.0]
        self.assertAlmostEqual(app.feature_distance(a, a), 0.0, places=5)

    def test_feature_distance_mismatched_lengths_is_one(self) -> None:
        self.assertEqual(app.feature_distance([1.0], [1.0, 2.0]), 1.0)

    def test_feature_distance_orthogonal_is_half(self) -> None:
        # Orthogonal vectors → cosine = 0 → distance = 0.5
        self.assertAlmostEqual(app.feature_distance([1.0, 0.0], [0.0, 1.0]), 0.5, places=5)

    def test_feature_distance_opposite_is_one(self) -> None:
        # Anti-correlated vectors → cosine = -1 → distance = 1.0
        self.assertAlmostEqual(app.feature_distance([1.0], [-1.0]), 1.0, places=5)

    def test_feature_distance_both_zero_is_zero(self) -> None:
        # Two blank features (all zeros) match each other — both are "empty cells"
        self.assertAlmostEqual(app.feature_distance([0.0, 0.0], [0.0, 0.0]), 0.0, places=5)

    def test_cell_feature_length_is_384(self) -> None:
        img = Image.new("RGB", (1600, 1300), (10, 10, 10))
        feature = app.cell_feature(img, row=0, col=1)
        self.assertEqual(len(feature), 16 * 24)

    def test_cell_feature_bright_cell_has_positive_max(self) -> None:
        # Bright content in a cell → after centering + normalization, max > 0
        img = Image.new("RGB", (1600, 1300), (10, 10, 10))
        draw = ImageDraw.Draw(img)
        cell_w = 1600 / app.COLS
        cell_h = 1300 / app.ROWS
        cx1 = int(20 * cell_w) + 5
        cy1 = int(6 * cell_h) + 5
        cx2 = int(21 * cell_w) - 5
        cy2 = int(7 * cell_h) - 5
        draw.rectangle((cx1, cy1, cx2, cy2), fill=(220, 220, 220))
        feature = app.cell_feature(img, row=6, col=20)
        self.assertGreater(max(feature), 0.0)

    def test_cell_feature_same_char_closer_than_different(self) -> None:
        # Two versions of the same glyph should be closer to each other
        # than to a clearly different cell.
        img_a = Image.new("RGB", (1600, 1300), (5, 5, 5))
        img_b = Image.new("RGB", (1600, 1300), (5, 5, 5))
        img_c = Image.new("RGB", (1600, 1300), (5, 5, 5))
        draw_a = ImageDraw.Draw(img_a)
        draw_b = ImageDraw.Draw(img_b)
        draw_c = ImageDraw.Draw(img_c)
        cell_w = 1600 / app.COLS
        cell_h = 1300 / app.ROWS
        row, col = 4, 10
        x0 = int(col * cell_w) + 4
        y0 = int(row * cell_h) + 4
        x1 = int((col + 1) * cell_w) - 4
        y1 = int((row + 1) * cell_h) - 4
        # img_a and img_b: near-identical small white rectangle (same character proxy)
        draw_a.rectangle((x0, y0, x1, y1), fill=(210, 210, 210))
        draw_b.rectangle((x0 + 1, y0 + 1, x1, y1), fill=(215, 215, 215))
        # img_c: completely different — large rectangle filling the whole cell
        draw_c.rectangle((x0 - 4, y0 - 4, x1 + 4, y1 + 4), fill=(5, 5, 5))  # blank (no glyph)
        feat_a = app.cell_feature(img_a, row, col)
        feat_b = app.cell_feature(img_b, row, col)
        feat_c = app.cell_feature(img_c, row, col)
        dist_same = app.feature_distance(feat_a, feat_b)
        dist_diff = app.feature_distance(feat_a, feat_c)
        self.assertLess(dist_same, dist_diff, "Similar glyphs should be closer than dissimilar ones")


class BlurScoreTests(unittest.TestCase):
    def _sharp_image(self) -> Image.Image:
        img = Image.new("RGB", (400, 300), (0, 0, 0))
        draw = ImageDraw.Draw(img)
        for x in range(0, 400, 8):
            draw.line((x, 0, x, 300), fill=(200, 200, 200), width=1)
        return img

    def _blurry_image(self) -> Image.Image:
        sharp = self._sharp_image()
        return sharp.filter(ImageFilter.GaussianBlur(radius=6))

    def test_sharp_image_returns_positive(self):
        score = app.compute_blur_score(self._sharp_image())
        self.assertGreater(score, 0.0)

    def test_blurry_image_below_threshold(self):
        score = app.compute_blur_score(self._blurry_image())
        self.assertLess(score, app.BLUR_THRESHOLD)

    def test_sharp_scores_higher_than_blurry(self):
        sharp_score = app.compute_blur_score(self._sharp_image())
        blurry_score = app.compute_blur_score(self._blurry_image())
        self.assertGreater(sharp_score, blurry_score)


class FuseGridsTests(unittest.TestCase):
    def _make_grid(self, fill: str = "") -> list[list[str]]:
        g = app.empty_grid()
        if fill:
            for r in range(app.ROWS):
                for c in range(app.FIRST_DATA_COL, app.LAST_DATA_COL + 1):
                    g[r][c] = fill
        return g

    def _set_cell(self, g, row, col, ch):
        g[row][col] = ch
        return g

    def test_all_agree(self):
        g1 = self._set_cell(self._make_grid(), 2, 5, "A")
        g2 = self._set_cell(self._make_grid(), 2, 5, "A")
        g3 = self._set_cell(self._make_grid(), 2, 5, "A")
        fused, summary = app.fuse_grids([g1, g2, g3])
        self.assertEqual(fused[2][5], "A")
        self.assertEqual(summary["agreed"], 1)
        self.assertEqual(summary["conflicted"], 0)

    def test_majority_wins(self):
        g1 = self._set_cell(self._make_grid(), 3, 10, "B")
        g2 = self._set_cell(self._make_grid(), 3, 10, "B")
        g3 = self._set_cell(self._make_grid(), 3, 10, "8")
        fused, summary = app.fuse_grids([g1, g2, g3])
        self.assertEqual(fused[3][10], "B")
        self.assertEqual(summary["conflicted"], 1)

    def test_fill_from_single_grid(self):
        g1 = self._make_grid()
        g2 = self._set_cell(self._make_grid(), 5, 7, "X")
        fused, summary = app.fuse_grids([g1, g2])
        self.assertEqual(fused[5][7], "X")
        self.assertEqual(summary["filled"], 1)

    def test_all_empty_cell_stays_empty(self):
        g1 = self._make_grid()
        g2 = self._make_grid()
        fused, summary = app.fuse_grids([g1, g2])
        self.assertEqual(fused[0][0], "")
        self.assertGreater(summary["empty"], 0)

    def test_output_dimensions(self):
        g1 = self._make_grid()
        g2 = self._make_grid()
        fused, _ = app.fuse_grids([g1, g2])
        self.assertEqual(len(fused), app.ROWS)
        self.assertTrue(all(len(row) == app.COLS for row in fused))

    def test_guard_columns_preserved(self):
        g1 = self._make_grid("Z")
        g2 = self._make_grid("Z")
        fused, _ = app.fuse_grids([g1, g2])
        for row in fused:
            self.assertEqual(row[0], "")
            self.assertEqual(row[app.COLS - 1], "")

    def test_requires_two_or_more_grids(self):
        with self.assertRaises((ValueError, Exception)):
            app.fuse_grids([self._make_grid()])

    def test_summary_grids_count(self):
        g1 = self._make_grid()
        g2 = self._make_grid()
        g3 = self._make_grid()
        _, summary = app.fuse_grids([g1, g2, g3])
        self.assertEqual(summary["grids"], 3)


class FuseGridsDeterminismTests(unittest.TestCase):
    def _cell_grid(self, row: int, col: int, char: str) -> list[list[str]]:
        g = app.empty_grid()
        g[row][col] = char
        return g

    def test_tie_broken_by_first_grid(self):
        """On a 1-1 tie the character from grid 0 (first grid) must win."""
        g1 = self._cell_grid(5, 10, "A")
        g2 = self._cell_grid(5, 10, "B")
        fused, summary = app.fuse_grids([g1, g2])
        self.assertEqual(fused[5][10], "A", "First grid wins on 1-1 tie")
        self.assertEqual(summary["conflicted"], 1)

    def test_majority_beats_first_grid(self):
        """A 2-1 majority overrides the first-grid preference."""
        g1 = self._cell_grid(5, 10, "A")
        g2 = self._cell_grid(5, 10, "B")
        g3 = self._cell_grid(5, 10, "B")
        fused, _ = app.fuse_grids([g1, g2, g3])
        self.assertEqual(fused[5][10], "B")

    def test_deterministic_on_repeated_calls(self):
        """Same inputs always produce the same winner (not hash-dependent)."""
        g1 = self._cell_grid(3, 7, "X")
        g2 = self._cell_grid(3, 7, "Y")
        results = {app.fuse_grids([g1, g2])[0][3][7] for _ in range(10)}
        self.assertEqual(len(results), 1, "Winner must be consistent across calls")


class RowDeduplicationTests(unittest.TestCase):
    """Tests for deduplicate_adjacent_rows (A1 fix)."""

    def _make_candidate(
        self, name: str, row_texts: list[str]
    ) -> tuple[str, list[list[str]], list[list[float]]]:
        grid = app.empty_grid()
        conf = app.empty_confidence_grid()
        for row_idx, text in enumerate(row_texts):
            for i, ch in enumerate(text):
                col = app.FIRST_DATA_COL + i
                if col <= app.LAST_DATA_COL and ch.strip():
                    grid[row_idx][col] = ch
                    conf[row_idx][col] = 0.9
        return (name, grid, conf)

    def _make_word_grid_from(
        self, candidate: tuple[str, list[list[str]], list[list[float]]], rows: list[int]
    ) -> tuple[list[list[str]], list[list[float]], list[str]]:
        """Simulate per-row selection: pick ``candidate`` for the given rows."""
        _, src_grid, src_conf = candidate
        word_grid = app.empty_grid()
        word_confidence = app.empty_confidence_grid()
        row_sources: list[str] = [""] * app.ROWS
        for r in rows:
            word_grid[r] = src_grid[r][:]
            word_confidence[r] = src_conf[r][:]
            row_sources[r] = candidate[0]
        return word_grid, word_confidence, row_sources

    def test_duplicate_row_replaced_with_value(self):
        """When row N+1 duplicates row N, deduplicate replaces it with the best non-duplicate."""
        label = "CRZALT"
        value = "FL204"
        # candidate_a has label in both rows 0 and 1 (the duplication bug)
        cand_a = self._make_candidate("a", [label, label] + [""] * 11)
        # candidate_b has label in row 0, value in row 1 (the correct reading)
        cand_b = self._make_candidate("b", [label, value] + [""] * 11)
        word_candidates = [cand_a, cand_b]

        # Simulate selection: row 0 from a (correct label position),
        # row 1 accidentally also from a (duplicate)
        word_grid, word_confidence, row_sources = self._make_word_grid_from(cand_a, [0, 1])

        app.deduplicate_adjacent_rows(word_grid, word_confidence, word_candidates, row_sources)

        row0 = "".join(word_grid[0][app.FIRST_DATA_COL : app.LAST_DATA_COL + 1]).strip()
        row1 = "".join(word_grid[1][app.FIRST_DATA_COL : app.LAST_DATA_COL + 1]).strip()
        self.assertNotEqual(row0, row1, "Row 1 should no longer duplicate row 0")

    def test_value_row_survives_dedup(self):
        """After deduplication the value candidate fills the value row."""
        cand_a = self._make_candidate("a", ["ECON", "ECON"] + [""] * 11)
        cand_b = self._make_candidate("b", ["ECON", ".860"] + [""] * 11)
        word_candidates = [cand_a, cand_b]

        word_grid, word_confidence, row_sources = self._make_word_grid_from(cand_a, [0, 1])
        app.deduplicate_adjacent_rows(word_grid, word_confidence, word_candidates, row_sources)

        row1 = "".join(word_grid[1][app.FIRST_DATA_COL : app.LAST_DATA_COL + 1]).strip()
        self.assertIn(".", row1, "The value '.860' should now occupy row 1")

    def test_non_duplicate_rows_untouched(self):
        """Rows that differ from their neighbour are not modified."""
        cand = self._make_candidate("a", ["CRZALT", "FL204"] + [""] * 11)
        word_candidates = [cand]
        word_grid, word_confidence, row_sources = self._make_word_grid_from(cand, [0, 1])

        snapshot0 = word_grid[0][:]
        snapshot1 = word_grid[1][:]
        app.deduplicate_adjacent_rows(word_grid, word_confidence, word_candidates, row_sources)

        self.assertEqual(word_grid[0], snapshot0, "Row 0 should be unchanged")
        self.assertEqual(word_grid[1], snapshot1, "Row 1 should be unchanged")

    def test_no_alternative_blanks_the_row(self):
        """When every candidate duplicates the row above, the row is blanked."""
        label = "OPTALT"
        # Only one candidate with the same text in both rows
        cand = self._make_candidate("a", [label, label] + [""] * 11)
        word_candidates = [cand]
        word_grid, word_confidence, row_sources = self._make_word_grid_from(cand, [0, 1])

        app.deduplicate_adjacent_rows(word_grid, word_confidence, word_candidates, row_sources)

        row1 = "".join(word_grid[1][app.FIRST_DATA_COL : app.LAST_DATA_COL + 1]).strip()
        self.assertEqual(row1, "", "Row with no non-duplicate alternative should be blanked")


class HousekeepingTests(unittest.TestCase):
    """Tests for D1 (cleanup_exports), D2 (recheck validation), D3 (blur band)."""

    def test_cleanup_exports_removes_old_docx(self):
        """cleanup_exports must delete aged-out .docx files (D1)."""
        with tempfile.TemporaryDirectory() as tmpdir:
            exports = Path(tmpdir)
            recent_docx = exports / "recent.docx"
            old_docx = exports / "old.docx"
            old_png = exports / "old.png"
            for p in (recent_docx, old_docx, old_png):
                p.write_bytes(b"x")
            aged = time.time() - app.EXPORT_MAX_AGE_SECONDS - 10
            os.utime(old_docx, (aged, aged))
            os.utime(old_png, (aged, aged))

            orig = app.EXPORTS
            app.EXPORTS = exports
            try:
                app.cleanup_exports()
            finally:
                app.EXPORTS = orig

            remaining = {p.name for p in exports.iterdir()}
            self.assertNotIn("old.docx", remaining, "Aged .docx must be deleted")
            self.assertNotIn("old.png", remaining, "Aged .png must be deleted")
            self.assertIn("recent.docx", remaining, "Recent .docx must survive")

    def test_whole_grid_recheck_calls_validate_field_formats(self):
        """whole_grid_focused_recheck must include validate_field_formats (D2)."""
        import inspect
        src = inspect.getsource(app.whole_grid_focused_recheck)
        self.assertIn("validate_field_formats", src)

    def test_blur_warning_threshold_above_blur_threshold(self):
        """BLUR_WARNING_THRESHOLD > BLUR_THRESHOLD defines the marginal band (D3)."""
        self.assertGreater(app.BLUR_WARNING_THRESHOLD, app.BLUR_THRESHOLD)

    def test_blur_warning_band_logic(self):
        """A score in the marginal band must produce blurWarning=True, blurry=False (D3)."""
        mid = (app.BLUR_THRESHOLD + app.BLUR_WARNING_THRESHOLD) / 2.0
        blurry = mid > 0 and mid < app.BLUR_THRESHOLD
        blur_warning = app.BLUR_THRESHOLD <= mid < app.BLUR_WARNING_THRESHOLD
        self.assertFalse(blurry)
        self.assertTrue(blur_warning)

    def test_below_blur_threshold_is_blurry_not_warning(self):
        """A score below BLUR_THRESHOLD must produce blurry=True, blurWarning=False (D3)."""
        score = app.BLUR_THRESHOLD / 2.0
        blurry = score > 0 and score < app.BLUR_THRESHOLD
        blur_warning = app.BLUR_THRESHOLD <= score < app.BLUR_WARNING_THRESHOLD
        self.assertTrue(blurry)
        self.assertFalse(blur_warning)


class AtlasTests(unittest.TestCase):
    """Tests for the B1 pre-seeded glyph-atlas engine."""

    def _get_atlas(self):
        atlas = app._build_glyph_atlas()
        if not atlas:
            self.skipTest("No atlas fonts available on this system")
        return atlas

    def _render_large(self, char, font, size=200):
        img = Image.new("L", (size, size), 0)
        draw = ImageDraw.Draw(img)
        bb = draw.textbbox((0, 0), char, font=font)
        x = size // 2 - (bb[2] - bb[0]) // 2 - bb[0]
        y = size // 2 - (bb[3] - bb[1]) // 2 - bb[1]
        draw.text((x, y), char, fill=255, font=font)
        return np.asarray(img, dtype=np.uint8)

    def test_atlas_has_entries_for_alphanumeric_chars(self):
        """Atlas must cover at least 80 % of alphanumeric OCR whitelist chars."""
        atlas = self._get_atlas()
        alphanumeric = [c for c in app.OCR_WHITELIST if c.isalnum()]
        found = sum(1 for c in alphanumeric if c in atlas)
        self.assertGreaterEqual(
            found, len(alphanumeric) * 0.80,
            "Atlas should cover at least 80 % of alphanumeric whitelist chars",
        )

    def test_atlas_templates_are_2d_uint8_arrays(self):
        """Every atlas template must be a 2-D uint8 ndarray with the correct shape."""
        atlas = self._get_atlas()
        expected_shape = (app._ATLAS_TMPL_H, app._ATLAS_TMPL_W)
        for char, templates in atlas.items():
            self.assertGreater(len(templates), 0, f"No variants for '{char}'")
            for t in templates:
                self.assertIsInstance(t, np.ndarray, f"Template for '{char}' is not ndarray")
                self.assertEqual(t.ndim, 2, f"Template for '{char}' is not 2-D")
                self.assertEqual(
                    t.shape, expected_shape,
                    f"Template shape for '{char}': got {t.shape}, expected {expected_shape}",
                )
                self.assertEqual(t.dtype, np.uint8, f"Template dtype for '{char}' is {t.dtype}")

    def test_atlas_blank_cell_returns_none(self):
        """A blank (all-zero) cell patch must produce no atlas match."""
        atlas = self._get_atlas()
        blank = np.zeros((64, 40), dtype=np.uint8)
        result = app._atlas_classify_patch(blank, atlas)
        self.assertIsNone(result, "Blank cell must return None")

    def test_atlas_confusable_pairs_never_wrong(self):
        """For confusable pairs (0/O, 8/B, 5/S, 1/I) the atlas must never
        return the wrong partner — correct or None are both acceptable."""
        atlas = self._get_atlas()
        # Use the first available atlas font for synthetic test cells
        from pathlib import Path
        from PIL import ImageFont
        font_path = next(
            (p for p in app._ATLAS_FONT_PATHS if Path(p).exists()), None
        )
        if font_path is None:
            self.skipTest("No atlas font found")
        font = ImageFont.truetype(font_path, 64)

        confusable_pairs = [("0", "O"), ("8", "B"), ("5", "S"), ("1", "I")]
        for a, b in confusable_pairs:
            for test_char, wrong_partner in [(a, b), (b, a)]:
                arr = self._render_large(test_char, font)
                arr_b = np.asarray(
                    Image.fromarray(arr).filter(ImageFilter.GaussianBlur(0.6)),
                    dtype=np.uint8,
                )
                result = app._atlas_classify_patch(arr_b, atlas)
                classified = result[0] if result is not None else None
                self.assertNotEqual(
                    classified, wrong_partner,
                    f"'{test_char}' must not be classified as '{wrong_partner}' "
                    f"(got {result!r})",
                )

    def test_fuse_atlas_grid_fill_only(self):
        """fuse_atlas_grid must NEVER overwrite any non-empty cell (fill-only)."""
        grid = app.empty_grid()
        conf = app.empty_confidence_grid()
        grid[0][1] = "A"; conf[0][1] = 0.90  # high-confidence — must survive
        grid[0][2] = "B"; conf[0][2] = 0.10  # low-confidence — must also survive
        # col 3 empty — atlas may fill it

        ag = app.empty_grid()
        ac = app.empty_confidence_grid()
        ag[0][1] = "X"; ac[0][1] = 0.80   # atlas > existing → still must NOT overwrite
        ag[0][2] = "Y"; ac[0][2] = 0.80   # atlas > existing → still must NOT overwrite
        ag[0][3] = "Z"; ac[0][3] = 0.80   # empty slot → atlas MUST fill

        out_g, out_c = app.fuse_atlas_grid(grid, conf, ag, ac)
        self.assertEqual(out_g[0][1], "A", "High-confidence non-empty cell must not be overwritten")
        self.assertAlmostEqual(out_c[0][1], 0.90)
        self.assertEqual(out_g[0][2], "B", "Low-confidence non-empty cell must not be overwritten")
        self.assertAlmostEqual(out_c[0][2], 0.10)
        self.assertEqual(out_g[0][3], "Z", "Empty cell must be filled by atlas")
        self.assertLessEqual(out_c[0][3], app._ATLAS_MAX_CONFIDENCE + 1e-9)


class FuzzyPhraseTests(unittest.TestCase):
    """B3: fuzzy phrase normalisation in normalize_mcdu_phrase."""

    def test_fuzzy_normalizes_lrcspo_to_lrc_spd(self):
        # "LRCSPO" is one character off from "LRCSPD" (→ "LRC SPD")
        result = app.normalize_mcdu_phrase("LRCSPO")
        self.assertEqual(result, "LRC SPD")

    def test_fuzzy_does_not_fire_on_already_correct(self):
        # Exact match in the table fires before the fuzzy loop; result unchanged
        result = app.normalize_mcdu_phrase("LRC SPD")
        self.assertEqual(result, "LRC SPD")

    def test_fuzzy_does_not_fire_on_unrelated(self):
        # "XYZABC" has no close match — must be returned unchanged
        result = app.normalize_mcdu_phrase("XYZABC")
        self.assertEqual(result, "XYZABC")


class TitleRowTests(unittest.TestCase):
    """A5 Part 2: snap_title_row snaps row 0 to the MCDU title vocabulary."""

    def _make_grid_with_title(self, text: str, start_col: int = 15) -> list[list[str]]:
        grid = app.empty_grid()
        for i, ch in enumerate(text):
            col = start_col + i
            if app.FIRST_DATA_COL <= col <= app.LAST_DATA_COL:
                grid[0][col] = ch if ch != " " else ""
        return grid

    def test_snap_title_row_corrects_one_char_typo(self):
        # "ACT RTA CR2" — last char wrong; should snap to "ACT RTA CRZ"
        grid = self._make_grid_with_title("ACT RTA CR2", start_col=14)
        out = app.snap_title_row(grid)
        title_text = "".join(out[0][c] or " " for c in range(14, 14 + 11)).strip()
        self.assertEqual(title_text, "ACT RTA CRZ")

    def test_snap_title_row_correct_title_unchanged(self):
        # Exact match — cells must be byte-identical
        grid = self._make_grid_with_title("ACT LRC D/D", start_col=14)
        out = app.snap_title_row(grid)
        for c in range(14, 14 + 11):
            self.assertEqual(out[0][c], grid[0][c])

    def test_snap_title_row_garbage_unchanged(self):
        # Random noise that doesn't match anything — must be untouched
        grid = self._make_grid_with_title("QQQQQQQQQQ", start_col=10)
        original = [row[:] for row in grid]
        out = app.snap_title_row(grid)
        self.assertEqual(out[0], original[0])


class CharSpacingTests(unittest.TestCase):
    """B4: apply_char_box_spacing inserts blank cells where char boxes reveal gaps."""

    def _make_geometry(self, cell_w: float = 40.0, cell_h: float = 60.0) -> dict:
        return {"cell_w": cell_w, "cell_h": cell_h, "origin_x": 0.0, "origin_y": 0.0}

    def _make_box(self, left: float, top: float, w: float, h: float, text: str = "X") -> dict:
        return {"left": left, "top": top, "width": w, "height": h, "text": text}

    def test_apply_char_box_spacing_inserts_gap(self):
        # Two boxes at adjacent cols 5 & 6 with centre gap of 70 px (> 1.7 × 40).
        # cx_a = 195 + 10 = 205 → col int(205/40) = 5
        # cx_b = 265 + 10 = 275 → col int(275/40) = 6
        # gap = 275 - 205 = 70 > 1.7 × 40 = 68 → blank inserted at col 6, B shifts to 7
        geo = self._make_geometry(cell_w=40.0, cell_h=60.0)
        grid = app.empty_grid()
        grid[0][5] = "A"
        grid[0][6] = "B"
        box_a = self._make_box(left=195.0, top=10.0, w=20.0, h=40.0, text="A")
        box_b = self._make_box(left=265.0, top=10.0, w=20.0, h=40.0, text="B")
        out = app.apply_char_box_spacing(grid, [box_a, box_b], geo)
        self.assertEqual(out[0][5], "A", "col 5 must keep A")
        self.assertEqual(out[0][6], "", "col 6 must be blank (gap inserted)")
        self.assertEqual(out[0][7], "B", "col 7 must have B shifted right")

    def test_apply_char_box_spacing_leaves_correct_spacing_unchanged(self):
        # Two boxes with close centres (< 1.7×cell_w) — no shift should happen
        geo = self._make_geometry(cell_w=40.0, cell_h=60.0)
        grid = app.empty_grid()
        grid[0][5] = "A"
        grid[0][6] = "B"
        box_a = self._make_box(left=195.0, top=10.0, w=20.0, h=40.0, text="A")  # cx=205
        box_b = self._make_box(left=235.0, top=10.0, w=20.0, h=40.0, text="B")  # cx=245, gap=40 < 68
        out = app.apply_char_box_spacing(grid, [box_a, box_b], geo)
        self.assertEqual(out[0][5], "A")
        self.assertEqual(out[0][6], "B")
        self.assertEqual(out[0][7], "")


class DateTokenSnapTests(unittest.TestCase):
    """_snap_date_token: MCDU date format correction with OCR confusion handling."""

    # --- real observed failures ---

    def test_13char_multi_confusion_real_failure(self):
        # "AUG07OCT02/25" misread as "AUGOZOGT02725":
        #   0→O (digit read as letter), 7→Z, C→G, /→7
        self.assertEqual(app._snap_date_token("AUGOZOGT02725"), "AUG07OCT02/25")

    def test_13char_zero_for_o_in_oct(self):
        # disambiguate_o_zero artefact: the "O" in OCT is after digit "7",
        # so it gets converted to "0", producing "AUG070CT02/25"
        self.assertEqual(app._snap_date_token("AUG070CT02/25"), "AUG07OCT02/25")

    def test_13char_sep_s_confused_with_5(self):
        # "S" in SEP read as "5"; slash read as "7"
        self.assertEqual(app._snap_date_token("5EP04OCT02725"), "SEP04OCT02/25")

    def test_8char_slash_read_as_7(self):
        # Simple 8-char date with "/" misread as "7"
        self.assertEqual(app._snap_date_token("JAN01725"), "JAN01/25")

    def test_13char_dec_c_confused_with_g(self):
        # "C" in DEC read as "G" (shape confusion)
        self.assertEqual(app._snap_date_token("AUG07DEG02/25"), "AUG07DEC02/25")

    # --- already-correct tokens must pass through unchanged ---

    def test_8char_correct_unchanged(self):
        self.assertIsNone(app._snap_date_token("JAN01/25"))

    def test_13char_correct_unchanged(self):
        self.assertIsNone(app._snap_date_token("AUG07OCT02/25"))

    def test_13char_sep_correct_unchanged(self):
        self.assertIsNone(app._snap_date_token("SEP04OCT02/25"))

    # --- tokens that do NOT fit the date shape must be rejected ---

    def test_non_date_13char_rejected(self):
        # Starts with "ACT" which is not a month
        self.assertIsNone(app._snap_date_token("ACTRTACRZABCD"))

    def test_wrong_length_11_rejected(self):
        self.assertIsNone(app._snap_date_token("AUG07OCT/25"))

    def test_wrong_length_14_rejected(self):
        self.assertIsNone(app._snap_date_token("AUGOZOGT027250"))

    def test_unknown_month_rejected(self):
        # "XYZ" is not a recognisable month under any confusion
        self.assertIsNone(app._snap_date_token("XYZ07OCT02/25"))

    def test_invalid_day_zero_rejected(self):
        # Day "00" is not a valid calendar day
        self.assertIsNone(app._snap_date_token("AUG00OCT02/25"))

    def test_totally_unrelated_token_rejected(self):
        self.assertIsNone(app._snap_date_token("FL390"))
        self.assertIsNone(app._snap_date_token("RECMDSPD"))
        self.assertIsNone(app._snap_date_token(".840"))


if __name__ == "__main__":
    unittest.main()
