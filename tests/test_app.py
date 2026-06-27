import math
import tempfile
import unittest
from pathlib import Path
from unittest.mock import patch

import app
import numpy as np
from PIL import Image, ImageDraw, ImageOps


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
        origin_x, origin_y = app.estimate_grid_origin(image)
        self.assertAlmostEqual(origin_x, target_x, delta=1.5)
        self.assertAlmostEqual(origin_y, target_y, delta=1.5)

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
        # max-channel makes magenta bright (220 > 90 threshold) → non-zero mask entries
        self.assertGreater(sum(feature), 0)


if __name__ == "__main__":
    unittest.main()
