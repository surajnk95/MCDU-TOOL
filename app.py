from __future__ import annotations

import base64
import csv
import io
import json
import os
import re
import shutil
import subprocess
import tempfile
import uuid
from http import HTTPStatus
from http.server import SimpleHTTPRequestHandler, ThreadingHTTPServer
from pathlib import Path
from typing import Any
from urllib.parse import urlparse

from docx import Document
from docx.enum.table import WD_CELL_VERTICAL_ALIGNMENT
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.shared import Pt
import math

import numpy as np
from PIL import Image, ImageEnhance, ImageFilter, ImageOps


ROOT = Path(__file__).resolve().parent
STATIC = ROOT / "static"
DATA = ROOT / "data"
EXPORTS = DATA / "exports"
CORRECTIONS = DATA / "corrections.json"
TEMPLATES = DATA / "templates.json"

ROWS = 13
COLS = 40
FIRST_DATA_COL = 1
LAST_DATA_COL = 38
SCREEN_W = 1600
MIN_SCREEN_H = 900
MAX_SCREEN_H = 1400


def find_tesseract() -> str:
    configured = os.environ.get("TESSERACT_CMD")
    if configured:
        return configured
    found = shutil.which("tesseract")
    if found:
        return found
    candidates = [
        Path("/opt/homebrew/bin/tesseract"),
        Path("/usr/local/bin/tesseract"),
        Path("C:/Program Files/Tesseract-OCR/tesseract.exe"),
        Path("C:/Program Files (x86)/Tesseract-OCR/tesseract.exe"),
    ]
    for candidate in candidates:
        if candidate.exists():
            return str(candidate)
    return "tesseract"


TESSERACT = find_tesseract()


def ensure_dirs() -> None:
    DATA.mkdir(exist_ok=True)
    EXPORTS.mkdir(exist_ok=True)
    if not CORRECTIONS.exists():
        CORRECTIONS.write_text("{}", encoding="utf-8")
    if not TEMPLATES.exists():
        TEMPLATES.write_text("{}", encoding="utf-8")


def json_response(handler: SimpleHTTPRequestHandler, status: int, payload: dict[str, Any]) -> None:
    body = json.dumps(payload).encode("utf-8")
    handler.send_response(status)
    handler.send_header("Content-Type", "application/json")
    handler.send_header("Content-Length", str(len(body)))
    handler.end_headers()
    handler.wfile.write(body)


def read_json_body(handler: SimpleHTTPRequestHandler) -> dict[str, Any]:
    length = int(handler.headers.get("Content-Length", "0"))
    if length <= 0:
        return {}
    raw = handler.rfile.read(length)
    return json.loads(raw.decode("utf-8"))


def load_image(data_url: str) -> Image.Image:
    if "," in data_url:
        data_url = data_url.split(",", 1)[1]
    raw = base64.b64decode(data_url)
    return ImageOps.exif_transpose(Image.open(io.BytesIO(raw))).convert("RGB")


def edge_length(a: dict[str, float], b: dict[str, float]) -> float:
    return math.hypot(a["x"] - b["x"], a["y"] - b["y"])


def screen_size_from_corners(corners: list[dict[str, float]]) -> tuple[int, int]:
    top = edge_length(corners[0], corners[1])
    right = edge_length(corners[1], corners[2])
    bottom = edge_length(corners[2], corners[3])
    left = edge_length(corners[3], corners[0])
    width = max(1.0, (top + bottom) * 0.5)
    height = max(1.0, (left + right) * 0.5)
    aspect = width / height
    normalized_height = int(round(SCREEN_W / aspect))
    return SCREEN_W, max(MIN_SCREEN_H, min(MAX_SCREEN_H, normalized_height))


def solve_linear_system(matrix: list[list[float]], values: list[float]) -> list[float]:
    n = len(values)
    augmented = [row[:] + [values[i]] for i, row in enumerate(matrix)]

    for pivot in range(n):
        max_row = max(range(pivot, n), key=lambda row: abs(augmented[row][pivot]))
        augmented[pivot], augmented[max_row] = augmented[max_row], augmented[pivot]
        pivot_value = augmented[pivot][pivot]
        if abs(pivot_value) < 1e-12:
            raise ValueError("Screen corners are too close together to calculate perspective.")

        for col in range(pivot, n + 1):
            augmented[pivot][col] /= pivot_value

        for row in range(n):
            if row == pivot:
                continue
            factor = augmented[row][pivot]
            for col in range(pivot, n + 1):
                augmented[row][col] -= factor * augmented[pivot][col]

    return [augmented[row][n] for row in range(n)]


def perspective_coefficients(src: list[dict[str, float]], dst_size: tuple[int, int]) -> list[float]:
    width, height = dst_size
    dst = [
        {"x": 0.0, "y": 0.0},
        {"x": float(width), "y": 0.0},
        {"x": float(width), "y": float(height)},
        {"x": 0.0, "y": float(height)},
    ]

    matrix: list[list[float]] = []
    values: list[float] = []
    for d, s in zip(dst, src):
        x = d["x"]
        y = d["y"]
        u = s["x"]
        v = s["y"]
        matrix.append([x, y, 1, 0, 0, 0, -u * x, -u * y])
        values.append(u)
        matrix.append([0, 0, 0, x, y, 1, -v * x, -v * y])
        values.append(v)
    return solve_linear_system(matrix, values)


def warp_screen(image: Image.Image, corners: list[dict[str, float]]) -> Image.Image:
    screen_size = screen_size_from_corners(corners)
    coeffs = perspective_coefficients(corners, screen_size)
    warped = image.transform(
        screen_size,
        Image.Transform.PERSPECTIVE,
        coeffs,
        Image.Resampling.BICUBIC,
    )
    return warped


def preprocess_for_ocr(image: Image.Image) -> Image.Image:
    gray = ImageOps.grayscale(image)
    gray = ImageOps.invert(gray)
    gray = ImageEnhance.Contrast(gray).enhance(2.8)
    gray = gray.filter(ImageFilter.SHARPEN)
    return gray


def cell_feature(image: Image.Image, row: int, col: int) -> list[float]:
    screen_w, screen_h = image.size
    cell_w = screen_w / COLS
    cell_h = screen_h / ROWS
    pad_x = cell_w * 0.08
    pad_y = cell_h * 0.08
    box = (
        int(max(0, col * cell_w - pad_x)),
        int(max(0, row * cell_h - pad_y)),
        int(min(screen_w, (col + 1) * cell_w + pad_x)),
        int(min(screen_h, (row + 1) * cell_h + pad_y)),
    )
    crop = ImageOps.grayscale(image.crop(box)).resize((16, 24), Image.Resampling.BILINEAR)
    arr = np.asarray(crop).astype(np.float32)
    threshold = max(90.0, float(arr.mean() + arr.std() * 0.65))
    mask = (arr > threshold).astype(np.float32)
    return mask.reshape(-1).tolist()


def feature_distance(a: list[float], b: list[float]) -> float:
    if len(a) != len(b):
        return 1.0
    arr_a = np.asarray(a, dtype=np.float32)
    arr_b = np.asarray(b, dtype=np.float32)
    return float(np.mean(np.abs(arr_a - arr_b)))


def load_templates() -> dict[str, list[list[float]]]:
    ensure_dirs()
    try:
        data = json.loads(TEMPLATES.read_text(encoding="utf-8"))
    except json.JSONDecodeError:
        data = {}
    templates: dict[str, list[list[float]]] = {}
    for char, features in data.items():
        if not isinstance(char, str) or not char:
            continue
        if not isinstance(features, list):
            continue
        templates[char[:1]] = [feature for feature in features if isinstance(feature, list)]
    return templates


def save_templates(templates: dict[str, list[list[float]]]) -> None:
    ensure_dirs()
    compact = {char: features[-30:] for char, features in templates.items() if features}
    TEMPLATES.write_text(json.dumps(compact), encoding="utf-8")


def classify_from_templates(feature: list[float], templates: dict[str, list[list[float]]]) -> tuple[str, float] | None:
    best_char = ""
    best_distance = 1.0
    for char, features in templates.items():
        for candidate in features:
            distance = feature_distance(feature, candidate)
            if distance < best_distance:
                best_char = char
                best_distance = distance
    if best_char and best_distance <= 0.22:
        return best_char, best_distance
    return None


def run_tesseract_tsv(image: Image.Image) -> list[dict[str, Any]]:
    ensure_dirs()
    with tempfile.TemporaryDirectory() as temp_dir:
        source = Path(temp_dir) / "screen.png"
        image.save(source)

        command = [
            TESSERACT,
            str(source),
            "stdout",
            "--psm",
            "11",
            "-l",
            "eng",
            "-c",
            "tessedit_char_whitelist=ABCDEFGHIJKLMNOPQRSTUVWXYZ0123456789°/.-<>",
            "tsv",
        ]
        try:
            completed = subprocess.run(command, capture_output=True, text=True, check=False)
        except FileNotFoundError as exc:
            raise RuntimeError(
                "Tesseract OCR was not found. Install Tesseract OCR, or set TESSERACT_CMD to the full tesseract.exe path."
            ) from exc
        if completed.returncode != 0:
            raise RuntimeError(completed.stderr.strip() or "Tesseract OCR failed.")

        rows = csv.DictReader(io.StringIO(completed.stdout), delimiter="\t")
        words: list[dict[str, Any]] = []
        for row in rows:
            text = (row.get("text") or "").strip()
            if not text:
                continue
            try:
                conf = float(row.get("conf") or -1)
            except ValueError:
                conf = -1
            if conf < 0:
                continue
            words.append(
                {
                    "text": text,
                    "conf": conf,
                    "left": int(row.get("left") or 0),
                    "top": int(row.get("top") or 0),
                    "width": int(row.get("width") or 0),
                    "height": int(row.get("height") or 0),
                }
            )
        return words


def run_tesseract_boxes(image: Image.Image) -> list[dict[str, Any]]:
    ensure_dirs()
    width, height = image.size
    with tempfile.TemporaryDirectory() as temp_dir:
        source = Path(temp_dir) / "screen.png"
        image.save(source)

        command = [
            TESSERACT,
            str(source),
            "stdout",
            "--psm",
            "6",
            "-l",
            "eng",
            "-c",
            "tessedit_char_whitelist=ABCDEFGHIJKLMNOPQRSTUVWXYZ0123456789°/.-<>",
            "makebox",
        ]
        try:
            completed = subprocess.run(command, capture_output=True, text=True, check=False)
        except FileNotFoundError:
            return []
        if completed.returncode != 0:
            return []

        boxes: list[dict[str, Any]] = []
        for line in completed.stdout.splitlines():
            parts = line.split()
            if len(parts) < 5:
                continue
            char = clean_ocr_text(parts[0])
            if not char:
                continue
            try:
                left = int(parts[1])
                bottom = int(parts[2])
                right = int(parts[3])
                top = int(parts[4])
            except ValueError:
                continue
            boxes.append(
                {
                    "text": char[:1],
                    "left": left,
                    "top": max(0, height - top),
                    "width": max(0, right - left),
                    "height": max(0, top - bottom),
                }
            )
        return boxes


def clean_ocr_text(text: str) -> str:
    text = text.replace("|", "I")
    text = re.sub(r"\s+", " ", text)
    return text.strip()


def empty_grid() -> list[list[str]]:
    return [["" for _ in range(COLS)] for _ in range(ROWS)]


def clamp_data_col(col: int) -> int:
    return max(FIRST_DATA_COL, min(LAST_DATA_COL, col))


def nearest_data_col(x: float, cell_w: float) -> int:
    return clamp_data_col(int(round(x / cell_w)))


def normalize_grid_guards(grid: list[list[str]]) -> list[list[str]]:
    normalized = [[cell for cell in row[:COLS]] + [""] * max(0, COLS - len(row)) for row in grid[:ROWS]]
    for row in normalized:
        row[0] = ""
        row[COLS - 1] = ""
    return normalized


def place_char(grid: list[list[str]], char_box: dict[str, Any], screen_size: tuple[int, int]) -> bool:
    text = clean_ocr_text(str(char_box["text"]))[:1]
    if not text:
        return False

    screen_w, screen_h = screen_size
    cell_w = screen_w / COLS
    cell_h = screen_h / ROWS
    box_w = float(char_box.get("width") or 0)
    box_h = float(char_box.get("height") or 0)
    if box_w > cell_w * 2.15 or box_h > cell_h * 1.35 or box_w < 2 or box_h < 2:
        return False

    center_x = float(char_box["left"]) + box_w * 0.5
    center_y = float(char_box["top"]) + box_h * 0.5
    row = max(0, min(ROWS - 1, int(center_y / cell_h)))
    col = clamp_data_col(int(center_x / cell_w))

    if grid[row][col] and grid[row][col] != text:
        for offset in (-1, 1, -2, 2):
            next_col = col + offset
            if FIRST_DATA_COL <= next_col <= LAST_DATA_COL and not grid[row][next_col]:
                col = next_col
                break
    if not grid[row][col]:
        grid[row][col] = text
        return True
    return False


def place_word(grid: list[list[str]], word: dict[str, Any], screen_size: tuple[int, int]) -> None:
    text = clean_ocr_text(str(word["text"]))
    if not text:
        return

    screen_w, screen_h = screen_size
    cell_w = screen_w / COLS
    cell_h = screen_h / ROWS
    row = max(0, min(ROWS - 1, int((word["top"] + word["height"] * 0.5) / cell_h)))
    start_col = nearest_data_col(float(word["left"]), cell_w)

    compact = text.replace(" ", "")
    if not compact:
        return
    while start_col <= LAST_DATA_COL and grid[row][start_col]:
        start_col += 1
    if start_col > LAST_DATA_COL:
        return
    max_len = LAST_DATA_COL - start_col + 1
    for index, char in enumerate(compact[:max_len]):
        col = start_col + index
        if not grid[row][col]:
            grid[row][col] = char


def connected_components(mask: np.ndarray) -> list[dict[str, Any]]:
    height, width = mask.shape
    seen = np.zeros(mask.shape, dtype=bool)
    components: list[dict[str, Any]] = []
    for start_y in range(height):
        for start_x in range(width):
            if seen[start_y, start_x] or not mask[start_y, start_x]:
                continue
            stack = [(start_x, start_y)]
            seen[start_y, start_x] = True
            xs: list[int] = []
            ys: list[int] = []
            while stack:
                x, y = stack.pop()
                xs.append(x)
                ys.append(y)
                for next_x, next_y in ((x + 1, y), (x - 1, y), (x, y + 1), (x, y - 1)):
                    if (
                        0 <= next_x < width
                        and 0 <= next_y < height
                        and not seen[next_y, next_x]
                        and mask[next_y, next_x]
                    ):
                        seen[next_y, next_x] = True
                        stack.append((next_x, next_y))
            count = len(xs)
            if count < 250:
                continue
            x1 = min(xs)
            x2 = max(xs) + 1
            y1 = min(ys)
            y2 = max(ys) + 1
            components.append(
                {
                    "count": count,
                    "x1": x1,
                    "y1": y1,
                    "x2": x2,
                    "y2": y2,
                    "fill": count / max(1, (x2 - x1) * (y2 - y1)),
                }
            )
    return components


def detect_display(payload: dict[str, Any]) -> dict[str, Any]:
    image = load_image(str(payload["image"]))
    scale = min(1.0, 900 / max(image.size))
    small = image.resize((max(1, round(image.width * scale)), max(1, round(image.height * scale))))
    arr = np.asarray(small).astype(np.float32)
    gray = arr[:, :, 0] * 0.299 + arr[:, :, 1] * 0.587 + arr[:, :, 2] * 0.114
    mask = gray < 58
    components = connected_components(mask)
    if not components:
        raise ValueError("Could not find a dark MCDU display region.")

    scored: list[tuple[float, dict[str, Any]]] = []
    image_area = small.width * small.height
    for component in components:
        width = component["x2"] - component["x1"]
        height = component["y2"] - component["y1"]
        area = width * height
        aspect = width / max(1, height)
        touches_edge = (
            component["x1"] <= 2
            or component["y1"] <= 2
            or component["x2"] >= small.width - 2
            or component["y2"] >= small.height - 2
        )
        if area < image_area * 0.08 or aspect < 0.75 or aspect > 3.2:
            continue
        score = component["count"] * min(1.0, component["fill"] * 1.4)
        if touches_edge:
            score *= 0.35
        scored.append((score, component))

    if not scored:
        raise ValueError("Could not isolate the black display. Drag the four corners manually.")

    _, best = max(scored, key=lambda item: item[0])
    pad_x = max(1, round((best["x2"] - best["x1"]) * 0.003))
    pad_y = max(1, round((best["y2"] - best["y1"]) * 0.003))
    x1 = max(0, best["x1"] + pad_x)
    y1 = max(0, best["y1"] + pad_y)
    x2 = min(small.width, best["x2"] - pad_x)
    y2 = min(small.height, best["y2"] - pad_y)
    inv = 1 / scale
    corners = [
        {"x": x1 * inv, "y": y1 * inv},
        {"x": x2 * inv, "y": y1 * inv},
        {"x": x2 * inv, "y": y2 * inv},
        {"x": x1 * inv, "y": y2 * inv},
    ]
    return {
        "corners": corners,
        "confidence": round(float(best["fill"]), 3),
        "displaySize": {"width": round((x2 - x1) * inv), "height": round((y2 - y1) * inv)},
    }


def flatten_display(payload: dict[str, Any]) -> dict[str, Any]:
    image = load_image(str(payload["image"]))
    corners = payload.get("corners")
    if not isinstance(corners, list) or len(corners) != 4:
        raise ValueError("Four screen corners are required.")

    warped = warp_screen(image, corners)
    preview_id = f"{uuid.uuid4().hex}.png"
    preview_path = EXPORTS / preview_id
    warped.save(preview_path)
    return {
        "previewUrl": f"/data/exports/{preview_id}",
        "width": warped.width,
        "height": warped.height,
    }


def load_corrections() -> dict[str, str]:
    ensure_dirs()
    try:
        data = json.loads(CORRECTIONS.read_text(encoding="utf-8"))
    except json.JSONDecodeError:
        data = {}
    return {str(key): str(value) for key, value in data.items()}


def save_corrections(corrections: dict[str, str]) -> None:
    ensure_dirs()
    CORRECTIONS.write_text(json.dumps(corrections, indent=2, sort_keys=True), encoding="utf-8")


def apply_corrections(grid: list[list[str]]) -> list[list[str]]:
    corrections = load_corrections()
    updated = [[cell for cell in row] for row in grid]
    for row_index, row in enumerate(updated):
        row_text = "".join(cell or " " for cell in row)
        normalized = re.sub(r"\s+", " ", row_text).strip()
        if normalized in corrections:
            corrected = corrections[normalized][:COLS].ljust(COLS)
            updated[row_index] = list(corrected)
    return normalize_grid_guards(updated)


def apply_templates(grid: list[list[str]], warped: Image.Image) -> list[list[str]]:
    templates = load_templates()
    if not templates:
        return grid
    updated = [[cell for cell in row] for row in grid]
    gray = np.asarray(ImageOps.grayscale(warped)).astype(np.float32)
    screen_w, screen_h = warped.size
    cell_w = screen_w / COLS
    cell_h = screen_h / ROWS
    for row in range(ROWS):
        for col in range(COLS):
            x1 = int(col * cell_w)
            x2 = int((col + 1) * cell_w)
            y1 = int(row * cell_h)
            y2 = int((row + 1) * cell_h)
            crop = gray[y1:y2, x1:x2]
            if crop.size == 0 or float(crop.max()) < 105:
                continue
            if col < FIRST_DATA_COL or col > LAST_DATA_COL:
                continue
            classified = classify_from_templates(cell_feature(warped, row, col), templates)
            if classified and not updated[row][col]:
                updated[row][col] = classified[0]
    return updated


def grid_character_count(grid: list[list[str]]) -> int:
    return sum(1 for row in grid for cell in row[FIRST_DATA_COL : LAST_DATA_COL + 1] if cell)


def remember_templates(payload: dict[str, Any]) -> dict[str, Any]:
    image = load_image(str(payload["image"]))
    corners = payload.get("corners")
    grid = payload.get("grid")
    if not isinstance(corners, list) or len(corners) != 4:
        raise ValueError("Four screen corners are required for template learning.")
    if not isinstance(grid, list) or len(grid) != ROWS:
        raise ValueError("A corrected 13-row grid is required for template learning.")

    warped = warp_screen(image, corners)
    templates = load_templates()
    learned = 0
    for row_index, row in enumerate(grid):
        if not isinstance(row, list):
            continue
        for col_index, value in enumerate(row[:COLS]):
            if col_index < FIRST_DATA_COL or col_index > LAST_DATA_COL:
                continue
            char = clean_ocr_text(str(value))[:1]
            if not char or char.isspace():
                continue
            templates.setdefault(char, []).append(cell_feature(warped, row_index, col_index))
            learned += 1
    save_templates(templates)
    return {"learned": learned, "characters": len(templates)}


def analyze(payload: dict[str, Any]) -> dict[str, Any]:
    image = load_image(str(payload["image"]))
    corners = payload.get("corners")
    if not isinstance(corners, list) or len(corners) != 4:
        raise ValueError("Four screen corners are required.")

    warped = warp_screen(image, corners)
    processed = preprocess_for_ocr(warped)
    screen_size = warped.size
    boxes = run_tesseract_boxes(processed)
    words = run_tesseract_tsv(processed)

    grid = empty_grid()
    for word in words:
        place_word(grid, word, screen_size)
    if grid_character_count(grid) < 8:
        for box in boxes:
            place_char(grid, box, screen_size)
    corrected_grid = normalize_grid_guards(apply_corrections(grid))

    preview_id = f"{uuid.uuid4().hex}.png"
    preview_path = EXPORTS / preview_id
    warped.save(preview_path)

    return {
        "grid": corrected_grid,
        "words": words,
        "boxes": boxes,
        "previewUrl": f"/data/exports/{preview_id}",
    }


def add_table_header(table: Any) -> None:
    header = table.rows[0].cells
    header[0].text = "Row"
    for col in range(COLS):
        label = str(col) if 1 <= col <= 38 else ""
        header[col + 1].text = label
        for paragraph in header[col + 1].paragraphs:
            paragraph.alignment = WD_ALIGN_PARAGRAPH.CENTER


def export_docx(payload: dict[str, Any]) -> dict[str, str]:
    grid = payload.get("grid")
    if not isinstance(grid, list) or len(grid) != ROWS:
        raise ValueError("A 13-row grid is required.")

    document = Document()
    document.add_heading("777-9 MCDU Grid Extraction", level=1)
    document.add_paragraph("Rows are numbered 1 to 13. Screen columns 2 to 39 are labelled 1 to 38.")

    table = document.add_table(rows=ROWS + 1, cols=COLS + 1)
    table.style = "Table Grid"
    add_table_header(table)

    for row_index in range(ROWS):
        cells = table.rows[row_index + 1].cells
        cells[0].text = str(row_index + 1)
        for col_index in range(COLS):
            value = ""
            if col_index in (0, COLS - 1):
                value = ""
            elif row_index < len(grid) and isinstance(grid[row_index], list) and col_index < len(grid[row_index]):
                value = str(grid[row_index][col_index])
            cells[col_index + 1].text = value.strip()
            cells[col_index + 1].vertical_alignment = WD_CELL_VERTICAL_ALIGNMENT.CENTER
            for paragraph in cells[col_index + 1].paragraphs:
                paragraph.alignment = WD_ALIGN_PARAGRAPH.CENTER
                for run in paragraph.runs:
                    run.font.size = Pt(7)

    filename = f"mcmdu-grid-{uuid.uuid4().hex[:8]}.docx"
    path = EXPORTS / filename
    document.save(path)
    return {"url": f"/data/exports/{filename}", "filename": filename}


def remember(payload: dict[str, Any]) -> dict[str, Any]:
    original_raw = str(payload.get("original", ""))
    corrected_raw = str(payload.get("corrected", ""))
    original = re.sub(r"\s+", " ", original_raw).strip()
    corrected_list = list(corrected_raw[:COLS].ljust(COLS))
    corrected_list[0] = " "
    corrected_list[COLS - 1] = " "
    corrected = "".join(corrected_list)
    if not original or not corrected.strip():
        return {"count": len(load_corrections()), "skipped": True}

    corrections = load_corrections()
    corrections[original] = corrected
    save_corrections(corrections)
    return {"count": len(corrections)}


class McmduHandler(SimpleHTTPRequestHandler):
    def translate_path(self, path: str) -> str:
        parsed = urlparse(path).path
        if parsed.startswith("/data/exports/"):
            return str((ROOT / parsed.lstrip("/")).resolve())
        if parsed == "/":
            return str(STATIC / "index.html")
        return str((STATIC / parsed.lstrip("/")).resolve())

    def do_POST(self) -> None:
        try:
            payload = read_json_body(self)
            if self.path == "/api/analyze":
                json_response(self, HTTPStatus.OK, analyze(payload))
            elif self.path == "/api/detect-display":
                json_response(self, HTTPStatus.OK, detect_display(payload))
            elif self.path == "/api/flatten-display":
                json_response(self, HTTPStatus.OK, flatten_display(payload))
            elif self.path == "/api/export-docx":
                json_response(self, HTTPStatus.OK, export_docx(payload))
            elif self.path == "/api/remember":
                json_response(self, HTTPStatus.OK, remember(payload))
            elif self.path == "/api/remember-templates":
                json_response(self, HTTPStatus.OK, remember_templates(payload))
            else:
                json_response(self, HTTPStatus.NOT_FOUND, {"error": "Unknown endpoint."})
        except Exception as exc:  # noqa: BLE001 - API boundary should return useful errors.
            print(f"{self.path} failed: {exc}")
            json_response(self, HTTPStatus.BAD_REQUEST, {"error": str(exc)})


def main() -> None:
    ensure_dirs()
    port = int(os.environ.get("PORT", "8766"))
    server = ThreadingHTTPServer(("127.0.0.1", port), McmduHandler)
    print(f"777-9 MCDU Grid Tool running at http://127.0.0.1:{port}")
    server.serve_forever()


if __name__ == "__main__":
    main()
