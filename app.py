from __future__ import annotations

import base64
import csv
import difflib
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
OCR_WHITELIST = "ABCDEFGHIJKLMNOPQRSTUVWXYZ0123456789°º˚/.-<>"
MCDU_WORD_HINTS = {
    "ACT",
    "RTE",
    "LEGS",
    "CRZ",
    "ALT",
    "SPD",
    "LRC",
    "D/D",
    "ECON",
    "RTA",
    "STEP",
    "ENG",
    "OUT",
    "FUEL",
    "RECMD",
    "MAX",
    "OPT",
    "TO",
    "SEL",
    "FL",
}
MCDU_VOCABULARY = {
    "ACT",
    "ALT",
    "ANRU",
    "CRZ",
    "DATA",
    "DES",
    "ECON",
    "ENG",
    "FUEL",
    "GPS",
    "INDEX",
    "IRU",
    "LEGS",
    "LRC",
    "MAX",
    "NAV",
    "NM",
    "OFF",
    "ON",
    "OPT",
    "OUT",
    "POS",
    "RADIO",
    "RECMD",
    "REF",
    "RTA",
    "RTE",
    "SEL",
    "SELECT",
    "SENSOR",
    "SPD",
    "STEP",
    "TO",
    "TRUE",
}
MCDU_PHRASE_REPLACEMENTS = (
    ("ACTRTACRZ", "ACT RTA CRZ"),
    ("ACTLRCD/D", "ACT LRC D/D"),
    ("MODLRCD/D", "MOD LRC D/D"),
    ("RTAPROGRESS", "RTA PROGRESS"),
    ("DESFORECAST", "DES FORECAST"),
    ("OFFPATHDES", "OFFPATH DES"),
    ("ENGOUT", "ENG OUT"),
    ("ALLENG", "ALL ENG"),
    ("CRZALT", "CRZ ALT"),
    ("RTASPD", "RTA SPD"),
    ("SELSPD", "SEL SPD"),
    ("RECHD", "RECMD"),
    ("NAX", "MAX"),
    ("TA/PUEL", "ETA/FUEL"),
    ("TA/FUEL", "ETA/FUEL"),
    ("PUEL", "FUEL"),
    ("KOFIETA/FUEL", "KBFI ETA/FUEL"),
    ("ETA/FUEL", "ETA/FUEL"),
    ("ETAFUEL", "ETA/FUEL"),
)


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


def word_quality(word: dict[str, Any]) -> float:
    text = clean_ocr_text(str(word.get("text", ""))).upper()
    compact = text.replace(" ", "")
    score = max(0.0, float(word.get("conf") or 0.0))
    score += min(len(compact), 12) * 3.0
    for hint in MCDU_WORD_HINTS:
        if hint in compact:
            score += 35.0
    if re.fullmatch(r"\d+/\d+", compact):
        score += 25.0
    if re.fullmatch(r"FL\d{2,3}", compact):
        score += 30.0
    return score


def run_tesseract_tsv(image: Image.Image) -> list[dict[str, Any]]:
    ensure_dirs()
    with tempfile.TemporaryDirectory() as temp_dir:
        source = Path(temp_dir) / "screen.png"
        image.save(source)

        candidates: list[list[dict[str, Any]]] = []
        last_error = ""
        for psm in ("11", "12", "6", "3", "4"):
            command = [
                TESSERACT,
                str(source),
                "stdout",
                "--psm",
                psm,
                "-l",
                "eng",
                "-c",
                f"tessedit_char_whitelist={OCR_WHITELIST}",
                "tsv",
            ]
            try:
                completed = subprocess.run(command, capture_output=True, text=True, check=False)
            except FileNotFoundError as exc:
                raise RuntimeError(
                    "Tesseract OCR was not found. Install Tesseract OCR, or set TESSERACT_CMD to the full tesseract.exe path."
                ) from exc
            if completed.returncode != 0:
                last_error = completed.stderr.strip()
                continue

            rows = csv.DictReader(io.StringIO(completed.stdout), delimiter="\t")
            words: list[dict[str, Any]] = []
            for row in rows:
                text = clean_ocr_text((row.get("text") or "").strip())
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
                        "psm": psm,
                    }
                )
            if words:
                candidates.append(words)
        if not candidates and last_error:
            raise RuntimeError(last_error or "Tesseract OCR failed.")
        if not candidates:
            return []

        def candidate_score(candidate: list[dict[str, Any]]) -> float:
            score = sum(word_quality(word) for word in candidate)
            row_count = len(
                {
                    int((float(word["top"]) + float(word["height"]) * 0.5) / max(1.0, image.height / ROWS))
                    for word in candidate
                }
            )
            score += row_count * 18.0
            score -= sum(20.0 for word in candidate if len(str(word["text"])) > 18)
            return score

        best = max(candidates, key=candidate_score)
        return sorted(best, key=lambda item: (int(item["top"]), int(item["left"])))


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
            f"tessedit_char_whitelist={OCR_WHITELIST}",
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
    text = text.replace("º", "°").replace("˚", "°").replace("Â°", "°")
    text = re.sub(r"(?<=\d{3})[oO07](?=/)", "°", text)
    text = re.sub(r"\b(\d{3})(?=/\d)", r"\1°", text)
    text = re.sub(r"\s+", " ", text)
    return text.strip()


def normalize_mcdu_phrase(text: str) -> str:
    text = clean_ocr_text(text).upper()
    compact = text.replace(" ", "")
    suffix = ">" if compact.endswith(">") else ""
    core = compact[:-1] if suffix else compact
    for source, replacement in MCDU_PHRASE_REPLACEMENTS:
        if core == source:
            return f"{replacement}{suffix}"
        if source in core and len(core) <= len(source) + 4:
            return f"{core.replace(source, replacement)}{suffix}"
    for source, replacement in (("RECHD", "RECMD"), ("NAX", "MAX"), ("PUEL", "FUEL"), ("KOFI", "KBFI")):
        if source in core:
            core = core.replace(source, replacement)
    original_core = compact[:-1] if suffix else compact
    if core != original_core:
        return f"{core}{suffix}"
    return text


def mcdu_row_score(text: str) -> float:
    normalized = normalize_mcdu_phrase(text)
    compact = normalized.replace(" ", "")
    score = 0.0
    for word in MCDU_VOCABULARY:
        if word in normalized.split() or word in compact:
            score += 8.0
    for hint in MCDU_WORD_HINTS:
        if hint in compact:
            score += 12.0
    if re.search(r"\d+/\d+", compact):
        score += 10.0
    if re.search(r"FL\d{2,3}", compact):
        score += 10.0
    score += compact.count("<") * 4.0 + compact.count(">") * 4.0
    score += min(len(compact), 20) * 0.4
    score -= len(re.findall(r"(?<![A-Z0-9])[A-Z0-9](?![A-Z0-9])", normalized)) * 2.5
    return score


def row_string_from_cells(row: list[str]) -> str:
    return "".join(cell or " " for cell in row[:COLS]).rstrip()


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


def calibrate_axis(centers: list[float], nominal_pitch: float, count: int) -> tuple[float, float]:
    if len(centers) < 6:
        return 0.0, nominal_pitch
    values = np.asarray(centers, dtype=np.float32)
    best = (float("inf"), 0.0, nominal_pitch)
    for scale in np.linspace(0.985, 1.015, 13):
        pitch = nominal_pitch * float(scale)
        indexes = np.clip(np.round(values / pitch - 0.5), 0, count - 1)
        offsets = values - (indexes + 0.5) * pitch
        origin = max(-nominal_pitch * 0.16, min(nominal_pitch * 0.16, float(np.median(offsets))))
        indexes = np.clip(np.round((values - origin) / pitch - 0.5), 0, count - 1)
        residuals = np.abs(values - (origin + (indexes + 0.5) * pitch))
        score = float(np.median(residuals)) + abs(origin) * 0.35 + abs(scale - 1.0) * nominal_pitch
        if score < best[0]:
            best = (score, origin, pitch)
    return best[1], best[2]


def calibrate_grid(
    char_boxes: list[dict[str, Any]],
    screen_size: tuple[int, int],
) -> dict[str, float]:
    screen_w, screen_h = screen_size
    x_centers = [
        float(box["left"]) + float(box["width"]) * 0.5
        for box in char_boxes
        if 2 <= float(box.get("width") or 0) <= screen_w / COLS * 2.1
    ]
    y_centers = [
        float(box["top"]) + float(box["height"]) * 0.5
        for box in char_boxes
        if 2 <= float(box.get("height") or 0) <= screen_h / ROWS * 1.4
    ]
    origin_x, cell_w = calibrate_axis(x_centers, screen_w / COLS, COLS)
    origin_y, cell_h = calibrate_axis(y_centers, screen_h / ROWS, ROWS)
    return {"origin_x": origin_x, "origin_y": origin_y, "cell_w": cell_w, "cell_h": cell_h}


def place_char(grid: list[list[str]], char_box: dict[str, Any], geometry: dict[str, float]) -> bool:
    text = clean_ocr_text(str(char_box["text"]))[:1]
    if not text:
        return False

    cell_w = geometry["cell_w"]
    cell_h = geometry["cell_h"]
    box_w = float(char_box.get("width") or 0)
    box_h = float(char_box.get("height") or 0)
    if box_w > cell_w * 2.15 or box_h > cell_h * 1.35 or box_w < 2 or box_h < 2:
        return False

    center_x = float(char_box["left"]) + box_w * 0.5
    center_y = float(char_box["top"]) + box_h * 0.5
    row = max(0, min(ROWS - 1, int((center_y - geometry["origin_y"]) / cell_h)))
    col = clamp_data_col(int((center_x - geometry["origin_x"]) / cell_w))

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


def place_word(
    grid: list[list[str]],
    scores: list[list[float]],
    word: dict[str, Any],
    geometry: dict[str, float],
) -> None:
    text = normalize_mcdu_phrase(str(word["text"]))
    if not text:
        return

    cell_w = geometry["cell_w"]
    cell_h = geometry["cell_h"]
    row = max(
        0,
        min(
            ROWS - 1,
            int((word["top"] + word["height"] * 0.5 - geometry["origin_y"]) / cell_h),
        ),
    )
    start_col = clamp_data_col(int(round((float(word["left"]) - geometry["origin_x"]) / cell_w)))

    compact = text.replace(" ", "")
    if not compact:
        return

    quality = word_quality(word)
    box_cols = max(1, int(round(float(word.get("width") or 1) / cell_w)))
    sequence = text if " " in text and len(text) <= box_cols + 2 else compact
    use_projected_spacing = " " not in sequence and box_cols > len(compact) + 1
    max_len = LAST_DATA_COL - start_col + 1

    for index, char in enumerate(sequence[:max_len]):
        if char.isspace():
            continue
        if use_projected_spacing:
            center_x = float(word["left"]) + float(word["width"]) * ((index + 0.5) / len(compact))
            col = clamp_data_col(int((center_x - geometry["origin_x"]) / cell_w))
        else:
            col = start_col + index
        if not grid[row][col] or quality > scores[row][col] + 10:
            grid[row][col] = char
            scores[row][col] = quality


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
                    "xs": xs,
                    "ys": ys,
                }
            )
    return components


def robust_line_fit(independent: np.ndarray, dependent: np.ndarray) -> tuple[float, float] | None:
    if independent.size < 12 or dependent.size != independent.size:
        return None
    keep = np.ones(independent.size, dtype=bool)
    slope = 0.0
    intercept = float(np.median(dependent))
    for _ in range(3):
        if int(np.sum(keep)) < 8:
            return None
        slope, intercept = np.polyfit(independent[keep], dependent[keep], 1)
        residuals = np.abs(dependent - (slope * independent + intercept))
        cutoff = max(1.5, float(np.percentile(residuals[keep], 72)))
        keep = residuals <= cutoff
    return float(slope), float(intercept)


def intersect_edge_lines(
    vertical: tuple[float, float],
    horizontal: tuple[float, float],
) -> tuple[float, float] | None:
    # vertical: x = a*y + b; horizontal: y = c*x + d
    a, b = vertical
    c, d = horizontal
    denominator = 1.0 - a * c
    if abs(denominator) < 1e-6:
        return None
    x = (a * d + b) / denominator
    y = c * x + d
    return float(x), float(y)


def component_corners(component: dict[str, Any], scale: float) -> list[dict[str, float]]:
    x1 = float(component["x1"])
    y1 = float(component["y1"])
    x2 = float(component["x2"])
    y2 = float(component["y2"])
    fallback = [(x1, y1), (x2, y1), (x2, y2), (x1, y2)]
    xs = np.asarray(component.get("xs") or [], dtype=np.float32)
    ys = np.asarray(component.get("ys") or [], dtype=np.float32)
    inside = (xs >= x1) & (xs <= x2) & (ys >= y1) & (ys <= y2)
    xs = xs[inside]
    ys = ys[inside]
    if xs.size < 100 or ys.size < 100:
        points = fallback
    else:
        width = max(1.0, x2 - x1)
        height = max(1.0, y2 - y1)
        unique_y = np.unique(ys.astype(np.int32))
        left_x: list[float] = []
        right_x: list[float] = []
        edge_y: list[float] = []
        for y in unique_y:
            row_x = xs[ys.astype(np.int32) == y]
            if row_x.size < 4:
                continue
            edge_y.append(float(y))
            left_x.append(float(np.percentile(row_x, 2)))
            right_x.append(float(np.percentile(row_x, 98)))

        unique_x = np.unique(xs.astype(np.int32))
        top_y: list[float] = []
        bottom_y: list[float] = []
        edge_x: list[float] = []
        for x in unique_x:
            col_y = ys[xs.astype(np.int32) == x]
            if col_y.size < 4:
                continue
            edge_x.append(float(x))
            top_y.append(float(np.percentile(col_y, 2)))
            bottom_y.append(float(np.percentile(col_y, 98)))

        edge_y_arr = np.asarray(edge_y, dtype=np.float32)
        edge_x_arr = np.asarray(edge_x, dtype=np.float32)
        left_line = robust_line_fit(edge_y_arr, np.asarray(left_x, dtype=np.float32))
        right_line = robust_line_fit(edge_y_arr, np.asarray(right_x, dtype=np.float32))
        top_line = robust_line_fit(edge_x_arr, np.asarray(top_y, dtype=np.float32))
        bottom_line = robust_line_fit(edge_x_arr, np.asarray(bottom_y, dtype=np.float32))
        intersections = (
            [
                intersect_edge_lines(left_line, top_line),
                intersect_edge_lines(right_line, top_line),
                intersect_edge_lines(right_line, bottom_line),
                intersect_edge_lines(left_line, bottom_line),
            ]
            if left_line and right_line and top_line and bottom_line
            else []
        )
        if not intersections or any(point is None for point in intersections):
            points = fallback
        else:
            points = [point for point in intersections if point is not None]
            margin_x = width * 0.12
            margin_y = height * 0.12
            corner_shift = max(
                max(abs(px - ex) / width, abs(py - ey) / height)
                for (px, py), (ex, ey) in zip(points, fallback)
            )
            if corner_shift > 0.20 or any(
                px < x1 - margin_x or px > x2 + margin_x or py < y1 - margin_y or py > y2 + margin_y
                for px, py in points
            ):
                points = fallback
            else:
                polygon_area = 0.5 * abs(
                    sum(
                        points[index][0] * points[(index + 1) % 4][1]
                        - points[(index + 1) % 4][0] * points[index][1]
                        for index in range(4)
                    )
                )
                if polygon_area < width * height * 0.62:
                    points = fallback

    inv = 1 / scale
    return [{"x": x * inv, "y": y * inv} for x, y in points]


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
        if area < image_area * 0.08 or aspect < 0.9 or aspect > 3.2:
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
    corner_component = dict(best)
    corner_component.update({"x1": x1, "y1": y1, "x2": x2, "y2": y2})
    corners = component_corners(corner_component, scale)
    inv = 1 / scale
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


def grid_row_text(grid: list[Any], row: int) -> str:
    values = grid[row] if row < len(grid) and isinstance(grid[row], list) else []
    return "".join(str(values[col])[:1] if col < len(values) and values[col] else " " for col in range(COLS))


def normalize_requirement_value(value: str) -> str:
    return re.sub(r"\s+", " ", clean_ocr_text(value)).strip().upper()


def requirement_matches(requirement: dict[str, Any], observed: str) -> bool:
    expected = str(requirement.get("expected", ""))
    ignore_case = bool(requirement.get("ignoreCase", True))
    ignore_spaces = bool(requirement.get("ignoreSpaces", False))
    check_type = str(requirement.get("type", "exact"))

    def normalized(value: str) -> str:
        value = clean_ocr_text(value)
        if ignore_spaces:
            value = re.sub(r"\s+", "", value)
        if ignore_case:
            value = value.upper()
        return value

    actual = normalized(observed)
    target = normalized(expected)
    if check_type == "blank":
        return not actual.strip()
    if check_type == "contains":
        return target in actual
    if check_type == "not_contains":
        return target not in actual
    if check_type == "fill":
        width = int(requirement["end"]) - int(requirement["start"]) + 1
        return actual == normalized(expected * width)
    return actual == target


def focused_grid_read(
    warped: Image.Image,
    row: int,
    start: int,
    end: int,
    expected: str,
) -> str:
    cell_w = warped.width / COLS
    cell_h = warped.height / ROWS
    x1 = max(0, int(start * cell_w))
    x2 = min(warped.width, int((end + 1) * cell_w))
    y1 = max(0, int((row - 1) * cell_h))
    y2 = min(warped.height, int(row * cell_h))
    crop = warped.crop((x1, y1, x2, y2))
    scale = 4
    enlarged = crop.resize((max(1, crop.width * scale), max(1, crop.height * scale)), Image.Resampling.LANCZOS)
    processed = preprocess_for_ocr(enlarged)
    width = end - start + 1
    focused = [" " for _ in range(width)]
    focused_cell_w = enlarged.width / max(1, width)

    try:
        words = run_tesseract_tsv(processed)
    except RuntimeError:
        words = []

    for word in words:
        text = normalize_mcdu_phrase(str(word.get("text", "")))
        if not text:
            continue
        compact = text.replace(" ", "")
        if not compact:
            continue
        word_left = float(word.get("left") or 0)
        word_width = float(word.get("width") or 1)
        start_index = max(0, min(width - 1, int(round(word_left / focused_cell_w))))
        span_cols = max(1, int(round(word_width / focused_cell_w)))
        sequence = text if " " in text and len(text) <= span_cols + 2 else compact
        use_projected_spacing = " " not in sequence and span_cols > len(compact) + 1
        for index, char in enumerate(sequence):
            if char.isspace():
                continue
            if use_projected_spacing:
                center_x = word_left + word_width * ((index + 0.5) / max(1, len(compact)))
                target = max(0, min(width - 1, int(center_x / focused_cell_w)))
            else:
                target = start_index + index
            if 0 <= target < width and focused[target] == " ":
                focused[target] = char

    if not "".join(focused).strip():
        for box in run_tesseract_boxes(processed):
            box_w = float(box.get("width") or 0)
            box_h = float(box.get("height") or 0)
            if box_w < 2 or box_h < 2 or box_w > focused_cell_w * 2.1:
                continue
            center_x = float(box["left"]) + box_w * 0.5
            index = max(0, min(width - 1, int(center_x / focused_cell_w)))
            char = clean_ocr_text(str(box.get("text", "")))[:1]
            if char and focused[index] == " ":
                focused[index] = char

    if expected == "-":
        gray = np.asarray(ImageOps.grayscale(crop)).astype(np.float32)
        for index in range(width):
            cx1 = int(index * crop.width / width)
            cx2 = int((index + 1) * crop.width / width)
            cy1 = int(crop.height * 0.36)
            cy2 = int(crop.height * 0.68)
            cell = gray[cy1:cy2, cx1:cx2]
            if cell.size == 0:
                continue
            bright = cell > 170
            if float(np.max(np.mean(bright, axis=1))) >= 0.30:
                focused[index] = "-"

    cell_reading = "".join(focused)
    raw_reading = " ".join(clean_ocr_text(str(word["text"])) for word in words).strip()
    if cell_reading.strip():
        return cell_reading
    return raw_reading[:width].ljust(width)


def grid_from_payload(value: Any) -> list[list[str]]:
    if not isinstance(value, list) or len(value) != ROWS:
        raise ValueError("A 13-row grid is required.")
    grid = empty_grid()
    for row_index, row in enumerate(value[:ROWS]):
        if not isinstance(row, list):
            continue
        for col_index, cell in enumerate(row[:COLS]):
            grid[row_index][col_index] = clean_ocr_text(str(cell))[:1] if cell else ""
    return normalize_grid_guards(grid)


def whole_grid_focused_recheck(
    warped: Image.Image,
    grid: list[list[str]],
    mode: str,
) -> tuple[list[list[str]], dict[str, int]]:
    mode = mode if mode in {"conservative", "balanced", "aggressive"} else "conservative"
    updated = normalize_grid_guards([[cell for cell in row] for row in grid])
    summary = {
        "rowsChecked": 0,
        "cellsFilled": 0,
        "cellsReplaced": 0,
        "dashesRecovered": 0,
        "conflictsKept": 0,
        "rowsSkipped": 0,
    }
    allowed = set("ABCDEFGHIJKLMNOPQRSTUVWXYZ0123456789<>/().°-+ ")

    for row in range(1, ROWS + 1):
        summary["rowsChecked"] += 1
        focused = focused_grid_read(warped, row, FIRST_DATA_COL, LAST_DATA_COL, "")
        focused = focused[: LAST_DATA_COL - FIRST_DATA_COL + 1].ljust(LAST_DATA_COL - FIRST_DATA_COL + 1)
        current_row = updated[row - 1]
        current_snapshot = current_row[:]
        current_compact = "".join(cell or "" for cell in current_row[FIRST_DATA_COL : LAST_DATA_COL + 1])
        focused_compact = "".join(char for char in focused if char and not char.isspace())
        current_text = row_string_from_cells(current_row)
        focused_text = f" {focused} "
        focused_score = mcdu_row_score(focused_text)
        current_score = mcdu_row_score(current_text)
        if (
            mode in {"balanced", "aggressive"}
            and focused_compact
            and focused_score >= 35.0
            and len(focused_compact) >= max(3, int(len(current_compact) * 0.55))
            and focused_score >= current_score + 10.0
        ):
            for offset, char in enumerate(focused):
                col = FIRST_DATA_COL + offset
                updated[row - 1][col] = "" if char.isspace() else clean_ocr_text(char)[:1]
            summary["cellsReplaced"] += sum(
                1
                for offset, char in enumerate(focused)
                if clean_ocr_text(char).strip()
                and current_snapshot[FIRST_DATA_COL + offset] != clean_ocr_text(char)[:1]
            )
            continue
        if mode != "aggressive" and (current_score >= 30.0 or focused_score < 25.0):
            continue
        if (
            mode != "aggressive"
            and current_compact
            and focused_compact
            and difflib.SequenceMatcher(None, current_compact, focused_compact).ratio() < 0.42
        ):
            summary["rowsSkipped"] += 1
            continue

        for offset, char in enumerate(focused):
            col = FIRST_DATA_COL + offset
            char = clean_ocr_text(char)[:1]
            if not char or char.isspace() or char not in allowed:
                continue
            current = updated[row - 1][col]
            if not current:
                has_left_context = any(updated[row - 1][nearby] for nearby in range(max(FIRST_DATA_COL, col - 2), col))
                has_right_context = any(updated[row - 1][nearby] for nearby in range(col + 1, min(LAST_DATA_COL, col + 2) + 1))
                if mode != "aggressive" and not (has_left_context and has_right_context):
                    continue
                updated[row - 1][col] = char
                if char == "-":
                    summary["dashesRecovered"] += 1
                else:
                    summary["cellsFilled"] += 1
                continue
            if current == char:
                continue
            if mode == "aggressive":
                updated[row - 1][col] = char
                summary["cellsReplaced"] += 1
            else:
                summary["conflictsKept"] += 1

    updated = recover_dash_lines(updated, warped) if mode in {"balanced", "aggressive"} else updated
    updated = disambiguate_o_zero(updated)
    updated = apply_corrections(updated)
    return normalize_grid_guards(updated), summary


def review_requirements(payload: dict[str, Any]) -> dict[str, Any]:
    requirements = payload.get("requirements")
    grid = payload.get("grid")
    if not isinstance(requirements, list) or not requirements:
        raise ValueError("Add at least one requirement before reviewing.")
    if not isinstance(grid, list) or len(grid) != ROWS:
        raise ValueError("Analyze or enter a 13-row grid before reviewing requirements.")

    warped: Image.Image | None = None
    image_data = str(payload.get("image", ""))
    corners = payload.get("corners")
    if image_data and isinstance(corners, list) and len(corners) == 4:
        warped = warp_screen(load_image(image_data), corners)

    results: list[dict[str, Any]] = []
    for index, raw_requirement in enumerate(requirements, 1):
        if not isinstance(raw_requirement, dict):
            continue
        try:
            row = int(raw_requirement["row"])
            start = int(raw_requirement["start"])
            end = int(raw_requirement["end"])
        except (KeyError, TypeError, ValueError):
            row = start = end = 0
        check_type = str(raw_requirement.get("type", "exact"))
        expected_raw = str(raw_requirement.get("expected", ""))
        valid_type = check_type in {"exact", "contains", "fill", "blank", "not_contains"}
        valid_expected = check_type == "blank" or bool(expected_raw)
        valid_fill = check_type != "fill" or len(expected_raw) == 1
        if not (1 <= row <= ROWS and 1 <= start <= end <= 38 and valid_type and valid_expected and valid_fill):
            results.append(
                {
                    "line": index,
                    "requirement": "Invalid requirement",
                    "status": "NEEDS REVIEW",
                    "expected": expected_raw,
                    "observed": "",
                    "rechecked": "",
                    "location": "",
                    "detail": "Check the row, column range, check type, and expected text.",
                }
            )
            continue

        requirement = dict(raw_requirement)
        width = end - start + 1
        expected_display = (
            expected_raw * width
            if check_type == "fill"
            else "Blank"
            if check_type == "blank"
            else expected_raw
        )
        observed = grid_row_text(grid, row - 1)[start : end + 1]
        passed = requirement_matches(requirement, observed)
        rechecked = ""
        detail = "Matched the extracted grid." if passed else "Initial grid reading did not match."

        if not passed and warped is not None:
            rechecked = focused_grid_read(warped, row, start, end, expected_raw)
            passed = requirement_matches(requirement, rechecked)
            detail = (
                "Passed after focused OCR recheck; confirm the focused reading."
                if passed
                else "Focused OCR recheck also did not match."
            )

        type_label = {
            "exact": "Exact text",
            "contains": "Contains text",
            "fill": "Fill range",
            "blank": "Blank range",
            "not_contains": "Must not contain",
        }[check_type]
        results.append(
            {
                "line": index,
                "requirement": type_label,
                "status": "PASS" if passed else "FAIL",
                "expected": expected_display,
                "observed": observed,
                "rechecked": rechecked,
                "location": f"Row {row}, columns {start}-{end}",
                "detail": detail,
            }
        )

    summary = {
        "total": len(results),
        "passed": sum(result["status"] == "PASS" for result in results),
        "failed": sum(result["status"] == "FAIL" for result in results),
        "needsReview": sum(result["status"] == "NEEDS REVIEW" for result in results),
    }
    return {"results": results, "summary": summary}


def compact_row_text(value: str) -> str:
    return re.sub(r"\s+", "", clean_ocr_text(value)).upper()


def fixed_row_text(value: str) -> str:
    return value[:COLS].ljust(COLS)


def correction_key(row: int, original: str) -> str:
    return f"row:{row + 1}|{fixed_row_text(original)}"


def apply_corrections(grid: list[list[str]]) -> list[list[str]]:
    corrections = load_corrections()
    updated = [[cell for cell in row] for row in grid]
    legacy_corrections = {
        key: value
        for key, value in corrections.items()
        if not key.startswith("row:") and not key.startswith("image:")
    }
    compact_corrections = {
        compact_row_text(key): value for key, value in legacy_corrections.items() if compact_row_text(key)
    }
    for row_index, row in enumerate(updated):
        row_text = fixed_row_text("".join(cell or " " for cell in row))
        exact_key = correction_key(row_index, row_text)
        if exact_key in corrections:
            updated[row_index] = list(fixed_row_text(corrections[exact_key]))
            continue

        normalized = re.sub(r"\s+", " ", row_text).strip()
        row_prefix = f"row:{row_index + 1}|"
        row_candidates = {
            key[len(row_prefix) :]: value for key, value in corrections.items() if key.startswith(row_prefix)
        }
        normalized_candidates = {
            re.sub(r"\s+", " ", key).strip(): value for key, value in row_candidates.items()
        }
        if normalized in normalized_candidates:
            corrected = fixed_row_text(normalized_candidates[normalized])
            updated[row_index] = list(corrected)
            continue

        if normalized in legacy_corrections:
            corrected = fixed_row_text(legacy_corrections[normalized])
            updated[row_index] = list(corrected)
            continue

        compact = compact_row_text(row_text)
        row_compact = {
            compact_row_text(key): value for key, value in row_candidates.items() if compact_row_text(key)
        }
        if compact in row_compact:
            updated[row_index] = list(fixed_row_text(row_compact[compact]))
            continue

        if compact in compact_corrections:
            corrected = fixed_row_text(compact_corrections[compact])
            updated[row_index] = list(corrected)
            continue

        if len(compact) >= 4:
            best_key = ""
            best_ratio = 0.0
            candidates = row_compact or compact_corrections
            for key in candidates:
                if abs(len(key) - len(compact)) > 4:
                    continue
                ratio = difflib.SequenceMatcher(None, compact, key).ratio()
                if ratio > best_ratio:
                    best_key = key
                    best_ratio = ratio
            if best_key and best_ratio >= 0.86:
                corrected = fixed_row_text(candidates[best_key])
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
            if not classified:
                continue
            char, distance = classified
            current = updated[row][col]
            if not current or (current in {"O", "0"} and char in {"O", "0"}) or distance <= 0.14:
                updated[row][col] = char
    return updated


def normalize_o_zero_token(token: str) -> str:
    if not token or not any(char in token for char in "O0"):
        return token

    letters_only = re.sub(r"[^A-Z0]", "", token.upper())
    as_word = letters_only.replace("0", "O")
    if as_word in MCDU_VOCABULARY:
        return token.replace("0", "O")

    if re.fullmatch(r"FL[O0-9]{2,3}", token):
        return "FL" + token[2:].replace("O", "0")

    result = list(token)
    digit_count = sum(value.isdigit() for value in token)
    letter_count = sum(value.isalpha() for value in token)
    for index, char in enumerate(result):
        if char not in {"O", "0"}:
            continue
        left = result[index - 1] if index > 0 else ""
        right = result[index + 1] if index + 1 < len(result) else ""
        numeric_context = (
            left.isdigit()
            or right.isdigit()
            or (digit_count > 0 and any(mark in token for mark in (".", "/", "°")))
            or digit_count > letter_count
        )
        result[index] = "0" if numeric_context else "O"
    return "".join(result)


def disambiguate_o_zero(grid: list[list[str]]) -> list[list[str]]:
    updated = [[cell for cell in row] for row in grid]
    for row in range(ROWS):
        col = FIRST_DATA_COL
        while col <= LAST_DATA_COL:
            if not updated[row][col]:
                col += 1
                continue
            start = col
            chars: list[str] = []
            while col <= LAST_DATA_COL and updated[row][col]:
                chars.append(updated[row][col])
                col += 1
            normalized = normalize_o_zero_token("".join(chars))
            for offset, char in enumerate(normalized):
                updated[row][start + offset] = char
    return normalize_grid_guards(updated)


def grid_character_count(grid: list[list[str]]) -> int:
    return sum(1 for row in grid for cell in row[FIRST_DATA_COL : LAST_DATA_COL + 1] if cell)


def verify_grid_with_char_boxes(
    word_grid: list[list[str]],
    char_boxes: list[dict[str, Any]],
    geometry: dict[str, float],
) -> list[list[str]]:
    verified = [[cell for cell in row] for row in word_grid]
    char_grid = empty_grid()
    for box in char_boxes:
        place_char(char_grid, box, geometry)

    for row in range(ROWS):
        for col in range(FIRST_DATA_COL, LAST_DATA_COL + 1):
            char_value = char_grid[row][col]
            if not char_value or verified[row][col]:
                continue
            nearby_same = any(
                verified[row][nearby] == char_value
                for nearby in range(max(FIRST_DATA_COL, col - 2), min(LAST_DATA_COL, col + 2) + 1)
            )
            if nearby_same:
                verified[row][col] = char_value
    return normalize_grid_guards(verified)


def recover_dash_lines(grid: list[list[str]], warped: Image.Image) -> list[list[str]]:
    updated = [[cell for cell in row] for row in grid]
    gray = np.asarray(ImageOps.grayscale(warped)).astype(np.float32)
    screen_w, screen_h = warped.size
    cell_w = screen_w / COLS
    cell_h = screen_h / ROWS

    for row in range(ROWS):
        y1 = int((row + 0.38) * cell_h)
        y2 = int((row + 0.68) * cell_h)
        if y2 <= y1:
            continue
        strip = gray[max(0, y1) : min(screen_h, y2), :]
        if strip.size == 0:
            continue

        active_cols: list[int] = []
        for col in range(FIRST_DATA_COL, LAST_DATA_COL + 1):
            x1 = int(col * cell_w + cell_w * 0.10)
            x2 = int((col + 1) * cell_w - cell_w * 0.10)
            cell = strip[:, max(0, x1) : min(screen_w, x2)]
            if cell.size == 0:
                continue
            bright = cell > 178
            bright_ratio = float(np.mean(bright))
            row_stroke = float(np.max(np.mean(bright, axis=1)))
            if 0.006 <= bright_ratio <= 0.16 and row_stroke >= 0.34:
                active_cols.append(col)

        run_start: int | None = None
        runs: list[tuple[int, int]] = []
        previous = -99
        for col in active_cols:
            if run_start is None or col != previous + 1:
                if run_start is not None and previous - run_start + 1 >= 9:
                    runs.append((run_start, previous))
                run_start = col
            previous = col
        if run_start is not None and previous - run_start + 1 >= 9:
            runs.append((run_start, previous))

        for start, end in runs:
            # Avoid turning a text-heavy row into dashes. Separator rows normally have
            # long mostly empty spans with only a few labels at the edges.
            occupied = sum(1 for col in range(start, end + 1) if updated[row][col])
            if occupied > max(3, (end - start + 1) // 4):
                continue
            for col in range(start, end + 1):
                if not updated[row][col]:
                    updated[row][col] = "-"
    return normalize_grid_guards(updated)


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
    geometry = calibrate_grid(boxes, screen_size)

    grid = empty_grid()
    word_scores = [[0.0 for _ in range(COLS)] for _ in range(ROWS)]
    for word in words:
        place_word(grid, word_scores, word, geometry)
    if grid_character_count(grid) < 8:
        for box in boxes:
            place_char(grid, box, geometry)
    grid = recover_dash_lines(grid, warped)
    grid = apply_templates(grid, warped)
    grid = disambiguate_o_zero(grid)
    corrected_grid = apply_corrections(grid)
    verification_summary = None
    verification = payload.get("verification")
    if isinstance(verification, dict) and bool(verification.get("enabled")):
        corrected_grid, verification_summary = whole_grid_focused_recheck(
            warped,
            corrected_grid,
            str(verification.get("mode", "conservative")),
        )

    preview_id = f"{uuid.uuid4().hex}.png"
    preview_path = EXPORTS / preview_id
    warped.save(preview_path)

    return {
        "grid": corrected_grid,
        "words": words,
        "boxes": boxes,
        "calibration": geometry,
        "previewUrl": f"/data/exports/{preview_id}",
        "verification": verification_summary,
    }


def refine_grid(payload: dict[str, Any]) -> dict[str, Any]:
    image = load_image(str(payload["image"]))
    corners = payload.get("corners")
    if not isinstance(corners, list) or len(corners) != 4:
        raise ValueError("Four screen corners are required.")
    grid = grid_from_payload(payload.get("grid"))
    warped = warp_screen(image, corners)
    refined, summary = whole_grid_focused_recheck(warped, grid, str(payload.get("mode", "conservative")))
    return {"grid": refined, "verification": summary}


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
    row = int(payload.get("row", 0))
    if not 0 <= row < ROWS:
        raise ValueError("Correction row must be between 1 and 13.")

    original = fixed_row_text(original_raw)
    corrected_list = list(fixed_row_text(corrected_raw))
    corrected_list[0] = " "
    corrected_list[COLS - 1] = " "
    corrected = "".join(corrected_list)
    if not corrected.strip():
        return {"count": len(load_corrections()), "skipped": True}

    corrections = load_corrections()
    corrections[correction_key(row, original)] = corrected
    save_corrections(corrections)
    return {"count": len(corrections)}


def remember_grid(payload: dict[str, Any]) -> dict[str, Any]:
    source_grid = payload.get("sourceGrid")
    corrected_grid = payload.get("grid")
    if not isinstance(source_grid, list) or not isinstance(corrected_grid, list):
        raise ValueError("Source and corrected 13-row grids are required.")

    corrections = load_corrections()
    corrections = {key: value for key, value in corrections.items() if not key.startswith("image:")}
    saved = 0
    for row in range(ROWS):
        original = grid_row_from_payload(source_grid, row)
        corrected = grid_row_from_payload(corrected_grid, row)
        if original == corrected:
            continue
        if corrected.strip():
            corrections[correction_key(row, original)] = corrected
        saved += 1
    save_corrections(corrections)
    return {"count": len(corrections), "saved": saved}


def grid_row_from_payload(grid: list[Any], row: int) -> str:
    values = grid[row] if row < len(grid) and isinstance(grid[row], list) else []
    cells = [str(values[col])[:1] if col < len(values) else "" for col in range(COLS)]
    cells[0] = ""
    cells[COLS - 1] = ""
    return fixed_row_text("".join(cell or " " for cell in cells))


class McmduHandler(SimpleHTTPRequestHandler):
    def end_headers(self) -> None:
        self.send_header("Cache-Control", "no-store, no-cache, must-revalidate")
        self.send_header("Pragma", "no-cache")
        self.send_header("Expires", "0")
        super().end_headers()

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
            elif self.path == "/api/remember-grid":
                json_response(self, HTTPStatus.OK, remember_grid(payload))
            elif self.path == "/api/remember-templates":
                json_response(self, HTTPStatus.OK, remember_templates(payload))
            elif self.path == "/api/review-requirements":
                json_response(self, HTTPStatus.OK, review_requirements(payload))
            elif self.path == "/api/refine-grid":
                json_response(self, HTTPStatus.OK, refine_grid(payload))
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
