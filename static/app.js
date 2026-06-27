const ROWS = 13;
const COLS = 40;
const FIRST_DATA_COL = 1;
const LAST_DATA_COL = 38;

const canvas = document.querySelector("#imageCanvas");
const ctx = canvas.getContext("2d");
const fileInput = document.querySelector("#fileInput");
const emptyState = document.querySelector("#emptyState");
const statusEl = document.querySelector("#status");
const analyzeButton = document.querySelector("#analyzeButton");
const exportButton = document.querySelector("#exportButton");
const rememberButton = document.querySelector("#rememberButton");
const resetButton = document.querySelector("#resetCorners");
const detectButton = document.querySelector("#detectButton");
const rotateCcwButton = document.querySelector("#rotateCcwButton");
const rotateCwButton = document.querySelector("#rotateCwButton");
const photoViewButton = document.querySelector("#photoViewButton");
const flatViewButton = document.querySelector("#flatViewButton");
const gridTable = document.querySelector("#gridTable");
const wholeGridVerify = document.querySelector("#wholeGridVerify");
const verificationMode = document.querySelector("#verificationMode");
const refineGridButton = document.querySelector("#refineGridButton");
const verificationSummary = document.querySelector("#verificationSummary");
const hybridOcr = document.querySelector("#hybridOcr");
const requirementRow = document.querySelector("#requirementRow");
const requirementStart = document.querySelector("#requirementStart");
const requirementEnd = document.querySelector("#requirementEnd");
const requirementType = document.querySelector("#requirementType");
const requirementText = document.querySelector("#requirementText");
const requirementIgnoreCase = document.querySelector("#requirementIgnoreCase");
const requirementIgnoreSpaces = document.querySelector("#requirementIgnoreSpaces");
const addRequirementButton = document.querySelector("#addRequirementButton");
const reviewRequirementsButton = document.querySelector("#reviewRequirementsButton");
const requirementList = document.querySelector("#requirementList");
const requirementsSummary = document.querySelector("#requirementsSummary");
const requirementsResults = document.querySelector("#requirementsResults");

const insetInputs = {
  left: document.querySelector("#leftInset"),
  right: document.querySelector("#rightInset"),
  top: document.querySelector("#topInset"),
  bottom: document.querySelector("#bottomInset"),
};

const state = {
  image: null,
  flatImage: null,
  flatUrl: "",
  gridAlignment: { x: 0, y: 0 },
  imageDataUrl: "",
  imageBounds: { x: 0, y: 0, width: 0, height: 0, scale: 1 },
  flatBounds: { x: 0, y: 0, width: 0, height: 0, scale: 1 },
  corners: [],
  candidates: [],
  selectedCandidateIndex: -1,
  dragging: -1,
  viewMode: "photo",
  sourceGrid: makeEmptyGrid(),
  confidenceGrid: Array.from({ length: ROWS }, () => Array.from({ length: COLS }, () => 0)),
  requirements: [],
};

function setStatus(message) {
  statusEl.textContent = message;
}

function makeEmptyGrid() {
  return Array.from({ length: ROWS }, () => Array.from({ length: COLS }, () => ""));
}

function normalizeGridGuards(grid) {
  const normalized = makeEmptyGrid();
  for (let row = 0; row < ROWS; row += 1) {
    for (let col = FIRST_DATA_COL; col <= LAST_DATA_COL; col += 1) {
      normalized[row][col] = grid[row]?.[col] || "";
    }
  }
  return normalized;
}

function fitCanvas() {
  const rect = canvas.getBoundingClientRect();
  const dpr = window.devicePixelRatio || 1;
  canvas.width = Math.max(1, Math.round(rect.width * dpr));
  canvas.height = Math.max(1, Math.round(rect.height * dpr));
  ctx.setTransform(dpr, 0, 0, dpr, 0, 0);
  draw();
}

function calculateImageBounds() {
  if (!state.image) {
    return { x: 0, y: 0, width: 0, height: 0, scale: 1 };
  }
  const rect = canvas.getBoundingClientRect();
  const scale = Math.min(rect.width / state.image.width, rect.height / state.image.height);
  const width = state.image.width * scale;
  const height = state.image.height * scale;
  return {
    x: (rect.width - width) / 2,
    y: (rect.height - height) / 2,
    width,
    height,
    scale,
  };
}

function calculateFitBounds(image) {
  if (!image) {
    return { x: 0, y: 0, width: 0, height: 0, scale: 1 };
  }
  const rect = canvas.getBoundingClientRect();
  const scale = Math.min(rect.width / image.width, rect.height / image.height);
  const width = image.width * scale;
  const height = image.height * scale;
  return {
    x: (rect.width - width) / 2,
    y: (rect.height - height) / 2,
    width,
    height,
    scale,
  };
}

function imageToCanvas(point) {
  const b = state.imageBounds;
  return {
    x: b.x + point.x * b.scale,
    y: b.y + point.y * b.scale,
  };
}

function canvasToImage(point) {
  const b = state.imageBounds;
  return {
    x: (point.x - b.x) / b.scale,
    y: (point.y - b.y) / b.scale,
  };
}

function defaultCorners() {
  const image = state.image;
  const marginX = image.width * 0.13;
  const marginY = image.height * 0.22;
  state.corners = [
    { x: marginX, y: marginY },
    { x: image.width - marginX, y: marginY },
    { x: image.width - marginX, y: image.height - marginY },
    { x: marginX, y: image.height - marginY },
  ];
}

function setCorners(corners) {
  if (!Array.isArray(corners) || corners.length !== 4) {
    return false;
  }
  state.corners = corners.map((corner) => ({
    x: Math.max(0, Math.min(state.image.width, Number(corner.x))),
    y: Math.max(0, Math.min(state.image.height, Number(corner.y))),
  }));
  return true;
}

function lerp(a, b, t) {
  return a + (b - a) * t;
}

function lerpPoint(a, b, t) {
  return { x: lerp(a.x, b.x, t), y: lerp(a.y, b.y, t) };
}

function projectPoint(corners, u, v) {
  const [topLeft, topRight, bottomRight, bottomLeft] = corners;
  const dx1 = topRight.x - bottomRight.x;
  const dx2 = bottomLeft.x - bottomRight.x;
  const dx3 = topLeft.x - topRight.x + bottomRight.x - bottomLeft.x;
  const dy1 = topRight.y - bottomRight.y;
  const dy2 = bottomLeft.y - bottomRight.y;
  const dy3 = topLeft.y - topRight.y + bottomRight.y - bottomLeft.y;
  const denominator = dx1 * dy2 - dx2 * dy1;

  let g = 0;
  let h = 0;
  if (Math.abs(denominator) > 1e-8) {
    g = (dx3 * dy2 - dx2 * dy3) / denominator;
    h = (dx1 * dy3 - dx3 * dy1) / denominator;
  }
  const a = topRight.x - topLeft.x + g * topRight.x;
  const b = bottomLeft.x - topLeft.x + h * bottomLeft.x;
  const d = topRight.y - topLeft.y + g * topRight.y;
  const e = bottomLeft.y - topLeft.y + h * bottomLeft.y;
  const scale = g * u + h * v + 1;
  return {
    x: (a * u + b * v + topLeft.x) / scale,
    y: (d * u + e * v + topLeft.y) / scale,
  };
}

function getInsetValues() {
  return {
    left: Number(insetInputs.left.value) / 100,
    right: Number(insetInputs.right.value) / 100,
    top: Number(insetInputs.top.value) / 100,
    bottom: Number(insetInputs.bottom.value) / 100,
  };
}

function getGridCorners() {
  const inset = getInsetValues();
  const u1 = inset.left;
  const u2 = 1 - inset.right;
  const v1 = inset.top;
  const v2 = 1 - inset.bottom;
  return [
    projectPoint(state.corners, u1, v1),
    projectPoint(state.corners, u2, v1),
    projectPoint(state.corners, u2, v2),
    projectPoint(state.corners, u1, v2),
  ];
}

function getRawCorners() {
  return state.corners.map((corner) => ({ x: corner.x, y: corner.y }));
}

function setViewMode(mode) {
  state.viewMode = mode;
  photoViewButton.classList.toggle("active", mode === "photo");
  flatViewButton.classList.toggle("active", mode === "flat");
  draw();
}

function invalidateFlattenedDisplay() {
  state.flatImage = null;
  state.flatUrl = "";
}

function draw() {
  const rect = canvas.getBoundingClientRect();
  ctx.clearRect(0, 0, rect.width, rect.height);

  if (!state.image) {
    return;
  }

  if (state.viewMode === "flat" && state.flatImage) {
    drawFlattenedView();
  } else {
    state.imageBounds = calculateImageBounds();
    const b = state.imageBounds;
    ctx.drawImage(state.image, b.x, b.y, b.width, b.height);

    drawGrid();
    drawHandles();
    drawCandidateOutlines();
  }
}

function drawGrid() {
  const corners = getGridCorners().map(imageToCanvas);
  const originX = state.gridAlignment.x || 0;
  const originY = state.gridAlignment.y || 0;

  ctx.save();
  ctx.lineWidth = 1;
  ctx.strokeStyle = "rgba(24, 154, 141, 0.92)";
  ctx.fillStyle = "#ffffff";
  ctx.font = "11px system-ui, sans-serif";
  ctx.textAlign = "center";
  ctx.textBaseline = "middle";

  for (let col = 0; col <= COLS; col += 1) {
    const u = originX + col / COLS;
    const top = projectPoint(corners, u, 0);
    const bottom = projectPoint(corners, u, 1);
    ctx.beginPath();
    ctx.moveTo(top.x, top.y);
    ctx.lineTo(bottom.x, bottom.y);
    ctx.stroke();
  }

  for (let row = 0; row <= ROWS; row += 1) {
    const v = originY + row / ROWS;
    const left = projectPoint(corners, 0, v);
    const right = projectPoint(corners, 1, v);
    ctx.beginPath();
    ctx.moveTo(left.x, left.y);
    ctx.lineTo(right.x, right.y);
    ctx.stroke();
  }

  ctx.fillStyle = "rgba(6, 79, 73, 0.9)";
  for (let row = 0; row < ROWS; row += 1) {
    const p = projectPoint(corners, originX - 0.035, originY + (row + 0.5) / ROWS);
    ctx.fillText(String(row + 1), p.x, p.y);
  }

  for (let col = 1; col <= 38; col += 1) {
    const p = projectPoint(corners, originX + (col + 0.5) / COLS, originY - 0.035);
    ctx.fillText(String(col), p.x, p.y);
  }
  ctx.restore();
}

function drawHandles() {
  const screenCorners = state.corners.map(imageToCanvas);

  ctx.save();
  ctx.lineWidth = 2;
  ctx.strokeStyle = "#f2994a";
  ctx.fillStyle = "#ffffff";
  ctx.beginPath();
  screenCorners.forEach((point, index) => {
    if (index === 0) {
      ctx.moveTo(point.x, point.y);
    } else {
      ctx.lineTo(point.x, point.y);
    }
  });
  ctx.closePath();
  ctx.stroke();

  screenCorners.forEach((point, index) => {
    ctx.beginPath();
    ctx.arc(point.x, point.y, 7, 0, Math.PI * 2);
    ctx.fill();
    ctx.stroke();
    ctx.fillStyle = "#172026";
    ctx.font = "11px system-ui, sans-serif";
    ctx.textAlign = "center";
    ctx.textBaseline = "middle";
    ctx.fillText(String(index + 1), point.x, point.y);
    ctx.fillStyle = "#ffffff";
  });
  ctx.restore();
}

function isPointInQuad(corners, point) {
  const n = corners.length;
  let inside = false;
  for (let i = 0, j = n - 1; i < n; j = i++) {
    const xi = corners[i].x;
    const yi = corners[i].y;
    const xj = corners[j].x;
    const yj = corners[j].y;
    if ((yi > point.y) !== (yj > point.y) &&
        point.x < ((xj - xi) * (point.y - yi)) / (yj - yi) + xi) {
      inside = !inside;
    }
  }
  return inside;
}

function drawCandidateOutlines() {
  if (state.candidates.length <= 1) {
    return;
  }
  ctx.save();
  state.candidates.forEach((candidate, index) => {
    if (index === state.selectedCandidateIndex) {
      return;
    }
    const canvasCorners = candidate.corners.map(imageToCanvas);
    ctx.strokeStyle = "#4a9ef2";
    ctx.lineWidth = 2;
    ctx.setLineDash([8, 5]);
    ctx.fillStyle = "rgba(74, 158, 242, 0.12)";
    ctx.beginPath();
    canvasCorners.forEach((pt, i) => {
      if (i === 0) ctx.moveTo(pt.x, pt.y);
      else ctx.lineTo(pt.x, pt.y);
    });
    ctx.closePath();
    ctx.stroke();
    ctx.fill();
    ctx.setLineDash([]);
    const cx = canvasCorners.reduce((s, p) => s + p.x, 0) / 4;
    const cy = canvasCorners.reduce((s, p) => s + p.y, 0) / 4;
    ctx.font = "bold 14px system-ui, sans-serif";
    ctx.textAlign = "center";
    ctx.textBaseline = "middle";
    ctx.fillStyle = "rgba(0,0,0,0.45)";
    ctx.fillText(`Screen ${index + 1} — click to select`, cx, cy + 1);
    ctx.fillStyle = "#4a9ef2";
    ctx.fillText(`Screen ${index + 1} — click to select`, cx, cy);
  });
  ctx.restore();
}

function selectCandidate(index) {
  state.selectedCandidateIndex = index;
  setCorners(state.candidates[index].corners);
  invalidateFlattenedDisplay();
  draw();
  setStatus(`Screen ${index + 1} selected — flattening`);
  flattenDisplay(false);
}

function drawFlattenedView() {
  state.flatBounds = calculateFitBounds(state.flatImage);
  const b = state.flatBounds;
  ctx.drawImage(state.flatImage, b.x, b.y, b.width, b.height);
  drawFlatGrid();
}

function drawFlatGrid() {
  const b = state.flatBounds;
  const inset = getInsetValues();
  const left = b.x + b.width * inset.left;
  const right = b.x + b.width * (1 - inset.right);
  const top = b.y + b.height * inset.top;
  const bottom = b.y + b.height * (1 - inset.bottom);
  const gridWidth = right - left;
  const gridHeight = bottom - top;

  ctx.save();
  ctx.lineWidth = 1;
  ctx.strokeStyle = "rgba(24, 154, 141, 0.95)";
  ctx.fillStyle = "rgba(6, 79, 73, 0.95)";
  ctx.font = "11px system-ui, sans-serif";
  ctx.textAlign = "center";
  ctx.textBaseline = "middle";

  for (let col = 0; col <= COLS; col += 1) {
    const x = left + (gridWidth * col) / COLS;
    ctx.beginPath();
    ctx.moveTo(x, top);
    ctx.lineTo(x, bottom);
    ctx.stroke();
  }

  for (let row = 0; row <= ROWS; row += 1) {
    const y = top + (gridHeight * row) / ROWS;
    ctx.beginPath();
    ctx.moveTo(left, y);
    ctx.lineTo(right, y);
    ctx.stroke();
  }

  for (let row = 0; row < ROWS; row += 1) {
    const y = top + (gridHeight * (row + 0.5)) / ROWS;
    ctx.fillText(String(row + 1), Math.max(10, left - 16), y);
  }

  for (let col = 1; col <= 38; col += 1) {
    const x = left + (gridWidth * (col + 0.5)) / COLS;
    ctx.fillText(String(col), x, Math.max(10, top - 14));
  }
  ctx.restore();
}

function pointerPosition(event) {
  const rect = canvas.getBoundingClientRect();
  return {
    x: event.clientX - rect.left,
    y: event.clientY - rect.top,
  };
}

function findHandle(point) {
  const handles = state.corners.map(imageToCanvas);
  let nearest = -1;
  let nearestDistance = 14;
  handles.forEach((handle, index) => {
    const distance = Math.hypot(handle.x - point.x, handle.y - point.y);
    if (distance < nearestDistance) {
      nearest = index;
      nearestDistance = distance;
    }
  });
  return nearest;
}

function renderGridTable(grid = makeEmptyGrid(), confidenceGrid = state.confidenceGrid) {
  gridTable.textContent = "";
  const thead = document.createElement("thead");
  const header = document.createElement("tr");
  header.appendChild(document.createElement("th")).textContent = "Row";
  for (let col = 0; col < COLS; col += 1) {
    const th = document.createElement("th");
    th.textContent = col >= 1 && col <= 38 ? String(col) : "";
    header.appendChild(th);
  }
  thead.appendChild(header);
  gridTable.appendChild(thead);

  const tbody = document.createElement("tbody");
  for (let row = 0; row < ROWS; row += 1) {
    const tr = document.createElement("tr");
    const rowHead = document.createElement("td");
    rowHead.textContent = String(row + 1);
    tr.appendChild(rowHead);
    for (let col = 0; col < COLS; col += 1) {
      const td = document.createElement("td");
      td.contentEditable = "true";
      td.dataset.row = String(row);
      td.dataset.col = String(col);
      td.textContent = col >= FIRST_DATA_COL && col <= LAST_DATA_COL ? grid[row]?.[col]?.trim() || "" : "";
      const confidence = Number(confidenceGrid?.[row]?.[col] || 0);
      if (td.textContent && confidence > 0 && confidence < 0.62) {
        td.classList.add("low-confidence");
        td.title = `Low OCR confidence: ${Math.round(confidence * 100)}%`;
      }
      if (col < FIRST_DATA_COL || col > LAST_DATA_COL) {
        td.contentEditable = "false";
        td.classList.add("guard-cell");
      }
      td.addEventListener("input", () => {
        const value = td.textContent.replace(/\s+/g, "").slice(-1).toUpperCase();
        if (td.textContent !== value) {
          td.textContent = value;
        }
        td.classList.remove("low-confidence");
        td.removeAttribute("title");
        exportButton.disabled = false;
        rememberButton.disabled = false;
      });
      td.addEventListener("paste", (event) => {
        event.preventDefault();
        const text = event.clipboardData?.getData("text")?.replace(/\s+/g, "").toUpperCase() || "";
        const cells = Array.from(gridTable.querySelectorAll(`td[data-row="${row}"][contenteditable="true"]`));
        let targetCol = col;
        for (const char of text) {
          if (targetCol > LAST_DATA_COL) {
            break;
          }
          const target = cells.find((cell) => Number(cell.dataset.col) === targetCol);
          if (target) {
            target.textContent = char;
            target.classList.remove("low-confidence");
            target.removeAttribute("title");
          }
          targetCol += 1;
        }
        exportButton.disabled = false;
        rememberButton.disabled = false;
      });
      tr.appendChild(td);
    }
    tbody.appendChild(tr);
  }
  gridTable.appendChild(tbody);
}

function renderVerificationSummary(summary) {
  if (!summary) {
    verificationSummary.hidden = true;
    verificationSummary.textContent = "";
    return;
  }
  verificationSummary.hidden = false;
  const parts = [
    `${summary.rowsChecked || 0} rows checked`,
    `${summary.cellsFilled || 0} cells filled`,
    `${summary.cellsReplaced || 0} cells replaced`,
    `${summary.dashesRecovered || 0} dashes recovered`,
    `${summary.conflictsKept || 0} conflicts kept`,
    `${summary.rowsSkipped || 0} noisy rows skipped`,
  ];
  verificationSummary.textContent = parts.join(", ");
}

function getCurrentGrid() {
  const grid = makeEmptyGrid();
  gridTable.querySelectorAll("td[contenteditable='true']").forEach((cell) => {
    const row = Number(cell.dataset.row);
    const col = Number(cell.dataset.col);
    if (col >= FIRST_DATA_COL && col <= LAST_DATA_COL) {
      grid[row][col] = cell.textContent.replace(/\s+/g, "").slice(0, 1).toUpperCase();
    }
  });
  return normalizeGridGuards(grid);
}

function gridRowText(row) {
  return normalizeGridGuards([row])[0].map((cell) => (cell || " ").slice(0, 1)).join("").slice(0, COLS).padEnd(COLS, " ");
}

async function postJson(url, body) {
  const response = await fetch(url, {
    method: "POST",
    headers: { "Content-Type": "application/json" },
    body: JSON.stringify(body),
  });
  const payload = await response.json();
  if (!response.ok) {
    throw new Error(payload.error || "Request failed.");
  }
  return payload;
}

async function applyRotation(degrees) {
  if (!state.image || degrees === 0) {
    return;
  }
  const img = state.image;
  const swap = degrees === 90 || degrees === 270;
  const offscreen = document.createElement("canvas");
  offscreen.width = swap ? img.naturalHeight : img.naturalWidth;
  offscreen.height = swap ? img.naturalWidth : img.naturalHeight;
  const c = offscreen.getContext("2d");
  c.translate(offscreen.width / 2, offscreen.height / 2);
  // PIL rotate(n) is CCW; canvas rotate is CW — negate to match.
  c.rotate((-degrees * Math.PI) / 180);
  c.drawImage(img, -img.naturalWidth / 2, -img.naturalHeight / 2);
  state.imageDataUrl = offscreen.toDataURL("image/jpeg", 0.92);
  await new Promise((resolve, reject) => {
    const newImg = new Image();
    newImg.addEventListener("load", () => {
      state.image = newImg;
      state.imageBounds = calculateImageBounds();
      resolve();
    }, { once: true });
    newImg.addEventListener("error", reject, { once: true });
    newImg.src = state.imageDataUrl;
  });
  state.flatImage = null;
  state.flatUrl = "";
  state.candidates = [];
  state.selectedCandidateIndex = -1;
}

async function autoDetectDisplay(skipOrientationProbe = false) {
  if (!state.imageDataUrl) {
    return false;
  }
  detectButton.disabled = true;
  rotateCcwButton.disabled = true;
  rotateCwButton.disabled = true;
  setStatus("Detecting display");
  try {
    const result = await postJson("/api/detect-display", {
      image: state.imageDataUrl,
      skipOrientationProbe,
    });

    // Auto-rotate: server probed orientation and found the image needed correction.
    // Rotate the client image to match the rotated image the server detected on.
    const rotation = typeof result.rotation === "number" ? result.rotation : 0;
    if (rotation !== 0) {
      setStatus(`Auto-rotating ${rotation}°`);
      await applyRotation(rotation);
    }

    state.candidates = Array.isArray(result.candidates) ? result.candidates : [];
    state.selectedCandidateIndex = typeof result.bestIndex === "number" ? result.bestIndex : 0;
    if (setCorners(result.corners)) {
      Object.values(insetInputs).forEach((input) => {
        input.value = "0";
      });
      draw();
      const size = result.displaySize ? ` ${result.displaySize.width}x${result.displaySize.height}` : "";
      const method = result.perspectiveRefined ? "Perspective display detected" : "Display detected";
      const multiNote = state.candidates.length > 1
        ? ` — ${state.candidates.length} screens found, click another to switch`
        : "";
      const rotateNote = rotation !== 0 ? ` (auto-rotated ${rotation}°)` : "";
      setStatus(`${method}${size}${multiNote}${rotateNote}`);
      await flattenDisplay(false);
      return true;
    }
    setStatus("Display detection returned invalid corners");
  } catch (error) {
    defaultCorners();
    draw();
    setStatus(error.message);
  } finally {
    detectButton.disabled = false;
    rotateCcwButton.disabled = !state.image;
    rotateCwButton.disabled = !state.image;
  }
  return false;
}

async function flattenDisplay(showStatus = true) {
  if (!state.imageDataUrl || !state.corners.length) {
    return false;
  }
  flatViewButton.disabled = true;
  if (showStatus) {
    setStatus("Flattening display");
  }
  try {
    const result = await postJson("/api/flatten-display", {
      image: state.imageDataUrl,
      corners: getRawCorners(),
    });
    const image = new Image();
    await new Promise((resolve, reject) => {
      image.addEventListener("load", resolve, { once: true });
      image.addEventListener("error", reject, { once: true });
      image.src = `${result.previewUrl}?t=${Date.now()}`;
    });
    state.flatImage = image;
    state.flatUrl = result.previewUrl;
    state.gridAlignment = {
      x: Number(result.gridAlignment?.x) || 0,
      y: Number(result.gridAlignment?.y) || 0,
    };
    photoViewButton.disabled = false;
    flatViewButton.disabled = false;
    setViewMode("flat");
    setStatus(`Flattened display ${result.width}x${result.height}`);
    return true;
  } catch (error) {
    setStatus(error.message);
  } finally {
    flatViewButton.disabled = false;
  }
  return false;
}

fileInput.addEventListener("change", () => {
  const file = fileInput.files?.[0];
  if (!file) {
    return;
  }
  const reader = new FileReader();
  reader.addEventListener("load", () => {
    const image = new Image();
    image.addEventListener("load", () => {
      state.image = image;
      state.imageDataUrl = String(reader.result);
      defaultCorners();
      emptyState.classList.add("hidden");
      detectButton.disabled = false;
      rotateCcwButton.disabled = false;
      rotateCwButton.disabled = false;
      photoViewButton.disabled = false;
      flatViewButton.disabled = false;
      analyzeButton.disabled = false;
      exportButton.disabled = true;
      rememberButton.disabled = true;
      refineGridButton.disabled = false;
      state.flatImage = null;
      state.flatUrl = "";
      state.gridAlignment = { x: 0, y: 0 };
      state.candidates = [];
      state.selectedCandidateIndex = -1;
      state.confidenceGrid = Array.from({ length: ROWS }, () => Array.from({ length: COLS }, () => 0));
      setViewMode("photo");
      renderGridTable();
      setStatus("Image loaded");
      fitCanvas();
      autoDetectDisplay();
    });
    image.src = String(reader.result);
  });
  reader.readAsDataURL(file);
});

canvas.addEventListener("pointerdown", (event) => {
  if (!state.image || state.viewMode === "flat") {
    return;
  }
  const point = pointerPosition(event);
  const handle = findHandle(point);
  if (handle >= 0) {
    state.dragging = handle;
    canvas.setPointerCapture(event.pointerId);
  }
});

canvas.addEventListener("pointermove", (event) => {
  if (state.dragging < 0 || !state.image || state.viewMode === "flat") {
    return;
  }
  const point = canvasToImage(pointerPosition(event));
  state.corners[state.dragging] = {
    x: Math.max(0, Math.min(state.image.width, point.x)),
    y: Math.max(0, Math.min(state.image.height, point.y)),
  };
  invalidateFlattenedDisplay();
  draw();
});

canvas.addEventListener("pointerup", () => {
  state.dragging = -1;
});

canvas.addEventListener("click", (event) => {
  if (!state.image || state.viewMode !== "photo" || state.candidates.length <= 1) {
    return;
  }
  const point = pointerPosition(event);
  if (findHandle(point) >= 0) {
    return;
  }
  for (let i = 0; i < state.candidates.length; i += 1) {
    if (i === state.selectedCandidateIndex) {
      continue;
    }
    const canvasCorners = state.candidates[i].corners.map(imageToCanvas);
    if (isPointInQuad(canvasCorners, point)) {
      selectCandidate(i);
      return;
    }
  }
});

Object.values(insetInputs).forEach((input) => {
  input.addEventListener("input", () => {
    state.gridAlignment = { x: 0, y: 0 };
    invalidateFlattenedDisplay();
    draw();
  });
});

resetButton.addEventListener("click", () => {
  if (!state.image) {
    return;
  }
  defaultCorners();
  state.flatImage = null;
  state.flatUrl = "";
  state.gridAlignment = { x: 0, y: 0 };
  state.candidates = [];
  state.selectedCandidateIndex = -1;
  setViewMode("photo");
  Object.values(insetInputs).forEach((input) => {
    input.value = "0";
  });
  draw();
  setStatus("Corners reset");
});

detectButton.addEventListener("click", () => {
  autoDetectDisplay();
});

rotateCcwButton.addEventListener("click", async () => {
  if (!state.image) {
    return;
  }
  rotateCcwButton.disabled = true;
  rotateCwButton.disabled = true;
  setStatus("Rotating 90° counter-clockwise");
  await applyRotation(90);
  draw();
  await autoDetectDisplay(true);
});

rotateCwButton.addEventListener("click", async () => {
  if (!state.image) {
    return;
  }
  rotateCcwButton.disabled = true;
  rotateCwButton.disabled = true;
  setStatus("Rotating 90° clockwise");
  await applyRotation(270);
  draw();
  await autoDetectDisplay(true);
});

photoViewButton.addEventListener("click", () => {
  setViewMode("photo");
  setStatus("Photo view");
});

flatViewButton.addEventListener("click", () => {
  if (state.flatImage) {
    setViewMode("flat");
    setStatus("Flattened view");
  } else {
    flattenDisplay();
  }
});

analyzeButton.addEventListener("click", async () => {
  if (!state.image) {
    return;
  }
  analyzeButton.disabled = true;
  setStatus("Analyzing OCR");
  try {
    const result = await postJson("/api/analyze", {
      image: state.imageDataUrl,
      corners: getGridCorners(),
      verification: {
        enabled: wholeGridVerify.checked,
        mode: verificationMode.value,
      },
      hybridOcr: hybridOcr.checked,
    });
    state.sourceGrid = normalizeGridGuards(result.grid);
    state.confidenceGrid = result.confidenceGrid || Array.from({ length: ROWS }, () => Array.from({ length: COLS }, () => 0));
    renderGridTable(state.sourceGrid, state.confidenceGrid);
    renderVerificationSummary(result.verification);
    exportButton.disabled = false;
    rememberButton.disabled = false;
    refineGridButton.disabled = false;
    const boxCount = Array.isArray(result.boxes) ? result.boxes.length : 0;
    const refined = result.verification ? `, ${result.verification.cellsFilled + result.verification.cellsReplaced} focused updates` : "";
    const preprocessing = result.preprocessing ? `, ${result.preprocessing} image pass` : "";
    const engines = result.ocrEngines?.used?.join(" + ") || "tesseract";
    const paddleNote = result.ocrEngines?.paddleError && hybridOcr.checked
      ? `; Paddle unavailable (${result.ocrEngines.paddleError})`
      : "";
    setStatus(`OCR complete using ${engines}: ${boxCount} character boxes, ${result.words.length} text blocks${preprocessing}${refined}${paddleNote}`);
  } catch (error) {
    setStatus(error.message);
  } finally {
    analyzeButton.disabled = false;
  }
});

refineGridButton.addEventListener("click", async () => {
  if (!state.image) {
    return;
  }
  refineGridButton.disabled = true;
  analyzeButton.disabled = true;
  setStatus("Running focused recheck across the grid");
  try {
    const result = await postJson("/api/refine-grid", {
      image: state.imageDataUrl,
      corners: getGridCorners(),
      grid: getCurrentGrid(),
      mode: verificationMode.value,
    });
    state.sourceGrid = normalizeGridGuards(result.grid);
    state.confidenceGrid = Array.from({ length: ROWS }, () => Array.from({ length: COLS }, () => 0));
    renderGridTable(state.sourceGrid, state.confidenceGrid);
    renderVerificationSummary(result.verification);
    exportButton.disabled = false;
    rememberButton.disabled = false;
    setStatus(`Focused recheck complete: ${result.verification.cellsFilled + result.verification.cellsReplaced} updates`);
  } catch (error) {
    setStatus(error.message);
  } finally {
    refineGridButton.disabled = false;
    analyzeButton.disabled = false;
  }
});

exportButton.addEventListener("click", async () => {
  setStatus("Creating Word file");
  exportButton.disabled = true;
  try {
    const result = await postJson("/api/export-docx", {
      grid: getCurrentGrid(),
    });
    const link = document.createElement("a");
    link.href = result.url;
    link.download = result.filename;
    link.click();
    setStatus("Word file exported");
  } catch (error) {
    setStatus(error.message);
  } finally {
    exportButton.disabled = false;
  }
});

rememberButton.addEventListener("click", async () => {
  const current = getCurrentGrid();
  const changes = [];
  for (let row = 0; row < ROWS; row += 1) {
    const original = gridRowText(state.sourceGrid[row] || []);
    const corrected = gridRowText(current[row] || []);
    if (corrected.trim() && original !== corrected) {
      changes.push({ row, original, corrected });
    }
  }

  rememberButton.disabled = true;
  setStatus(`Remembering ${changes.length} correction rows`);
  try {
    const sourceGrid = state.sourceGrid;
    await postJson("/api/remember-grid", {
      sourceGrid,
      grid: current,
    });
    const learned = await postJson("/api/remember-templates", {
      image: state.imageDataUrl,
      corners: getGridCorners(),
      sourceGrid,
      grid: current,
    });
    state.sourceGrid = current;
    setStatus(`Reusable corrections saved, ${learned.learned} visual samples learned`);
  } catch (error) {
    setStatus(error.message);
  } finally {
    rememberButton.disabled = false;
  }
});

function requirementDescription(requirement) {
  const typeLabels = {
    exact: "equals",
    contains: "contains",
    fill: "is filled with",
    blank: "is blank",
    not_contains: "does not contain",
  };
  const value = requirement.type === "blank" ? "" : ` "${requirement.expected}"`;
  return `Row ${requirement.row}, columns ${requirement.start}-${requirement.end} ${typeLabels[requirement.type]}${value}`;
}

function renderRequirementList() {
  requirementList.textContent = "";
  state.requirements.forEach((requirement, index) => {
    const item = document.createElement("div");
    item.className = "requirement-item";
    const text = document.createElement("span");
    text.textContent = requirementDescription(requirement);
    const remove = document.createElement("button");
    remove.type = "button";
    remove.textContent = "Remove";
    remove.dataset.requirementIndex = String(index);
    item.append(text, remove);
    requirementList.appendChild(item);
  });
  reviewRequirementsButton.disabled = state.requirements.length === 0;
}

function syncRequirementFields() {
  const type = requirementType.value;
  const needsText = type !== "blank";
  requirementText.disabled = !needsText;
  if (!needsText) {
    requirementText.value = "";
  }
  if (type === "fill") {
    requirementText.maxLength = 1;
  } else {
    requirementText.removeAttribute("maxlength");
  }
}

function autoSetRequirementEnd() {
  if (requirementType.value !== "exact") {
    return;
  }
  const start = Number(requirementStart.value);
  const length = Math.max(1, requirementText.value.length);
  requirementEnd.value = String(Math.min(38, start + length - 1));
}

function renderRequirementResults(result) {
  const summary = result.summary || {};
  requirementsSummary.hidden = false;
  requirementsSummary.textContent = `${summary.passed || 0} passed, ${summary.failed || 0} failed, ${summary.needsReview || 0} need confirmation`;
  requirementsResults.hidden = false;
  requirementsResults.textContent = "";

  const table = document.createElement("table");
  table.className = "review-table";
  const header = document.createElement("tr");
  ["Status", "Requirement", "Expected", "Grid reading", "Focused recheck", "Location", "Reason"].forEach((label) => {
    const th = document.createElement("th");
    th.textContent = label;
    header.appendChild(th);
  });
  const thead = document.createElement("thead");
  thead.appendChild(header);
  table.appendChild(thead);

  const tbody = document.createElement("tbody");
  (result.results || []).forEach((item) => {
    const tr = document.createElement("tr");
    tr.className = `review-${String(item.status).toLowerCase().replaceAll(" ", "-")}`;
    [
      item.status,
      item.requirement,
      item.expected,
      item.observed,
      item.rechecked || "",
      item.location,
      item.detail,
    ].forEach((value) => {
      const td = document.createElement("td");
      td.textContent = value || "";
      tr.appendChild(td);
    });
    tbody.appendChild(tr);
  });
  table.appendChild(tbody);
  requirementsResults.appendChild(table);
}

addRequirementButton.addEventListener("click", () => {
  const row = Number(requirementRow.value);
  const start = Number(requirementStart.value);
  const end = Number(requirementEnd.value);
  const type = requirementType.value;
  const expected = requirementText.value;
  if (!Number.isInteger(row) || row < 1 || row > 13 || start < 1 || end < start || end > 38) {
    setStatus("Enter row 1-13 and a valid column range 1-38");
    return;
  }
  if (type !== "blank" && !expected) {
    setStatus("Enter the expected text or character");
    return;
  }
  if (type === "fill" && expected.length !== 1) {
    setStatus("Fill range requires exactly one character");
    return;
  }
  state.requirements.push({
    row,
    start,
    end,
    type,
    expected,
    ignoreCase: requirementIgnoreCase.checked,
    ignoreSpaces: requirementIgnoreSpaces.checked,
  });
  renderRequirementList();
  setStatus(`${state.requirements.length} requirement${state.requirements.length === 1 ? "" : "s"} ready`);
});

requirementList.addEventListener("click", (event) => {
  const button = event.target.closest("button[data-requirement-index]");
  if (!button) {
    return;
  }
  state.requirements.splice(Number(button.dataset.requirementIndex), 1);
  renderRequirementList();
});

requirementType.addEventListener("change", () => {
  syncRequirementFields();
  autoSetRequirementEnd();
});
requirementStart.addEventListener("input", autoSetRequirementEnd);
requirementText.addEventListener("input", autoSetRequirementEnd);

reviewRequirementsButton.addEventListener("click", async () => {
  reviewRequirementsButton.disabled = true;
  setStatus("Reviewing requirements against the grid");
  try {
    const result = await postJson("/api/review-requirements", {
      requirements: state.requirements,
      grid: getCurrentGrid(),
      image: state.imageDataUrl,
      corners: state.image ? getGridCorners() : [],
    });
    renderRequirementResults(result);
    setStatus(`${result.summary.passed} passed, ${result.summary.failed} failed`);
  } catch (error) {
    setStatus(error.message);
  } finally {
    reviewRequirementsButton.disabled = false;
  }
});

window.addEventListener("resize", fitCanvas);
renderGridTable();
fitCanvas();
syncRequirementFields();
renderRequirementList();
