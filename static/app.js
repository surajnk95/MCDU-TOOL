const ROWS = 13;
const COLS = 40;

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
const photoViewButton = document.querySelector("#photoViewButton");
const flatViewButton = document.querySelector("#flatViewButton");
const gridTable = document.querySelector("#gridTable");

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
  imageDataUrl: "",
  imageBounds: { x: 0, y: 0, width: 0, height: 0, scale: 1 },
  flatBounds: { x: 0, y: 0, width: 0, height: 0, scale: 1 },
  corners: [],
  dragging: -1,
  viewMode: "photo",
  sourceGrid: makeEmptyGrid(),
};

function setStatus(message) {
  statusEl.textContent = message;
}

function makeEmptyGrid() {
  return Array.from({ length: ROWS }, () => Array.from({ length: COLS }, () => ""));
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

function bilinear(corners, u, v) {
  const top = lerpPoint(corners[0], corners[1], u);
  const bottom = lerpPoint(corners[3], corners[2], u);
  return lerpPoint(top, bottom, v);
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
    bilinear(state.corners, u1, v1),
    bilinear(state.corners, u2, v1),
    bilinear(state.corners, u2, v2),
    bilinear(state.corners, u1, v2),
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
  }
}

function drawGrid() {
  const corners = getGridCorners().map(imageToCanvas);

  ctx.save();
  ctx.lineWidth = 1;
  ctx.strokeStyle = "rgba(24, 154, 141, 0.92)";
  ctx.fillStyle = "#ffffff";
  ctx.font = "11px system-ui, sans-serif";
  ctx.textAlign = "center";
  ctx.textBaseline = "middle";

  for (let col = 0; col <= COLS; col += 1) {
    const u = col / COLS;
    const top = bilinear(corners, u, 0);
    const bottom = bilinear(corners, u, 1);
    ctx.beginPath();
    ctx.moveTo(top.x, top.y);
    ctx.lineTo(bottom.x, bottom.y);
    ctx.stroke();
  }

  for (let row = 0; row <= ROWS; row += 1) {
    const v = row / ROWS;
    const left = bilinear(corners, 0, v);
    const right = bilinear(corners, 1, v);
    ctx.beginPath();
    ctx.moveTo(left.x, left.y);
    ctx.lineTo(right.x, right.y);
    ctx.stroke();
  }

  ctx.fillStyle = "rgba(6, 79, 73, 0.9)";
  for (let row = 0; row < ROWS; row += 1) {
    const p = bilinear(corners, -0.035, (row + 0.5) / ROWS);
    ctx.fillText(String(row + 1), p.x, p.y);
  }

  for (let col = 1; col <= 38; col += 1) {
    const p = bilinear(corners, (col + 0.5) / COLS, -0.035);
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

function renderGridTable(grid = makeEmptyGrid()) {
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
      td.textContent = grid[row]?.[col]?.trim() || "";
      td.addEventListener("input", () => {
        exportButton.disabled = false;
        rememberButton.disabled = false;
      });
      tr.appendChild(td);
    }
    tbody.appendChild(tr);
  }
  gridTable.appendChild(tbody);
}

function getCurrentGrid() {
  const grid = makeEmptyGrid();
  gridTable.querySelectorAll("td[contenteditable='true']").forEach((cell) => {
    const row = Number(cell.dataset.row);
    const col = Number(cell.dataset.col);
    grid[row][col] = cell.textContent.slice(0, 8).trim();
  });
  return grid;
}

function gridRowText(row) {
  return row.map((cell) => (cell || " ").slice(0, 1)).join("").slice(0, COLS).padEnd(COLS, " ");
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

async function autoDetectDisplay() {
  if (!state.imageDataUrl) {
    return false;
  }
  detectButton.disabled = true;
  setStatus("Detecting display");
  try {
    const result = await postJson("/api/detect-display", {
      image: state.imageDataUrl,
    });
    if (setCorners(result.corners)) {
      Object.values(insetInputs).forEach((input) => {
        input.value = "0";
      });
      draw();
      const size = result.displaySize ? ` ${result.displaySize.width}x${result.displaySize.height}` : "";
      setStatus(`Display detected${size}`);
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
      photoViewButton.disabled = false;
      flatViewButton.disabled = false;
      analyzeButton.disabled = false;
      exportButton.disabled = true;
      rememberButton.disabled = true;
      state.flatImage = null;
      state.flatUrl = "";
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

Object.values(insetInputs).forEach((input) => {
  input.addEventListener("input", draw);
});

resetButton.addEventListener("click", () => {
  if (!state.image) {
    return;
  }
  defaultCorners();
  state.flatImage = null;
  state.flatUrl = "";
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
    });
    state.sourceGrid = result.grid;
    renderGridTable(result.grid);
    exportButton.disabled = false;
    rememberButton.disabled = false;
    const boxCount = Array.isArray(result.boxes) ? result.boxes.length : 0;
    setStatus(`OCR complete: ${boxCount} character boxes, ${result.words.length} text blocks`);
  } catch (error) {
    setStatus(error.message);
  } finally {
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
    if (original && corrected && original !== corrected) {
      changes.push({ original, corrected });
    }
  }

  rememberButton.disabled = true;
  setStatus(`Remembering ${changes.length} correction rows`);
  try {
    for (const change of changes) {
      await postJson("/api/remember", change);
    }
    const learned = await postJson("/api/remember-templates", {
      image: state.imageDataUrl,
      corners: getGridCorners(),
      grid: current,
    });
    state.sourceGrid = current;
    setStatus(`Corrections remembered, ${learned.learned} character samples learned`);
  } catch (error) {
    setStatus(error.message);
  } finally {
    rememberButton.disabled = false;
  }
});

window.addEventListener("resize", fitCanvas);
renderGridTable();
fitCanvas();
