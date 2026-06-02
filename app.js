const COLORS = {
  oxford: "#002147",
  deepBlue: "#1f5a86",
  ink: "#17324b",
  muted: "#637487",
  paper: "#f7f8f6",
  grid: "#d9e0e3",
  month: "#e6eef2",
};

const MORANDI = {
  4: "#c7dceb",
  5: "#e5cdb8",
  6: "#bfd8ba",
  7: "#d8c8e2",
  8: "#c7decf",
  9: "#e3c7c0",
  10: "#c3d3e8",
  11: "#dde2bf",
};

const GRADE_COLORS = [
  "#e5f0dc",
  "#eee7fb",
  "#f4e7dc",
  "#e2f1f4",
  "#f2edcf",
  "#e6eefa",
  "#e1f0e8",
  "#f1e2e8",
  "#e9edf0",
];

const TERM_COLORS = {
  summer: "#dcecf6",
  first: "#e9f1dc",
  winter: "#f4e6dc",
  second: "#eee7f6",
  fallback: "#e9edf0",
};

const STAGE_COLORS = {
  "规划季": "#cfe5ee",
  "申请季": "#d7e8d2",
};

const state = {
  model: null,
  filename: "学生时间规划表",
  scale: 0.32,
  selectedCardId: null,
  nextCardId: 1,
};

const els = {
  fileInput: document.getElementById("fileInput"),
  titleInput: document.getElementById("titleInput"),
  scaleInput: document.getElementById("scaleInput"),
  qualityInput: document.getElementById("qualityInput"),
  feishuUrl: document.getElementById("feishuUrl"),
  loadFeishuSheets: document.getElementById("loadFeishuSheets"),
  loadFeishuData: document.getElementById("loadFeishuData"),
  feishuSheet: document.getElementById("feishuSheet"),
  addCard: document.getElementById("addCard"),
  deleteCard: document.getElementById("deleteCard"),
  selectionBox: document.getElementById("selectionBox"),
  downloadPng: document.getElementById("downloadPng"),
  downloadPdf: document.getElementById("downloadPdf"),
  statusText: document.getElementById("statusText"),
  stage: document.getElementById("timelineStage"),
};

function extractSpreadsheetToken(input) {
  const text = String(input || "").trim();
  if (!text) return "";
  const match = text.match(/\/(?:sheets|spreadsheet)\/([A-Za-z0-9_-]+)/);
  if (match) return match[1];
  return text;
}

function readCell(ws, row, col) {
  const ref = XLSX.utils.encode_cell({ r: row - 1, c: col - 1 });
  const cell = ws[ref];
  return cell ? cell.v : "";
}

function usedBounds(ws) {
  const range = XLSX.utils.decode_range(ws["!ref"]);
  return {
    maxRow: range.e.r + 1,
    maxCol: range.e.c + 1,
  };
}

function mergeLookup(ws) {
  const map = new Map();
  const covered = new Map();
  for (const merge of ws["!merges"] || []) {
    const startRow = merge.s.r + 1;
    const startCol = merge.s.c + 1;
    const endRow = merge.e.r + 1;
    const endCol = merge.e.c + 1;
    map.set(`${startRow}:${startCol}`, { endRow, endCol });
    for (let row = startRow; row <= endRow; row += 1) {
      for (let col = startCol; col <= endCol; col += 1) {
        covered.set(`${row}:${col}`, { startRow, startCol, endRow, endCol });
      }
    }
  }
  return { map, covered };
}

function measureText(text, size, bold = true) {
  const canvas = measureText.canvas || (measureText.canvas = document.createElement("canvas"));
  const ctx = canvas.getContext("2d");
  ctx.font = `${bold ? 800 : 400} ${size}px "PingFang SC", "STHeiti", "Microsoft YaHei", Arial`;
  return ctx.measureText(String(text || "")).width;
}

function calculateWidths(values, dataStart, dataEnd) {
  const widths = { 1: 92, 2: 176, 3: 82 };
  const minW = { 4: 290, 5: 190, 6: 320, 7: 310, 8: 300, 9: 200, 10: 390, 11: 250 };
  const maxW = { 4: 370, 5: 250, 6: 400, 7: 390, 8: 380, 9: 270, 10: 520, 11: 330 };
  for (let col = 4; col <= 11; col += 1) {
    let best = measureText(values.get(`2:${col}`) || "", 31, true) + 40;
    for (let row = dataStart; row <= dataEnd; row += 1) {
      const value = values.get(`${row}:${col}`);
      if (!value) continue;
      String(value).split("\n").forEach((line) => {
        best = Math.max(best, Math.min(measureText(line || " ", 30, true) + 42, maxW[col]));
      });
    }
    widths[col] = Math.max(minW[col], Math.min(maxW[col], Math.round(best)));
  }
  return Array.from({ length: 11 }, (_, index) => widths[index + 1]);
}

function mergedRangesForCol(values, merges, col, dataStart, dataEnd) {
  const ranges = [];
  const seen = new Set();
  for (let row = dataStart; row <= dataEnd; row += 1) {
    if (seen.has(row)) continue;
    const label = values.get(`${row}:${col}`);
    if (!label) continue;
    const merge = merges.map.get(`${row}:${col}`) || { endRow: row, endCol: col };
    const endRow = Math.min(merge.endRow, dataEnd);
    ranges.push({ label: String(label), row, endRow });
    for (let item = row; item <= endRow; item += 1) seen.add(item);
  }
  return ranges;
}

function termColor(label) {
  const text = String(label || "");
  if (text.includes("暑期") || text.includes("暑假") || /summer/i.test(text)) return TERM_COLORS.summer;
  if (text.includes("上学期") || text.includes("秋季") || /fall|autumn/i.test(text)) return TERM_COLORS.first;
  if (text.includes("寒假") || text.includes("寒期") || /winter/i.test(text)) return TERM_COLORS.winter;
  if (text.includes("下学期") || text.includes("春季") || /spring/i.test(text)) return TERM_COLORS.second;
  return TERM_COLORS.fallback;
}

function parseWorkbook(buffer, filename) {
  const wb = XLSX.read(buffer, { type: "array", cellDates: false });
  const ws = wb.Sheets[wb.SheetNames[0]];
  const bounds = usedBounds(ws);
  const values = new Map();
  for (let row = 1; row <= bounds.maxRow; row += 1) {
    for (let col = 1; col <= Math.max(bounds.maxCol, 11); col += 1) {
      const value = readCell(ws, row, col);
      if (value !== undefined && value !== null && value !== "") values.set(`${row}:${col}`, value);
    }
  }

  const dataStart = 3;
  const dataEnd = bounds.maxRow;
  const rowCount = dataEnd - dataStart + 1;
  const title = String(values.get("1:1") || filename.replace(/\.(xlsx|xls)$/i, ""));
  const merges = mergeLookup(ws);
  const colW = calculateWidths(values, dataStart, dataEnd);
  const margin = 82;
  const titleH = 286;
  const headerH = 112;
  const rowH = 138;
  const monthRailW = 68;
  const dataW = colW.reduce((sum, item) => sum + item, 0);
  const width = dataW + monthRailW + margin * 2;
  const height = margin + titleH + headerH + rowH * rowCount + 142;
  const x0 = margin;
  const y0 = margin + titleH;
  const xs = [x0];
  for (let i = 0; i < colW.length - 1; i += 1) xs.push(xs[xs.length - 1] + colW[i]);

  const model = {
    title,
    dataStart,
    dataEnd,
    rowCount,
    margin,
    titleH,
    headerH,
    rowH,
    monthRailW,
    dataW,
    width,
    height,
    x0,
    y0,
    colW,
    xs,
    values,
    merges,
    cards: [],
  };

  collectCards(model);
  return model;
}

function parseOnlineValues(values, filename, sheetTitle) {
  const matrix = Array.isArray(values) ? values : [];
  const rowCount = Math.max(matrix.length, 3);
  const dataEnd = rowCount;
  const valueMap = new Map();
  for (let row = 1; row <= rowCount; row += 1) {
    const sourceRow = matrix[row - 1] || [];
    for (let col = 1; col <= 11; col += 1) {
      const value = sourceRow[col - 1];
      if (value !== undefined && value !== null && value !== "") {
        valueMap.set(`${row}:${col}`, value);
      }
    }
  }
  const title = String(valueMap.get("1:1") || sheetTitle || filename || "学生时间规划表");
  const merges = inferOnlineMerges(valueMap, 3, dataEnd);
  const model = buildModelFromValues(valueMap, merges, dataEnd, title);
  collectCards(model);
  return model;
}

function buildModelFromValues(values, merges, dataEnd, title) {
  const dataStart = 3;
  const rowCount = dataEnd - dataStart + 1;
  const colW = calculateWidths(values, dataStart, dataEnd);
  const margin = 82;
  const titleH = 286;
  const headerH = 112;
  const rowH = 138;
  const monthRailW = 68;
  const dataW = colW.reduce((sum, item) => sum + item, 0);
  const width = dataW + monthRailW + margin * 2;
  const height = margin + titleH + headerH + rowH * rowCount + 142;
  const x0 = margin;
  const y0 = margin + titleH;
  const xs = [x0];
  for (let i = 0; i < colW.length - 1; i += 1) xs.push(xs[xs.length - 1] + colW[i]);
  return {
    title,
    dataStart,
    dataEnd,
    rowCount,
    margin,
    titleH,
    headerH,
    rowH,
    monthRailW,
    dataW,
    width,
    height,
    x0,
    y0,
    colW,
    xs,
    values,
    merges,
    cards: [],
  };
}

function inferOnlineMerges(values, dataStart, dataEnd) {
  const map = new Map();
  const covered = new Map();
  [1, 2].forEach((col) => {
    let row = dataStart;
    while (row <= dataEnd) {
      const label = values.get(`${row}:${col}`);
      if (!label) {
        row += 1;
        continue;
      }
      let endRow = row;
      while (endRow + 1 <= dataEnd && !values.get(`${endRow + 1}:${col}`)) {
        endRow += 1;
      }
      if (endRow > row) {
        map.set(`${row}:${col}`, { endRow, endCol: col });
        for (let rr = row; rr <= endRow; rr += 1) {
          covered.set(`${rr}:${col}`, { startRow: row, startCol: col, endRow, endCol: col });
        }
      }
      row = endRow + 1;
    }
  });
  return { map, covered };
}

function collectCards(model) {
  const drawn = new Set();
  for (let row = model.dataStart; row <= model.dataEnd; row += 1) {
    for (let col = 4; col <= 11; col += 1) {
      if (drawn.has(`${row}:${col}`)) continue;
      const value = model.values.get(`${row}:${col}`);
      if (value === undefined || value === null || value === "") continue;
      const merge = model.merges.map.get(`${row}:${col}`) || { endRow: row, endCol: col };
      for (let rr = row; rr <= merge.endRow; rr += 1) {
        for (let cc = col; cc <= merge.endCol; cc += 1) drawn.add(`${rr}:${cc}`);
      }
      model.cards.push({
        id: `card-${state.nextCardId++}`,
        row,
        col,
        endRow: Math.min(merge.endRow, model.dataEnd),
        endCol: Math.min(merge.endCol, 11),
        text: String(value),
        fill: MORANDI[col] || "#d8d8d0",
      });
    }
  }
}

function el(className, style = {}, text = "") {
  const node = document.createElement("div");
  node.className = className;
  Object.entries(style).forEach(([key, value]) => {
    node.style[key] = typeof value === "number" ? `${value}px` : value;
  });
  if (text !== "") node.textContent = text;
  return node;
}

function addCell(parent, className, x, y, w, h, text, size, background) {
  const node = el(`cell ${className}`, {
    left: x,
    top: y,
    width: w,
    height: h,
    fontSize: size,
    background,
  }, text);
  parent.appendChild(node);
  return node;
}

function getTaskLeft(model) {
  return model.xs[3];
}

function getTaskRight(model) {
  return model.x0 + model.dataW;
}

function getBodyTop(model) {
  return model.y0 + model.headerH;
}

function getBodyBottom(model) {
  return getBodyTop(model) + model.rowH * model.rowCount;
}

function cardBox(model, card) {
  const x = model.xs[card.col - 1];
  const y = getBodyTop(model) + (card.row - model.dataStart) * model.rowH;
  const w = model.colW.slice(card.col - 1, card.endCol).reduce((sum, item) => sum + item, 0);
  const h = (card.endRow - card.row + 1) * model.rowH;
  return { x, y, w, h };
}

function columnAt(model, x) {
  const clamped = Math.max(getTaskLeft(model), Math.min(getTaskRight(model) - 1, x));
  for (let col = 4; col <= 11; col += 1) {
    const left = model.xs[col - 1];
    const right = left + model.colW[col - 1];
    if (clamped >= left && clamped < right) return col;
  }
  return 11;
}

function rowAt(model, y) {
  const clamped = Math.max(getBodyTop(model), Math.min(getBodyBottom(model) - 1, y));
  return model.dataStart + Math.floor((clamped - getBodyTop(model)) / model.rowH);
}

function setSelectedCard(id) {
  state.selectedCardId = id;
  const card = state.model?.cards.find((item) => item.id === id);
  els.deleteCard.disabled = !card;
  els.selectionBox.textContent = card
    ? `已选中：第 ${card.row - state.model.dataStart + 1} 个月份，${card.col - 3} - ${card.endCol - 3} 列`
    : "未选中色块";
  document.querySelectorAll(".card").forEach((node) => {
    node.classList.toggle("selected", node.dataset.cardId === id);
  });
}

function updateCardElement(model, card, node) {
  const box = cardBox(model, card);
  const w = box.w - 20;
  const h = box.h - 20;
  node.style.left = `${box.x + 10}px`;
  node.style.top = `${box.y + 10}px`;
  node.style.width = `${w}px`;
  node.style.height = `${h}px`;
  node.style.fontSize = `${fitCardFont(card.text, w - 30, h - 16)}px`;
}

function beginCardGesture(event, card, node, mode) {
  event.preventDefault();
  event.stopPropagation();
  setSelectedCard(card.id);
  const model = state.model;
  const start = {
    x: event.clientX,
    y: event.clientY,
    row: card.row,
    col: card.col,
    endRow: card.endRow,
    endCol: card.endCol,
  };
  const widthCols = start.endCol - start.col;
  const heightRows = start.endRow - start.row;

  function move(pointerEvent) {
    const dx = (pointerEvent.clientX - start.x) / state.scale;
    const dy = (pointerEvent.clientY - start.y) / state.scale;

    if (mode === "move") {
      const startBox = cardBox(model, start);
      const nextCol = columnAt(model, startBox.x + dx);
      const nextRow = rowAt(model, startBox.y + dy);
      card.col = Math.max(4, Math.min(11 - widthCols, nextCol));
      card.endCol = card.col + widthCols;
      card.row = Math.max(model.dataStart, Math.min(model.dataEnd - heightRows, nextRow));
      card.endRow = card.row + heightRows;
    }

    if (mode === "left") {
      card.col = Math.min(card.endCol, columnAt(model, model.xs[start.col - 1] + dx));
    }

    if (mode === "right") {
      const rightEdge = model.xs[start.endCol - 1] + model.colW[start.endCol - 1] + dx;
      card.endCol = Math.max(card.col, columnAt(model, rightEdge));
    }

    if (mode === "top") {
      card.row = Math.min(card.endRow, rowAt(model, getBodyTop(model) + (start.row - model.dataStart) * model.rowH + dy));
    }

    if (mode === "bottom") {
      const bottomEdge = getBodyTop(model) + (start.endRow - model.dataStart + 1) * model.rowH + dy;
      card.endRow = Math.max(card.row, rowAt(model, bottomEdge));
    }

    card.fill = MORANDI[card.col] || card.fill;
    updateCardElement(model, card, node);
    setSelectedCard(card.id);
  }

  function end() {
    window.removeEventListener("pointermove", move);
    window.removeEventListener("pointerup", end);
    renderDom();
    setSelectedCard(card.id);
  }

  window.addEventListener("pointermove", move);
  window.addEventListener("pointerup", end, { once: true });
}

function addNewCard() {
  if (!state.model) return;
  const card = {
    id: `card-${state.nextCardId++}`,
    row: state.model.dataStart,
    col: 4,
    endRow: state.model.dataStart,
    endCol: 4,
    text: "新增计划",
    fill: MORANDI[4],
  };
  state.model.cards.push(card);
  renderDom();
  setSelectedCard(card.id);
}

function deleteSelectedCard() {
  if (!state.model || !state.selectedCardId) return;
  state.model.cards = state.model.cards.filter((card) => card.id !== state.selectedCardId);
  state.selectedCardId = null;
  renderDom();
  setSelectedCard(null);
}

function renderDom() {
  const model = state.model;
  if (!model) return;
  els.stage.className = "timeline-stage";
  els.stage.innerHTML = "";
  els.stage.style.width = `${model.width}px`;
  els.stage.style.height = `${model.height}px`;
  els.stage.style.transform = `scale(${state.scale})`;

  els.stage.appendChild(el("header-bg", { left: 0, top: 0, width: model.width, height: model.margin + model.titleH - 20 }));
  els.stage.appendChild(el("header-line", { left: 0, top: model.margin + model.titleH - 18, width: model.width }));
  els.stage.appendChild(el("title-text", {
    left: model.margin,
    top: 72,
    width: model.width - model.margin * 2,
    fontSize: 92,
  }, model.title));
  els.stage.appendChild(el("logo-text", {
    left: model.width - model.margin - 620,
    top: 76,
    width: 600,
    fontSize: 40,
  }, "星屿国际教育\nAstral Academy"));

  const tableW = model.dataW + model.monthRailW;
  els.stage.appendChild(el("panel-bg", {
    left: model.x0 - 18,
    top: model.y0 - 18,
    width: tableW + 36,
    height: model.headerH + model.rowH * model.rowCount + 36,
  }));

  const headers = Array.from({ length: 11 }, (_, index) => String(model.values.get(`2:${index + 1}`) || ""));
  headers[0] = "阶段";
  headers.forEach((text, index) => {
    addCell(els.stage, "header-cell", model.xs[index], model.y0, model.colW[index], model.headerH, text, 31, COLORS.oxford);
  });
  addCell(els.stage, "header-cell", model.x0 + model.dataW, model.y0, model.monthRailW, model.headerH, "", 24, COLORS.month);

  for (let i = 0; i < model.rowCount; i += 1) {
    const y = model.y0 + model.headerH + i * model.rowH;
    els.stage.appendChild(el("cell", {
      left: model.x0,
      top: y,
      width: tableW,
      height: model.rowH,
      background: i % 2 === 0 ? "#fbfcfb" : "#f1f4f2",
    }));
  }

  mergedRangesForCol(model.values, model.merges, 2, model.dataStart, model.dataEnd).forEach((range) => {
    const y = model.y0 + model.headerH + (range.row - model.dataStart) * model.rowH;
    const h = (range.endRow - range.row + 1) * model.rowH;
    addCell(els.stage, "grade-cell", model.xs[1], y, model.colW[1], h, range.label, 34, termColor(range.label));
  });

  mergedRangesForCol(model.values, model.merges, 1, model.dataStart, model.dataEnd).forEach((range) => {
    const y = model.y0 + model.headerH + (range.row - model.dataStart) * model.rowH;
    const h = (range.endRow - range.row + 1) * model.rowH;
    addCell(els.stage, "stage-cell", model.xs[0], y, model.colW[0], h, range.label, 31, STAGE_COLORS[range.label] || "#d9e5ea");
  });

  for (let row = model.dataStart; row <= model.dataEnd; row += 1) {
    const y = model.y0 + model.headerH + (row - model.dataStart) * model.rowH;
    addCell(els.stage, "month-cell", model.xs[2], y, model.colW[2], model.rowH, String(model.values.get(`${row}:3`) || ""), 36, COLORS.month);
    addCell(els.stage, "month-rail", model.x0 + model.dataW, y, model.monthRailW, model.rowH, String(model.values.get(`${row}:3`) || ""), 26, COLORS.month);
  }

  model.cards.forEach((card) => {
    const box = cardBox(model, card);
    const w = box.w - 20;
    const h = box.h - 20;
    const node = el("card", {
      left: box.x + 10,
      top: box.y + 10,
      width: w,
      height: h,
      background: card.fill,
      fontSize: fitCardFont(card.text, w - 30, h - 16),
    });
    const textNode = el("card-text", {}, card.text);
    textNode.contentEditable = "true";
    textNode.spellcheck = false;
    textNode.addEventListener("pointerdown", (event) => {
      event.stopPropagation();
      setSelectedCard(card.id);
    });
    textNode.addEventListener("input", () => {
      card.text = textNode.textContent || "";
      node.style.fontSize = `${fitCardFont(card.text, w - 30, h - 16)}px`;
    });
    node.appendChild(textNode);

    const moveHandle = el("move-handle");
    moveHandle.title = "拖动色块";
    moveHandle.addEventListener("pointerdown", (event) => beginCardGesture(event, card, node, "move"));
    node.appendChild(moveHandle);

    ["left", "right", "top", "bottom"].forEach((mode) => {
      const handle = el(`resize-handle resize-${mode}`);
      handle.title = "调整大小";
      handle.addEventListener("pointerdown", (event) => beginCardGesture(event, card, node, mode));
      node.appendChild(handle);
    });

    node.dataset.cardId = card.id;
    node.addEventListener("pointerdown", (event) => {
      if (event.target === node) beginCardGesture(event, card, node, "move");
    });
    node.classList.toggle("selected", card.id === state.selectedCardId);
    els.stage.appendChild(node);
  });

  const fy = model.y0 + model.headerH + model.rowH * model.rowCount + 48;
  els.stage.appendChild(el("footer-line", { left: model.margin, top: fy, width: model.width - model.margin * 2 }));
  els.stage.appendChild(el("footer-left", { left: model.margin, top: fy + 24, width: 520, fontSize: 30 }, "星屿国际教育 Astral Academy"));
  els.stage.appendChild(el("footer-right", { left: model.width - model.margin - 720, top: fy + 24, width: 720, fontSize: 28 }, model.title));
  if (state.selectedCardId && !model.cards.some((card) => card.id === state.selectedCardId)) {
    state.selectedCardId = null;
  }
  setSelectedCard(state.selectedCardId);
}

function wrapText(ctx, text, maxWidth) {
  const tokens = [];
  String(text || "").split("\n").forEach((line, index) => {
    if (index > 0) tokens.push("\n");
    let current = "";
    let mode = "";
    for (const ch of line) {
      const nextMode = /^[\w /&.\-:_()]$/.test(ch) ? "ascii" : "cjk";
      if (current && nextMode !== mode) {
        tokens.push(current);
        current = ch;
      } else {
        current += ch;
      }
      mode = nextMode;
    }
    if (current) tokens.push(current);
  });

  const lines = [];
  let current = "";
  tokens.forEach((token) => {
    if (token === "\n") {
      lines.push(current);
      current = "";
      return;
    }
    const trial = current + token;
    if (ctx.measureText(trial).width <= maxWidth || !current) {
      if (!current && ctx.measureText(token).width > maxWidth) {
        for (const ch of token) {
          const charTrial = current + ch;
          if (ctx.measureText(charTrial).width <= maxWidth || !current) current = charTrial;
          else {
            lines.push(current.trimEnd());
            current = ch;
          }
        }
      } else {
        current = trial;
      }
      return;
    }
    lines.push(current.trimEnd());
    current = token.trimStart();
  });
  if (current) lines.push(current.trimEnd());
  return lines;
}

function fitCardFont(text, maxW, maxH) {
  const canvas = measureText.canvas || (measureText.canvas = document.createElement("canvas"));
  const ctx = canvas.getContext("2d");
  for (let size = 32; size >= 23; size -= 1) {
    ctx.font = `800 ${size}px "PingFang SC", "STHeiti", "Microsoft YaHei", Arial`;
    const lines = wrapText(ctx, text, maxW);
    if (lines.length * size * 1.16 <= maxH) return size;
  }
  return 23;
}

function roundRect(ctx, x, y, w, h, radius) {
  const r = Math.min(radius, w / 2, h / 2);
  ctx.beginPath();
  ctx.moveTo(x + r, y);
  ctx.lineTo(x + w - r, y);
  ctx.quadraticCurveTo(x + w, y, x + w, y + r);
  ctx.lineTo(x + w, y + h - r);
  ctx.quadraticCurveTo(x + w, y + h, x + w - r, y + h);
  ctx.lineTo(x + r, y + h);
  ctx.quadraticCurveTo(x, y + h, x, y + h - r);
  ctx.lineTo(x, y + r);
  ctx.quadraticCurveTo(x, y, x + r, y);
  ctx.closePath();
}

function drawCentered(ctx, x, y, w, h, text, size, color, vertical = false) {
  ctx.fillStyle = color;
  ctx.font = `800 ${size}px "PingFang SC", "STHeiti", "Microsoft YaHei", Arial`;
  ctx.textAlign = "center";
  ctx.textBaseline = "middle";
  if (vertical) {
    const chars = String(text || "").split("");
    const lineH = size * 1.12;
    let cy = y + h / 2 - ((chars.length - 1) * lineH) / 2;
    chars.forEach((ch) => {
      ctx.fillText(ch, x + w / 2, cy);
      cy += lineH;
    });
    return;
  }
  String(text || "").split("\n").forEach((line, index, lines) => {
    const lineH = size * 1.18;
    const cy = y + h / 2 - ((lines.length - 1) * lineH) / 2 + index * lineH;
    ctx.fillText(line, x + w / 2, cy);
  });
}

function renderToCanvas(multiplier = 2) {
  const model = state.model;
  const safeMultiplier = getSafeMultiplier(model, multiplier);
  const canvas = document.createElement("canvas");
  canvas.width = Math.round(model.width * safeMultiplier);
  canvas.height = Math.round(model.height * safeMultiplier);
  const ctx = canvas.getContext("2d");
  ctx.scale(safeMultiplier, safeMultiplier);

  ctx.fillStyle = COLORS.paper;
  ctx.fillRect(0, 0, model.width, model.height);
  ctx.fillStyle = COLORS.oxford;
  ctx.fillRect(0, 0, model.width, model.margin + model.titleH - 20);
  ctx.fillStyle = "#c6d5dc";
  ctx.fillRect(0, model.margin + model.titleH - 18, model.width, 4);
  drawCentered(ctx, model.margin, 72, model.width - model.margin * 2, 110, model.title, 92, "#ffffff");
  ctx.textAlign = "right";
  ctx.textBaseline = "top";
  ctx.fillStyle = "#ffffff";
  ctx.font = `800 40px "PingFang SC", "STHeiti", "Microsoft YaHei", Arial`;
  ctx.fillText("星屿国际教育", model.width - model.margin - 20, 76);
  ctx.font = `800 35px "PingFang SC", "STHeiti", "Microsoft YaHei", Arial`;
  ctx.fillText("Astral Academy", model.width - model.margin - 20, 128);

  const tableW = model.dataW + model.monthRailW;
  ctx.fillStyle = "#ffffff";
  roundRect(ctx, model.x0 - 18, model.y0 - 18, tableW + 36, model.headerH + model.rowH * model.rowCount + 36, 18);
  ctx.fill();
  ctx.strokeStyle = "#e1e6e8";
  ctx.lineWidth = 2;
  ctx.stroke();

  const headers = Array.from({ length: 11 }, (_, index) => String(model.values.get(`2:${index + 1}`) || ""));
  headers[0] = "阶段";
  headers.forEach((text, index) => {
    ctx.fillStyle = COLORS.oxford;
    ctx.fillRect(model.xs[index], model.y0, model.colW[index], model.headerH);
    drawCentered(ctx, model.xs[index] + 8, model.y0, model.colW[index] - 16, model.headerH, text, 31, "#ffffff");
  });
  ctx.fillStyle = COLORS.month;
  ctx.fillRect(model.x0 + model.dataW, model.y0, model.monthRailW, model.headerH);

  for (let i = 0; i < model.rowCount; i += 1) {
    const y = model.y0 + model.headerH + i * model.rowH;
    ctx.fillStyle = i % 2 === 0 ? "#fbfcfb" : "#f1f4f2";
    ctx.fillRect(model.x0, y, tableW, model.rowH);
    ctx.fillStyle = COLORS.month;
    ctx.fillRect(model.x0 + model.dataW, y, model.monthRailW, model.rowH);
    ctx.strokeStyle = COLORS.grid;
    ctx.lineWidth = i % 2 === 0 ? 2 : 1;
    ctx.beginPath();
    ctx.moveTo(model.x0, y);
    ctx.lineTo(model.x0 + tableW, y);
    ctx.stroke();
  }

  mergedRangesForCol(model.values, model.merges, 2, model.dataStart, model.dataEnd).forEach((range) => {
    const y = model.y0 + model.headerH + (range.row - model.dataStart) * model.rowH;
    const h = (range.endRow - range.row + 1) * model.rowH;
    ctx.fillStyle = termColor(range.label);
    ctx.fillRect(model.xs[1], y, model.colW[1], h);
    ctx.strokeStyle = "#c9d3d8";
    ctx.lineWidth = 2;
    ctx.strokeRect(model.xs[1], y, model.colW[1], h);
    drawCentered(ctx, model.xs[1], y, model.colW[1], h, range.label, 34, COLORS.ink);
  });

  mergedRangesForCol(model.values, model.merges, 1, model.dataStart, model.dataEnd).forEach((range) => {
    const y = model.y0 + model.headerH + (range.row - model.dataStart) * model.rowH;
    const h = (range.endRow - range.row + 1) * model.rowH;
    ctx.fillStyle = STAGE_COLORS[range.label] || "#d9e5ea";
    ctx.fillRect(model.xs[0], y, model.colW[0], h);
    ctx.strokeStyle = "#c9d3d8";
    ctx.lineWidth = 2;
    ctx.strokeRect(model.xs[0], y, model.colW[0], h);
    drawCentered(ctx, model.xs[0], y, model.colW[0], h, range.label, 31, COLORS.oxford, true);
  });

  model.cards.forEach((card) => {
    const x = model.xs[card.col - 1];
    const y = model.y0 + model.headerH + (card.row - model.dataStart) * model.rowH;
    const w = model.colW.slice(card.col - 1, card.endCol).reduce((sum, item) => sum + item, 0);
    const h = (card.endRow - card.row + 1) * model.rowH;
    const pad = 10;
    ctx.fillStyle = card.fill;
    roundRect(ctx, x + pad, y + pad, w - 2 * pad, h - 2 * pad, 12);
    ctx.fill();
    ctx.strokeStyle = "#ffffff";
    ctx.lineWidth = 2;
    ctx.stroke();
    const fontSize = fitCardFont(card.text, w - 50, h - 36);
    ctx.font = `800 ${fontSize}px "PingFang SC", "STHeiti", "Microsoft YaHei", Arial`;
    ctx.fillStyle = COLORS.ink;
    ctx.textAlign = "left";
    ctx.textBaseline = "top";
    const lines = wrapText(ctx, card.text, w - 50);
    const lineH = fontSize * 1.14;
    let ty = y + h / 2 - (lines.length * lineH) / 2;
    lines.forEach((line) => {
      ctx.fillText(line, x + pad + 15, ty);
      ty += lineH;
    });
  });

  for (let row = model.dataStart; row <= model.dataEnd; row += 1) {
    const y = model.y0 + model.headerH + (row - model.dataStart) * model.rowH;
    const month = String(model.values.get(`${row}:3`) || "");
    ctx.fillStyle = COLORS.month;
    ctx.fillRect(model.xs[2], y, model.colW[2], model.rowH);
    ctx.strokeStyle = "#c9d3d8";
    ctx.lineWidth = 2;
    ctx.strokeRect(model.xs[2], y, model.colW[2], model.rowH);
    drawCentered(ctx, model.xs[2], y, model.colW[2], model.rowH, month, 36, COLORS.oxford);
    ctx.strokeStyle = "#d8e1e5";
    ctx.lineWidth = 1;
    ctx.strokeRect(model.x0 + model.dataW, y, model.monthRailW, model.rowH);
    drawCentered(ctx, model.x0 + model.dataW, y, model.monthRailW, model.rowH, month, 26, "#8ea1ad");
  }

  const fy = model.y0 + model.headerH + model.rowH * model.rowCount + 48;
  ctx.fillStyle = "#c9d6dc";
  ctx.fillRect(model.margin, fy, model.width - model.margin * 2, 2);
  ctx.textAlign = "left";
  ctx.textBaseline = "top";
  ctx.fillStyle = COLORS.deepBlue;
  ctx.font = `800 30px "PingFang SC", "STHeiti", "Microsoft YaHei", Arial`;
  ctx.fillText("星屿国际教育 Astral Academy", model.margin, fy + 24);
  ctx.textAlign = "right";
  ctx.fillStyle = COLORS.muted;
  ctx.font = `400 28px "PingFang SC", "STHeiti", "Microsoft YaHei", Arial`;
  ctx.fillText(model.title, model.width - model.margin, fy + 24);

  return canvas;
}

function getSafeMultiplier(model, requested) {
  const maxPixels = 36000000;
  const maxSide = 12000;
  const byArea = Math.sqrt(maxPixels / (model.width * model.height));
  const bySide = Math.min(maxSide / model.width, maxSide / model.height);
  return Math.max(1, Math.min(Number(requested) || 2, byArea, bySide));
}

function canvasToBlob(canvas, type = "image/png", quality) {
  return new Promise((resolve, reject) => {
    canvas.toBlob((blob) => {
      if (blob) resolve(blob);
      else reject(new Error("浏览器没有生成导出文件。"));
    }, type, quality);
  });
}

function downloadBlob(blob, filename) {
  const url = URL.createObjectURL(blob);
  const link = document.createElement("a");
  link.download = filename;
  link.href = url;
  document.body.appendChild(link);
  link.click();
  link.remove();
  window.setTimeout(() => URL.revokeObjectURL(url), 1000);
}

async function withExportState(message, action) {
  const oldStatus = els.statusText.textContent;
  els.downloadPng.disabled = true;
  els.downloadPdf.disabled = true;
  els.statusText.textContent = message;
  try {
    await action();
    els.statusText.textContent = "导出完成。";
  } catch (error) {
    console.error(error);
    els.statusText.textContent = `导出失败：${error.message || "请换成标准倍率后重试。"}`;
  } finally {
    if (state.model) {
      els.downloadPng.disabled = false;
      els.downloadPdf.disabled = false;
    }
    window.setTimeout(() => {
      if (state.model && els.statusText.textContent === "导出完成。") {
        els.statusText.textContent = oldStatus;
      }
    }, 1800);
  }
}

function safeFilename(name, ext) {
  return `${String(name || "学生时间规划表").replace(/[\\/:*?"<>| ]+/g, "")}_可编辑导出.${ext}`;
}

function setReady(ready) {
  els.titleInput.disabled = !ready;
  els.addCard.disabled = !ready;
  els.deleteCard.disabled = true;
  els.downloadPng.disabled = !ready;
  els.downloadPdf.disabled = !ready;
  if (!ready) {
    state.selectedCardId = null;
    els.selectionBox.textContent = "未选中色块";
  }
}

async function feishuRequest(params) {
  const query = new URLSearchParams(params);
  const response = await fetch(`/.netlify/functions/feishu-sheets?${query.toString()}`);
  const data = await response.json().catch(() => ({}));
  if (!response.ok || data.error) {
    throw new Error(data.error || `飞书接口请求失败：${response.status}`);
  }
  return data;
}

function setFeishuSheets(sheets) {
  els.feishuSheet.innerHTML = "";
  if (!sheets.length) {
    const option = document.createElement("option");
    option.value = "";
    option.textContent = "没有读取到子表";
    els.feishuSheet.appendChild(option);
    els.feishuSheet.disabled = true;
    els.loadFeishuData.disabled = true;
    return;
  }
  sheets.forEach((sheet) => {
    const option = document.createElement("option");
    option.value = sheet.sheetId;
    option.textContent = sheet.title || sheet.sheetId;
    option.dataset.title = sheet.title || "";
    els.feishuSheet.appendChild(option);
  });
  els.feishuSheet.disabled = false;
  els.loadFeishuData.disabled = false;
}

els.fileInput.addEventListener("change", async (event) => {
  const file = event.target.files[0];
  if (!file) return;
  try {
    els.statusText.textContent = "正在读取 Excel 并生成可编辑预览...";
    const buffer = await file.arrayBuffer();
    state.filename = file.name.replace(/\.(xlsx|xls)$/i, "");
    state.selectedCardId = null;
    state.nextCardId = 1;
    state.model = parseWorkbook(buffer, file.name);
    els.titleInput.value = state.model.title;
    setReady(true);
    renderDom();
    els.statusText.textContent = "已生成。可以编辑文字，也可以新增、删除、拖动或拉伸色块。";
  } catch (error) {
    console.error(error);
    setReady(false);
    els.statusText.textContent = "读取失败，请确认 Excel 结构与现有模板一致。";
  }
});

els.titleInput.addEventListener("input", () => {
  if (!state.model) return;
  state.model.title = els.titleInput.value || state.filename;
  renderDom();
});

els.scaleInput.addEventListener("input", () => {
  state.scale = Number(els.scaleInput.value);
  renderDom();
});

els.addCard.addEventListener("click", addNewCard);
els.deleteCard.addEventListener("click", deleteSelectedCard);

els.loadFeishuSheets.addEventListener("click", async () => {
  const token = extractSpreadsheetToken(els.feishuUrl.value);
  if (!token) {
    els.statusText.textContent = "请先粘贴飞书在线表格链接。";
    return;
  }
  els.loadFeishuSheets.disabled = true;
  els.statusText.textContent = "正在读取飞书子表列表...";
  try {
    const data = await feishuRequest({ action: "sheets", token });
    setFeishuSheets(data.sheets || []);
    els.statusText.textContent = "已读取子表，请选择后点击“生成表格”。";
  } catch (error) {
    console.error(error);
    els.statusText.textContent = `读取子表失败：${error.message}`;
  } finally {
    els.loadFeishuSheets.disabled = false;
  }
});

els.loadFeishuData.addEventListener("click", async () => {
  const token = extractSpreadsheetToken(els.feishuUrl.value);
  const sheetId = els.feishuSheet.value;
  if (!token || !sheetId) {
    els.statusText.textContent = "请先读取并选择一个子表。";
    return;
  }
  els.loadFeishuData.disabled = true;
  els.statusText.textContent = "正在读取飞书在线表格数据...";
  try {
    const selected = els.feishuSheet.selectedOptions[0];
    const data = await feishuRequest({ action: "values", token, sheetId, range: "A1:K120" });
    state.filename = selected?.dataset.title || selected?.textContent || "飞书在线表格";
    state.selectedCardId = null;
    state.nextCardId = 1;
    state.model = parseOnlineValues(data.values || [], state.filename, state.filename);
    els.titleInput.value = state.model.title;
    setReady(true);
    renderDom();
    els.statusText.textContent = "已从飞书在线表格生成。可继续编辑色块并导出 PNG/PDF。";
  } catch (error) {
    console.error(error);
    els.statusText.textContent = `读取表格失败：${error.message}`;
  } finally {
    els.loadFeishuData.disabled = !els.feishuSheet.value;
  }
});

els.downloadPng.addEventListener("click", () => {
  withExportState("正在生成 PNG...", async () => {
    const canvas = renderToCanvas(Number(els.qualityInput.value));
    const blob = await canvasToBlob(canvas, "image/png");
    downloadBlob(blob, safeFilename(state.model.title, "png"));
  });
});

els.downloadPdf.addEventListener("click", () => {
  withExportState("正在生成 PDF...", async () => {
    if (!window.jspdf?.jsPDF) {
      throw new Error("PDF 组件未加载，请刷新页面后重试。");
    }
    const canvas = renderToCanvas(Number(els.qualityInput.value));
    const { jsPDF } = window.jspdf;
    const orientation = canvas.width >= canvas.height ? "landscape" : "portrait";
    const pdfWidth = orientation === "landscape" ? 1600 : 1200;
    const pdfHeight = pdfWidth * (canvas.height / canvas.width);
    const pdf = new jsPDF({
      orientation,
      unit: "pt",
      format: [pdfWidth, pdfHeight],
      compress: true,
    });
    pdf.addImage(canvas, "PNG", 0, 0, pdfWidth, pdfHeight, undefined, "FAST");
    pdf.save(safeFilename(state.model.title, "pdf"));
  });
});
