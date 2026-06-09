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

const TASK_COLORS = [
  "#c7dceb",
  "#e5cdb8",
  "#bfd8ba",
  "#d8c8e2",
  "#c7decf",
  "#e3c7c0",
  "#c3d3e8",
  "#dde2bf",
];

const HEADER_FONT_SIZE = 31;
const BODY_FONT_SIZE = 30;
const MONTH_FONT_SIZE = 34;
const MIN_ROW_H = 98;
const MAX_ROW_H = 260;

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
  previewPng: document.getElementById("previewPng"),
  previewPdf: document.getElementById("previewPdf"),
  statusText: document.getElementById("statusText"),
  stage: document.getElementById("timelineStage"),
};

function extractSpreadsheetToken(input) {
  const text = String(input || "").trim();
  if (!text) return "";
  const match = text.match(/\/(?:sheets|spreadsheet)\/([A-Za-z0-9_-]+)/);
  if (match) return match[1];
  const wikiMatch = text.match(/\/wiki\/([A-Za-z0-9_-]+)/);
  if (wikiMatch) return text;
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

function taskColor(col) {
  return TASK_COLORS[(Math.max(4, col) - 4) % TASK_COLORS.length];
}

function detectLastCol(values, merges, dataStart, dataEnd, fallback = 3) {
  let lastCol = Math.max(3, fallback || 3);
  values.forEach((value, key) => {
    if (value === undefined || value === null || String(value).trim() === "") return;
    const [row, col] = key.split(":").map(Number);
    if (row >= 2 && row <= dataEnd) lastCol = Math.max(lastCol, col);
  });
  if (merges?.map) {
    merges.map.forEach((merge, key) => {
      const [row, col] = key.split(":").map(Number);
      if (row <= dataEnd && merge.endRow >= 2 && col <= merge.endCol) {
        lastCol = Math.max(lastCol, merge.endCol);
      }
    });
  }
  return Math.max(3, lastCol);
}

function calculateWidths(values, dataStart, dataEnd, lastCol) {
  const widths = { 1: 92, 2: 176, 3: 82 };
  for (let col = 4; col <= lastCol; col += 1) {
    let best = measureText(values.get(`2:${col}`) || "", HEADER_FONT_SIZE, true) + 42;
    for (let row = dataStart; row <= dataEnd; row += 1) {
      const value = values.get(`${row}:${col}`);
      if (!value) continue;
      String(value).split("\n").forEach((line) => {
        best = Math.max(best, Math.min(measureText(line || " ", BODY_FONT_SIZE, true) + 54, 560));
      });
    }
    widths[col] = Math.max(180, Math.min(560, Math.round(best)));
  }
  return Array.from({ length: lastCol }, (_, index) => widths[index + 1] || 180);
}

function estimateWrappedLineCount(text, width) {
  const canvas = measureText.canvas || (measureText.canvas = document.createElement("canvas"));
  const ctx = canvas.getContext("2d");
  ctx.font = `800 ${BODY_FONT_SIZE}px "PingFang SC", "STHeiti", "Microsoft YaHei", Arial`;
  return wrapText(ctx, text, Math.max(80, width)).length || 1;
}

function calculateRowHeights(values, merges, dataStart, dataEnd, colW, lastCol) {
  const rowHeights = Array.from({ length: dataEnd - dataStart + 1 }, () => MIN_ROW_H);
  for (let row = dataStart; row <= dataEnd; row += 1) {
    for (let col = 3; col <= lastCol; col += 1) {
      const value = values.get(`${row}:${col}`);
      if (!value) continue;
      const merge = merges.map.get(`${row}:${col}`) || { endRow: row, endCol: col };
      const colSpanW = colW.slice(col - 1, Math.min(merge.endCol, lastCol)).reduce((sum, item) => sum + item, 0);
      const rowSpan = Math.max(1, Math.min(merge.endRow, dataEnd) - row + 1);
      const textW = col >= 4 ? colSpanW - 50 : colSpanW - 18;
      const lines = estimateWrappedLineCount(String(value), textW);
      const needed = Math.min(MAX_ROW_H * rowSpan, Math.max(MIN_ROW_H, lines * BODY_FONT_SIZE * 1.22 + 34));
      const perRow = Math.ceil(needed / rowSpan);
      for (let target = row; target <= Math.min(merge.endRow, dataEnd); target += 1) {
        rowHeights[target - dataStart] = Math.max(rowHeights[target - dataStart], Math.min(MAX_ROW_H, perRow));
      }
    }
  }
  return rowHeights;
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
    for (let col = 1; col <= bounds.maxCol; col += 1) {
      const value = readCell(ws, row, col);
      if (value !== undefined && value !== null && value !== "") values.set(`${row}:${col}`, value);
    }
  }

  const title = String(values.get("1:1") || filename.replace(/\.(xlsx|xls)$/i, ""));
  const merges = mergeLookup(ws);
  const model = buildModelFromValues(values, merges, bounds.maxRow, title);
  collectCards(model);
  return model;
}

function parseOnlineValues(values, filename, sheetTitle, feishuMerges = []) {
  const normalizedMerges = normalizeFeishuMerges(feishuMerges);
  const matrix = trimOnlineRows(Array.isArray(values) ? values : [], normalizedMerges);
  const rowCount = Math.max(matrix.length, 3);
  const dataEnd = rowCount;
  const valueMap = new Map();
  for (let row = 1; row <= rowCount; row += 1) {
    const sourceRow = matrix[row - 1] || [];
    for (let col = 1; col <= sourceRow.length; col += 1) {
      const value = sourceRow[col - 1];
      if (value !== undefined && value !== null && value !== "") {
        valueMap.set(`${row}:${col}`, value);
      }
    }
  }
  const title = String(valueMap.get("1:1") || sheetTitle || filename || "学生时间规划表");
  const valueCols = [...valueMap.keys()].map((key) => Number(key.split(":")[1]) || 0);
  const detectedLastCol = Math.max(3, ...valueCols, ...normalizedMerges.map((merge) => merge.endCol || 0));
  const merges = normalizedMerges.length
    ? buildOnlineMergeLookup(normalizedMerges, dataEnd, detectedLastCol)
    : inferOnlineMerges(valueMap, 3, dataEnd);
  const model = buildModelFromValues(valueMap, merges, dataEnd, title);
  collectCards(model);
  return model;
}

function trimOnlineRows(matrix, merges = []) {
  const rows = matrix.map((row) => Array.isArray(row) ? row : []);
  let lastDataRow = Math.min(rows.length, 2);
  for (let index = 2; index < rows.length; index += 1) {
    const hasContent = rows[index].some((cell) => {
      if (cell === undefined || cell === null) return false;
      return String(cell).trim() !== "";
    });
    if (hasContent) lastDataRow = index + 1;
  }
  merges.forEach((merge) => {
    if (merge.startRow <= lastDataRow && merge.endRow > lastDataRow) {
      lastDataRow = merge.endRow;
    }
  });
  return rows.slice(0, Math.max(lastDataRow, 3));
}

function firstNumber(...values) {
  for (const value of values) {
    if (value !== undefined && value !== null && value !== "") return Number(value);
  }
  return null;
}

function columnLettersToNumber(letters) {
  return String(letters || "").toUpperCase().split("").reduce((sum, ch) => sum * 26 + ch.charCodeAt(0) - 64, 0);
}

function parseA1Range(range) {
  const match = String(range || "").match(/([A-Z]+)(\d+)(?::([A-Z]+)(\d+))?/i);
  if (!match) return null;
  const startCol = columnLettersToNumber(match[1]);
  const startRow = Number(match[2]);
  const endCol = match[3] ? columnLettersToNumber(match[3]) : startCol;
  const endRow = match[4] ? Number(match[4]) : startRow;
  return { startRow, startCol, endRow, endCol };
}

function normalizeFeishuMerges(rawMerges) {
  if (!Array.isArray(rawMerges)) return [];
  return rawMerges.map((merge) => {
    const a1Range = parseA1Range(merge.range || merge.range_text || merge.rangeText || merge.coord || merge.coordinate);
    if (a1Range) return a1Range;
    const startRowIndex = firstNumber(merge.start_row_index, merge.startRowIndex, merge.start_row, merge.startRow, merge.row);
    const startColIndex = firstNumber(merge.start_column_index, merge.startColumnIndex, merge.start_col, merge.startCol, merge.column, merge.col);
    if (startRowIndex === null || startColIndex === null) return null;
    const rowCount = firstNumber(merge.row_count, merge.rowCount, merge.rows);
    const colCount = firstNumber(merge.column_count, merge.columnCount, merge.col_count, merge.colCount, merge.cols);
    const endRowIndex = firstNumber(merge.end_row_index, merge.endRowIndex, merge.end_row, merge.endRow);
    const endColIndex = firstNumber(merge.end_column_index, merge.endColumnIndex, merge.end_col, merge.endCol);
    const startRow = startRowIndex + 1;
    const startCol = startColIndex + 1;
    const endRow = rowCount !== null ? startRow + rowCount - 1 : (endRowIndex !== null ? endRowIndex + 1 : startRow);
    const endCol = colCount !== null ? startCol + colCount - 1 : (endColIndex !== null ? endColIndex + 1 : startCol);
    if (endRow < startRow || endCol < startCol) return null;
    return { startRow, startCol, endRow, endCol };
  }).filter(Boolean);
}

function buildOnlineMergeLookup(merges, dataEnd, lastCol = Infinity) {
  const map = new Map();
  const covered = new Map();
  merges.forEach((merge) => {
    const startRow = Math.max(merge.startRow, 1);
    const startCol = Math.max(merge.startCol, 1);
    const endRow = Math.min(merge.endRow, dataEnd);
    const endCol = Math.min(merge.endCol, lastCol);
    if (endRow < startRow || endCol < startCol || startCol > lastCol) return;
    map.set(`${startRow}:${startCol}`, { endRow, endCol });
    for (let row = startRow; row <= endRow; row += 1) {
      for (let col = startCol; col <= endCol; col += 1) {
        covered.set(`${row}:${col}`, { startRow, startCol, endRow, endCol });
      }
    }
  });
  return { map, covered };
}

function buildModelFromValues(values, merges, dataEnd, title, fallbackLastCol = 3) {
  const dataStart = 3;
  const rowCount = dataEnd - dataStart + 1;
  const lastCol = detectLastCol(values, merges, dataStart, dataEnd, fallbackLastCol);
  const colW = calculateWidths(values, dataStart, dataEnd, lastCol);
  const rowHeights = calculateRowHeights(values, merges, dataStart, dataEnd, colW, lastCol);
  const margin = 82;
  const titleH = 286;
  const headerH = 112;
  const monthRailW = 68;
  const dataW = colW.reduce((sum, item) => sum + item, 0);
  const bodyH = rowHeights.reduce((sum, item) => sum + item, 0);
  const width = dataW + monthRailW + margin * 2;
  const height = margin + titleH + headerH + bodyH + 142;
  const x0 = margin;
  const y0 = margin + titleH;
  const xs = [x0];
  for (let i = 0; i < colW.length - 1; i += 1) xs.push(xs[xs.length - 1] + colW[i]);
  return {
    title,
    dataStart,
    dataEnd,
    rowCount,
    lastCol,
    taskStartCol: 4,
    margin,
    titleH,
    headerH,
    rowHeights,
    bodyH,
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
    for (let col = model.taskStartCol; col <= model.lastCol; col += 1) {
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
        endCol: Math.min(merge.endCol, model.lastCol),
        text: String(value),
        fill: taskColor(col),
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
  return model.xs[model.taskStartCol - 1] || model.xs[model.lastCol - 1];
}

function getTaskRight(model) {
  return model.x0 + model.dataW;
}

function getBodyTop(model) {
  return model.y0 + model.headerH;
}

function getRowHeight(model, row) {
  return model.rowHeights[row - model.dataStart] || MIN_ROW_H;
}

function getRowTop(model, row) {
  let y = getBodyTop(model);
  for (let current = model.dataStart; current < row; current += 1) {
    y += getRowHeight(model, current);
  }
  return y;
}

function getRowsHeight(model, startRow, endRow) {
  let height = 0;
  for (let row = startRow; row <= endRow; row += 1) {
    height += getRowHeight(model, row);
  }
  return height;
}

function getBodyBottom(model) {
  return getBodyTop(model) + model.bodyH;
}

function cardBox(model, card) {
  const x = model.xs[card.col - 1];
  const y = getRowTop(model, card.row);
  const w = model.colW.slice(card.col - 1, card.endCol).reduce((sum, item) => sum + item, 0);
  const h = getRowsHeight(model, card.row, card.endRow);
  return { x, y, w, h };
}

function columnAt(model, x) {
  const clamped = Math.max(getTaskLeft(model), Math.min(getTaskRight(model) - 1, x));
  for (let col = model.taskStartCol; col <= model.lastCol; col += 1) {
    const left = model.xs[col - 1];
    const right = left + model.colW[col - 1];
    if (clamped >= left && clamped < right) return col;
  }
  return model.lastCol;
}

function rowAt(model, y) {
  const clamped = Math.max(getBodyTop(model), Math.min(getBodyBottom(model) - 1, y));
  let top = getBodyTop(model);
  for (let row = model.dataStart; row <= model.dataEnd; row += 1) {
    const bottom = top + getRowHeight(model, row);
    if (clamped >= top && clamped < bottom) return row;
    top = bottom;
  }
  return model.dataEnd;
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
  node.style.fontSize = `${BODY_FONT_SIZE}px`;
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
      card.col = Math.max(model.taskStartCol, Math.min(model.lastCol - widthCols, nextCol));
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
      card.row = Math.min(card.endRow, rowAt(model, getRowTop(model, start.row) + dy));
    }

    if (mode === "bottom") {
      const bottomEdge = getRowTop(model, start.endRow) + getRowHeight(model, start.endRow) + dy;
      card.endRow = Math.max(card.row, rowAt(model, bottomEdge));
    }

    card.fill = taskColor(card.col);
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
    col: state.model.taskStartCol,
    endRow: state.model.dataStart,
    endCol: state.model.taskStartCol,
    text: "新增计划",
    fill: taskColor(state.model.taskStartCol),
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
    height: model.headerH + model.bodyH + 36,
  }));

  const headers = Array.from({ length: model.lastCol }, (_, index) => String(model.values.get(`2:${index + 1}`) || ""));
  headers[0] = "阶段";
  headers.forEach((text, index) => {
    addCell(els.stage, "header-cell", model.xs[index], model.y0, model.colW[index], model.headerH, text, HEADER_FONT_SIZE, COLORS.oxford);
  });
  addCell(els.stage, "header-cell", model.x0 + model.dataW, model.y0, model.monthRailW, model.headerH, "", 24, COLORS.month);

  for (let i = 0; i < model.rowCount; i += 1) {
    const row = model.dataStart + i;
    const y = getRowTop(model, row);
    els.stage.appendChild(el("cell", {
      left: model.x0,
      top: y,
      width: tableW,
      height: getRowHeight(model, row),
      background: i % 2 === 0 ? "#fbfcfb" : "#f1f4f2",
    }));
  }

  mergedRangesForCol(model.values, model.merges, 2, model.dataStart, model.dataEnd).forEach((range) => {
    const y = getRowTop(model, range.row);
    const h = getRowsHeight(model, range.row, range.endRow);
    addCell(els.stage, "grade-cell", model.xs[1], y, model.colW[1], h, range.label, BODY_FONT_SIZE, termColor(range.label));
  });

  mergedRangesForCol(model.values, model.merges, 1, model.dataStart, model.dataEnd).forEach((range) => {
    const y = getRowTop(model, range.row);
    const h = getRowsHeight(model, range.row, range.endRow);
    addCell(els.stage, "stage-cell", model.xs[0], y, model.colW[0], h, range.label, BODY_FONT_SIZE, STAGE_COLORS[range.label] || "#d9e5ea");
  });

  for (let row = model.dataStart; row <= model.dataEnd; row += 1) {
    const y = getRowTop(model, row);
    const h = getRowHeight(model, row);
    addCell(els.stage, "month-cell", model.xs[2], y, model.colW[2], h, String(model.values.get(`${row}:3`) || ""), MONTH_FONT_SIZE, COLORS.month);
    addCell(els.stage, "month-rail", model.x0 + model.dataW, y, model.monthRailW, h, String(model.values.get(`${row}:3`) || ""), 26, COLORS.month);
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
      fontSize: BODY_FONT_SIZE,
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
      node.style.fontSize = `${BODY_FONT_SIZE}px`;
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

  const fy = model.y0 + model.headerH + model.bodyH + 48;
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
  return BODY_FONT_SIZE;
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
  roundRect(ctx, model.x0 - 18, model.y0 - 18, tableW + 36, model.headerH + model.bodyH + 36, 18);
  ctx.fill();
  ctx.strokeStyle = "#e1e6e8";
  ctx.lineWidth = 2;
  ctx.stroke();

  const headers = Array.from({ length: model.lastCol }, (_, index) => String(model.values.get(`2:${index + 1}`) || ""));
  headers[0] = "阶段";
  headers.forEach((text, index) => {
    ctx.fillStyle = COLORS.oxford;
    ctx.fillRect(model.xs[index], model.y0, model.colW[index], model.headerH);
    drawCentered(ctx, model.xs[index] + 8, model.y0, model.colW[index] - 16, model.headerH, text, HEADER_FONT_SIZE, "#ffffff");
  });
  ctx.fillStyle = COLORS.month;
  ctx.fillRect(model.x0 + model.dataW, model.y0, model.monthRailW, model.headerH);

  for (let i = 0; i < model.rowCount; i += 1) {
    const row = model.dataStart + i;
    const y = getRowTop(model, row);
    const h = getRowHeight(model, row);
    ctx.fillStyle = i % 2 === 0 ? "#fbfcfb" : "#f1f4f2";
    ctx.fillRect(model.x0, y, tableW, h);
    ctx.fillStyle = COLORS.month;
    ctx.fillRect(model.x0 + model.dataW, y, model.monthRailW, h);
    ctx.strokeStyle = COLORS.grid;
    ctx.lineWidth = i % 2 === 0 ? 2 : 1;
    ctx.beginPath();
    ctx.moveTo(model.x0, y);
    ctx.lineTo(model.x0 + tableW, y);
    ctx.stroke();
  }

  mergedRangesForCol(model.values, model.merges, 2, model.dataStart, model.dataEnd).forEach((range) => {
    const y = getRowTop(model, range.row);
    const h = getRowsHeight(model, range.row, range.endRow);
    ctx.fillStyle = termColor(range.label);
    ctx.fillRect(model.xs[1], y, model.colW[1], h);
    ctx.strokeStyle = "#c9d3d8";
    ctx.lineWidth = 2;
    ctx.strokeRect(model.xs[1], y, model.colW[1], h);
    drawCentered(ctx, model.xs[1], y, model.colW[1], h, range.label, BODY_FONT_SIZE, COLORS.ink);
  });

  mergedRangesForCol(model.values, model.merges, 1, model.dataStart, model.dataEnd).forEach((range) => {
    const y = getRowTop(model, range.row);
    const h = getRowsHeight(model, range.row, range.endRow);
    ctx.fillStyle = STAGE_COLORS[range.label] || "#d9e5ea";
    ctx.fillRect(model.xs[0], y, model.colW[0], h);
    ctx.strokeStyle = "#c9d3d8";
    ctx.lineWidth = 2;
    ctx.strokeRect(model.xs[0], y, model.colW[0], h);
    drawCentered(ctx, model.xs[0], y, model.colW[0], h, range.label, BODY_FONT_SIZE, COLORS.oxford, true);
  });

  model.cards.forEach((card) => {
    const x = model.xs[card.col - 1];
    const y = getRowTop(model, card.row);
    const w = model.colW.slice(card.col - 1, card.endCol).reduce((sum, item) => sum + item, 0);
    const h = getRowsHeight(model, card.row, card.endRow);
    const pad = 10;
    ctx.fillStyle = card.fill;
    roundRect(ctx, x + pad, y + pad, w - 2 * pad, h - 2 * pad, 12);
    ctx.fill();
    ctx.strokeStyle = "#ffffff";
    ctx.lineWidth = 2;
    ctx.stroke();
    const fontSize = BODY_FONT_SIZE;
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
    const y = getRowTop(model, row);
    const h = getRowHeight(model, row);
    const month = String(model.values.get(`${row}:3`) || "");
    ctx.fillStyle = COLORS.month;
    ctx.fillRect(model.xs[2], y, model.colW[2], h);
    ctx.strokeStyle = "#c9d3d8";
    ctx.lineWidth = 2;
    ctx.strokeRect(model.xs[2], y, model.colW[2], h);
    drawCentered(ctx, model.xs[2], y, model.colW[2], h, month, MONTH_FONT_SIZE, COLORS.oxford);
    ctx.strokeStyle = "#d8e1e5";
    ctx.lineWidth = 1;
    ctx.strokeRect(model.x0 + model.dataW, y, model.monthRailW, h);
    drawCentered(ctx, model.x0 + model.dataW, y, model.monthRailW, h, month, 26, "#8ea1ad");
  }

  const fy = model.y0 + model.headerH + model.bodyH + 48;
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
  return true;
}

function openBlob(blob) {
  const url = URL.createObjectURL(blob);
  const opened = window.open(url, "_blank", "noopener,noreferrer");
  if (!opened) {
    window.location.href = url;
  }
  window.setTimeout(() => URL.revokeObjectURL(url), 60000);
}

async function downloadOrOpenBlob(blob, filename) {
  try {
    downloadBlob(blob, filename);
    return "download";
  } catch (error) {
    console.warn(error);
    openBlob(blob);
    return "open";
  }
}

function makePdfBlob(canvas) {
  if (!window.jspdf?.jsPDF) {
    throw new Error("PDF 组件未加载，请刷新页面后重试。");
  }
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
  return pdf.output("blob");
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
  els.previewPng.disabled = !ready;
  els.previewPdf.disabled = !ready;
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

function setFeishuSheets(sheets, preferredSheetId = "") {
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
  if (preferredSheetId && sheets.some((sheet) => sheet.sheetId === preferredSheetId)) {
    els.feishuSheet.value = preferredSheetId;
  }
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
    setFeishuSheets(data.sheets || [], data.preferredSheetId || "");
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
    const data = await feishuRequest({ action: "values", token, sheetId, range: "A1:Z80" });
    state.filename = selected?.dataset.title || selected?.textContent || "飞书在线表格";
    state.selectedCardId = null;
    state.nextCardId = 1;
    state.model = parseOnlineValues(data.values || [], state.filename, state.filename, data.merges || []);
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
    await downloadOrOpenBlob(blob, safeFilename(state.model.title, "png"));
  });
});

els.downloadPdf.addEventListener("click", () => {
  withExportState("正在生成 PDF...", async () => {
    const canvas = renderToCanvas(Number(els.qualityInput.value));
    const blob = makePdfBlob(canvas);
    await downloadOrOpenBlob(blob, safeFilename(state.model.title, "pdf"));
  });
});

els.previewPng.addEventListener("click", () => {
  withExportState("正在打开 PNG...", async () => {
    const canvas = renderToCanvas(Number(els.qualityInput.value));
    const blob = await canvasToBlob(canvas, "image/png");
    openBlob(blob);
  });
});

els.previewPdf.addEventListener("click", () => {
  withExportState("正在打开 PDF...", async () => {
    const canvas = renderToCanvas(Number(els.qualityInput.value));
    openBlob(makePdfBlob(canvas));
  });
});
