"use strict";
// ─── Shared Excel Utilities ───────────────────────────────────────────────

const ExcelJS = require("exceljs");

/**
 * Load a workbook from disk (formula results cached in cell.value.result).
 */
async function loadWorkbook(filePath) {
  const wb = new ExcelJS.Workbook();
  await wb.xlsx.readFile(filePath);
  return wb;
}

async function saveWorkbook(wb, filePath) {
  await wb.xlsx.writeFile(filePath);
}

/** Extract a plain scalar from a cell value (handles formula objects & rich text). */
function getCellValue(cell) {
  if (!cell) return null;
  const v = cell.value;
  if (v === null || v === undefined) return null;
  if (typeof v === "object") {
    if ("result" in v) return v.result; // formula cell with cached result
    if ("text" in v) return v.text; // rich-text cell
  }
  return v;
}

/**
 * Scan all rows for the first row containing ≥ 2 CO-qualified columns.
 * Recognises:
 *   "317532B.1"  → CO 1  (trailing .N notation)
 *   "CO1", "CO 2" → CO N (explicit label)
 * Returns { coNum: 1-based-colIndex }
 */
function findCoColumns(ws) {
  let result = null;
  ws.eachRow({ includeEmpty: false }, (row) => {
    if (result) return;
    const coCols = {};
    row.eachCell({ includeEmpty: false }, (cell, colNum) => {
      const val = String(getCellValue(cell) || "").trim();
      let m = val.match(/\.(\d+)$/);
      if (m) {
        const n = parseInt(m[1]);
        if (n >= 1 && n <= 10) {
          coCols[n] = colNum;
          return;
        }
      }
      m = val.match(/^CO\s*(\d+)$/i);
      if (m) coCols[parseInt(m[1])] = colNum;
    });
    if (Object.keys(coCols).length >= 2) result = coCols;
  });
  return result || {};
}

/**
 * Find the first row whose any cell text matches the given regex (case-insensitive).
 * Returns the 1-based row index or null.
 */
function findRowByPattern(ws, pattern) {
  const re = new RegExp(pattern, "i");
  let found = null;
  ws.eachRow({ includeEmpty: false }, (row, rowNum) => {
    if (found) return;
    row.eachCell({ includeEmpty: false }, (cell) => {
      if (found) return;
      if (re.test(String(getCellValue(cell) || "").trim())) found = rowNum;
    });
  });
  return found;
}

/**
 * Locate the header row that contains an SR.NO / S.NO column.
 * Returns { headerRowIdx, srCol } or throws.
 */
function findHeaderRow(ws) {
  let res = null;
  ws.eachRow({ includeEmpty: false }, (row, rowNum) => {
    if (res) return;
    row.eachCell({ includeEmpty: false }, (cell, colNum) => {
      if (res) return;
      const clean = String(getCellValue(cell) || "")
        .replace(/[\s.]/g, "")
        .toUpperCase();
      if (["SRNO", "SNO", "SLNO", "SERIALNO", "NO"].includes(clean)) {
        res = { headerRowIdx: rowNum, srCol: colNum };
      }
    });
  });
  if (!res)
    throw new Error(
      "Cannot find header row with 'SR.NO' or 'S.NO' in the marks file.",
    );
  return res;
}

/**
 * Safely parse a cell value as an integer attainment level (1–3).
 * Rejects mixed strings like "CO1,CO2".
 */
function parseAttainmentCell(v) {
  if (v === null || v === undefined) return null;
  const s = String(v).trim();
  if (!s) return null;
  // reject if it contains non-numeric characters other than decimal point/sign
  if (/[a-zA-Z]/.test(s)) return null;
  try {
    const n = parseInt(parseFloat(s));
    if (n >= 1 && n <= 3) return n;
  } catch (_) {}
  return null;
}

/**
 * Map percentage of students above average → SPPU attainment level (0–3).
 */
function pctToAttainment(pct) {
  if (pct >= 60) return 3;
  if (pct >= 50) return 2;
  if (pct >= 40) return 1;
  return 0;
}

/**
 * Check whether a row is a student data row (SR.NO is a positive integer).
 */
function isStudentRow(ws, rowNum, srCol) {
  const raw = getCellValue(ws.getRow(rowNum).getCell(srCol));
  try {
    return parseInt(parseFloat(String(raw || ""))) >= 1;
  } catch (_) {
    return false;
  }
}

/**
 * Compute attainment from raw marks: average → count above avg → pct → level.
 */
function computeAttainmentFromMarks(marks) {
  if (!marks.length) throw new Error("No valid student marks found.");
  const avg = marks.reduce((a, b) => a + b, 0) / marks.length;
  const above = marks.filter((x) => x > avg).length;
  const pct = (above / marks.length) * 100;
  return {
    attainment: pctToAttainment(pct),
    avg,
    above,
    pct,
    total: marks.length,
  };
}

module.exports = {
  loadWorkbook,
  saveWorkbook,
  getCellValue,
  findCoColumns,
  findRowByPattern,
  findHeaderRow,
  parseAttainmentCell,
  pctToAttainment,
  isStudentRow,
  computeAttainmentFromMarks,
};
