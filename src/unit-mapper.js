"use strict";
// ─── Unit Test Attainment Mapper ──────────────────────────────────────────
// Equivalent to co_unit_test_attainment_mapper.py

const {
  loadWorkbook,
  saveWorkbook,
  getCellValue,
  findCoColumns,
  findRowByPattern,
  findHeaderRow,
  parseAttainmentCell,
  isStudentRow,
  computeAttainmentFromMarks,
} = require("./excel-utils");

const UT_PATTERNS = [
  /\bUT\s*(\d+)\b/i,
  /\bUNIT\s*TEST\s*(\d+)\b/i,
  /\bUNIT\s*(\d+)\b/i,
];

/**
 * Extract unit-test attainment from a marks file.
 * Returns:
 *   { unitTests: { 1: { attainment, cos }, 2: {...} } }
 */
async function extractUnitTestAttainment(marksFile) {
  const wb = await loadWorkbook(marksFile);
  const ws = wb.worksheets[0];

  const { headerRowIdx, srCol } = findHeaderRow(ws);
  const headerRow = ws.getRow(headerRowIdx);

  const utCols = {}; // { utNum: colIndex }
  const utCosFromHeader = {}; // COs embedded in header text

  headerRow.eachCell({ includeEmpty: false }, (cell, colNum) => {
    const val = String(getCellValue(cell) || "").trim();
    for (const pat of UT_PATTERNS) {
      const m = val.match(pat);
      if (m) {
        const utNum = parseInt(m[1]);
        utCols[utNum] = colNum;
        // CO info embedded in header, e.g. "UT1 (CO1, CO2)"
        const headerCos = [...val.matchAll(/CO\s*(\d+)/gi)].map((x) =>
          parseInt(x[1]),
        );
        if (headerCos.length) utCosFromHeader[utNum] = headerCos;
        break;
      }
    }
  });

  if (!Object.keys(utCols).length) {
    const headers = [];
    headerRow.eachCell({ includeEmpty: false }, (c) =>
      headers.push(getCellValue(c)),
    );
    throw new Error(
      `Could not identify Unit Test columns.\nHeaders: ${headers.join(", ")}\n` +
        `Expected 'UT1', 'UT2', 'Unit Test 1', etc.`,
    );
  }

  // Start with header-embedded COs, then let dedicated mapping rows override
  const coMapPerUt = { ...utCosFromHeader };
  const maxRow = ws.rowCount;

  for (let r = headerRowIdx + 1; r <= maxRow; r++) {
    if (isStudentRow(ws, r, srCol)) continue;
    for (const [utNumStr, colIdx] of Object.entries(utCols)) {
      const raw = getCellValue(ws.getRow(r).getCell(colIdx));
      const text = String(raw || "");
      const cos = [...text.matchAll(/CO\s*(\d+)/gi)].map((m) => parseInt(m[1]));
      if (cos.length) coMapPerUt[parseInt(utNumStr)] = cos;
    }
  }

  const unitTests = {};

  for (const [utNumStr, colIdx] of Object.entries(utCols)) {
    const utNum = parseInt(utNumStr);

    let attainment = null;
    for (let r = headerRowIdx + 1; r <= maxRow; r++) {
      if (isStudentRow(ws, r, srCol)) continue;
      const v = parseAttainmentCell(getCellValue(ws.getRow(r).getCell(colIdx)));
      if (v !== null) {
        attainment = v;
        break;
      }
    }

    if (attainment === null) {
      const marks = [];
      for (let r = headerRowIdx + 1; r <= maxRow; r++) {
        if (!isStudentRow(ws, r, srCol)) continue;
        const v = parseFloat(
          String(getCellValue(ws.getRow(r).getCell(colIdx)) ?? ""),
        );
        if (!isNaN(v)) marks.push(v);
      }
      if (marks.length)
        attainment = computeAttainmentFromMarks(marks).attainment;
    }

    unitTests[utNum] = {
      attainment: attainment ?? 1,
      cos: coMapPerUt[utNum] || [],
    };
  }

  return { unitTests };
}

/**
 * Write unit-test attainment into the output CO workbook (in-place).
 */
async function writeUnitTestAttainment(outputFile, unitData) {
  const wb = await loadWorkbook(outputFile);
  const ws = wb.worksheets[0];

  const coCols = findCoColumns(ws);
  if (!Object.keys(coCols).length)
    throw new Error("Cannot find CO columns in template.");
  const allCoNums = Object.keys(coCols)
    .map(Number)
    .sort((a, b) => a - b);

  for (const [utNumStr, info] of Object.entries(unitData.unitTests)) {
    const utNum = parseInt(utNumStr);
    const { attainment, cos } = info;
    if (!cos.length) continue;

    // Locate "Unit Test N" row
    let utRow = findRowByPattern(ws, `Unit\\s+Test\\s+${utNum}`);
    if (!utRow) utRow = findRowByPattern(ws, `UT${utNum}`);
    if (!utRow) continue;

    const mappedCos = new Set(cos);

    for (const coNum of allCoNums) {
      if (!coCols[coNum]) continue;
      const cell = ws.getRow(utRow).getCell(coCols[coNum]);
      if (mappedCos.has(coNum)) {
        cell.value = attainment;
      } else {
        cell.value = null;
      }
    }
  }

  await saveWorkbook(wb, outputFile);
}

module.exports = { extractUnitTestAttainment, writeUnitTestAttainment };
