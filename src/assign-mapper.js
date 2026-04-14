"use strict";
// ─── Assignment Attainment Mapper ─────────────────────────────────────────
// Equivalent to co_assign_attainment_mapper.py

const fs = require("fs");
const path = require("path");
const ExcelJS = require("exceljs");
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

const ASSIGN_PATTERNS = [
  /\bASSIGN\s*\(?(\d+)\)?/i,
  /\bASSIGNMENT\s*\(?(\d+)\)?/i,
  /\bASS\s*\(?(\d+)\)/i,
];

/**
 * Extract assignment attainment from a marks file.
 * Returns:
 *   { assignments: { 1: { attainment, cos }, 2: {...}, ... } }
 */
async function extractAssignAttainment(marksFile) {
  const wb = await loadWorkbook(marksFile);
  const ws = wb.worksheets[0];

  const { headerRowIdx, srCol } = findHeaderRow(ws);

  // Collect ASSIGN columns from the header row
  const headerRow = ws.getRow(headerRowIdx);
  const assignCols = {}; // { assignNum: colIndex }

  headerRow.eachCell({ includeEmpty: false }, (cell, colNum) => {
    const val = String(getCellValue(cell) || "").trim();
    for (const pat of ASSIGN_PATTERNS) {
      const m = val.match(pat);
      if (m) {
        assignCols[parseInt(m[1])] = colNum;
        break;
      }
    }
  });

  if (!Object.keys(assignCols).length) {
    const headers = [];
    headerRow.eachCell({ includeEmpty: false }, (c) =>
      headers.push(getCellValue(c)),
    );
    throw new Error(
      `Could not identify ASSIGN columns.\nHeaders found: ${headers.join(", ")}\n` +
        `Expected headers like 'Assign1', 'Assign2', etc.`,
    );
  }

  // Scan non-student rows for explicit CO mapping text
  const coMapPerAssign = {};
  const maxRow = ws.rowCount;

  for (let r = headerRowIdx + 1; r <= maxRow; r++) {
    if (isStudentRow(ws, r, srCol)) continue;
    for (const [assignNum, colIdx] of Object.entries(assignCols)) {
      const raw = getCellValue(ws.getRow(r).getCell(colIdx));
      const text = String(raw || "");
      const cos = [...text.matchAll(/CO\s*(\d+)/gi)].map((m) => parseInt(m[1]));
      if (cos.length) coMapPerAssign[parseInt(assignNum)] = cos;
    }
  }

  // Extract attainment per assignment
  const assignments = {};

  for (const [assignNumStr, colIdx] of Object.entries(assignCols)) {
    const assignNum = parseInt(assignNumStr);

    // Try to find an explicit attainment row
    let attainment = null;
    for (let r = headerRowIdx + 1; r <= maxRow; r++) {
      if (isStudentRow(ws, r, srCol)) continue;
      const v = parseAttainmentCell(getCellValue(ws.getRow(r).getCell(colIdx)));
      if (v !== null) {
        attainment = v;
        break;
      }
    }

    // Compute from raw marks if no explicit row
    if (attainment === null) {
      const marks = [];
      for (let r = headerRowIdx + 1; r <= maxRow; r++) {
        if (!isStudentRow(ws, r, srCol)) continue;
        const v = parseFloat(
          String(getCellValue(ws.getRow(r).getCell(colIdx)) ?? ""),
        );
        if (!isNaN(v)) marks.push(v);
      }
      if (marks.length) {
        const stats = computeAttainmentFromMarks(marks);
        attainment = stats.attainment;
      }
    }

    assignments[assignNum] = {
      attainment: attainment ?? 1,
      cos: coMapPerAssign[assignNum] || [],
    };
  }

  return { assignments };
}

/**
 * Write assignment attainment values into outputFile (already a working copy of the CO template).
 */
async function writeAssignAttainment(outputFile, assignData) {
  const wb = await loadWorkbook(outputFile);
  const ws = wb.worksheets[0];

  const coCols = findCoColumns(ws);
  if (!Object.keys(coCols).length)
    throw new Error("Cannot find CO columns in template.");

  const allCoNums = Object.keys(coCols)
    .map(Number)
    .sort((a, b) => a - b);

  // Find the "Theory Assignment" / "Assign" row
  let assignRow = findRowByPattern(ws, /\bAssign(ment)?\b/);
  if (!assignRow) assignRow = findRowByPattern(ws, /\bTheory\s+Assignment\b/);
  if (!assignRow)
    throw new Error("Cannot find 'Assignment' row in CO template.");

  const allMappedCos = new Set();

  for (const [numStr, info] of Object.entries(assignData.assignments)) {
    const { attainment, cos } = info;
    if (!cos.length) continue;
    for (const coNum of cos) {
      if (coCols[coNum]) {
        ws.getRow(assignRow).getCell(coCols[coNum]).value = attainment;
        allMappedCos.add(coNum);
      }
    }
  }

  // Clear unmapped CO cells in the assignment row
  for (const coNum of allCoNums) {
    if (!allMappedCos.has(coNum) && coCols[coNum]) {
      ws.getRow(assignRow).getCell(coCols[coNum]).value = null;
    }
  }

  await saveWorkbook(wb, outputFile);
}

module.exports = { extractAssignAttainment, writeAssignAttainment };
