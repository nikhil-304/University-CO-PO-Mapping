"use strict";
// ─── CES Attainment Mapper ────────────────────────────────────────────────
// Equivalent to co_ces_attainment_mapper.py

const {
  loadWorkbook,
  saveWorkbook,
  getCellValue,
  findCoColumns,
  findRowByPattern,
} = require("./excel-utils");

function parseLevelCell(v) {
  if (v === null || v === undefined) return null;
  try {
    const n = parseInt(parseFloat(String(v).trim()));
    if (n >= 1 && n <= 3) return n;
  } catch (_) {}
  return null;
}

/**
 * Extract CES attainment from a Course Exit Survey file.
 * Returns: { coLevels: { 1: 3, 2: 2, ... } }
 */
async function extractCesAttainment(cesFile) {
  const wb = await loadWorkbook(cesFile);
  const ws = wb.worksheets[0];

  // Find the summary header row: contains "Level" AND ("Excellent"/"Average"/"Good")
  let headerRowIdx = null;
  let levelCol = null;

  ws.eachRow({ includeEmpty: false }, (row, rowNum) => {
    if (headerRowIdx) return;
    let foundLevelCol = null;
    row.eachCell({ includeEmpty: false }, (cell, colNum) => {
      if (/^Level$/i.test(String(getCellValue(cell) || "").trim())) {
        foundLevelCol = colNum;
      }
    });
    if (foundLevelCol !== null) {
      // Verify the row also contains a quality descriptor
      const rowText = [];
      row.eachCell({ includeEmpty: false }, (c) =>
        rowText.push(String(getCellValue(c) || "").toUpperCase()),
      );
      const joined = rowText.join(" ");
      if (
        joined.includes("EXCELLENT") ||
        joined.includes("AVERAGE") ||
        joined.includes("GOOD")
      ) {
        headerRowIdx = rowNum;
        levelCol = foundLevelCol;
      }
    }
  });

  if (!headerRowIdx || !levelCol) {
    throw new Error(
      "Cannot find CES summary header row.\n" +
        "Expected a row containing 'Level' alongside 'Excellent', 'Average', or 'Good'.",
    );
  }

  const coLevels = {};
  let coNum = 1;
  const maxRow = ws.rowCount;

  for (let r = headerRowIdx + 1; r <= maxRow; r++) {
    const row = ws.getRow(r);
    // Skip fully empty rows
    let hasAny = false;
    row.eachCell({ includeEmpty: false }, () => {
      hasAny = true;
    });
    if (!hasAny) continue;

    const v = parseLevelCell(getCellValue(row.getCell(levelCol)));
    if (v !== null) {
      coLevels[coNum] = v;
      coNum++;
    }
  }

  if (!Object.keys(coLevels).length) {
    throw new Error("No attainment levels found in CES summary table.");
  }

  return { coLevels };
}

/**
 * Write CES attainment levels into the output CO workbook (in-place).
 */
async function writeCesAttainment(outputFile, cesData) {
  const wb = await loadWorkbook(outputFile);
  const ws = wb.worksheets[0];

  const coCols = findCoColumns(ws);
  if (!Object.keys(coCols).length)
    throw new Error("Cannot find CO columns in template.");

  // Locate the "Course Exit Survey" / CES row
  let cesRow = findRowByPattern(ws, /Course\s+Exit\s+Survey/);
  if (!cesRow) cesRow = findRowByPattern(ws, /\bCES\b/);
  if (!cesRow)
    throw new Error("Cannot find 'Course Exit Survey' row in template.");

  for (const [coNumStr, level] of Object.entries(cesData.coLevels)) {
    const coNum = parseInt(coNumStr);
    if (coCols[coNum]) {
      ws.getRow(cesRow).getCell(coCols[coNum]).value = level;
    }
  }

  await saveWorkbook(wb, outputFile);
}

module.exports = { extractCesAttainment, writeCesAttainment };
