"use strict";
// ─── CO → PO Attainment Mapper ────────────────────────────────────────────
// Equivalent to co_to_po_attainment_mapper.py
//
// Formula chain replicated:
//   Row 13 : avg_direct  = AVERAGE(non-blank from rows 10, 11, 12) per CO
//   Row 14 : A           = avg_direct × 0.3
//   Row 18 : B           = row17 × 0.7
//   Row 19 : D           = (A + B) × 0.7  (but python does (row14 + row18) * 0.7?
//                          actually D = (A + B) where A=row14 already the 0.3 weighted)
//              Wait — let me re-read: Row14=A=avg_direct*0.3, Row18=B=row17*0.7,
//              Row19=D=(A+B)*0.7 → D = (avg_direct*0.3 + endsem*0.7)*0.7
//   Row 22 : I           = row21 × 0.3
//   Row 23 : CO Attain.  = D + I

const fs = require("fs");
const {
  loadWorkbook,
  saveWorkbook,
  getCellValue,
  findCoColumns,
  findRowByPattern,
} = require("./excel-utils");

function toNum(v) {
  if (v === null || v === undefined) return null;
  const n = parseFloat(String(v));
  return isNaN(n) ? null : n;
}

/**
 * Read CO_Attainment_Final.xlsx, compute D+I per CO (replicating formula chain).
 * Returns: { coLevels: { 1: 2.405, 2: 2.51, ... } }
 */
async function extractCoAttainment(coCompleteFile) {
  const wb = await loadWorkbook(coCompleteFile);
  const ws = wb.worksheets[0];

  const coCols = findCoColumns(ws);
  if (!Object.keys(coCols).length)
    throw new Error("Cannot find CO columns in CO_Attainment_Final.xlsx");

  // Locate required rows
  const rowTheory = findRowByPattern(ws, /Theory\s+Assignment/);
  const rowUt1 = findRowByPattern(ws, /Unit\s+Test\s+1/);
  const rowUt2 = findRowByPattern(ws, /Unit\s+Test\s+2/);
  const rowEndsem = findRowByPattern(ws, /End\s+Sem\s+Result/);
  const rowCes = findRowByPattern(ws, /Course\s+Exit\s+Survey/);

  if (!rowEndsem)
    throw new Error("Cannot find 'End Sem Result' row in CO file.");
  if (!rowCes)
    throw new Error("Cannot find 'Course Exit Survey' row in CO file.");

  const getVal = (row, col) =>
    row ? toNum(getCellValue(ws.getRow(row).getCell(col))) : null;

  const coLevels = {};

  for (const [coNumStr, colIdx] of Object.entries(coCols)) {
    const coNum = parseInt(coNumStr);

    // avg_direct = average of non-null values from theory, ut1, ut2 rows
    const directVals = [];
    for (const r of [rowTheory, rowUt1, rowUt2]) {
      const v = getVal(r, colIdx);
      if (v !== null) directVals.push(v);
    }
    const avgDirect = directVals.length
      ? directVals.reduce((a, b) => a + b, 0) / directVals.length
      : 0;

    const A = avgDirect * 0.3; // Row 14
    const endsem = getVal(rowEndsem, colIdx) ?? 0;
    const B = endsem * 0.7; // Row 18
    const D = (A + B) * 0.7; // Row 19  ← matches python
    const ces = getVal(rowCes, colIdx) ?? 0;
    const I = ces * 0.3; // Row 22
    const att = D + I; // Row 23

    coLevels[coNum] = Math.round(att * 1e10) / 1e10;
  }

  if (!Object.keys(coLevels).length)
    throw new Error("No CO attainment values computed.");

  return { coLevels };
}

/**
 * Copy PO static template → poOutput, then write CO attainment values into it.
 */
async function writePoAttainment(poStatic, poOutput, coData) {
  fs.copyFileSync(poStatic, poOutput);

  const wb = await loadWorkbook(poOutput);
  const ws = wb.worksheets[0];

  // Find header row containing "CO Attainment"
  let headerRow = null;
  let coAttainCol = null;

  ws.eachRow({ includeEmpty: false }, (row, rowNum) => {
    if (headerRow) return;
    if (rowNum <= 3) return; // skip title rows
    row.eachCell({ includeEmpty: false }, (cell, colNum) => {
      if (headerRow) return;
      if (/CO\s*Attainment/i.test(String(getCellValue(cell) || ""))) {
        headerRow = rowNum;
        coAttainCol = colNum;
      }
    });
  });

  if (!headerRow || !coAttainCol) {
    throw new Error(
      "Cannot find 'CO Attainment' column in PO_Attainment_Static.xlsx",
    );
  }

  // CO1 is one row below header, CO2 two rows below, etc.
  for (const [coNumStr, value] of Object.entries(coData.coLevels)) {
    const coNum = parseInt(coNumStr);
    const targetRow = headerRow + coNum;
    ws.getRow(targetRow).getCell(coAttainCol).value = value;
  }

  await saveWorkbook(wb, poOutput);
}

module.exports = { extractCoAttainment, writePoAttainment };
