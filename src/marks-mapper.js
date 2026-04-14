"use strict";
// ─── INSEM / ENDSEM Marks Attainment Mapper ──────────────────────────────
// Equivalent to co_attainment_mapper.py

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

const INSEM_PATTERNS = [
  /\bIN\s*SEM\b/i,
  /\bINSEM\b/i,
  /IN\s*\(UNIT\s*1/i,
  /IN\s*\(/i,
  /\bINTERNAL\s+(EXAM|SEM|MARKS|TEST|ASSESSMENT)\b/i,
  /^IN\b/i,
];
const ENDSEM_PATTERNS = [
  /\bEND\s*SEM\b/i,
  /\bENDSEM\b/i,
  /END\s*\(UNIT/i,
  /END\s*\(/i,
  /\bEXTERNAL\s+(EXAM|SEM|MARKS|ASSESSMENT)\b/i,
  /^END\b/i,
  /\bEND\s+RESULT\b/i,
];

function matchesAny(text, patterns) {
  return patterns.some((p) => p.test(text));
}

/**
 * Extract INSEM + ENDSEM attainment from a marks file.
 * Returns:
 *   { insemAttainment, endsemAttainment, insemCos, endsemCos }
 */
async function extractMarksAttainment(marksFile) {
  const wb = await loadWorkbook(marksFile);
  const ws = wb.worksheets[0];

  const { headerRowIdx, srCol } = findHeaderRow(ws);
  const headerRow = ws.getRow(headerRowIdx);

  let inCol = null,
    endCol = null;
  headerRow.eachCell({ includeEmpty: false }, (cell, colNum) => {
    const val = String(getCellValue(cell) || "").trim();
    if (inCol === null && matchesAny(val, INSEM_PATTERNS)) inCol = colNum;
    if (endCol === null && matchesAny(val, ENDSEM_PATTERNS)) endCol = colNum;
  });

  if (inCol === null || endCol === null) {
    const headers = [];
    headerRow.eachCell({ includeEmpty: false }, (c) =>
      headers.push(getCellValue(c)),
    );
    throw new Error(
      `Could not identify IN/END columns.\nHeaders: ${headers.join(", ")}\n` +
        `Need a column with 'INSEM'/'IN SEM' and one with 'ENDSEM'/'END SEM'.`,
    );
  }

  const maxRow = ws.rowCount;
  let attainmentRow = null,
    coMapRow = null;

  for (let r = headerRowIdx + 1; r <= maxRow; r++) {
    if (isStudentRow(ws, r, srCol)) continue;

    // CO mapping row — label-based
    if (coMapRow === null) {
      const row = ws.getRow(r);
      let hasCOmap = false;
      row.eachCell({ includeEmpty: false }, (cell) => {
        const lbl = String(getCellValue(cell) || "").toUpperCase();
        if (
          lbl.includes("CO") &&
          (lbl.includes("MAPPING") || lbl.includes("MAP"))
        )
          hasCOmap = true;
      });
      if (hasCOmap) coMapRow = r;
    }

    // CO mapping row — value-based fallback
    if (coMapRow === null) {
      const inV = String(getCellValue(ws.getRow(r).getCell(inCol)) || "");
      const endV = String(getCellValue(ws.getRow(r).getCell(endCol)) || "");
      if (/CO\s*\d+/i.test(inV) && /CO\s*\d+/i.test(endV)) coMapRow = r;
    }

    // Attainment row — both IN and END have level 1-3
    if (attainmentRow === null) {
      const inAtt = parseAttainmentCell(
        getCellValue(ws.getRow(r).getCell(inCol)),
      );
      const endAtt = parseAttainmentCell(
        getCellValue(ws.getRow(r).getCell(endCol)),
      );
      if (inAtt !== null && endAtt !== null) attainmentRow = r;
    }
  }

  // Read explicit attainment values or compute from raw marks
  let insemAttainment = attainmentRow
    ? parseAttainmentCell(getCellValue(ws.getRow(attainmentRow).getCell(inCol)))
    : null;
  let endsemAttainment = attainmentRow
    ? parseAttainmentCell(
        getCellValue(ws.getRow(attainmentRow).getCell(endCol)),
      )
    : null;

  if (insemAttainment === null || endsemAttainment === null) {
    const insemMarks = [],
      endsemMarks = [];
    for (let r = headerRowIdx + 1; r <= maxRow; r++) {
      if (!isStudentRow(ws, r, srCol)) continue;
      const iV = parseFloat(
        String(getCellValue(ws.getRow(r).getCell(inCol)) ?? ""),
      );
      const eV = parseFloat(
        String(getCellValue(ws.getRow(r).getCell(endCol)) ?? ""),
      );
      if (!isNaN(iV)) insemMarks.push(iV);
      if (!isNaN(eV)) endsemMarks.push(eV);
    }
    if (insemAttainment === null && insemMarks.length)
      insemAttainment = computeAttainmentFromMarks(insemMarks).attainment;
    if (endsemAttainment === null && endsemMarks.length)
      endsemAttainment = computeAttainmentFromMarks(endsemMarks).attainment;
  }

  // Parse CO mappings
  let insemCos = [];
  let endsemCos = [];
  if (coMapRow) {
    const inText = String(
      getCellValue(ws.getRow(coMapRow).getCell(inCol)) || "",
    );
    const endText = String(
      getCellValue(ws.getRow(coMapRow).getCell(endCol)) || "",
    );
    insemCos = [...inText.matchAll(/CO\s*(\d+)/gi)].map((m) => parseInt(m[1]));
    endsemCos = [...endText.matchAll(/CO\s*(\d+)/gi)].map((m) =>
      parseInt(m[1]),
    );
  }

  return {
    insemAttainment: insemAttainment ?? 1,
    endsemAttainment: endsemAttainment ?? 1,
    insemCos,
    endsemCos,
  };
}

/**
 * Write INSEM + ENDSEM attainment into the output CO workbook (in-place).
 */
async function writeMarksAttainment(outputFile, marksData) {
  const wb = await loadWorkbook(outputFile);
  const ws = wb.worksheets[0];

  const coCols = findCoColumns(ws);
  if (!Object.keys(coCols).length)
    throw new Error("Cannot find CO columns in template.");
  const allCoNums = Object.keys(coCols)
    .map(Number)
    .sort((a, b) => a - b);

  let { insemAttainment, endsemAttainment, insemCos, endsemCos } = marksData;

  // Fall back to default CO split if no CO mapping found
  if (!insemCos.length && !endsemCos.length) {
    const split = Math.max(Math.floor(allCoNums.length / 3), 1);
    insemCos = allCoNums.slice(0, split);
    endsemCos = allCoNums.slice(split);
  }

  const insemRow =
    findRowByPattern(ws, /\bIn\s*Sem\s*Exam\b/) ||
    findRowByPattern(ws, /\bIn\s*Sem\b/);
  const endsemRow =
    findRowByPattern(ws, /\bEnd\s*Sem\s*Result\b/) ||
    findRowByPattern(ws, /\bEnd\s*Sem\b/);

  if (!insemRow || !endsemRow)
    throw new Error(
      "Cannot find 'In Sem Exam' or 'End Sem Result' rows in template.",
    );

  const insemSet = new Set(insemCos);

  for (const coNum of allCoNums) {
    if (!coCols[coNum]) continue;
    // INSEM row: only write mapped COs, clear the rest
    ws.getRow(insemRow).getCell(coCols[coNum]).value = insemSet.has(coNum)
      ? insemAttainment
      : null;
    // ENDSEM row: write all COs
    ws.getRow(endsemRow).getCell(coCols[coNum]).value = endsemAttainment;
  }

  await saveWorkbook(wb, outputFile);
}

module.exports = { extractMarksAttainment, writeMarksAttainment };
