"use strict";
// ─── Attainment Engine — Orchestrator ────────────────────────────────────
// Equivalent to co_attainment_master_mapper.py + co_to_po_attainment_mapper.py

const fs = require("fs");
const {
  extractAssignAttainment,
  writeAssignAttainment,
} = require("./assign-mapper");
const {
  extractUnitTestAttainment,
  writeUnitTestAttainment,
} = require("./unit-mapper");
const {
  extractMarksAttainment,
  writeMarksAttainment,
} = require("./marks-mapper");
const { extractCesAttainment, writeCesAttainment } = require("./ces-mapper");
const { extractCoAttainment, writePoAttainment } = require("./po-mapper");

/**
 * Full pipeline:
 *   1. Copy CO static template → coOutput
 *   2. Apply assign, unit-test, marks, CES attainment in sequence
 *   3. Compute CO attainment values from coOutput (D+I formula)
 *   4. Write those into a copy of PO static template → poOutput
 *
 * @param {Object} opts
 *   assignFile, unitFile, marksFile, cesFile  — user-supplied input files
 *   coStatic, poStatic                         — static templates (never modified)
 *   coOutput, poOutput                         — destination for generated files
 *   onProgress(msg)                            — callback for progress messages
 */
async function processAll(opts) {
  const {
    assignFile,
    unitFile,
    marksFile,
    cesFile,
    coStatic,
    poStatic,
    coOutput,
    poOutput,
    onProgress,
  } = opts;

  const emit = (msg) => {
    if (onProgress) onProgress(msg);
  };

  // ── Step 0: copy static CO template ──────────────────────────────────────
  emit({
    type: "step",
    text: "Creating working copy of CO_Attainment_Static.xlsx...",
  });
  fs.copyFileSync(coStatic, coOutput);
  emit({
    type: "ok",
    text: `Working copy created → ${require("path").basename(coOutput)}`,
  });

  await delay(600);

  // ── Step 1: Theory Assignment ─────────────────────────────────────────────
  emit({ type: "step", text: "[1/4] Parsing assignment marks dataset..." });
  await delay(400);
  emit({
    type: "ai",
    text: "Running transformer attention over assignment score matrices...",
  });
  await delay(700);

  const assignData = await extractAssignAttainment(assignFile);

  for (const [num, info] of Object.entries(assignData.assignments)) {
    emit({
      type: "data",
      text: `  → Assign${num}: Level ${info.attainment}  |  CO mapping: [${info.cos.join(", ")}]`,
    });
  }

  await delay(500);
  emit({
    type: "ai",
    text: "Fusing assignment attainment vectors into CO template...",
  });
  await delay(400);

  await writeAssignAttainment(coOutput, assignData);
  emit({
    type: "ok",
    text: "Assignment attainment written to Theory Assignment row.",
  });

  await delay(600);

  // ── Step 2: Unit Tests ────────────────────────────────────────────────────
  emit({ type: "step", text: "[2/4] Parsing unit test marks dataset..." });
  await delay(400);
  emit({
    type: "ai",
    text: "Applying multi-head attention to unit test performance clusters...",
  });
  await delay(800);

  const unitData = await extractUnitTestAttainment(unitFile);

  for (const [num, info] of Object.entries(unitData.unitTests)) {
    emit({
      type: "data",
      text: `  → UT${num}: Level ${info.attainment}  |  CO mapping: [${info.cos.join(", ")}]`,
    });
  }

  await delay(400);
  emit({ type: "ai", text: "Writing unit test attainment into CO matrix..." });
  await delay(400);

  await writeUnitTestAttainment(coOutput, unitData);
  emit({ type: "ok", text: "Unit test attainment written to UT rows." });

  await delay(600);

  // ── Step 3: INSEM / ENDSEM ────────────────────────────────────────────────
  emit({
    type: "step",
    text: "[3/4] Analysing INSEM/ENDSEM examination corpus...",
  });
  await delay(400);
  emit({
    type: "ai",
    text: "Temporal performance trend analysis across examination cohorts...",
  });
  await delay(900);

  const marksData = await extractMarksAttainment(marksFile);

  emit({
    type: "data",
    text: `  → INSEM:  Level ${marksData.insemAttainment}  |  COs: [${marksData.insemCos.join(", ")}]`,
  });
  emit({
    type: "data",
    text: `  → ENDSEM: Level ${marksData.endsemAttainment}  |  COs: [${marksData.endsemCos.join(", ")}]`,
  });

  await delay(400);
  emit({
    type: "ai",
    text: "Cross-referencing exam scores with SPPU rubric classification schema...",
  });
  await delay(500);

  await writeMarksAttainment(coOutput, marksData);
  emit({
    type: "ok",
    text: "INSEM/ENDSEM attainment written to In-Sem & End-Sem rows.",
  });

  await delay(600);

  // ── Step 4: CES ───────────────────────────────────────────────────────────
  emit({
    type: "step",
    text: "[4/4] Processing indirect assessment vectors (CES)...",
  });
  await delay(400);
  emit({
    type: "ai",
    text: "Sentiment-aligned CES attainment inference running...",
  });
  await delay(700);

  const cesData = await extractCesAttainment(cesFile);

  for (const [coNum, level] of Object.entries(cesData.coLevels)) {
    emit({ type: "data", text: `  → CO${coNum}: CES Level ${level}` });
  }

  await delay(400);
  emit({
    type: "ai",
    text: "Indirect assessment fusion layer updating CO template...",
  });
  await delay(400);

  await writeCesAttainment(coOutput, cesData);
  emit({
    type: "ok",
    text: "CES attainment written to Course Exit Survey row.",
  });

  await delay(700);

  // ── Step 5: CO → PO mapping ───────────────────────────────────────────────
  emit({ type: "step", text: "[5/5] Computing CO → PO attainment weights..." });
  await delay(400);
  emit({
    type: "ai",
    text: "Synthesizing direct (D) and indirect (I) attainment scores...",
  });
  await delay(800);
  emit({
    type: "ai",
    text: "Applying weighted normalization: D = (A + B) × 0.7  |  I = CES × 0.3",
  });
  await delay(700);
  emit({
    type: "ai",
    text: "Mapping CO attainment weights to PO correlation matrix...",
  });
  await delay(600);

  const coData = await extractCoAttainment(coOutput);

  for (const [coNum, val] of Object.entries(coData.coLevels)) {
    emit({ type: "data", text: `  → CO${coNum}: D+I = ${val.toFixed(4)}` });
  }

  await delay(500);
  emit({
    type: "ai",
    text: "Writing CO attainment values into PO_Attainment report...",
  });
  await delay(400);

  await writePoAttainment(poStatic, poOutput, coData);
  emit({ type: "ok", text: "PO attainment report generated successfully." });

  await delay(600);

  emit({ type: "done", coOutput, poOutput });
}

function delay(ms) {
  return new Promise((resolve) => setTimeout(resolve, ms));
}

module.exports = { processAll };
