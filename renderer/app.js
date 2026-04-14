"use strict";
// ── ATTAIN-AI Terminal Frontend ──────────────────────────────────────────

const output = document.getElementById("output");
const cmdInput = document.getElementById("cmd-input");

// session state
let busy = false;
let sessionFiles = { assign: null, unit: null, marks: null, ces: null };
let lastOutputs = { co: null, po: null };
let cmdHistory = [];
let historyIdx = -1;

// ── Utilities ──────────────────────────────────────────────────────────────

function ts() {
  const d = new Date();
  const p = (n) => String(n).padStart(2, "0");
  return `${p(d.getHours())}:${p(d.getMinutes())}:${p(d.getSeconds())}`;
}

function scrollBottom() {
  output.scrollTop = output.scrollHeight;
}

/**
 * Append a line to the terminal output.
 * @param {string} text   - the text to display
 * @param {string} type   - cmd | ai | step | ok | data | done | err | warn | info | ask | banner
 * @param {Object} opts   - { noTs, typing }
 */
function print(text, type = "info", opts = {}) {
  const ICONS = {
    cmd: "❯",
    ai: "◆",
    step: "▶",
    ok: "✓",
    data: "·",
    done: "★",
    err: "✗",
    warn: "⚠",
    info: " ",
    ask: "?",
    banner: " ",
  };

  const line = document.createElement("div");
  line.className = "line";

  if (!opts.noTs) {
    const tSpan = document.createElement("span");
    tSpan.className = "line-ts";
    tSpan.textContent = `[${ts()}]`;
    line.appendChild(tSpan);
  }

  const icon = document.createElement("span");
  icon.className = `line-body pfx-${type}`;
  icon.textContent = (ICONS[type] ?? " ") + " ";
  line.appendChild(icon);

  const body = document.createElement("span");
  body.className = `line-body txt-${type}`;

  if (opts.typing) {
    line.classList.add("typing-line");
    body.textContent = text;
  } else {
    body.textContent = text;
  }

  line.appendChild(body);
  output.appendChild(line);
  scrollBottom();
  return line;
}

function printBlank() {
  const d = document.createElement("div");
  d.style.height = "4px";
  output.appendChild(d);
  scrollBottom();
}

function printDivider() {
  const hr = document.createElement("hr");
  hr.className = "divider";
  output.appendChild(hr);
  scrollBottom();
}

function removeTypingCursor(lineEl) {
  lineEl && lineEl.classList.remove("typing-line");
}

function printBanner() {
  const ASCII = `
╔══════════════════════════════════════════════════════════════════════╗
║                                                                      ║
║   ░█████╗░████████╗████████╗░█████╗░██╗███╗░░██╗                     ║
║   ██╔══██╗╚══██╔══╝╚══██╔══╝██╔══██╗██║████╗░██║                     ║
║   ███████║░░░██║░░░░░░██║░░░███████║██║██╔██╗██║                     ║
║   ██╔══██║░░░██║░░░░░░██║░░░██╔══██║██║██║╚████║                     ║
║   ██║░░██║░░░██║░░░░░░██║░░░██║░░██║██║██║░╚███║                     ║
║   ╚═╝░░╚═╝░░╚═╝░░░░░░╚═╝░░░╚═╝░░╚═╝╚═╝╚═╝░░╚══╝                      ║
║                                                                      ║
║   ░█████╗░██╗                                                        ║
║   ██╔══██╗██║                                                        ║
║   ███████║██║   Academic Intelligence Engine  ·  SPPU                ║
║   ██╔══██║██║   CO-PO Attainment Mapping System  v2.0                ║
║   ██║░░██║██║   Made by Nikhil Shrivastava                           ║
║   ╚═╝░░╚═╝╚═╝                                                        ║
║                                                                      ║
╚══════════════════════════════════════════════════════════════════════╝`;

  const pre = document.createElement("pre");
  pre.className = "ascii-banner";
  pre.textContent = ASCII;
  output.appendChild(pre);

  printBlank();
  print(
    "Multi-layer attainment inference  |  Weighted CO-PO fusion  |  SPPU rubric",
    "info",
  );
  printDivider();
  scrollBottom();
}

// ── Boot sequence ──────────────────────────────────────────────────────────

async function boot() {
  printBanner();
  await sleep(300);

  const l1 = print("Initializing neural inference sub-systems...", "ai", {
    typing: true,
  });
  await sleep(900);
  removeTypingCursor(l1);

  const l2 = print(
    "Loading CO-PO attainment prediction models (384M parameters)...",
    "ai",
    { typing: true },
  );
  await sleep(1100);
  removeTypingCursor(l2);

  const l3 = print("Calibrating SPPU rubric classification schema...", "ai", {
    typing: true,
  });
  await sleep(700);
  removeTypingCursor(l3);

  // check templates
  const tpl = await window.api.checkTemplates();
  if (tpl.coExists && tpl.poExists) {
    print("CO_Attainment_Static.xlsx  ✓  loaded (template)", "ok");
    print("PO_Attainment_Static.xlsx  ✓  loaded (template)", "ok");
  } else {
    if (!tpl.coExists)
      print("CO_Attainment_Static.xlsx not found in templates/", "err");
    if (!tpl.poExists)
      print("PO_Attainment_Static.xlsx not found in templates/", "err");
    print(
      "Place both static template files in the templates/ directory.",
      "warn",
    );
  }

  await sleep(400);
  printDivider();
  print("All systems nominal. Multi-layer fusion engine ready.", "ok");
  printBlank();
  print("Type  help  to see available commands.", "info");
  printBlank();
}

// ── Command handler ────────────────────────────────────────────────────────

async function handleCommand(raw) {
  const cmd = raw.trim().toLowerCase();
  if (!cmd) return;

  cmdHistory.unshift(raw.trim());
  historyIdx = -1;

  // echo the typed command
  print(raw.trim(), "cmd");

  if (cmd === "help") {
    cmdHelp();
  } else if (cmd === "run" || cmd === "start") {
    await cmdRun();
  } else if (cmd === "clear" || cmd === "cls") {
    output.innerHTML = "";
    printBanner();
  } else if (cmd === "status") {
    cmdStatus();
  } else if (cmd === "reset") {
    sessionFiles = { assign: null, unit: null, marks: null, ces: null };
    lastOutputs = { co: null, po: null };
    print("Session reset. All file selections cleared.", "ok");
  } else if (cmd.startsWith("download") || cmd === "dl") {
    cmdDownload();
  } else if (cmd === "exit" || cmd === "quit") {
    print("Shutting down ATTAIN-AI...", "ai");
    await sleep(600);
    window.close();
  } else {
    print(
      `Unknown command: '${raw.trim()}'. Type 'help' for available commands.`,
      "err",
    );
  }
}

// ── Commands ───────────────────────────────────────────────────────────────

function cmdHelp() {
  printBlank();
  const lines = [
    ["run", "Start the full CO-PO attainment analysis pipeline"],
    ["status", "Show currently staged input files"],
    ["download", "Download generated output files (after run)"],
    ["reset", "Clear all staged files and outputs"],
    ["clear", "Clear the terminal screen"],
    ["help", "Show this help message"],
    ["exit", "Close the application"],
  ];
  print("Available commands:", "info");
  for (const [c, d] of lines) {
    print(`  ${c.padEnd(12)} ${d}`, "data");
  }
  printBlank();
}

function cmdStatus() {
  printBlank();
  print("Session file status:", "info");
  const entries = [
    ["Assignment marks", sessionFiles.assign],
    ["Unit test marks", sessionFiles.unit],
    ["INSEM/ENDSEM marks", sessionFiles.marks],
    ["Course Exit Survey", sessionFiles.ces],
  ];
  for (const [label, f] of entries) {
    if (f) {
      print(`  ${label.padEnd(22)} ✓  ${shortPath(f)}`, "ok");
    } else {
      print(`  ${label.padEnd(22)} —  (not selected)`, "warn");
    }
  }

  if (lastOutputs.co || lastOutputs.po) {
    printBlank();
    print("Generated outputs:", "info");
    if (lastOutputs.co)
      print(`  CO report  →  ${shortPath(lastOutputs.co)}`, "ok");
    if (lastOutputs.po)
      print(`  PO report  →  ${shortPath(lastOutputs.po)}`, "ok");
    print("Type 'download' to save the output files.", "info");
  }
  printBlank();
}

function shortPath(p) {
  if (!p) return "";
  const parts = p.replace(/\\/g, "/").split("/");
  return parts.slice(-1)[0];
}

async function cmdRun() {
  if (busy) {
    print("Processing is already in progress. Please wait...", "warn");
    return;
  }
  busy = true;
  cmdInput.disabled = true;

  try {
    printBlank();
    print(
      "Starting ATTAIN-AI academic attainment analysis pipeline...",
      "step",
    );
    await sleep(400);
    printDivider();

    // ── File selection ─────────────────────────────────────────────────────
    const fileDefs = [
      {
        key: "assign",
        num: 1,
        label: "ASSIGNMENT MARKS",
        desc: "Theory assignment Excel file",
        detail: "Columns: Assign1, Assign2, Assign3  (with CO mappings)",
      },
      {
        key: "unit",
        num: 2,
        label: "UNIT TEST MARKS",
        desc: "Unit test Excel file",
        detail: "Columns: UT1, UT2  (with CO mappings)",
      },
      {
        key: "marks",
        num: 3,
        label: "MAIN MARKS  ( INSEM / ENDSEM )",
        desc: "In-Sem and End-Sem examination Excel file",
        detail: "Columns: IN SEM, END SEM  (with CO mappings)",
      },
      {
        key: "ces",
        num: 4,
        label: "COURSE EXIT SURVEY  ( CES )",
        desc: "CES attainment summary Excel file",
        detail: "Contains a Level column for each CO (1 – 3)",
      },
    ];

    for (const def of fileDefs) {
      // ── file header block ──
      printBlank();
      printBlank();
      printFileHeader(
        def.num,
        fileDefs.length,
        def.label,
        def.desc,
        def.detail,
      );
      printBlank();

      print(`Press  ENTER  to open the file picker when ready…`, "ask");
      await waitForEnter();
      print(`Opening file picker…`, "info");
      await sleep(300);
      const p = await window.api.openFile({
        title: `[${def.num}/4] Select — ${def.label}`,
      });

      if (!p) {
        printBlank();
        print("File selection cancelled. Run aborted.", "err");
        busy = false;
        cmdInput.disabled = false;
        return;
      }

      sessionFiles[def.key] = p;

      // success row
      const chip = document.createElement("div");
      chip.style.cssText =
        "margin-left:26px; margin-top:6px; margin-bottom:4px;";
      chip.innerHTML = `<span class="file-chip">&#x2714; ${shortPath(p)}</span>`;
      output.appendChild(chip);
      scrollBottom();
      await sleep(250);
    }

    printBlank();
    printDivider();
    print("All 4 input files staged successfully.", "ok");
    printBlank();

    // ── Confirmation prompt ────────────────────────────────────────────────
    print("Proceed with full CO-PO attainment analysis?  [y / n]", "ask");
    const confirmed = await waitForYesNo();
    if (!confirmed) {
      printBlank();
      print("Aborted. Files remain staged — type 'run' to try again.", "warn");
      busy = false;
      cmdInput.disabled = false;
      return;
    }

    printBlank();
    print("Confirmed. Initiating inference pipeline...", "ok");
    await sleep(400);

    // ── AI warm-up messages ────────────────────────────────────────────────
    const warmup = [
      "Bootstrapping attention-based attainment inference model...",
      "Loading pre-trained SPPU CO rubric embeddings...",
      "Constructing student performance feature tensors...",
      "Warming up CO-PO correlation weight matrices...",
    ];
    for (const msg of warmup) {
      const l = print(msg, "ai", { typing: true });
      await sleep(650 + Math.random() * 400);
      removeTypingCursor(l);
    }

    printDivider();

    // ── Progress listener ──────────────────────────────────────────────────
    window.api.removeProgressListeners();
    window.api.onProgress((msg) => {
      if (msg.type === "step") {
        printBlank();
        print(msg.text, "step");
      } else if (msg.type === "ai") {
        const l = print(msg.text, "ai", { typing: true });
        setTimeout(() => removeTypingCursor(l), 500);
      } else if (msg.type === "ok") {
        print(msg.text, "ok");
      } else if (msg.type === "data") {
        print(msg.text, "data");
      } else if (msg.type === "warn") {
        print(msg.text, "warn");
      } else if (msg.type === "done") {
        lastOutputs.co = msg.coOutput;
        lastOutputs.po = msg.poOutput;
      }
    });

    // ── Invoke engine ──────────────────────────────────────────────────────
    const result = await window.api.processFiles({
      assignFile: sessionFiles.assign,
      unitFile: sessionFiles.unit,
      marksFile: sessionFiles.marks,
      cesFile: sessionFiles.ces,
    });

    // Ensure lastOutputs is set (from IPC return value as fallback)
    if (result && result.coOutput) lastOutputs.co = result.coOutput;
    if (result && result.poOutput) lastOutputs.po = result.poOutput;

    // ── Completion ─────────────────────────────────────────────────────────
    await sleep(300);
    printDivider();
    print("═══════════════════════════════════════════════", "done", {
      noTs: true,
    });
    print("  ATTAIN-AI ANALYSIS COMPLETE", "done", { noTs: true });
    print("═══════════════════════════════════════════════", "done", {
      noTs: true,
    });
    printBlank();
    print("Both attainment reports have been generated.", "ok");
    print("Type 'download' to save the files to your chosen location.", "info");
    printBlank();

    // Download buttons
    renderDownloadButtons();
  } catch (err) {
    printDivider();
    print(`Error: ${err.message || err}`, "err");
    print("Please verify all input files and try again.", "warn");
  } finally {
    busy = false;
    cmdInput.disabled = false;
    cmdInput.focus();
  }
}

function renderDownloadButtons() {
  const block = document.createElement("div");
  block.className = "dl-block";

  if (lastOutputs.co) {
    const btn = document.createElement("button");
    btn.className = "dl-btn";
    btn.innerHTML = `<span class="dl-icon">⬇</span> Save CO_Attainment_Final.xlsx <span class="dl-sub">CO attainment report</span>`;
    btn.onclick = () =>
      downloadFile(lastOutputs.co, "CO_Attainment_Final.xlsx");
    block.appendChild(btn);
  }

  if (lastOutputs.po) {
    const btn = document.createElement("button");
    btn.className = "dl-btn";
    btn.innerHTML = `<span class="dl-icon">⬇</span> Save PO_Attainment_Final.xlsx <span class="dl-sub">PO attainment report</span>`;
    btn.onclick = () =>
      downloadFile(lastOutputs.po, "PO_Attainment_Final.xlsx");
    block.appendChild(btn);
  }

  output.appendChild(block);
  scrollBottom();
}

async function downloadFile(sourcePath, defaultName) {
  print(`Saving ${defaultName}...`, "ai");
  const dest = await window.api.saveOutput({ sourcePath, defaultName });
  if (dest) {
    print(`Saved → ${dest}`, "ok");
  } else {
    print("Save cancelled.", "warn");
  }
}

function cmdDownload() {
  if (!lastOutputs.co && !lastOutputs.po) {
    print(
      "No outputs available yet. Run 'run' first to generate reports.",
      "warn",
    );
    return;
  }
  renderDownloadButtons();
}

// ── Input handling ─────────────────────────────────────────────────────────

cmdInput.addEventListener("keydown", async (e) => {
  if (e.key === "Enter") {
    const val = cmdInput.value;
    // Only clear & handle when not busy (or status override).
    // When busy=true, other handlers (waitForYesNo) read the value first.
    if (!busy || val.trim().toLowerCase() === "status") {
      cmdInput.value = "";
      await handleCommand(val);
    }
  } else if (e.key === "ArrowUp") {
    e.preventDefault();
    if (historyIdx < cmdHistory.length - 1) {
      historyIdx++;
      cmdInput.value = cmdHistory[historyIdx];
    }
  } else if (e.key === "ArrowDown") {
    e.preventDefault();
    if (historyIdx > 0) {
      historyIdx--;
      cmdInput.value = cmdHistory[historyIdx];
    } else {
      historyIdx = -1;
      cmdInput.value = "";
    }
  }
});

// Window controls
document.getElementById("wc-close").onclick = () => window.api.winClose();
document.getElementById("wc-min").onclick = () => window.api.winMinimize();
document.getElementById("wc-max").onclick = () => window.api.winMaximize();

// ── Helpers ────────────────────────────────────────────────────────────────

function sleep(ms) {
  return new Promise((r) => setTimeout(r, ms));
}

/**
 * Print a bordered block header for each file selection step.
 */
function printFileHeader(num, total, label, desc, detail) {
  const W = 58;
  const bar = "─".repeat(W);
  const pad = (s) => {
    const space = W - 2 - s.length;
    return "│ " + s + " ".repeat(Math.max(0, space)) + " │";
  };
  const lines = [
    "┌" + bar + "┐",
    pad(""),
    pad(`  [ ${num} / ${total} ]  ${label}`),
    pad(`  ${desc}`),
    pad(`  ${detail}`),
    pad(""),
    "└" + bar + "┘",
  ];
  const pre = document.createElement("pre");
  pre.className = "ascii-banner";
  pre.style.borderLeft = "3px solid var(--cyan)";
  pre.style.paddingLeft = "8px";
  pre.style.margin = "4px 0";
  pre.textContent = lines.join("\n");
  output.appendChild(pre);
  scrollBottom();
}

/**
 * Re-enable input, wait for ANY Enter key press, then disable again.
 * Used to let the user read the context panel before the file dialog fires.
 */
function waitForEnter() {
  return new Promise((resolve) => {
    cmdInput.disabled = false;
    cmdInput.value = "";
    cmdInput.focus();

    function onKey(e) {
      if (e.key !== "Enter") return;
      e.stopImmediatePropagation();
      cmdInput.value = "";
      cmdInput.disabled = true;
      cmdInput.removeEventListener("keydown", onKey, { capture: true });
      resolve();
    }

    cmdInput.addEventListener("keydown", onKey, { capture: true });
  });
}

/**
 * Enable input, wait for y/yes or n/no, then disable again.
 * Uses capture phase so this fires BEFORE the global keydown clears the value.
 */
function waitForYesNo() {
  return new Promise((resolve) => {
    cmdInput.disabled = false;
    cmdInput.value = "";
    cmdInput.focus();

    function onKey(e) {
      if (e.key !== "Enter") return;
      e.stopImmediatePropagation(); // prevent global handler from clearing value
      const val = cmdInput.value.trim().toLowerCase();
      cmdInput.value = "";
      if (val === "y" || val === "yes") {
        print("y", "cmd");
        cmdInput.disabled = true;
        cmdInput.removeEventListener("keydown", onKey, { capture: true });
        resolve(true);
      } else if (val === "n" || val === "no") {
        print("n", "cmd");
        cmdInput.disabled = true;
        cmdInput.removeEventListener("keydown", onKey, { capture: true });
        resolve(false);
      } else {
        print("Type  y  to proceed  or  n  to abort", "warn");
      }
    }

    // capture: true fires BEFORE the bubbling-phase global keydown handler
    cmdInput.addEventListener("keydown", onKey, { capture: true });
  });
}

// ── Start ──────────────────────────────────────────────────────────────────

// Script is at end of <body>, so DOM is ready. Also guard with DOMContentLoaded.
if (document.readyState === "loading") {
  document.addEventListener("DOMContentLoaded", () =>
    boot().then(() => cmdInput.focus()),
  );
} else {
  boot().then(() => cmdInput.focus());
}

// Keep input focused whenever terminal area is clicked
document.getElementById("terminal").addEventListener("click", () => {
  if (!busy) cmdInput.focus();
});
