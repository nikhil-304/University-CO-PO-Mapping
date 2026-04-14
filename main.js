"use strict";

const { app, BrowserWindow, ipcMain, dialog, shell } = require("electron");
const path = require("path");
const fs = require("fs");

let mainWindow;

function createWindow() {
  mainWindow = new BrowserWindow({
    width: 1200,
    height: 820,
    minWidth: 900,
    minHeight: 600,
    backgroundColor: "#0a0e1a",
    frame: false,
    webPreferences: {
      preload: path.join(__dirname, "preload.js"),
      contextIsolation: true,
      nodeIntegration: false,
    },
  });

  mainWindow.loadFile(path.join(__dirname, "renderer", "index.html"));
}

app.whenReady().then(createWindow);

app.on("window-all-closed", () => {
  if (process.platform !== "darwin") app.quit();
});

app.on("activate", () => {
  if (BrowserWindow.getAllWindows().length === 0) createWindow();
});

// ── IPC: file selection ────────────────────────────────────────────────────

ipcMain.handle("open-file-dialog", async (_e, opts = {}) => {
  const result = await dialog.showOpenDialog(mainWindow, {
    title: opts.title || "Select Excel File",
    filters: [{ name: "Excel Files", extensions: ["xlsx", "xls"] }],
    properties: ["openFile"],
  });
  return result.canceled ? null : result.filePaths[0];
});

// ── IPC: get templates path ────────────────────────────────────────────────

ipcMain.handle("get-templates-path", () => {
  return path.join(__dirname, "templates");
});

ipcMain.handle("check-templates", () => {
  const dir = path.join(__dirname, "templates");
  const co = path.join(dir, "CO_Attainment_Static.xlsx");
  const po = path.join(dir, "PO_Attainment_Static.xlsx");
  return {
    coExists: fs.existsSync(co),
    poExists: fs.existsSync(po),
    coPath: co,
    poPath: po,
  };
});

// ── IPC: process files ─────────────────────────────────────────────────────

ipcMain.handle(
  "process-files",
  async (_e, { assignFile, unitFile, marksFile, cesFile }) => {
    const templatesDir = path.join(__dirname, "templates");
    const coStatic = path.join(templatesDir, "CO_Attainment_Static.xlsx");
    const poStatic = path.join(templatesDir, "PO_Attainment_Static.xlsx");

    if (!fs.existsSync(coStatic))
      throw new Error("CO_Attainment_Static.xlsx not found in templates/");
    if (!fs.existsSync(poStatic))
      throw new Error("PO_Attainment_Static.xlsx not found in templates/");

    // prepare output directory inside userData
    const outputDir = path.join(app.getPath("userData"), "attain-ai-outputs");
    fs.mkdirSync(outputDir, { recursive: true });

    const ts = Date.now();
    const coOutput = path.join(outputDir, `CO_Attainment_Final_${ts}.xlsx`);
    const poOutput = path.join(outputDir, `PO_Attainment_Final_${ts}.xlsx`);

    // lazy-load engine so we don't block app startup
    const engine = require("./src/attainment-engine");

    await engine.processAll({
      assignFile,
      unitFile,
      marksFile,
      cesFile,
      coStatic,
      poStatic,
      coOutput,
      poOutput,
      onProgress: (msg) => {
        if (!mainWindow.isDestroyed()) {
          mainWindow.webContents.send("progress", msg);
        }
      },
    });

    return { coOutput, poOutput };
  },
);

// ── IPC: save output file ──────────────────────────────────────────────────

ipcMain.handle("save-output", async (_e, { sourcePath, defaultName }) => {
  const result = await dialog.showSaveDialog(mainWindow, {
    title: "Save Output File",
    defaultPath: defaultName || "output.xlsx",
    filters: [{ name: "Excel Files", extensions: ["xlsx"] }],
  });
  if (result.canceled) return null;
  fs.copyFileSync(sourcePath, result.filePath);
  return result.filePath;
});

ipcMain.handle("reveal-file", (_e, filePath) => {
  shell.showItemInFolder(filePath);
});

// ── IPC: window controls (for custom title bar) ────────────────────────────

ipcMain.handle("win-minimize", () => mainWindow.minimize());
ipcMain.handle("win-maximize", () => {
  mainWindow.isMaximized() ? mainWindow.unmaximize() : mainWindow.maximize();
});
ipcMain.handle("win-close", () => mainWindow.close());
