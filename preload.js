"use strict";

const { contextBridge, ipcRenderer } = require("electron");

contextBridge.exposeInMainWorld("api", {
  openFile: (opts) => ipcRenderer.invoke("open-file-dialog", opts),
  checkTemplates: () => ipcRenderer.invoke("check-templates"),
  processFiles: (payload) => ipcRenderer.invoke("process-files", payload),
  saveOutput: (payload) => ipcRenderer.invoke("save-output", payload),
  revealFile: (filePath) => ipcRenderer.invoke("reveal-file", filePath),
  onProgress: (cb) => ipcRenderer.on("progress", (_e, msg) => cb(msg)),
  removeProgressListeners: () => ipcRenderer.removeAllListeners("progress"),
  winMinimize: () => ipcRenderer.invoke("win-minimize"),
  winMaximize: () => ipcRenderer.invoke("win-maximize"),
  winClose: () => ipcRenderer.invoke("win-close"),
});
