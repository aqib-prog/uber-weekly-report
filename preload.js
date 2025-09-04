// preload.js - Updated with manual setup APIs
const { contextBridge, ipcRenderer } = require("electron");

contextBridge.exposeInMainWorld("api", {
  openLogin: () => ipcRenderer.invoke("open-login"),
  saveSession: () => ipcRenderer.invoke("save-session"),
  smokeEarnings: () => ipcRenderer.invoke("smoke-earnings"),
  hasSession: () => ipcRenderer.invoke("has-session"),

  // NEW: Manual setup approach (these were missing!)
  openUberForManualSetup: () =>
    ipcRenderer.invoke("open-uber-for-manual-setup"),
  runAutomation: () => ipcRenderer.invoke("run-automation"),

  // Keep for backwards compatibility
  runWeekly: (payload) => ipcRenderer.invoke("run-weekly", payload),
  generatePdf: (excelFilePath) =>
    ipcRenderer.invoke("generate-pdf", excelFilePath),
  downloadFile: (filePath) => ipcRenderer.invoke("download-file", filePath),
});
