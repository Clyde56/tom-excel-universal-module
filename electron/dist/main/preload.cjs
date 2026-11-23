"use strict";

// src/preload/preload.ts
var import_electron = require("electron");
import_electron.contextBridge.exposeInMainWorld("tomExcel", {
  openExternal: (payload) => import_electron.ipcRenderer.invoke("shell:open-external", payload),
  getPathForFile: (file) => import_electron.webUtils.getPathForFile(file),
  openWorkbook: (payload) => import_electron.ipcRenderer.invoke("workbook:open", payload),
  chooseWorkbookFile: (payload) => import_electron.ipcRenderer.invoke("workbook:choose-open", payload),
  inspectWorkbook: (payload) => import_electron.ipcRenderer.invoke("workbook:inspect", payload),
  chooseIconFile: () => import_electron.ipcRenderer.invoke("icon:choose"),
  chooseIconFilePath: () => import_electron.ipcRenderer.invoke("icon:choose-file"),
  readIconFile: (payload) => import_electron.ipcRenderer.invoke("icon:read-file", payload),
  listOfficeImageMsoIcons: (payload) => import_electron.ipcRenderer.invoke("icon:list-mso", payload),
  persistOfficeImageMsoIcon: (payload) => import_electron.ipcRenderer.invoke("icon:persist-mso", payload),
  toggleUnsupportedImageMsoIcon: (payload) => import_electron.ipcRenderer.invoke("icon:toggle-unsupported-mso", payload),
  chooseSaveTarget: (payload) => import_electron.ipcRenderer.invoke("workbook:choose-save-target", payload),
  saveWorkbookCopy: (payload) => import_electron.ipcRenderer.invoke("workbook:save-copy", payload),
  saveWorkbookCopyToPath: (payload) => import_electron.ipcRenderer.invoke("workbook:save-copy-to-path", payload),
  createInstallerPackage: (payload) => import_electron.ipcRenderer.invoke("workbook:create-installer-package", payload),
  previewWorkbook: (payload) => import_electron.ipcRenderer.invoke("workbook:preview", payload),
  getAiSettings: () => import_electron.ipcRenderer.invoke("ai:get-settings"),
  saveAiSettings: (payload) => import_electron.ipcRenderer.invoke("ai:save-settings", payload),
  testAiConnection: (payload) => import_electron.ipcRenderer.invoke("ai:test-connection", payload),
  planRibbonWithAi: (payload) => import_electron.ipcRenderer.invoke("ai:plan-ribbon", payload),
  repairMissingMacroBindingsWithAi: (payload) => import_electron.ipcRenderer.invoke("ai:repair-missing-bindings", payload),
  planRibbonWithAiStream: (payload) => import_electron.ipcRenderer.invoke("ai:plan-ribbon-stream", payload),
  cancelAiPlanRibbonStream: (payload) => import_electron.ipcRenderer.invoke("ai:cancel-plan-ribbon-stream", payload),
  submitAiAgentToolResult: (payload) => import_electron.ipcRenderer.invoke("ai:agent-tool-result", payload),
  onAiPlanRibbonStreamEvent: (callback) => {
    const listener = (_event, payload) => callback(payload);
    import_electron.ipcRenderer.on("ai:plan-ribbon-stream-event", listener);
    return () => import_electron.ipcRenderer.removeListener("ai:plan-ribbon-stream-event", listener);
  },
  onVbaAccessRequest: (callback) => {
    const listener = (_event, payload) => callback(payload);
    import_electron.ipcRenderer.on("vba-access:request", listener);
    return () => import_electron.ipcRenderer.removeListener("vba-access:request", listener);
  },
  respondVbaAccessRequest: (payload) => import_electron.ipcRenderer.invoke("vba-access:respond", payload)
});
