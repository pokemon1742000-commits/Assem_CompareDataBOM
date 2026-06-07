const { contextBridge, ipcRenderer } = require('electron');

contextBridge.exposeInMainWorld('inventoryApi', {
  getAppVersion: () => ipcRenderer.invoke('app:version'),
  openGithub: () => ipcRenderer.invoke('app:openGithub'),
  openExcel: () => ipcRenderer.invoke('excel:open'),
  readExcelSheets: (selections) => ipcRenderer.invoke('excel:readSheets', selections),
  exportExcel: (payload) => ipcRenderer.invoke('excel:export', payload),
  exportCompare: (payload) => ipcRenderer.invoke('excel:exportCompare', payload),
  exportDiscrepancy: (payload) => ipcRenderer.invoke('excel:exportDiscrepancy', payload),
  loadRecent: () => ipcRenderer.invoke('recent:load'),
  saveRecent: (payload) => ipcRenderer.invoke('recent:save', payload),
  clearRecent: () => ipcRenderer.invoke('recent:clear'),
  checkForUpdates: () => ipcRenderer.invoke('update:check'),
  onUpdateStatus: (callback) => {
    ipcRenderer.on('update:status', (_event, message) => callback(message));
  }
});
