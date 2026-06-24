// Preload: exposes a small, safe API to the renderer over the context bridge.
const { contextBridge, ipcRenderer } = require('electron');

contextBridge.exposeInMainWorld('api', {
  debug: !!process.env.PDF_EDITOR_DEBUG,
  openFile: () => ipcRenderer.invoke('open-file'),
  saveFileDialog: (defaultName) => ipcRenderer.invoke('save-file-dialog', defaultName),
  saveXlsxDialog: (defaultName) => ipcRenderer.invoke('save-xlsx-dialog', defaultName),
  writeFile: (filePath, data) => ipcRenderer.invoke('write-file', filePath, data),
  readFile: (filePath) => ipcRenderer.invoke('read-file', filePath),
  setTitle: (title) => ipcRenderer.send('set-title', title),
  approveClose: () => ipcRenderer.send('close-approved'),
  onMenuAction: (callback) =>
    ipcRenderer.on('menu-action', (_evt, action) => callback(action)),
});
