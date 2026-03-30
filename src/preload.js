const { contextBridge, ipcRenderer } = require('electron');

contextBridge.exposeInMainWorld('electronAPI', {
  openFile: () => ipcRenderer.invoke('open-file'),
  convertMdToDocx: (content) => ipcRenderer.invoke('convert-md-to-docx', content),
  saveFile: (base64Data, originalName) => ipcRenderer.invoke('save-file', base64Data, originalName),
});
