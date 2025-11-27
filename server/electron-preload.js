const { contextBridge, ipcRenderer } = require('electron');

contextBridge.exposeInMainWorld('electronAPI', {
  // Now pass folderPath as argument to generateWorksheets
  generateWorksheets: (folderPath) => ipcRenderer.invoke('generateWorksheets', folderPath),
  // Folder picker dialog
  chooseOutputFolder: () => ipcRenderer.invoke('chooseOutputFolder'),
  openOneDriveFolder: () => ipcRenderer.invoke('openOneDriveFolder')
});