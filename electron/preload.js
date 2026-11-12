// Мост между страницей программы и Electron: страница узнаёт, что файл сохранён,
// и может попросить открыть его. Больше ничего странице не доступно.
const { contextBridge, ipcRenderer } = require("electron");

contextBridge.exposeInMainWorld("desktop", {
  onFileSaved: (callback) => ipcRenderer.on("file-saved", (event, info) => callback(info)),
  openFile: (filePath) => ipcRenderer.invoke("open-file", filePath),
  showFile: (filePath) => ipcRenderer.invoke("show-file", filePath),
});
