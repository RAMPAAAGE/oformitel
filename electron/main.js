const { app, BrowserWindow, dialog, ipcMain, shell } = require("electron");
const { spawn } = require("child_process");
const fs = require("fs");
const path = require("path");


let flaskProcess;
let applicationIsClosing = false;
let startupErrorWasShown = false;
let startupTimer;
let serverOutput = "";
let serverIsReady = false;


function getBackendFolder() {
  if (app.isPackaged) {
    return path.join(process.resourcesPath, "backend");
  }

  return path.join(__dirname, "..", "backend");
}


function getPythonCommand() {
  if (process.platform === "win32") {
    return "python";
  }

  return "python3";
}


function getBackendStart() {
  const backendFolder = getBackendFolder();

  if (app.isPackaged) {
    let fileName = "oformitel-backend";
    if (process.platform === "win32") {
      fileName += ".exe";
    }

    return {
      command: path.join(backendFolder, fileName),
      arguments: [],
      folder: backendFolder,
    };
  }

  return {
    command: getPythonCommand(),
    arguments: [path.join(backendFolder, "app.py")],
    folder: backendFolder,
  };
}


function showStartupError(message) {
  if (startupErrorWasShown || applicationIsClosing) {
    return;
  }

  startupErrorWasShown = true;
  dialog.showErrorBox("Не удалось запустить программу", message);
  app.quit();
}


function startFlask() {
  const backend = getBackendStart();
  const environment = {
    ...process.env,
    OFORMITEL_DATA_FOLDER: app.getPath("userData"),
    OFORMITEL_PORT: "0", // 0 значит "возьми любой свободный порт"
  };

  flaskProcess = spawn(backend.command, backend.arguments, {
    cwd: backend.folder,
    env: environment,
    stdio: ["ignore", "pipe", "ignore"],
    windowsHide: true,
  });

  flaskProcess.on("error", () => {
    showStartupError("Не удалось запустить внутреннюю часть программы.");
  });

  flaskProcess.on("exit", () => {
    if (!applicationIsClosing) {
      showStartupError("Внутренняя часть программы неожиданно завершила работу.");
    }
  });

  flaskProcess.stdout.on("data", checkServerOutput);
  startupTimer = setTimeout(() => {
    showStartupError("Сервер программы не ответил за 30 секунд.");
  }, 30000);
}


function checkServerOutput(data) {
  if (serverIsReady) {
    return;
  }

  serverOutput += data.toString();
  const match = serverOutput.match(/READY (\d+)/);

  if (match) {
    serverIsReady = true;
    clearTimeout(startupTimer);
    createWindow(Number(match[1]));
  }
}


// Готовые файлы сохраняются сразу в «Загрузки», без окна «Сохранить как».
// Открыть или показать в папке можно только файл, который сохранила сама программа.
const savedFiles = new Set();

function freeFilePath(folder, fileName) {
  const extension = path.extname(fileName);
  const name = path.basename(fileName, extension);
  let filePath = path.join(folder, fileName);
  let number = 2;
  while (fs.existsSync(filePath)) {
    filePath = path.join(folder, name + " (" + number + ")" + extension);
    number += 1;
  }
  return filePath;
}

function saveDownloadsToFolder(window) {
  window.webContents.session.on("will-download", (event, item) => {
    // OFORMITEL_DOWNLOADS - другая папка для сохранения (нужно автотестам)
    const folder = process.env.OFORMITEL_DOWNLOADS || app.getPath("downloads");
    const filePath = freeFilePath(folder, item.getFilename());
    item.setSavePath(filePath);
    item.once("done", (doneEvent, state) => {
      const saved = state === "completed";
      if (saved) {
        savedFiles.add(filePath);
      }
      window.webContents.send("file-saved", { saved: saved, path: filePath, name: path.basename(filePath) });
    });
  });
}

ipcMain.handle("open-file", (event, filePath) => {
  if (!savedFiles.has(filePath)) {
    return "Файл не найден";
  }
  // автотесты проверяют, что файл сохранён, но не открывают Word
  if (process.env.OFORMITEL_DONT_OPEN) {
    return "";
  }
  return shell.openPath(filePath);
});

ipcMain.handle("show-file", (event, filePath) => {
  if (savedFiles.has(filePath)) {
    shell.showItemInFolder(filePath);
  }
});

function createWindow(port) {
  const window = new BrowserWindow({
    width: 1000,
    height: 760,
    minWidth: 720,
    minHeight: 600,
    title: "Оформитель документов",
    icon: path.join(__dirname, "build", "icon.png"),
    webPreferences: {
      contextIsolation: true,
      preload: path.join(__dirname, "preload.js"),
    },
  });

  saveDownloadsToFolder(window);
  window.loadURL("http://127.0.0.1:" + port);
}


function stopFlask() {
  applicationIsClosing = true;
  clearTimeout(startupTimer);

  if (flaskProcess && !flaskProcess.killed) {
    flaskProcess.kill();
  }
}


// Порядок работы: Electron запускает Flask, Flask пишет "READY <порт>",
// и только после этого открывается окно программы.
app.whenReady().then(startFlask);

app.on("window-all-closed", () => {
  app.quit();
});

app.on("before-quit", stopFlask);
