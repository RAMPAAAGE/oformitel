const { app, BrowserWindow, dialog } = require("electron");
const { spawn } = require("child_process");
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
    OFORMITEL_PORT: "0",
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
    },
  });

  window.loadURL("http://127.0.0.1:" + port);
}


function stopFlask() {
  applicationIsClosing = true;
  clearTimeout(startupTimer);

  if (flaskProcess && !flaskProcess.killed) {
    flaskProcess.kill();
  }
}


function startApplication() {
  startFlask();
}


app.whenReady().then(startApplication);

app.on("window-all-closed", () => {
  app.quit();
});

app.on("before-quit", stopFlask);
