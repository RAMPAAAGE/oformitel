const { existsSync } = require("fs");
const { spawnSync } = require("child_process");
const path = require("path");

let command = "python3";
const projectFolder = path.join(__dirname, "..");
const localPython = process.platform === "win32"
  ? path.join(projectFolder, "backend", ".venv", "Scripts", "python.exe")
  : path.join(projectFolder, "backend", ".venv", "bin", "python");

if (existsSync(localPython)) {
  command = localPython;
}

const args = [path.join(__dirname, "..", "backend", "build_backend.py")];

if (process.platform === "win32") {
  command = "python";
}

const result = spawnSync(command, args, { stdio: "inherit" });

if (result.error) {
  console.error("Не удалось запустить Python для сборки backend.");
  process.exit(1);
}

if (result.signal) {
  console.error("Процесс сборки завершён сигналом: " + result.signal);
  process.exit(1);
}

process.exit(result.status ?? 0);
