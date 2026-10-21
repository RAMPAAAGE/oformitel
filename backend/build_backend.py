import os
import subprocess
import sys


BASE_DIR = os.path.dirname(os.path.abspath(__file__))
ELECTRON_FOLDER = os.path.join(BASE_DIR, "..", "electron")
BUILD_FOLDER = os.environ.get("OFORMITEL_BUILD_FOLDER", os.path.join(BASE_DIR, "build"))
DIST_FOLDER = os.environ.get("OFORMITEL_DIST_FOLDER", os.path.join(BASE_DIR, "dist"))


def make_data_argument(source, destination):
    return source + os.pathsep + destination


def build_backend():
    os.makedirs(BUILD_FOLDER, exist_ok=True)
    os.makedirs(DIST_FOLDER, exist_ok=True)
    command = [
        sys.executable,
        "-m",
        "PyInstaller",
        "--noconfirm",
        "--clean",
        "--onefile",
        "--name",
        "oformitel-backend",
        "--distpath",
        DIST_FOLDER,
        "--workpath",
        BUILD_FOLDER,
        "--specpath",
        BUILD_FOLDER,
        "--collect-all",
        "docx",
        "--collect-all",
        "pptx",
        "--add-data",
        make_data_argument(os.path.join(BASE_DIR, "profiles"), "profiles"),
        "--add-data",
        make_data_argument(os.path.join(ELECTRON_FOLDER, "build", "index.html"), "electron"),
        "--add-data",
        make_data_argument(os.path.join(ELECTRON_FOLDER, "build", "style.css"), "electron"),
        "app.py",
    ]

    subprocess.run(command, check=True, cwd=BASE_DIR)


if __name__ == "__main__":
    build_backend()
