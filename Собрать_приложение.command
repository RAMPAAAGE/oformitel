#!/bin/bash

set -euo pipefail

PROJECT_DIR="$(cd "$(dirname "$0")" && pwd)"
BACKEND_DIR="$PROJECT_DIR/backend"
ELECTRON_DIR="$PROJECT_DIR/electron"
VENV_DIR="$BACKEND_DIR/.venv"
VENV_PYTHON="$VENV_DIR/bin/python3"
BUILD_DIR="$(dirname "$PROJECT_DIR")/билды"
BREW_BIN=""
PYTHON_BIN=""
NPM_BIN=""


find_brew() {
    if command -v brew >/dev/null 2>&1; then
        BREW_BIN="$(command -v brew)"
        return
    fi

    if [ -x /opt/homebrew/bin/brew ]; then
        BREW_BIN="/opt/homebrew/bin/brew"
        return
    fi

    if [ -x /usr/local/bin/brew ]; then
        BREW_BIN="/usr/local/bin/brew"
    fi
}


ensure_brew() {
    find_brew

    if [ -n "$BREW_BIN" ]; then
        return
    fi

    echo "Для установки недостающих программ нужен Homebrew."
    read -r -p "Установить Homebrew сейчас? [y/N] " answer

    case "$answer" in
        y|Y|yes|YES)
            /bin/bash -c "$(curl -fsSL https://raw.githubusercontent.com/Homebrew/install/HEAD/install.sh)"
            find_brew
            ;;
        *)
            echo "Сборка остановлена: без Homebrew нельзя установить недостающие программы."
            exit 1
            ;;
    esac

    if [ -z "$BREW_BIN" ]; then
        echo "Homebrew установлен, но его не удалось найти. Открой новый Терминал и запусти скрипт ещё раз."
        exit 1
    fi

    eval "$("$BREW_BIN" shellenv)"
}


python_is_suitable() {
    "$1" -c 'import sys; raise SystemExit(0 if sys.version_info >= (3, 10) else 1)'
}


ensure_python() {
    PYTHON_BIN="$(command -v python3 || true)"

    if [ -n "$PYTHON_BIN" ] && python_is_suitable "$PYTHON_BIN"; then
        echo "Python уже установлен: $($PYTHON_BIN --version)"
        return
    fi

    ensure_brew
    echo "Устанавливаю Python через Homebrew..."
    "$BREW_BIN" install python
    PYTHON_BIN="$("$BREW_BIN" --prefix python)/bin/python3"

    if ! python_is_suitable "$PYTHON_BIN"; then
        echo "Не удалось подготовить подходящую версию Python."
        exit 1
    fi
}


node_is_suitable() {
    "$1" -e 'process.exit(Number(process.versions.node.split(".")[0]) >= 20 ? 0 : 1)'
}


ensure_node() {
    local node_bin
    node_bin="$(command -v node || true)"

    if [ -n "$node_bin" ] && node_is_suitable "$node_bin"; then
        NPM_BIN="$(command -v npm)"
        echo "Node.js уже установлен: $($node_bin --version)"
        return
    fi

    ensure_brew
    echo "Устанавливаю Node.js через Homebrew..."
    "$BREW_BIN" install node
    eval "$("$BREW_BIN" shellenv)"
    NPM_BIN="$(command -v npm || true)"

    if [ -z "$NPM_BIN" ]; then
        echo "Node.js установился, но npm не найден. Открой новый Терминал и запусти скрипт ещё раз."
        exit 1
    fi
}


prepare_python_packages() {
    if [ ! -x "$VENV_PYTHON" ]; then
        echo "Создаю виртуальное окружение Python..."
        "$PYTHON_BIN" -m venv "$VENV_DIR"
    fi

    if "$VENV_PYTHON" -c "import flask, docx, pptx, PyInstaller" >/dev/null 2>&1; then
        echo "Python-библиотеки уже установлены."
        return
    fi

    echo "Устанавливаю Python-библиотеки проекта..."
    "$VENV_PYTHON" -m pip install -r "$BACKEND_DIR/requirements.txt" -r "$BACKEND_DIR/requirements-build.txt"
}


prepare_electron_packages() {
    if (
        cd "$ELECTRON_DIR"
        "$NPM_BIN" ls electron electron-builder --depth=0 >/dev/null 2>&1
    ); then
        echo "Electron-библиотеки уже установлены."
        return
    fi

    echo "Устанавливаю Electron-библиотеки проекта..."
    (
        cd "$ELECTRON_DIR"
        "$NPM_BIN" install
    )
}


build_application() {
    mkdir -p "$BUILD_DIR"
    echo "Собираю приложение. Это может занять несколько минут..."

    (
        cd "$ELECTRON_DIR"
        PATH="$VENV_DIR/bin:$PATH" \
        PYINSTALLER_CONFIG_DIR="$PROJECT_DIR/.pyinstaller-cache" \
        "$NPM_BIN" run dist
    )
}


show_result() {
    local dmg_file
    dmg_file="$(find "$BUILD_DIR" -maxdepth 1 -type f -name '*.dmg' -print -quit)"

    if [ -z "$dmg_file" ]; then
        echo "Сборка завершилась без файла .dmg. Проверь сообщения выше."
        exit 1
    fi

    echo
    echo "Готово! Установочный файл создан здесь:"
    echo "$dmg_file"
    open "$BUILD_DIR"
}


close_script_tab() {
    local terminal_tty
    terminal_tty="$(tty 2>/dev/null || true)"

    if [ "${TERM_PROGRAM:-}" != "Apple_Terminal" ] || [ -z "$terminal_tty" ]; then
        return
    fi

    osascript - "$terminal_tty" 2>/dev/null <<'APPLESCRIPT' || true
on run argv
    tell application "Terminal"
        repeat with terminalWindow in windows
            repeat with terminalTab in tabs of terminalWindow
                if tty of terminalTab is item 1 of argv then
                    close terminalTab
                    return
                end if
            end repeat
        end repeat
    end tell
end run
APPLESCRIPT
}


echo "=== Сборка Оформителя документов для macOS ==="
ensure_python
ensure_node
prepare_python_packages
prepare_electron_packages
build_application
show_result
close_script_tab
