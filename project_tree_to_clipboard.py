#!/usr/bin/env python3
"""
project_tree_to_clipboard.py

Кладёт скрипт в корень проекта и запускает:
    python project_tree_to_clipboard.py

Скрипт строит дерево файлов текущей папки (рекурсивно) и копирует
результат в буфер обмена. Работает на Windows, macOS и Linux
без установки сторонних библиотек (использует системные утилиты
clip / pbcopy / xclip / xsel / wl-copy).

Настройки (что игнорировать / насколько глубоко идти) — ниже,
в блоке CONFIG.
"""

import os
import platform
import subprocess
import sys
from pathlib import Path

# ------------------------- CONFIG -------------------------

# Папки, которые не нужно показывать в дереве
IGNORE_DIRS = {
    ".git", ".svn", ".hg",
    "node_modules", "__pycache__", ".pytest_cache",
    "venv", ".venv", "env",
    "dist", "build", "target", "out",
    ".idea", ".vscode", ".mypy_cache",
    ".next", ".nuxt", ".cache",
    "coverage", ".DS_Store",
}

# Расширения/файлы, которые тоже стоит скрыть (по желанию)
IGNORE_FILES = {
    ".DS_Store", "Thumbs.db",
}

# Максимальная глубина вложенности (None = без ограничений)
MAX_DEPTH = None

# ------------------------------------------------------------


def build_tree(root: Path, prefix: str = "", depth: int = 0, lines=None):
    if lines is None:
        lines = []

    if MAX_DEPTH is not None and depth > MAX_DEPTH:
        return lines

    try:
        entries = sorted(
            root.iterdir(),
            key=lambda p: (p.is_file(), p.name.lower())
        )
    except PermissionError:
        return lines

    entries = [
        e for e in entries
        if e.name not in IGNORE_DIRS and e.name not in IGNORE_FILES
    ]

    for i, entry in enumerate(entries):
        is_last = (i == len(entries) - 1)
        connector = "└── " if is_last else "├── "
        lines.append(f"{prefix}{connector}{entry.name}")

        if entry.is_dir():
            extension = "    " if is_last else "│   "
            build_tree(entry, prefix + extension, depth + 1, lines)

    return lines


def copy_to_clipboard(text: str) -> bool:
    system = platform.system()

    try:
        if system == "Windows":
            # clip.exe принимает cp866/utf-16, поэтому кодируем аккуратно
            process = subprocess.Popen(
                ["clip"], stdin=subprocess.PIPE, close_fds=True
            )
            process.communicate(input=text.encode("utf-16-le"))
            return process.returncode == 0

        elif system == "Darwin":  # macOS
            process = subprocess.Popen(
                ["pbcopy"], stdin=subprocess.PIPE
            )
            process.communicate(input=text.encode("utf-8"))
            return process.returncode == 0

        else:  # Linux и прочие *nix
            candidates = [
                ["xclip", "-selection", "clipboard"],
                ["xsel", "--clipboard", "--input"],
                ["wl-copy"],
            ]
            for cmd in candidates:
                try:
                    process = subprocess.Popen(
                        cmd, stdin=subprocess.PIPE
                    )
                    process.communicate(input=text.encode("utf-8"))
                    if process.returncode == 0:
                        return True
                except FileNotFoundError:
                    continue
            return False

    except Exception as e:
        print(f"Ошибка при копировании в буфер: {e}", file=sys.stderr)
        return False


def main():
    root = Path.cwd()
    header = f"{root.name}/"
    tree_lines = [header] + build_tree(root)
    tree_text = "\n".join(tree_lines)

    print(tree_text)
    print()

    ok = copy_to_clipboard(tree_text)
    if ok:
        print(">>> Дерево скопировано в буфер обмена.")
    else:
        print(">>> Не удалось скопировать автоматически.")
        system = platform.system()
        if system == "Linux":
            print("    Установи xclip или wl-clipboard, например:")
            print("    sudo apt install xclip")
        # На всякий случай сохраним в файл рядом
        out_file = root / "project_tree.txt"
        out_file.write_text(tree_text, encoding="utf-8")
        print(f"    Дерево сохранено в файл: {out_file}")


if __name__ == "__main__":
    main()
