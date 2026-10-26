#!/bin/bash

# удаляет сохранённые документы и титульные листы (папки uploads и outputs)
# приложение Electron хранит их в ~/Library/Application Support/oformitel-dokumentov
# профили не трогает, файл .gitkeep тоже оставляет

PROJECT_DIR="$(cd "$(dirname "$0")" && pwd)"

if [ "$#" -gt 0 ]; then
    FOLDERS=("$@")
else
    FOLDERS=(
        "$HOME/Library/Application Support/oformitel-dokumentov/uploads"
        "$HOME/Library/Application Support/oformitel-dokumentov/outputs"
        "$PROJECT_DIR/backend/uploads"
        "$PROJECT_DIR/backend/outputs"
    )
fi

total=0
for folder in "${FOLDERS[@]}"; do
    if [ ! -d "$folder" ]; then
        echo "Папки нет, пропускаю: $folder"
        continue
    fi
    count=$(find "$folder" -mindepth 1 -maxdepth 1 -type f ! -name ".gitkeep" | wc -l | tr -d ' ')
    find "$folder" -mindepth 1 -maxdepth 1 -type f ! -name ".gitkeep" -delete
    echo "Удалено файлов: $count из $folder"
    total=$((total + count))
done

echo "Готово. Всего удалено: $total"
read -n 1 -s -r -p "Нажмите любую клавишу, чтобы закрыть окно..."
echo
