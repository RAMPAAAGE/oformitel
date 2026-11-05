# Профили оформления - это "рецепты": какой нужен шрифт, размер, интервал и поля.
# Каждый профиль хранится в отдельном файле .json в папке profiles.

import copy
import json
import os

# Типы документов. От типа зависит, какие разделы обязательно должны быть в работе.
# Программа только предупреждает, что раздела нет, и не пишет его за автора.
DOCUMENT_TYPES = {
    "research": {
        "label": "Исследовательская работа",
        "title_page": True,
        "toc": True,
        "sections": ["введение", "заключение", "список литературы"],
    },
    "referat": {
        "label": "Реферат",
        "title_page": True,
        "toc": True,
        "sections": ["введение", "заключение", "список литературы"],
    },
    "practical": {
        "label": "Практическая работа",
        "title_page": True,
        "toc": False,
        "sections": ["цель", "ход работы", "результат", "заключение"],
    },
    "report": {
        "label": "Доклад",
        "title_page": False,
        "toc": False,
        "sections": [],
    },
    "free": {
        "label": "Только оформление",
        "title_page": False,
        "toc": False,
        "sections": [],
    },
}


def load_profile(path):
    if not os.path.isfile(path):
        raise ValueError("Файл профиля не найден.")
    try:
        with open(path, "r", encoding="utf-8") as file:
            return json.load(file)
    except json.JSONDecodeError:
        raise ValueError("Файл профиля содержит неверные данные JSON.")


def list_profiles(folder="profiles"):
    # список профилей для выпадающего списка: имя файла и название
    if not os.path.isdir(folder):
        return []
    profiles = []
    for file_name in sorted(os.listdir(folder)):
        path = os.path.join(folder, file_name)
        if file_name.lower().endswith(".json") and os.path.isfile(path):
            name = load_profile(path).get("name", "Без названия")
            profiles.append({"file": file_name, "name": name})
    return profiles


def create_profile(
    name,
    font_name,
    font_size_pt,
    line_spacing,
    first_line_indent_cm,
    margin_top_cm,
    margin_bottom_cm,
    margin_left_cm,
    margin_right_cm,
    document_type="free",
    profiles_folder="profiles",
):
    if not name:
        raise ValueError("Укажите название профиля.")
    if not font_name:
        raise ValueError("Укажите название шрифта.")
    if document_type not in DOCUMENT_TYPES:
        raise ValueError("Выбран неизвестный тип документа.")

    os.makedirs(profiles_folder, exist_ok=True)
    profile = {
        "name": name,
        "font_name": font_name,
        "font_size_pt": font_size_pt,
        "line_spacing": line_spacing,
        "first_line_indent_cm": first_line_indent_cm,
        "margin_top_cm": margin_top_cm,
        "margin_bottom_cm": margin_bottom_cm,
        "margin_left_cm": margin_left_cm,
        "margin_right_cm": margin_right_cm,
        "document_type": document_type,
        "structure": copy.deepcopy(DOCUMENT_TYPES[document_type]),
    }

    # ищем свободное имя: profile_1.json, profile_2.json, ...
    number = 1
    while os.path.exists(os.path.join(profiles_folder, "profile_" + str(number) + ".json")):
        number += 1
    path = os.path.join(profiles_folder, "profile_" + str(number) + ".json")

    with open(path, "w", encoding="utf-8") as file:
        json.dump(profile, file, ensure_ascii=False, indent=4)
    return path
