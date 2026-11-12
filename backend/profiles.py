import copy
import json
import os

# от типа документа зависит, какие разделы в работе обязательны
DOCUMENT_TYPES = {
    "research": {
        "label": "Исследовательская работа",
        "title_page": True,
        "toc": True,
        "sections": ["введение", "заключение", "список литературы"],
    },
    "project": {
        "label": "Проектная работа",
        "title_page": True,
        "toc": True,
        "sections": ["введение", "заключение", "список литературы"],
    },
    "business_plan": {
        "label": "Бизнес-план",
        "title_page": True,
        "toc": True,
        "sections": [
            "описание бизнес-идеи", "анализ состояния рынка", "ценностное предложение и анализ целевой аудитории",
            "бизнес-модель", "операционный план", "финансовая модель",
        ],
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
    # для выпадающих списков и для формы изменения профиля: имя файла и все настройки
    if not os.path.isdir(folder):
        return []
    profiles = []
    for file_name in sorted(os.listdir(folder)):
        path = os.path.join(folder, file_name)
        if file_name.lower().endswith(".json") and os.path.isfile(path):
            data = load_profile(path)
            profiles.append({"file": file_name, "name": data.get("name", "Без названия"), "data": data})
    return profiles


def make_profile(name, font_name, font_size_pt, line_spacing, first_line_indent_cm,
                 margin_top_cm, margin_bottom_cm, margin_left_cm, margin_right_cm, document_type="free"):
    if not name:
        raise ValueError("Укажите название профиля.")
    if not font_name:
        raise ValueError("Укажите название шрифта.")
    if document_type not in DOCUMENT_TYPES:
        raise ValueError("Выбран неизвестный тип документа.")
    return {
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


def save_profile(profile, path):
    with open(path, "w", encoding="utf-8") as file:
        json.dump(profile, file, ensure_ascii=False, indent=4)


def create_profile(profiles_folder="profiles", **values):
    profile = make_profile(**values)
    os.makedirs(profiles_folder, exist_ok=True)
    number = 1
    while os.path.exists(os.path.join(profiles_folder, "profile_" + str(number) + ".json")):
        number += 1
    path = os.path.join(profiles_folder, "profile_" + str(number) + ".json")
    save_profile(profile, path)
    return path


def update_profile(path, **values):
    if not os.path.isfile(path):
        raise ValueError("Профиль не найден.")
    profile = make_profile(**values)
    # отметка нужна, чтобы при запуске программа не заменила изменённый встроенный профиль стандартным
    profile["edited"] = True
    save_profile(profile, path)
    return path
