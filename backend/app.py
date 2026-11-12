import json
import os
import re
import shutil
import sys
import time

from flask import Flask, jsonify, redirect, render_template, request, send_from_directory, url_for
from werkzeug.utils import secure_filename
from werkzeug.serving import make_server

from docx import Document
from engine import (
    CRITERIA_NAMES,
    DOCUMENT_TYPES,
    analyze_layout,
    check_and_format_document,
    check_file_size,
    create_profile,
    list_profiles,
    load_profile,
    update_profile,
    validate_docx_file,
)
from pptx_engine import analyze_presentation, validate_pptx_file
from template import PART_TIPS, WORK_TYPES, create_template, default_plan
from title_page import create_title_page_document


def get_program_folder():
    if getattr(sys, "frozen", False):
        return sys._MEIPASS
    return os.path.dirname(os.path.abspath(__file__))


def get_data_folder():
    folder = os.environ.get("OFORMITEL_DATA_FOLDER")
    if folder:
        return folder
    if getattr(sys, "frozen", False):
        return os.path.join(os.path.expanduser("~"), ".oformitel")
    return PROGRAM_FOLDER


PROGRAM_FOLDER = get_program_folder()
DEFAULT_PROFILES_FOLDER = os.path.join(PROGRAM_FOLDER, "profiles")
DATA_FOLDER = get_data_folder()
MAX_SHOWN_CHANGES = 300
PROFILES_FOLDER = os.path.join(DATA_FOLDER, "profiles")
UPLOADS_FOLDER = os.path.join(DATA_FOLDER, "uploads")
OUTPUTS_FOLDER = os.path.join(DATA_FOLDER, "outputs")
# сохранённые пользователем структуры работ («мои шаблоны»)
MY_TEMPLATES_FOLDER = os.path.join(DATA_FOLDER, "my_templates")

if getattr(sys, "frozen", False):
    TEMPLATES_FOLDER = os.path.join(PROGRAM_FOLDER, "electron")
else:
    TEMPLATES_FOLDER = os.path.join(PROGRAM_FOLDER, "..", "electron", "build")

app = Flask(
    __name__,
    template_folder=TEMPLATES_FOLDER,
    static_folder=OUTPUTS_FOLDER,
    static_url_path="/outputs",
)


def create_work_folders():
    os.makedirs(PROFILES_FOLDER, exist_ok=True)
    os.makedirs(UPLOADS_FOLDER, exist_ok=True)
    os.makedirs(OUTPUTS_FOLDER, exist_ok=True)
    os.makedirs(MY_TEMPLATES_FOLDER, exist_ok=True)
    copy_default_profiles()


# так встроенные профили назывались в старых версиях; их копии у пользователя дублируют новые
OLD_BUILTIN_FILES = ["gost_7_32.json", "inzhenery_budushchego.json"]


def copy_default_profiles():
    if os.path.abspath(PROFILES_FOLDER) == os.path.abspath(DEFAULT_PROFILES_FOLDER):
        return
    if not os.path.isdir(DEFAULT_PROFILES_FOLDER):
        return
    for file_name in OLD_BUILTIN_FILES:
        old_path = os.path.join(PROFILES_FOLDER, file_name)
        if os.path.isfile(old_path) and not load_profile(old_path).get("edited"):
            os.remove(old_path)
    for file_name in os.listdir(DEFAULT_PROFILES_FOLDER):
        if not file_name.lower().endswith(".json"):
            continue
        source_path = os.path.join(DEFAULT_PROFILES_FOLDER, file_name)
        destination_path = os.path.join(PROFILES_FOLDER, file_name)
        # встроенные профили обновляем при каждом запуске, чтобы исправления в новой версии
        # дошли до пользователя. Но если пользователь сам изменил профиль - не трогаем
        if os.path.isfile(destination_path) and load_profile(destination_path).get("edited"):
            continue
        if os.path.isfile(source_path):
            shutil.copyfile(source_path, destination_path)


def is_builtin_profile(file_name):
    return os.path.isfile(os.path.join(DEFAULT_PROFILES_FOLDER, file_name))


def get_free_file_name(folder, file_name):
    name, extension = os.path.splitext(file_name)
    free_name = file_name
    number = 2
    while os.path.exists(os.path.join(folder, free_name)):
        free_name = name + "_" + str(number) + extension
        number += 1
    return free_name


def title_page_label(file_name):
    # «Название работы» — 13.11 00:27; название программа записывает в свойства файла при создании
    path = os.path.join(OUTPUTS_FOLDER, file_name)
    created = time.strftime("%d.%m %H:%M", time.localtime(os.path.getmtime(path)))
    try:
        work_title = Document(path).core_properties.title
    except Exception:
        work_title = ""
    if not work_title:
        work_title = file_name
    return "«" + work_title + "» — " + created


def list_title_pages():
    result = []
    for file_name in os.listdir(OUTPUTS_FOLDER):
        if file_name.startswith("titulnyi_list") and file_name.lower().endswith(".docx"):
            file_time = os.path.getmtime(os.path.join(OUTPUTS_FOLDER, file_name))
            result.append({"file": file_name, "label": title_page_label(file_name), "time": file_time})
    result.sort(key=lambda item: item["time"], reverse=True)
    return result


def list_my_templates():
    result = []
    for file_name in sorted(os.listdir(MY_TEMPLATES_FOLDER)):
        if file_name.endswith(".json"):
            try:
                with open(os.path.join(MY_TEMPLATES_FOLDER, file_name), encoding="utf-8") as file:
                    data = json.load(file)
                result.append({"file": file_name, "name": data["name"], "plan": data["plan"]})
            except (ValueError, KeyError):
                continue
    return result


def get_my_template_path(form):
    file_name = form.get("template_file", "")
    path = os.path.join(MY_TEMPLATES_FOLDER, file_name)
    if file_name != os.path.basename(file_name) or not os.path.isfile(path):
        raise ValueError("Шаблон не найден.")
    return path


def show_main_page(active_tab="document", error="", message="", select_profile=""):
    profiles = list_profiles(PROFILES_FOLDER)
    return render_template(
        "index.html",
        profiles=profiles,
        builtin_files=[profile["file"] for profile in profiles if is_builtin_profile(profile["file"])],
        document_types=DOCUMENT_TYPES,
        work_types=WORK_TYPES,
        part_tips=PART_TIPS,
        my_templates=list_my_templates(),
        title_pages=list_title_pages(),
        active_tab=active_tab,
        error=error,
        message=message,
        select_profile=select_profile,
    )


def save_upload(uploaded_file, extension, default_name):
    file_name = secure_filename(uploaded_file.filename)
    if not file_name.lower().endswith(extension):
        file_name = default_name
    input_name = get_free_file_name(UPLOADS_FOLDER, file_name)
    input_path = os.path.join(UPLOADS_FOLDER, input_name)
    uploaded_file.save(input_path)

    size_is_ok, _ = check_file_size(input_path)
    if not size_is_ok:
        os.remove(input_path)
        raise ValueError("Размер загружаемого файла не должен превышать 5 МБ.")

    return input_name, input_path


def get_profile_values(form):
    try:
        values = {
            "name": form["name"].strip(),
            "font_name": form["font_name"].strip(),
            "font_size_pt": int(form["font_size_pt"]),
            "line_spacing": float(form["line_spacing"]),
            "first_line_indent_cm": float(form["first_line_indent_cm"]),
            "margin_top_cm": float(form["margin_top_cm"]),
            "margin_bottom_cm": float(form["margin_bottom_cm"]),
            "margin_left_cm": float(form["margin_left_cm"]),
            "margin_right_cm": float(form["margin_right_cm"]),
            "document_type": form.get("document_type", "free"),
        }
    except (KeyError, ValueError):
        raise ValueError("Заполните все числовые поля профиля корректно.")

    if not values["name"] or not values["font_name"]:
        raise ValueError("Укажите название профиля и шрифт.")
    if not 1 <= values["font_size_pt"] <= 100:
        raise ValueError("Размер шрифта должен быть от 1 до 100 пт.")
    if not 0.1 <= values["line_spacing"] <= 5:
        raise ValueError("Межстрочный интервал должен быть от 0.1 до 5.")
    if not 0 <= values["first_line_indent_cm"] <= 10:
        raise ValueError("Отступ первой строки должен быть от 0 до 10 см.")
    for margin_name in ("margin_top_cm", "margin_bottom_cm", "margin_left_cm", "margin_right_cm"):
        if not 0.1 <= values[margin_name] <= 20:
            raise ValueError("Поля страницы должны быть от 0.1 до 20 см.")
    return values


def get_profile_path(form):
    profile_file = form.get("profile_file", "")
    profile_path = os.path.join(PROFILES_FOLDER, profile_file)
    if profile_file != os.path.basename(profile_file) or not os.path.isfile(profile_path):
        raise ValueError("Выбранный профиль не найден.")
    return profile_path


def get_title_path(title_file):
    # None - титульный лист не выбран
    if not title_file:
        return None
    title_path = os.path.join(OUTPUTS_FOLDER, title_file)
    if title_file != os.path.basename(title_file) or not title_file.startswith("titulnyi_list") \
            or not os.path.isfile(title_path):
        raise ValueError("Выбранный титульный лист не найден.")
    return title_path


def get_template_plan(form):
    # настройки шаблона приходят из формы одной строкой JSON
    try:
        plan = json.loads(form.get("plan", ""))
    except ValueError:
        raise ValueError("Не удалось прочитать настройки шаблона.")
    result = default_plan(plan.get("type", "research"))
    for flag in ["toc", "intro", "numbered", "conclusion", "literature", "appendix", "picture", "table", "hints"]:
        if flag in plan:
            result[flag] = bool(plan[flag])
    chapters = []
    for chapter in plan.get("chapters", []):
        subsections = [str(name)[:200] for name in chapter.get("subsections", [])][:20]
        chapters.append({
            "name": str(chapter.get("name", ""))[:200],
            "subsections": subsections,
            "hint": str(chapter.get("hint", ""))[:1000],
        })
    result["chapters"] = chapters[:30]
    result["title"] = str(plan.get("title", "none"))
    return result


def get_filled_list(form, field_name):
    result = []
    for value in form.getlist(field_name):
        if value.strip():
            result.append(value.strip())
    return result


def get_groups(form):
    # группа - ученики одного класса: у поля класса имя group_0_who, у фамилий - group_0_name и т.д.
    numbers = []
    for key in form.keys():
        match = re.fullmatch(r"group_(\d+)_who", key)
        if match:
            numbers.append(int(match.group(1)))
    groups = []
    for number in sorted(numbers):
        prefix = "group_" + str(number) + "_"
        names = get_filled_list(form, prefix + "name")
        if names:
            groups.append({
                "who": form.get(prefix + "who", "").strip(),
                "school": form.get(prefix + "school", "").strip(),
                "names": names,
            })
    return groups


def get_supervisors(form):
    supervisors = []
    for position, name in zip(form.getlist("supervisor_position"), form.getlist("supervisor_name")):
        if name.strip():
            supervisors.append({"position": position.strip(), "name": name.strip()})
    return supervisors


def get_title_page_values(form):
    return {
        "organization": form.get("organization", "").strip(),
        "work_title": form.get("work_title", "").strip(),
        "groups": get_groups(form),
        "supervisors": get_supervisors(form),
        "city": form.get("city", "").strip(),
        "year": form.get("year", "").strip(),
    }


create_work_folders()


@app.get("/")
def index():
    message = ""
    active_tab = "document"
    # после сохранения профиль сразу выбирается в панели «Требования конференции»
    select_profile = request.args.get("select", "")
    if request.args.get("profile_created") == "1":
        message = "Профиль «" + request.args.get("name", "") + "» создан и выбран."
        active_tab = "profiles"
    if request.args.get("profile_saved") == "1":
        message = "Изменения в профиле «" + request.args.get("name", "") + "» сохранены."
        active_tab = "profiles"
    if request.args.get("profile_deleted") == "1":
        message = "Профиль удалён."
        active_tab = "profiles"
    return show_main_page(active_tab=active_tab, message=message, select_profile=select_profile)


@app.get("/style.css")
def style_sheet():
    return send_from_directory(TEMPLATES_FOLDER, "style.css")


@app.post("/apply")
def apply_profile():
    try:
        uploaded_file = request.files.get("document")
        if uploaded_file is None or uploaded_file.filename == "":
            raise ValueError("Выберите документ .docx.")
        if not uploaded_file.filename.lower().endswith(".docx"):
            raise ValueError("Можно выбрать только файл с расширением .docx.")

        profile_path = get_profile_path(request.form)

        title_path = get_title_path(request.form.get("title_file", ""))

        input_name, input_path = save_upload(uploaded_file, ".docx", "document.docx")
        try:
            validate_docx_file(input_path)

            skip_title = request.form.get("skip_title") == "1"
            if not title_path and not skip_title and not analyze_layout(Document(input_path))["title"]:
                return jsonify({
                    "no_title": True,
                    "message": "В работе нет титульного листа. Создайте его здесь же — после этого работа оформится "
                               "автоматически. Или оформите работу без титульного листа.",
                })

            output_name = get_free_file_name(OUTPUTS_FOLDER, "formatted_" + input_name)
            output_path = os.path.join(OUTPUTS_FOLDER, output_name)
            result = check_and_format_document(input_path, profile_path, output_path, title_path,
                                               request.form.get("document_type"))
        finally:
            # исходный файл больше не нужен: результат уже сохранён в outputs
            os.remove(input_path)

        warnings = result["warnings"]

        if result["is_correct"]:
            return jsonify({
                "correct": True,
                "message": "Документ уже оформлен по профилю «" + result["profile_name"] + "». Исправления не требуются.",
                "warnings": warnings,
            })

        message = ""
        size_is_ok, size_mb = check_file_size(output_path)
        if not size_is_ok:
            message = "Внимание: размер файла " + str(round(size_mb, 1)) + " МБ, это больше допустимых 5 МБ."

        download_name = "оформлено_" + os.path.basename(uploaded_file.filename)
        return jsonify({
            "correct": False,
            "url": url_for("download_file", file_name=output_name, name=download_name),
            "download_name": download_name,
            "criteria": [
                {"n": n, "name": CRITERIA_NAMES[n], "details": result["details"].get(n, "")}
                for n in result["failed"]
            ],
            "title_inserted": result["title_inserted"],
            "changes": result["changes"][:MAX_SHOWN_CHANGES],
            "changes_total": len(result["changes"]),
            "warnings": warnings,
            "message": message,
        })

    except ValueError as error:
        return jsonify({"error": str(error)}), 400
    except Exception:
        return jsonify({"error": "Не удалось обработать документ. Проверьте, что это неповреждённый файл .docx."}), 400


@app.post("/save_profile")
def save_profile_route():
    try:
        values = get_profile_values(request.form)
        profile_file = request.form.get("profile_file", "")
        if profile_file == "":
            path = create_profile(profiles_folder=PROFILES_FOLDER, **values)
            return redirect(url_for("index", profile_created="1", name=values["name"],
                                    select=os.path.basename(path)))
        path = update_profile(get_profile_path(request.form), **values)
        return redirect(url_for("index", profile_saved="1", name=values["name"], select=os.path.basename(path)))
    except ValueError as error:
        return show_main_page(active_tab="profiles", error=str(error)), 400
    except Exception:
        return show_main_page(
            active_tab="profiles",
            error="Не удалось сохранить профиль. Проверьте заполненные поля.",
        ), 400


@app.post("/delete_profile")
def delete_profile():
    try:
        profile_path = get_profile_path(request.form)
        if is_builtin_profile(os.path.basename(profile_path)):
            raise ValueError("Встроенный профиль удалить нельзя, но можно вернуть ему стандартные настройки.")
        os.remove(profile_path)
        return redirect(url_for("index", profile_deleted="1"))
    except ValueError as error:
        return show_main_page(active_tab="profiles", error=str(error)), 400


@app.post("/reset_profile")
def reset_profile():
    try:
        profile_path = get_profile_path(request.form)
        file_name = os.path.basename(profile_path)
        if not is_builtin_profile(file_name):
            raise ValueError("Стандартные настройки есть только у встроенных профилей.")
        shutil.copyfile(os.path.join(DEFAULT_PROFILES_FOLDER, file_name), profile_path)
        return redirect(url_for("index", profile_saved="1", name=load_profile(profile_path)["name"],
                                select=file_name))
    except (ValueError, shutil.SameFileError) as error:
        if isinstance(error, shutil.SameFileError):
            error = "Программа запущена из исходников: встроенные профили меняются прямо в папке проекта."
        return show_main_page(active_tab="profiles", error=str(error)), 400


@app.get("/download/<path:file_name>")
def download_file(file_name):
    # имя для сохранения передаём в заголовке: окно программы берёт его оттуда
    name = request.args.get("name") or file_name
    return send_from_directory(OUTPUTS_FOLDER, file_name, as_attachment=True, download_name=name)


@app.post("/create_template")
def create_template_route():
    try:
        plan = get_template_plan(request.form)
        profile = load_profile(get_profile_path(request.form))
        title_path = None
        if plan["title"] not in ("none", "placeholder"):
            title_path = get_title_path(plan["title"])
        output_name = get_free_file_name(OUTPUTS_FOLDER, "shablon.docx")
        create_template(os.path.join(OUTPUTS_FOLDER, output_name), profile, plan, title_path)
        download_name = "Шаблон — " + WORK_TYPES[plan["type"]]["label"] + ".docx"
        return jsonify({
            "url": url_for("download_file", file_name=output_name, name=download_name),
            "download_name": download_name,
        })
    except ValueError as error:
        return jsonify({"error": str(error)}), 400
    except Exception:
        return jsonify({"error": "Не удалось создать шаблон. Проверьте настройки."}), 400


@app.post("/save_my_template")
def save_my_template():
    try:
        name = request.form.get("name", "").strip()[:100]
        if not name:
            raise ValueError("Придумайте название шаблона.")
        plan = get_template_plan(request.form)
        file_name = get_free_file_name(MY_TEMPLATES_FOLDER, "template.json")
        with open(os.path.join(MY_TEMPLATES_FOLDER, file_name), "w", encoding="utf-8") as file:
            json.dump({"name": name, "plan": plan}, file, ensure_ascii=False, indent=4)
        return jsonify({"file": file_name, "name": name, "plan": plan})
    except ValueError as error:
        return jsonify({"error": str(error)}), 400


@app.post("/delete_my_template")
def delete_my_template():
    try:
        os.remove(get_my_template_path(request.form))
        return jsonify({"deleted": True})
    except ValueError as error:
        return jsonify({"error": str(error)}), 400


@app.post("/create_title_page")
def create_title_page():
    try:
        fields = get_title_page_values(request.form)
        output_name = get_free_file_name(OUTPUTS_FOLDER, "titulnyi_list.docx")
        output_path = os.path.join(OUTPUTS_FOLDER, output_name)
        profile = load_profile(get_profile_path(request.form))
        create_title_page_document(fields, output_path, profile)
        return jsonify({
            "file": output_name,
            "label": title_page_label(output_name),
            "url": url_for("download_file", file_name=output_name,
                           name="Титульный лист — " + fields["work_title"][:80] + ".docx"),
        })
    except ValueError as error:
        return jsonify({"error": str(error)}), 400
    except Exception:
        return jsonify({"error": "Не удалось создать титульный лист. Проверьте заполненные поля."}), 400


@app.post("/analyze_pptx")
def analyze_pptx():
    try:
        uploaded_file = request.files.get("presentation")
        if uploaded_file is None or uploaded_file.filename == "":
            raise ValueError("Выберите файл презентации .pptx.")
        if not uploaded_file.filename.lower().endswith(".pptx"):
            raise ValueError("Можно выбрать только файл с расширением .pptx.")

        input_name, input_path = save_upload(uploaded_file, ".pptx", "presentation.pptx")
        try:
            validate_pptx_file(input_path)
            issues = analyze_presentation(input_path)
        finally:
            os.remove(input_path)  # презентацию только читаем, хранить её не нужно
        return jsonify({"issues": issues})
    except ValueError as error:
        return jsonify({"error": str(error)}), 400
    except Exception:
        return jsonify({"error": "Не удалось обработать презентацию. Проверьте, что это неповреждённый файл .pptx."}), 400


def run_server():
    host = os.environ.get("OFORMITEL_HOST", "127.0.0.1")
    port = int(os.environ.get("OFORMITEL_PORT", "5000"))
    server = make_server(host, port, app, threaded=True)
    print("READY " + str(server.server_port), flush=True)
    server.serve_forever()


if __name__ == "__main__":
    run_server()
