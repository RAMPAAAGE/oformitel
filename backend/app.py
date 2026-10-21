import os
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
    validate_docx_file,
)
from pptx_engine import analyze_presentation, validate_pptx_file
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
    copy_default_profiles()


def copy_default_profiles():
    if os.path.abspath(PROFILES_FOLDER) == os.path.abspath(DEFAULT_PROFILES_FOLDER):
        return
    if not os.path.isdir(DEFAULT_PROFILES_FOLDER):
        return
    for file_name in os.listdir(DEFAULT_PROFILES_FOLDER):
        if not file_name.lower().endswith(".json"):
            continue
        source_path = os.path.join(DEFAULT_PROFILES_FOLDER, file_name)
        destination_path = os.path.join(PROFILES_FOLDER, file_name)
        if os.path.isfile(source_path) and not os.path.exists(destination_path):
            shutil.copyfile(source_path, destination_path)


def get_free_file_name(folder, file_name):
    name, extension = os.path.splitext(file_name)
    free_name = file_name
    number = 2
    while os.path.exists(os.path.join(folder, free_name)):
        free_name = name + "_" + str(number) + extension
        number += 1
    return free_name


def title_page_label(file_name):
    file_time = os.path.getmtime(os.path.join(OUTPUTS_FOLDER, file_name))
    created = time.strftime("%d.%m %H:%M", time.localtime(file_time))
    return file_name + " (создан " + created + ")"


def list_title_pages():
    # созданные на соседней вкладке титульные листы, свежие первыми
    result = []
    for file_name in os.listdir(OUTPUTS_FOLDER):
        if file_name.startswith("titulnyi_list") and file_name.lower().endswith(".docx"):
            file_time = os.path.getmtime(os.path.join(OUTPUTS_FOLDER, file_name))
            result.append({"file": file_name, "label": title_page_label(file_name), "time": file_time})
    result.sort(key=lambda item: item["time"], reverse=True)
    return result


def show_main_page(active_tab="document", error="", message=""):
    return render_template(
        "index.html",
        profiles=list_profiles(PROFILES_FOLDER),
        document_types=DOCUMENT_TYPES,
        title_pages=list_title_pages(),
        active_tab=active_tab,
        error=error,
        message=message,
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


def get_title_page_values(form):
    return {
        "organization": form.get("organization", "").strip(),
        "work_title": form.get("work_title", "").strip(),
        "participants": form.get("participants", "").strip(),
        "supervisor": form.get("supervisor", "").strip(),
        "city": form.get("city", "").strip(),
        "year": form.get("year", "").strip(),
    }


create_work_folders()


@app.get("/")
def index():
    message = ""
    if request.args.get("profile_created") == "1":
        message = "Новый профиль сохранён. Теперь его можно выбрать для документа."
    return show_main_page(message=message)


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

        profile_file = request.form.get("profile_file", "")
        profile_path = os.path.join(PROFILES_FOLDER, profile_file)
        if profile_file != os.path.basename(profile_file) or not os.path.isfile(profile_path):
            raise ValueError("Выбранный профиль не найден.")

        # титульный лист для вставки в начало документа (необязательно)
        title_path = None
        title_file = request.form.get("title_file", "")
        if title_file:
            title_path = os.path.join(OUTPUTS_FOLDER, title_file)
            if title_file != os.path.basename(title_file) or not title_file.startswith("titulnyi_list") \
                    or not os.path.isfile(title_path):
                raise ValueError("Выбранный титульный лист не найден.")

        input_name, input_path = save_upload(uploaded_file, ".docx", "document.docx")
        validate_docx_file(input_path)

        skip_title = request.form.get("skip_title") == "1"
        if not title_path and not skip_title and not analyze_layout(Document(input_path))["title"]:
            return jsonify({
                "no_title": True,
                "message": "Титульный лист в документе не найден. Создайте его на вкладке «Титульный лист» "
                           "и выберите в списке, либо оформите документ без него.",
            })

        output_name = get_free_file_name(OUTPUTS_FOLDER, "formatted_" + input_name)
        output_path = os.path.join(OUTPUTS_FOLDER, output_name)
        result = check_and_format_document(input_path, profile_path, output_path, title_path)

        warnings = result["warnings"]
        if skip_title:
            warnings = [w for w in warnings if not w.startswith("Титульный лист не найден")]

        if result["is_gost"]:
            return jsonify({
                "gost": True,
                "message": "Документ оформлен по ГОСТу. Исправления не требуются.",
                "warnings": warnings,
            })

        message = ""
        size_is_ok, size_mb = check_file_size(output_path)
        if not size_is_ok:
            message = "Внимание: размер файла " + str(round(size_mb, 1)) + " МБ, это больше допустимых 5 МБ."

        return jsonify({
            "gost": False,
            "url": url_for("static", filename=output_name),
            "criteria": [{"n": n, "name": CRITERIA_NAMES[n]} for n in result["failed"]],
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


@app.post("/new_profile")
def new_profile():
    try:
        values = get_profile_values(request.form)
        create_profile(profiles_folder=PROFILES_FOLDER, **values)
        return redirect(url_for("index", profile_created="1"))
    except ValueError as error:
        return show_main_page(active_tab="document", error=str(error)), 400
    except Exception:
        return show_main_page(
            active_tab="document",
            error="Не удалось сохранить профиль. Проверьте заполненные поля.",
        ), 400


@app.post("/create_title_page")
def create_title_page():
    try:
        fields = get_title_page_values(request.form)
        output_name = get_free_file_name(OUTPUTS_FOLDER, "titulnyi_list.docx")
        output_path = os.path.join(OUTPUTS_FOLDER, output_name)
        create_title_page_document(fields, output_path)
        return jsonify({
            "file": output_name,
            "label": title_page_label(output_name),
            "url": url_for("static", filename=output_name),
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
        validate_pptx_file(input_path)
        return jsonify({"issues": analyze_presentation(input_path)})
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
