import io
import json
import os
import sys
import tempfile
import unittest

from docx import Document
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.shared import Cm, Pt, RGBColor
from pptx import Presentation


PROJECT_FOLDER = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
BACKEND_FOLDER = os.path.join(PROJECT_FOLDER, "backend")
sys.path.insert(0, BACKEND_FOLDER)

os.environ["OFORMITEL_DATA_FOLDER"] = tempfile.mkdtemp(prefix="oformitel-tests-")

from app import OUTPUTS_FOLDER, PROFILES_FOLDER, app, copy_default_profiles  # noqa: E402
from profiles import create_profile  # noqa: E402
from template import WORK_TYPES, default_plan  # noqa: E402
from engine import check_and_format_document  # noqa: E402
from fixes import fix_toc  # noqa: E402
from checks import missing_references  # noqa: E402
from layout import analyze_layout  # noqa: E402
from pptx_engine import analyze_presentation  # noqa: E402


def output_file(url):
    # "/download/formatted_x.docx?name=..." -> "formatted_x.docx"
    return os.path.basename(url.split("?")[0])


def make_docx():
    file_data = io.BytesIO()
    document = Document()
    document.add_paragraph("Проверочный документ")
    document.add_paragraph("Текст для проверки обработки файла.")
    document.save(file_data)
    file_data.seek(0)
    return file_data


def make_badly_formatted_docx():
    file_data = io.BytesIO()
    document = Document()
    section = document.sections[0]
    section.top_margin = Cm(1)
    section.bottom_margin = Cm(1)
    section.left_margin = Cm(1)
    section.right_margin = Cm(1)
    paragraph = document.add_paragraph("Документ для проверки пользовательского профиля.")
    run = paragraph.runs[0]
    run.font.name = "Arial"
    run.font.size = Pt(10)
    paragraph.paragraph_format.line_spacing = 1
    paragraph.paragraph_format.first_line_indent = Cm(0)
    document.save(file_data)
    file_data.seek(0)
    return file_data


def make_pptx():
    file_data = io.BytesIO()
    presentation = Presentation()
    slide = presentation.slides.add_slide(presentation.slide_layouts[0])
    slide.shapes.title.text = "Проверочный слайд."
    presentation.save(file_data)
    file_data.seek(0)
    return file_data


def make_pptx_with_traps(path):
    presentation = Presentation()
    slide = presentation.slides.add_slide(presentation.slide_layouts[0])
    slide.shapes.title.text = "Заголовок с точкой."
    slide.placeholders[1].text = "а" * 401
    presentation.save(path)


def make_toc_docx():
    document = Document()
    document.add_paragraph("Оглавление")
    entry = document.add_paragraph("Введение........................................3")
    entry.paragraph_format.left_indent = Cm(0.8)
    entry = document.add_paragraph("Заключение.....................................10")
    entry.paragraph_format.left_indent = Cm(0.8)
    document.add_paragraph()
    document.add_paragraph("Введение")
    document.add_paragraph("Текст введения длиной больше ста символов нужен только для того, чтобы анализатор распознал начало основного текста документа.")
    return document


class SmokeTests(unittest.TestCase):
    def setUp(self):
        self.client = app.test_client()

    def test_main_page_opens(self):
        response = self.client.get("/")

        self.assertEqual(response.status_code, 200)

    def test_title_page_is_created(self):
        response = self.client.post(
            "/create_title_page",
            data={
                "organization": "Тестовая организация",
                "work_title": "Проверочная работа",
                "group_0_who": "ученик 10 класса",
                "group_0_name": "Иванов И. И.",
                "supervisor_position": "учитель",
                "supervisor_name": "Петров П. П.",
                "city": "Москва",
                "year": "2026",
                "profile_file": "01_inzhenery_budushchego.json",
            },
        )

        self.assertEqual(response.status_code, 200)
        self.assertIn("file", response.get_json())

    def test_docx_is_processed(self):
        response = self.client.post(
            "/apply",
            data={
                "document": (make_docx(), "test.docx"),
                "profile_file": "01_inzhenery_budushchego.json",
                "skip_title": "1",
            },
            content_type="multipart/form-data",
        )

        self.assertEqual(response.status_code, 200)
        self.assertIn("correct", response.get_json())

    def test_custom_profile_is_saved_and_applied(self):
        profile_values = {
            "name": "Проверочный профиль",
            "font_name": "Arial",
            "font_size_pt": "11",
            "line_spacing": "2",
            "first_line_indent_cm": "0.5",
            "margin_top_cm": "1.1",
            "margin_bottom_cm": "1.2",
            "margin_left_cm": "1.3",
            "margin_right_cm": "1.4",
            "document_type": "research",
        }
        response = self.client.post("/save_profile", data=dict(profile_values, profile_file=""))
        self.assertEqual(response.status_code, 302)

        created_profiles = []
        for file_name in os.listdir(PROFILES_FOLDER):
            if not file_name.startswith("profile_") or not file_name.endswith(".json"):
                continue
            path = os.path.join(PROFILES_FOLDER, file_name)
            with open(path, "r", encoding="utf-8") as profile_file:
                profile = json.load(profile_file)
            if profile.get("name") == profile_values["name"]:
                created_profiles.append(file_name)
        self.assertEqual(len(created_profiles), 1)
        with open(os.path.join(PROFILES_FOLDER, created_profiles[0]), "r", encoding="utf-8") as profile_file:
            saved_profile = json.load(profile_file)
        self.assertEqual(saved_profile["document_type"], "research")
        self.assertTrue(saved_profile["structure"]["toc"])

        response = self.client.post(
            "/apply",
            data={
                "document": (make_badly_formatted_docx(), "custom-profile.docx"),
                "profile_file": created_profiles[0],
                "skip_title": "1",
            },
            content_type="multipart/form-data",
        )
        self.assertEqual(response.status_code, 200)
        result = response.get_json()
        self.assertIn("url", result)
        self.assertEqual(result["download_name"], "оформлено_custom-profile.docx")
        self.assertTrue(any("обязательный титульный лист" in warning for warning in result["warnings"]))
        self.assertTrue(any("обязательное оглавление" in warning for warning in result["warnings"]))

        output_name = output_file(result["url"])
        output = Document(os.path.join(OUTPUTS_FOLDER, output_name))
        section = output.sections[0]
        self.assertAlmostEqual(section.top_margin.cm, 1.1, places=1)
        self.assertAlmostEqual(section.bottom_margin.cm, 1.2, places=1)
        self.assertAlmostEqual(section.left_margin.cm, 1.3, places=1)
        self.assertAlmostEqual(section.right_margin.cm, 1.4, places=1)
        paragraph = output.paragraphs[0]
        run = paragraph.runs[0]
        self.assertEqual(run.font.name, "Arial")
        self.assertEqual(run.font.size.pt, 11)
        self.assertAlmostEqual(paragraph.paragraph_format.line_spacing, 2)
        self.assertAlmostEqual(paragraph.paragraph_format.first_line_indent.cm, 0.5, places=1)

    def test_invalid_custom_profile_values_are_rejected(self):
        response = self.client.post(
            "/save_profile",
            data={
                "name": "Невалидный профиль",
                "font_name": "Arial",
                "font_size_pt": "-10",
                "line_spacing": "0",
                "first_line_indent_cm": "-1",
                "margin_top_cm": "-1",
                "margin_bottom_cm": "-1",
                "margin_left_cm": "-1",
                "margin_right_cm": "-1",
            },
        )

        self.assertEqual(response.status_code, 400)
        self.assertIn("Размер шрифта", response.get_data(as_text=True))

    def test_pptx_is_analyzed(self):
        response = self.client.post(
            "/analyze_pptx",
            data={"presentation": (make_pptx(), "test.pptx")},
            content_type="multipart/form-data",
        )

        self.assertEqual(response.status_code, 200)
        self.assertEqual(len(response.get_json()["issues"]), 1)

    def test_title_page_has_each_participant_on_own_line(self):
        response = self.client.post(
            "/create_title_page",
            data={
                "organization": "Тестовая организация",
                "work_title": "Проверочная работа",
                "group_0_who": "ученики 10 класса",
                "group_0_school": "Школа № 1",
                "group_0_name": ["Иванов И. И.", "", "Сидоров С. С."],
                "group_3_who": "ученица 9 класса",
                "group_3_name": "Петрова Е. Е.",
                "supervisor_position": ["учитель физики", "учитель химии"],
                "supervisor_name": ["Петров П. П.", "Смирнова А. А."],
                "city": "Москва",
                "year": "2026",
                "profile_file": "01_inzhenery_budushchego.json",
            },
        )

        self.assertEqual(response.status_code, 200)
        file_path = os.path.join(OUTPUTS_FOLDER, response.get_json()["file"])
        lines = [paragraph.text for paragraph in Document(file_path).paragraphs]
        start = lines.index("Участники:")
        self.assertEqual(lines[start:start + 12], [
            "Участники:", "ученики 10 класса", "Школа №\u00a01", "Иванов И. И.", "Сидоров С. С.",
            "ученица 9 класса", "Петрова Е. Е.", "",
            "Научные руководители:", "учитель физики", "Петров П. П.", "учитель химии",
        ])

    def test_justify_and_table_caption_are_fixed(self):
        document = Document()
        document.add_paragraph("Обычный абзац текста, выровненный по левому краю (табл. 1).")
        table = document.add_table(rows=2, cols=2)
        table.cell(0, 0).text = "Ячейка"
        with tempfile.TemporaryDirectory() as folder:
            source = os.path.join(folder, "source.docx")
            result = os.path.join(folder, "result.docx")
            document.save(source)
            profile = os.path.join(PROFILES_FOLDER, "01_inzhenery_budushchego.json")

            report = check_and_format_document(source, profile, result)

            self.assertIn(11, report["failed"])
            self.assertIn(12, report["failed"])
            fixed = Document(result)
            self.assertEqual(fixed.paragraphs[0].alignment, WD_ALIGN_PARAGRAPH.JUSTIFY)
            self.assertEqual(fixed.paragraphs[1].text, "Таблица 1")
            self.assertEqual(check_and_format_document(result, profile, result + "2")["failed"], [])

    def test_missing_reference_to_picture_is_reported(self):
        document = Document()
        document.add_paragraph("Текст без ссылки на таблицу.")
        document.add_paragraph("Таблица 1 - Результаты")
        document.add_table(rows=1, cols=1)

        self.assertEqual(missing_references(document, analyze_layout(document)), ["таблица 1"])

    def test_builtin_profile_can_be_edited_and_reset(self):
        values = {
            "profile_file": "02_gost_7_32.json",
            "name": "ГОСТ (мой)",
            "font_name": "Times New Roman",
            "font_size_pt": "12",
            "line_spacing": "1.5",
            "first_line_indent_cm": "1.25",
            "margin_top_cm": "2",
            "margin_bottom_cm": "2",
            "margin_left_cm": "3",
            "margin_right_cm": "1.5",
            "document_type": "research",
        }
        path = os.path.join(PROFILES_FOLDER, "02_gost_7_32.json")

        self.assertEqual(self.client.post("/save_profile", data=values).status_code, 302)
        with open(path, encoding="utf-8") as file:
            self.assertEqual(json.load(file)["font_size_pt"], 12)
        copy_default_profiles()
        with open(path, encoding="utf-8") as file:
            self.assertEqual(json.load(file)["font_size_pt"], 12)

        self.assertEqual(self.client.post("/reset_profile", data={"profile_file": "02_gost_7_32.json"}).status_code, 302)
        with open(path, encoding="utf-8") as file:
            self.assertEqual(json.load(file)["font_size_pt"], 14)

        self.assertEqual(self.client.post("/delete_profile", data={"profile_file": "02_gost_7_32.json"}).status_code, 400)
        self.assertTrue(os.path.isfile(path))

    def test_new_profile_is_created_and_selected(self):
        before = set(os.listdir(PROFILES_FOLDER))
        response = self.client.post("/save_profile", data={
            "profile_file": "", "name": "Моя конференция", "font_name": "Times New Roman", "font_size_pt": "14",
            "line_spacing": "1.5", "first_line_indent_cm": "1.25", "margin_top_cm": "2", "margin_bottom_cm": "2",
            "margin_left_cm": "3", "margin_right_cm": "1.5", "document_type": "research",
        })
        created = set(os.listdir(PROFILES_FOLDER)) - before
        self.assertEqual(len(created), 1)
        new_file = created.pop()
        self.assertIn("select=" + new_file, response.headers["Location"])

        page = self.client.get(response.headers["Location"]).get_data(as_text=True)
        self.assertIn("Профиль «Моя конференция» создан и выбран.", page)
        self.assertIn('value="' + new_file + '">Моя конференция', page)
        os.remove(os.path.join(PROFILES_FOLDER, new_file))

    def test_my_template_is_saved_listed_and_deleted(self):
        plan = default_plan("research")
        plan["chapters"] = [{"name": "Компьютерные игры", "subsections": ["История игр"]}]
        response = self.client.post("/save_my_template", data={"name": "Про игры", "plan": json.dumps(plan)})
        saved = response.get_json()
        self.assertEqual(saved["plan"]["chapters"][0]["name"], "Компьютерные игры")
        self.assertIn("Мой шаблон: Про игры", self.client.get("/").get_data(as_text=True))

        response = self.client.post("/delete_my_template", data={"template_file": saved["file"]})
        self.assertTrue(response.get_json()["deleted"])
        self.assertNotIn("Мой шаблон: Про игры", self.client.get("/").get_data(as_text=True))

    def test_user_profile_can_be_deleted(self):
        path = create_profile(
            profiles_folder=PROFILES_FOLDER, name="Удаляемый", font_name="Arial", font_size_pt=12, line_spacing=1,
            first_line_indent_cm=1, margin_top_cm=2, margin_bottom_cm=2, margin_left_cm=2, margin_right_cm=2,
        )
        response = self.client.post("/delete_profile", data={"profile_file": os.path.basename(path)})

        self.assertEqual(response.status_code, 302)
        self.assertFalse(os.path.exists(path))

    def test_every_template_passes_all_checks(self):
        profile = os.path.join(PROFILES_FOLDER, "01_inzhenery_budushchego.json")
        for work_type in WORK_TYPES:
            for hints in [True, False]:
                plan = default_plan(work_type)
                plan["hints"] = hints
                plan["appendix"] = True
                response = self.client.post("/create_template", data={
                    "plan": json.dumps(plan), "profile_file": "01_inzhenery_budushchego.json",
                })
                self.assertEqual(response.status_code, 200)
                path = os.path.join(OUTPUTS_FOLDER, output_file(response.get_json()["url"]))
                with tempfile.TemporaryDirectory() as folder:
                    report = check_and_format_document(path, profile, os.path.join(folder, "out.docx"))
                self.assertTrue(report["is_correct"], (work_type, hints, report))
                if work_type in ("research", "project", "referat"):
                    self.assertEqual(report["warnings"], [], (work_type, hints))

    def test_business_plan_template_fits_shag_v_biznes(self):
        plan = default_plan("business_plan")
        response = self.client.post("/create_template", data={
            "plan": json.dumps(plan), "profile_file": "04_shag_v_biznes.json",
        })
        path = os.path.join(OUTPUTS_FOLDER, output_file(response.get_json()["url"]))
        profile = os.path.join(PROFILES_FOLDER, "04_shag_v_biznes.json")
        with tempfile.TemporaryDirectory() as folder:
            report = check_and_format_document(path, profile, os.path.join(folder, "out.docx"))
        self.assertTrue(report["is_correct"], report)
        self.assertEqual(report["warnings"], [])

        document = Document(path)
        layout = analyze_layout(document)
        self.assertEqual(layout["first_section"].text, "Описание бизнес-идеи")

    def test_letter_page_becomes_a4(self):
        document = Document()
        document.add_paragraph("Текст на американском листе Letter.")
        with tempfile.TemporaryDirectory() as folder:
            source = os.path.join(folder, "source.docx")
            document.save(source)
            profile = os.path.join(PROFILES_FOLDER, "01_inzhenery_budushchego.json")
            report = check_and_format_document(source, profile, os.path.join(folder, "out.docx"))
            section = Document(os.path.join(folder, "out.docx")).sections[0]
        self.assertIn(1, report["failed"])
        self.assertAlmostEqual(section.page_width.cm, 21, places=1)
        self.assertAlmostEqual(section.page_height.cm, 29.7, places=1)

    def test_inserted_title_page_keeps_its_sizes(self):
        response = self.client.post(
            "/create_title_page",
            data={
                "organization": "Тестовая организация",
                "work_title": "Проверочная работа",
                "group_0_who": "ученики 10 класса",
                "group_0_name": ["Иванов И. И.", "Сидоров С. С.", "Петров П. П."],
                "supervisor_position": "учитель",
                "supervisor_name": "Смирнова А. А.",
                "city": "Москва",
                "year": "2026",
                "profile_file": "01_inzhenery_budushchego.json",
            },
        )
        title_file = response.get_json()["file"]
        response = self.client.post(
            "/apply",
            data={
                "document": (make_badly_formatted_docx(), "work.docx"),
                "profile_file": "01_inzhenery_budushchego.json",
                "title_file": title_file,
            },
            content_type="multipart/form-data",
        )
        output = Document(os.path.join(OUTPUTS_FOLDER, output_file(response.get_json()["url"])))

        title = [p for p in output.paragraphs if p.text == "ПРОВЕРОЧНАЯ РАБОТА"][0]
        self.assertEqual(title.runs[0].font.size.pt, 16)
        self.assertTrue(title.runs[0].font.bold)

    def test_underline_color_and_spaces_are_fixed(self):
        document = Document()
        paragraph = document.add_paragraph("Текст с подчёркиванием и цветом.")
        paragraph.runs[0].font.underline = True
        paragraph.runs[0].font.color.rgb = RGBColor(255, 0, 0)
        paragraph.paragraph_format.space_after = Pt(12)
        paragraph.paragraph_format.left_indent = Cm(1)
        with tempfile.TemporaryDirectory() as folder:
            source = os.path.join(folder, "source.docx")
            document.save(source)
            profile = os.path.join(PROFILES_FOLDER, "01_inzhenery_budushchego.json")

            report = check_and_format_document(source, profile, os.path.join(folder, "out.docx"))

            self.assertIn(13, report["failed"])
            self.assertIn(14, report["failed"])
            fixed = Document(os.path.join(folder, "out.docx")).paragraphs[0]
            self.assertFalse(fixed.runs[0].font.underline)
            self.assertEqual(fixed.runs[0].font.color.rgb, RGBColor(0, 0, 0))
            self.assertEqual(fixed.paragraph_format.space_after.pt, 0)
            self.assertEqual(fixed.paragraph_format.left_indent.cm, 0)

    def test_empty_title_page_is_rejected(self):
        response = self.client.post("/create_title_page", data={})

        self.assertEqual(response.status_code, 400)

    def test_stylesheet_is_available_to_packaged_app(self):
        response = self.client.get("/style.css")

        self.assertEqual(response.status_code, 200)
        self.assertIn(b"--accent", response.data)
        response.close()

    def test_toc_trap_keeps_page_number_on_same_line(self):
        document = make_toc_docx()
        fix_toc(document)

        layout = analyze_layout(document)
        entries = [entry for entry in layout["toc"] if entry.text.strip()]
        self.assertEqual([entry.text for entry in entries], ["Введение\t3", "Заключение\t10"])
        self.assertTrue(all(entry.paragraph_format.left_indent.cm == 0 for entry in entries))

    def test_uploaded_file_is_deleted_after_processing(self):
        from app import UPLOADS_FOLDER

        self.client.post(
            "/apply",
            data={
                "document": (make_docx(), "delete-me.docx"),
                "profile_file": "01_inzhenery_budushchego.json",
                "skip_title": "1",
            },
            content_type="multipart/form-data",
        )

        left_files = [name for name in os.listdir(UPLOADS_FOLDER) if name.startswith("delete-me")]
        self.assertEqual(left_files, [])

    def test_paragraph_with_missing_style_does_not_crash(self):
        # бывают документы, где у абзаца указан стиль, которого нет в файле
        document = Document()
        paragraph = document.add_paragraph("Абзац со стилем, которого нет в документе.")
        paragraph._p.get_or_add_pPr().get_or_add_pStyle().val = "NoSuchStyle"
        file_data = io.BytesIO()
        document.save(file_data)
        file_data.seek(0)

        response = self.client.post(
            "/apply",
            data={
                "document": (file_data, "missing-style.docx"),
                "profile_file": "01_inzhenery_budushchego.json",
                "skip_title": "1",
            },
            content_type="multipart/form-data",
        )

        self.assertEqual(response.status_code, 200)
        self.assertNotIn("error", response.get_json())

    def test_correct_document_message_names_the_profile(self):
        with tempfile.TemporaryDirectory() as folder:
            first = os.path.join(folder, "first.docx")
            second = os.path.join(folder, "second.docx")
            with open(first, "wb") as file:
                file.write(make_badly_formatted_docx().read())
            profile = os.path.join(PROFILES_FOLDER, "01_inzhenery_budushchego.json")
            check_and_format_document(first, profile, second)  # первый раз исправляем

            with open(second, "rb") as file:
                response = self.client.post(
                    "/apply",
                    data={
                        "document": (io.BytesIO(file.read()), "second.docx"),
                        "profile_file": "01_inzhenery_budushchego.json",
                        "skip_title": "1",
                    },
                    content_type="multipart/form-data",
                )

        result = response.get_json()
        self.assertTrue(result["correct"])
        self.assertIn("Инженеры будущего", result["message"])

    def test_punctuation_keeps_file_extensions_and_tag_names(self):
        from punctuation import find_edits

        good = "Установщик бывает двух видов: .exe для Windows и тег w:p для абзаца."
        edits, _ = find_edits(good, False, True, None)
        self.assertEqual(edits, [])

        bad = "Это  ошибка ,и тут тоже:пропущен пробел"
        edits, _ = find_edits(bad, False, True, None)
        self.assertTrue(len(edits) >= 3)

    def test_gost_profile_explains_margin_change(self):
        response = self.client.post(
            "/apply",
            data={
                "document": (make_badly_formatted_docx(), "gost.docx"),
                "profile_file": "02_gost_7_32.json",
                "skip_title": "1",
            },
            content_type="multipart/form-data",
        )
        criteria = response.get_json()["criteria"]
        margins = [c for c in criteria if c["n"] == 1][0]
        self.assertEqual(margins["details"], "формат листа → A4, верхнее 1 → 2 см, нижнее 1 → 2 см, левое 1 → 3 см, правое 1 → 1,5 см")

    def test_presentation_traps_are_reported(self):
        with tempfile.TemporaryDirectory() as folder:
            path = os.path.join(folder, "traps.pptx")
            make_pptx_with_traps(path)
            issues = analyze_presentation(path)

        issue_types = [issue["issue_type"] for issue in issues]
        self.assertIn("Точка в конце заголовка", issue_types)
        self.assertIn("Слишком длинный текст на слайде", issue_types)


if __name__ == "__main__":
    unittest.main()
