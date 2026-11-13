import struct
import zlib
from io import BytesIO

from docx import Document
from docx.enum.style import WD_STYLE_TYPE
from docx.enum.table import WD_TABLE_ALIGNMENT
from docx.enum.text import WD_ALIGN_PARAGRAPH, WD_LINE_SPACING, WD_TAB_ALIGNMENT, WD_TAB_LEADER
from docx.oxml import OxmlElement
from docx.oxml.ns import qn
from docx.shared import Cm, Pt, RGBColor

from fixes import fix_numbers, insert_title_page
from title_page import add_title_page, set_page

# Типы работ: подсказка для пользователя и то, из чего по умолчанию состоит шаблон.
# Ключи совпадают с типами документов в profiles.py
WORK_TYPES = {
    "research": {
        "label": "Исследовательская работа",
        "tip": "Проверяет гипотезу с помощью эксперимента, опроса или наблюдения.\n"
               "Обязательно: титульный лист, оглавление, введение (актуальность, цель, задачи, объект и предмет, "
               "гипотеза, методы), теоретическая и практическая главы, заключение с выводами по задачам, "
               "список литературы.",
        "toc": True, "intro": True, "numbered": True, "conclusion": True, "literature": True,
        "chapters": [
            {"name": "Теоретическая часть", "subsections": ["Название подраздела", "Название подраздела"]},
            {"name": "Практическая часть", "subsections": ["Описание эксперимента", "Результаты"]},
        ],
        "intro_items": [
            "Актуальность: почему тема важна сейчас и кому будет полезен результат.",
            "Цель работы: одно предложение о том, что нужно получить в итоге.",
            "Задачи работы: шаги к цели (изучить, провести, сравнить, сделать выводы).",
            "Объект исследования: что изучается в целом.",
            "Предмет исследования: какая сторона объекта рассматривается.",
            "Гипотеза: предположение, которое проверяется в работе.",
            "Методы исследования: анализ литературы, эксперимент, наблюдение, опрос, сравнение.",
        ],
    },
    "project": {
        "label": "Проектная работа",
        "tip": "Создаёт продукт: программу, устройство, макет, мероприятие.\n"
               "Обязательно: титульный лист, оглавление, введение (проблема, цель, задачи, продукт, для кого он), "
               "анализ существующих решений, описание разработки, тестирование, заключение, список литературы.",
        "toc": True, "intro": True, "numbered": True, "conclusion": True, "literature": True,
        "chapters": [
            {"name": "Анализ проблемы и существующих решений", "subsections": []},
            {"name": "Разработка продукта", "subsections": ["Выбор инструментов", "Описание продукта"]},
            {"name": "Тестирование и результаты", "subsections": []},
        ],
        "intro_items": [
            "Проблема: что не так сейчас и почему это важно.",
            "Цель проекта: какой продукт будет создан.",
            "Задачи проекта: шаги к цели.",
            "Продукт: что получится в итоге и для кого он.",
        ],
    },
    "business_plan": {
        "label": "Бизнес-план",
        "tip": "Подробный план, как идея станет работающим бизнесом (конференция «Шаг в бизнес»).\n"
               "Обязательно: титульный лист, оглавление, описание бизнес-идеи, анализ рынка, ценностное "
               "предложение и целевая аудитория, бизнес-модель, операционный план, финансовая модель, приложения. "
               "Введение и заключение не нужны.",
        "toc": True, "intro": False, "numbered": False, "conclusion": False, "literature": False,
        "appendix": True, "appendix_name": "Приложения",
        "chapters": [
            {"name": "Описание бизнес-идеи", "subsections": [],
             "hint": "Актуальность идеи, цели, задачи и суть бизнес-плана; особенности товара или услуги "
                     "и их преимущества перед конкурентами."},
            {"name": "Анализ состояния рынка", "subsections": [],
             "hint": "Целевой рынок, его структура и тенденции, основные игроки и конкуренция, структура сбыта. "
                     "Покажите, что у продукта есть рынок."},
            {"name": "Ценностное предложение и анализ целевой аудитории", "subsections": [],
             "hint": "Сегменты целевой аудитории и ценностное предложение для каждого из них."},
            {"name": "Бизнес-модель", "subsections": [],
             "hint": "Бизнес-модель по шаблону Остервальдера и Пинье (9 блоков)."},
            {"name": "Операционный план", "subsections": [],
             "hint": "План сбыта, производства и персонала, общие издержки, цели на ближайший период."},
            {"name": "Финансовая модель", "subsections": [],
             "hint": "Таблица доходов, расходов и прибыли: расчёт выручки и продаж, затрат и прибыли с учётом "
                     "налогового режима, показатели эффективности. Расчёты должны быть реалистичными."},
        ],
        "intro_items": [],
    },
    "referat": {
        "label": "Реферат",
        "tip": "Обзор литературы по теме своими словами, без собственного эксперимента.\n"
               "Обязательно: титульный лист, оглавление, введение (актуальность, цель, задачи), "
               "2–3 главы по теме, заключение, список литературы (обычно не меньше 5 источников).",
        "toc": True, "intro": True, "numbered": True, "conclusion": True, "literature": True,
        "chapters": [
            {"name": "Название первой главы", "subsections": []},
            {"name": "Название второй главы", "subsections": []},
        ],
        "intro_items": [
            "Актуальность: почему тема важна.",
            "Цель работы: что нужно выяснить.",
            "Задачи работы: какие вопросы рассматриваются в главах.",
        ],
    },
    "practical": {
        "label": "Практическая работа",
        "tip": "Отчёт о работе, выполненной по заданию (лабораторная, практикум).\n"
               "Обязательно: титульный лист, цель, ход работы, результат, заключение (вывод). "
               "Оглавление и список литературы обычно не нужны.",
        "toc": False, "intro": False, "numbered": False, "conclusion": True, "literature": False,
        "chapters": [
            {"name": "Цель", "subsections": []},
            {"name": "Ход работы", "subsections": []},
            {"name": "Результат", "subsections": []},
        ],
        "intro_items": [],
    },
    "theses": {
        "label": "Тезисы",
        "tip": "Краткое изложение работы на 1–2 страницы: зачем она, что сделано и что получилось. "
               "Их подают на конференцию вместе с работой или вместо неё на первом этапе.\n"
               "Шрифт и интервал для тезисов берутся из профиля (у «Инженеров будущего» — 12 пт, одинарный).",
        "toc": False, "intro": False, "numbered": False, "conclusion": False, "literature": True,
        "chapters": [{"name": "Название работы", "subsections": []}],
        "intro_items": [],
    },
    "report": {
        "label": "Доклад",
        "tip": "Текст выступления: коротко о главном.\n"
               "Титульный лист и оглавление обычно не нужны, в конце — список литературы.",
        "toc": False, "intro": False, "numbered": False, "conclusion": False, "literature": True,
        "chapters": [{"name": "Название доклада", "subsections": []}],
        "intro_items": [],
    },
    "free": {
        "label": "Только оформление",
        "tip": "Любой документ без требований к структуре: программа проверит только оформление.",
        "toc": False, "intro": False, "numbered": False, "conclusion": False, "literature": False,
        "chapters": [{"name": "Раздел", "subsections": []}],
        "intro_items": [],
    },
}

# подсказки при наведении на пункты конструктора
PART_TIPS = {
    "title": "Первая страница работы: школа, название, участники, руководители, город и год. "
             "Номер страницы на ней не ставится.",
    "toc": "Список разделов с номерами страниц. Word заполнит его сам: при открытии файла "
           "согласитесь обновить поля.",
    "intro": "Зачем нужна работа: актуальность, цель, задачи и т.д. Обычно 1–2 страницы. "
             "С введения начинается нумерация страниц.",
    "chapters": "Основная часть. Каждая глава начинается с новой страницы, подразделы идут друг за другом.",
    "numbered": "«1. Теоретическая часть», «1.1. Подраздел». Для практической работы и доклада обычно не нужно.",
    "conclusion": "Выводы по каждой задаче, подтвердилась ли гипотеза, достигнута ли цель.",
    "literature": "Источники по ГОСТ: книги, статьи, сайты с датой обращения. На них ссылаются в тексте.",
    "appendix": "Большие таблицы, код программы, анкеты — то, что мешает читать основной текст.",
    "picture": "Пример рисунка с подписью «Рисунок 1 – Название» под ним и ссылкой (рис. 1) в тексте.",
    "table": "Пример таблицы Word с подписью «Таблица 1 – Название» над ней и ссылкой (табл. 1) в тексте.",
    "hints": "Под каждым заголовком будет подсказка, что там писать. Без галочки — пустые строки.",
}

TITLE_FIELDS = {
    "organization": "Государственное бюджетное общеобразовательное учреждение города Москвы «Школа № ___»",
    "work_title": "Название вашей работы",
    "groups": [{"who": "ученики 10 класса", "school": "ГБОУ Школа № ___", "names": ["Фамилия Имя Отчество"]}],
    "supervisors": [{"position": "учитель ___ ГБОУ Школа № ___", "name": "Фамилия Имя Отчество"}],
    "city": "Москва",
    "year": "2026",
}


def default_plan(work_type):
    kind = WORK_TYPES.get(work_type, WORK_TYPES["research"])
    return {
        "type": work_type,
        "title": "placeholder" if work_type not in ("theses", "report", "free") else "none",
        "toc": kind["toc"],
        "intro": kind["intro"],
        "numbered": kind["numbered"],
        "chapters": kind["chapters"],
        "conclusion": kind["conclusion"],
        "literature": kind["literature"],
        "appendix": kind.get("appendix", False),
        "picture": True,
        "table": True,
        "hints": True,
    }


def remove_theme_fonts(style):
    # "шрифт темы" перебивает обычный шрифт, поэтому его убираем
    fonts = style.element.rPr.rFonts
    for attribute in ["w:asciiTheme", "w:hAnsiTheme", "w:eastAsiaTheme", "w:cstheme"]:
        fonts.attrib.pop(qn(attribute), None)


def set_russian_language(style):
    # без этого Word считает текст английским и подчёркивает каждое слово
    language = OxmlElement("w:lang")
    language.set(qn("w:val"), "ru-RU")
    language.set(qn("w:eastAsia"), "ru-RU")
    style.element.get_or_add_rPr().append(language)


def setup_styles(document, profile):
    text_width = 21 - profile["margin_left_cm"] - profile["margin_right_cm"]

    normal = document.styles["Normal"]
    normal.font.name = profile["font_name"]
    normal.font.size = Pt(profile["font_size_pt"])
    remove_theme_fonts(normal)
    set_russian_language(normal)
    normal_format = normal.paragraph_format
    normal_format.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY
    normal_format.line_spacing_rule = WD_LINE_SPACING.MULTIPLE
    normal_format.line_spacing = profile["line_spacing"]
    normal_format.first_line_indent = Cm(profile["first_line_indent_cm"])
    normal_format.space_before = Pt(0)
    normal_format.space_after = Pt(0)

    for name in ["Heading 1", "Heading 2"]:
        heading = document.styles[name]
        heading.font.name = profile["font_name"]
        heading.font.size = Pt(profile["font_size_pt"])
        heading.font.bold = True
        heading.font.italic = False
        heading.font.color.rgb = RGBColor(0, 0, 0)
        remove_theme_fonts(heading)
        heading_format = heading.paragraph_format
        heading_format.line_spacing_rule = WD_LINE_SPACING.MULTIPLE
        heading_format.line_spacing = profile["line_spacing"]
        heading_format.space_before = Pt(0)
        heading_format.space_after = Pt(0)
        heading_format.keep_with_next = True
    # по правилам весь текст по ширине, заголовки тоже (короткая строка по ширине стоит у левого края)
    document.styles["Heading 1"].paragraph_format.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY
    document.styles["Heading 1"].paragraph_format.first_line_indent = Cm(0)
    document.styles["Heading 2"].paragraph_format.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY
    document.styles["Heading 2"].paragraph_format.first_line_indent = Cm(profile["first_line_indent_cm"])

    # стили строк оглавления: Word возьмёт их, когда будет обновлять оглавление
    for level in [1, 2]:
        toc = document.styles.add_style("toc " + str(level), WD_STYLE_TYPE.PARAGRAPH)
        toc.base_style = normal
        toc_format = toc.paragraph_format
        toc_format.alignment = WD_ALIGN_PARAGRAPH.LEFT
        toc_format.first_line_indent = Cm(0)
        toc_format.left_indent = Cm(0.5 * (level - 1))
        toc_format.tab_stops.add_tab_stop(Cm(text_width), WD_TAB_ALIGNMENT.RIGHT, WD_TAB_LEADER.DOTS)


def add_text(document, text, hints):
    # без подсказок - пустая строка, в которую сразу можно печатать
    return document.add_paragraph(text if hints else "")


def add_chapter(document, text):
    has_content = len(document.paragraphs) > 0
    heading = document.add_paragraph(text, style="Heading 1")
    heading.paragraph_format.page_break_before = has_content
    return heading


def add_field_part(paragraph, kind, instruction=None):
    run = paragraph.add_run()
    if instruction is not None:
        text = OxmlElement("w:instrText")
        text.set(qn("xml:space"), "preserve")
        text.text = instruction
        run._r.append(text)
        return
    mark = OxmlElement("w:fldChar")
    mark.set(qn("w:fldCharType"), kind)
    run._r.append(mark)


def add_toc(document, entries):
    has_content = len(document.paragraphs) > 0
    toc_heading = document.add_paragraph("Оглавление")
    toc_heading.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY
    toc_heading.paragraph_format.first_line_indent = Cm(0)
    toc_heading.paragraph_format.page_break_before = has_content
    toc_heading.runs[0].font.bold = True

    # оглавление - это поле Word: его начало стоит в первой строке, конец - в последней
    for i, (text, level, page) in enumerate(entries):
        entry = document.add_paragraph(style="toc " + str(level))
        if i == 0:
            add_field_part(entry, "begin")
            add_field_part(entry, None, ' TOC \\o "1-2" \\h \\z \\u ')
            add_field_part(entry, "separate")
        entry.add_run(text + "\t" + str(page))
        if i == len(entries) - 1:
            add_field_part(entry, "end")

    # Word обновит оглавление (и номера страниц в нём) при открытии файла
    update = OxmlElement("w:updateFields")
    update.set(qn("w:val"), "true")
    document.settings.element.append(update)


def make_chart_png(width=600, height=340):
    # простая столбчатая диаграмма без сторонних библиотек: картинку собираем по пикселям
    bars = [(70, 120), (190, 200), (310, 160), (430, 250)]
    rows = []
    for y in range(height):
        row = bytearray([0])
        for x in range(width):
            color = (255, 255, 255)
            if x == 40 or y == height - 30:
                color = (60, 60, 60)
            for left, bar_height in bars:
                if left <= x < left + 80 and height - 30 - bar_height <= y < height - 30:
                    color = (90, 120, 160)
            row += bytes(color)
        rows.append(bytes(row))

    def chunk(name, data):
        body = name + data
        return struct.pack(">I", len(data)) + body + struct.pack(">I", zlib.crc32(body) & 0xFFFFFFFF)

    header = struct.pack(">IIBBBBB", width, height, 8, 2, 0, 0, 0)
    png = b"\x89PNG\r\n\x1a\n" + chunk(b"IHDR", header)
    png += chunk(b"IDAT", zlib.compress(b"".join(rows)))
    png += chunk(b"IEND", b"")
    return BytesIO(png)


def add_picture(document):
    document.add_paragraph("На каждый рисунок в тексте должна быть ссылка (рис. 1). Подпись ставится под рисунком.")
    picture = document.add_paragraph()
    picture.alignment = WD_ALIGN_PARAGRAPH.CENTER
    picture.paragraph_format.first_line_indent = Cm(0)
    picture.paragraph_format.keep_with_next = True
    picture.add_run().add_picture(make_chart_png(), width=Cm(12))

    caption = document.add_paragraph("Рисунок 1 – Название рисунка")
    caption.alignment = WD_ALIGN_PARAGRAPH.CENTER
    caption.paragraph_format.first_line_indent = Cm(0)


def add_table(document):
    document.add_paragraph("Таблицы делаются средствами Word: «Вставка» → «Таблица». На каждую таблицу "
                           "в тексте тоже нужна ссылка (табл. 1). Подпись ставится над таблицей.")
    caption = document.add_paragraph("Таблица 1 – Название таблицы")
    caption.alignment = WD_ALIGN_PARAGRAPH.LEFT
    caption.paragraph_format.first_line_indent = Cm(0)
    caption.paragraph_format.keep_with_next = True

    rows = [
        ["Показатель", "До", "После"],
        ["Показатель 1", "10", "15"],
        ["Показатель 2", "20", "28"],
    ]
    table = document.add_table(rows=len(rows), cols=len(rows[0]))
    table.style = "Table Grid"
    table.alignment = WD_TABLE_ALIGNMENT.CENTER
    for r, row in enumerate(rows):
        for c, value in enumerate(row):
            cell = table.cell(r, c)
            cell.text = value
            p = cell.paragraphs[0]
            p.alignment = WD_ALIGN_PARAGRAPH.CENTER
            p.paragraph_format.first_line_indent = Cm(0)
            p.paragraph_format.line_spacing_rule = WD_LINE_SPACING.SINGLE
            if r == 0:
                p.runs[0].font.bold = True
    document.add_paragraph("После таблицы продолжается обычный текст.")


def add_numbered_list(document, items):
    # каждому списку - своя нумерация, иначе второй список продолжит номера первого
    numbering = document.part.numbering_part.element
    style_num_id = document.styles["List Number"].element.pPr.numPr.numId.val
    abstract_id = numbering.num_having_numId(style_num_id).abstractNumId.val
    new_num = numbering.add_num(abstract_id)
    new_num.add_lvlOverride(ilvl=0).add_startOverride(1)

    for item in items:
        p = document.add_paragraph(item, style="List Number")
        num_settings = p._p.get_or_add_pPr().get_or_add_numPr()
        num_settings.get_or_add_ilvl().val = 0
        num_settings.get_or_add_numId().val = new_num.numId


def chapter_names(plan):
    # [(название главы, [названия подразделов], подсказка), ...] с номерами, если они нужны
    result = []
    for number, chapter in enumerate(plan["chapters"], 1):
        name = chapter["name"].strip() or "Глава"
        subsections = []
        for sub_number, sub_name in enumerate(chapter["subsections"], 1):
            sub_name = sub_name.strip() or "Подраздел"
            if plan["numbered"]:
                sub_name = str(number) + "." + str(sub_number) + ". " + sub_name
            subsections.append(sub_name)
        if plan["numbered"]:
            name = str(number) + ". " + name
        result.append((name, subsections, chapter.get("hint", "")))
    return result


def toc_entries(plan):
    # номера страниц примерные: каждый раздел с новой страницы, Word уточнит их при обновлении
    page = 1 + (plan["title"] != "none") + 1
    entries = []
    if plan["intro"]:
        entries.append(("Введение", 1, page))
        page += 1
    for name, subsections, hint in chapter_names(plan):
        entries.append((name, 1, page))
        for sub_name in subsections:
            entries.append((sub_name, 2, page))
        page += 1
    appendix_name = WORK_TYPES.get(plan["type"], {}).get("appendix_name", "Приложение 1")
    for flag, name in [("conclusion", "Заключение"), ("literature", "Список литературы"), ("appendix", appendix_name)]:
        if plan[flag]:
            entries.append((name, 1, page))
            page += 1
    return entries


def create_template(output_path, profile, plan, title_path=None):
    kind = WORK_TYPES.get(plan["type"], WORK_TYPES["research"])
    hints = plan["hints"]
    document = Document()
    set_page(document, profile)
    setup_styles(document, profile)

    if plan["title"] == "placeholder":
        add_title_page(document, TITLE_FIELDS, profile)
    if plan["toc"]:
        add_toc(document, toc_entries(plan))

    if plan["intro"]:
        add_chapter(document, "Введение")
        if hints and kind["intro_items"]:
            for item in kind["intro_items"]:
                document.add_paragraph(item)
        else:
            add_text(document, "Введение: зачем нужна работа, её цель и задачи.", hints)

    chapters = chapter_names(plan)
    for i, (name, subsections, hint) in enumerate(chapters):
        add_chapter(document, name)
        add_text(document, hint or "Текст главы. Каждая глава начинается с новой страницы.", hints)
        for sub_name in subsections:
            document.add_paragraph(sub_name, style="Heading 2")
            add_text(document, "Текст подраздела.", hints)
        if i == len(chapters) - 1:
            if plan["picture"]:
                add_picture(document)
            if plan["table"]:
                add_table(document)

    if plan["conclusion"]:
        add_chapter(document, "Заключение")
        add_text(document, "Выводы по каждой задаче: подтвердилась ли гипотеза, достигнута ли цель.", hints)
    if plan["literature"]:
        add_chapter(document, "Список литературы")
        if hints:
            add_numbered_list(document, [
                "Фамилия И. О. Название книги. – Москва : Издательство, 2023. – 200 с.",
                "Фамилия И. О. Название статьи // Название журнала. – 2024. – № 3. – С. 10–15.",
                "Название сайта. – URL: https://example.ru (дата обращения: 01.11.2026).",
            ])
        else:
            add_numbered_list(document, [""])
    if plan["appendix"]:
        add_chapter(document, kind.get("appendix_name", "Приложение 1"))
        add_text(document, "Материалы, которые мешают читать основной текст: большие таблицы, код, анкеты.", hints)

    if title_path:
        insert_title_page(document, title_path)
    fix_numbers(document, profile)
    document.save(output_path)
    return output_path
