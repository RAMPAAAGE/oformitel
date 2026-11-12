import math
import re

from docx import Document
from docx.enum.text import WD_ALIGN_PARAGRAPH, WD_LINE_SPACING
from docx.oxml import OxmlElement
from docx.oxml.ns import qn
from docx.shared import Cm, Pt

# Сколько строк размером 14 пт с одинарным интервалом помещается на страницу A4
# при полях 2 см. Берём с запасом, чтобы город и год не уехали на вторую страницу.
LINES_ON_PAGE = 42
# примерно столько букв помещается в одну строку при ширине текста 17 см
LETTERS_IN_LINE = 60
# участники и руководители пишутся колонкой у правого края, как в образце конференции
PEOPLE_COLUMN_CM = 7.75
LETTERS_IN_PEOPLE_LINE = 30


def add_paragraph(document, font_size_pt):
    # интервал и размер задаём у каждой строки, даже пустой: при вставке в другой документ
    # иначе возьмутся настройки его стиля (например, интервал 1,5) и лист не влезет на страницу
    paragraph = document.add_paragraph()
    paragraph.paragraph_format.line_spacing_rule = WD_LINE_SPACING.SINGLE
    paragraph.paragraph_format.space_before = Pt(0)
    paragraph.paragraph_format.space_after = Pt(0)
    paragraph.paragraph_format.first_line_indent = Cm(0)
    mark_size = OxmlElement("w:sz")
    mark_size.set(qn("w:val"), str(int(font_size_pt * 2)))
    mark = OxmlElement("w:rPr")
    mark.append(mark_size)
    paragraph._p.get_or_add_pPr().append(mark)
    return paragraph


def join_initials(match):
    space = "\u00a0" if match.group(2) else ""
    return match.group(1) + space + match.group(3) + "\u00a0"


def keep_together(text):
    # неразрывный пробел, чтобы "№" и инициалы не отрывались при переносе: "№ 654", "А.Д. Фридмана"
    text = re.sub(r"№\s*", "№\u00a0", text)
    return re.sub(r"\b([А-ЯЁ]\.)(\s*)([А-ЯЁ]\.)\s+(?=[А-ЯЁ])", join_initials, text)


def add_styled_paragraph(document, text, alignment, bold, font_name, font_size_pt, make_uppercase=False):
    paragraph = add_paragraph(document, font_size_pt)
    paragraph.alignment = alignment
    text = keep_together(text)

    if make_uppercase:
        text = text.upper()

    run = paragraph.add_run(text)
    run.font.name = font_name
    run.font.size = Pt(font_size_pt)
    run.font.bold = bold

    return paragraph


def add_empty_paragraphs(document, how_many, font_size_pt):
    for _ in range(how_many):
        add_paragraph(document, font_size_pt)


def count_lines(text, letters_in_line=LETTERS_IN_LINE):
    return max(1, math.ceil(len(text) / letters_in_line))


def check_required_fields(fields):
    for field_name in ["organization", "work_title", "city", "year"]:
        if not fields.get(field_name, "").strip():
            raise ValueError("Заполните все поля титульного листа.")
    if count_participants(fields) == 0:
        raise ValueError("Укажите хотя бы одного участника.")
    if not fields.get("supervisors"):
        raise ValueError("Укажите хотя бы одного руководителя.")


def count_participants(fields):
    count = 0
    for group in fields.get("groups", []):
        count += len(group["names"])
    return count


def people_lines(fields):
    # Участники:
    # ученики 10 класса        <- одна строка на группу (один класс или одна школа)
    # ГБОУ Школа № 654
    # Иванов Иван Иванович
    # ...
    # (пустая строка)
    # Научный руководитель:
    # учитель информатики ...
    # Петрова Мария Сергеевна
    lines = ["Участники:" if count_participants(fields) > 1 else "Участник:"]
    for group in fields["groups"]:
        if group["who"]:
            lines.append(group["who"])
        if group["school"]:
            lines.append(group["school"])
        lines += group["names"]
    lines.append("")
    lines.append("Научные руководители:" if len(fields["supervisors"]) > 1 else "Научный руководитель:")
    for supervisor in fields["supervisors"]:
        if supervisor["position"]:
            lines.append(supervisor["position"])
        lines.append(supervisor["name"])
    return lines


def set_page(document, profile):
    section = document.sections[0]
    # python-docx по умолчанию создаёт американский лист Letter, а нужен A4
    section.page_width = Cm(21)
    section.page_height = Cm(29.7)
    section.top_margin = Cm(profile["margin_top_cm"])
    section.bottom_margin = Cm(profile["margin_bottom_cm"])
    section.left_margin = Cm(profile["margin_left_cm"])
    section.right_margin = Cm(profile["margin_right_cm"])


def create_title_page_document(fields, output_path, profile):
    check_required_fields(fields)
    document = Document()
    set_page(document, profile)
    add_title_page(document, fields, profile)
    # по названию работы титульный лист потом узнаётся в списках программы
    document.core_properties.title = fields["work_title"]
    document.save(output_path)
    return output_path


def add_title_page(document, fields, profile):
    font = profile["font_name"]
    text_size = profile["font_size_pt"]
    title_size = text_size + 2
    text_width = 21 - profile["margin_left_cm"] - profile["margin_right_cm"]

    people = people_lines(fields)

    organization_lines = count_lines(fields["organization"])
    # название набрано заглавными и крупнее, поэтому в строку влезает меньше букв
    title_lines = count_lines(fields["work_title"], 45)
    people_line_count = 0
    for line in people:
        people_line_count += count_lines(line, LETTERS_IN_PEOPLE_LINE)

    lines_on_page = LINES_ON_PAGE * 14 // text_size

    # название - на трети высоты листа, как в образце
    lines_before_title = max(2, lines_on_page // 3 - organization_lines)
    # участники и руководители - внизу справа, над городом
    lines_after_people = 3
    lines_after_title = (lines_on_page - organization_lines - lines_before_title - title_lines
                         - people_line_count - lines_after_people - 1)
    if lines_after_title < 2:
        # много участников: поднимаем название выше
        lines_before_title = max(1, lines_before_title + lines_after_title - 2)
        lines_after_title = 2

    add_styled_paragraph(document, fields["organization"], WD_ALIGN_PARAGRAPH.CENTER, True, font, text_size)
    add_empty_paragraphs(document, lines_before_title, text_size)
    add_styled_paragraph(document, fields["work_title"], WD_ALIGN_PARAGRAPH.CENTER, True, font, title_size,
                         make_uppercase=True)
    add_empty_paragraphs(document, lines_after_title, text_size)

    for line in people:
        paragraph = add_styled_paragraph(document, line, WD_ALIGN_PARAGRAPH.LEFT, False, font, text_size)
        paragraph.paragraph_format.left_indent = Cm(text_width - PEOPLE_COLUMN_CM)

    add_empty_paragraphs(document, lines_after_people, text_size)
    city_and_year = fields["city"] + ", " + fields["year"]
    add_styled_paragraph(document, city_and_year, WD_ALIGN_PARAGRAPH.CENTER, True, font, text_size)
