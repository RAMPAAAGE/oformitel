# Исправления. Каждая функция исправляет один критерий.
# Их вызывают только для тех критериев, которые не прошли проверку.

import copy

from docx import Document
from docx.enum.style import WD_STYLE_TYPE
from docx.enum.text import WD_ALIGN_PARAGRAPH, WD_LINE_SPACING, WD_TAB_ALIGNMENT, WD_TAB_LEADER
from docx.oxml import OxmlElement
from docx.oxml.ns import qn
from docx.shared import Cm, Pt
from docx.text.paragraph import Paragraph

from checks import SPACING_ROLES, find_typed_nums, has_caption, has_leader, has_trailing_num, new_page, title_runs
from layout import all_paragraphs, analyze_layout, get_role
from word_helpers import (
    get_runs, get_sect, has_picture, has_sect, page_breaks, prev_par, section_number,
    split_dots, split_space_page, validate_docx_file,
)


# --- 1-5: поля, шрифт, размер, интервал, отступ ---

def apply_margins(doc, profile):
    for section in doc.sections:
        section.top_margin = Cm(profile["margin_top_cm"])
        section.bottom_margin = Cm(profile["margin_bottom_cm"])
        section.left_margin = Cm(profile["margin_left_cm"])
        section.right_margin = Cm(profile["margin_right_cm"])


def apply_font(doc, profile):
    # сначала меняем шрифт в стилях (заголовки берут шрифт из стиля), потом в самом тексте
    for style in doc.styles:
        if style.type == WD_STYLE_TYPE.PARAGRAPH:
            style.font.name = profile["font_name"]
            # "шрифт темы" перебивает обычный шрифт, поэтому его убираем
            for attribute in ["w:asciiTheme", "w:hAnsiTheme", "w:eastAsiaTheme", "w:cstheme"]:
                style.element.rPr.rFonts.attrib.pop(qn(attribute), None)
    layout = analyze_layout(doc)
    for p in all_paragraphs(doc, layout, True):
        for run in get_runs(p):
            run.font.name = profile["font_name"]


def apply_size(doc, profile):
    layout = analyze_layout(doc)
    for p in all_paragraphs(doc, layout, False):
        for run in get_runs(p):
            run.font.size = Pt(profile["font_size_pt"])


def apply_spacing(doc, profile):
    layout = analyze_layout(doc)
    for p in doc.paragraphs:
        if get_role(p, layout) in SPACING_ROLES:
            p.paragraph_format.line_spacing_rule = WD_LINE_SPACING.MULTIPLE
            p.paragraph_format.line_spacing = profile["line_spacing"]


def apply_indent(doc, profile):
    layout = analyze_layout(doc)
    for p in doc.paragraphs:
        role = get_role(p, layout)
        if role == "body":
            p.paragraph_format.first_line_indent = Cm(profile["first_line_indent_cm"])
        elif role in ("heading", "center"):
            p.paragraph_format.first_line_indent = Cm(0)


# --- 6: нумерация страниц ---
# Титульный лист и оглавление без номера, со введения - с номером.

def strip_trailing_num(p):
    # убираем с конца абзаца число, табуляции и перенос строки перед ним
    nodes = []
    for node in p._p.iter():
        if node.tag in (qn("w:t"), qn("w:br"), qn("w:tab")):
            nodes.append(node)
    nodes.reverse()  # идём с конца абзаца
    removed = False
    for node in nodes:
        if not removed:
            if node.tag != qn("w:t"):
                return
            node.getparent().remove(node)  # само число
            removed = True
        elif node.tag == qn("w:tab"):
            node.getparent().remove(node)
        elif node.tag == qn("w:br") and node.get(qn("w:type")) != "page":
            node.getparent().remove(node)
            return
        else:
            return


def remove_typed_nums(doc):
    # удаляем номера страниц, которые набраны руками
    layout = analyze_layout(doc)
    for p in find_typed_nums(doc, layout):
        if has_trailing_num(p):
            strip_trailing_num(p)
        elif has_sect(p._p) or page_breaks(p._p):
            for t in p._p.iter(qn("w:t")):
                t.text = ""  # разрыв оставляем, число убираем
        else:
            p._p.getparent().remove(p._p)


def add_page_field(p, profile):
    # номер страницы в Word - это поле из 4 частей: begin, код PAGE, separate, end
    for kind in ["begin", "code", "separate", "end"]:
        run = p.add_run()
        run.font.name = profile["font_name"]
        run.font.size = Pt(profile["font_size_pt"])
        if kind == "code":
            element = OxmlElement("w:instrText")
            element.set(qn("xml:space"), "preserve")
            element.text = " PAGE "
        else:
            element = OxmlElement("w:fldChar")
            element.set(qn("w:fldCharType"), kind)
        run._r.append(element)


def clear_footer(footer):
    # удаляем из колонтитула всё (и абзацы w:p, и блок w:sdt, который вставляет Word)
    # и оставляем один пустой абзац. Иначе номер страницы мог бы задвоиться
    element = footer._element
    for child in list(element):
        element.remove(child)
    new_p = OxmlElement("w:p")
    element.append(new_p)
    return Paragraph(new_p, footer)


def write_number(footer, profile):
    p = clear_footer(footer)
    p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    add_page_field(p, profile)


def find_sect_props(p):
    # настройки раздела, в котором лежит абзац
    own = get_sect(p._p)
    if own is not None:
        return own
    for element in p._p.itersiblings(qn("w:p")):
        found = get_sect(element)
        if found is not None:
            return found
    return p._p.getparent().find(qn("w:sectPr"))  # последний раздел документа


def remove_breaks_before(p):
    # ручные разрывы страницы перед абзацем убираем: их заменит разрыв раздела
    for element in p._p.itersiblings(preceding=True):
        if element.tag != qn("w:p"):
            break
        for br in page_breaks(element):
            br.getparent().remove(br)
        wrapped = Paragraph(element, p._parent)
        if wrapped.text.strip() != "" or has_picture(wrapped):
            break


def add_section_break(p):
    # ставим разрыв раздела перед абзацем, чтобы у нового раздела был свой колонтитул
    if p._p.getprevious() is None:
        return
    prev = prev_par(p)
    if prev is not None and has_sect(prev):
        return  # разрыв уже есть

    remove_breaks_before(p)
    props = find_sect_props(p)
    new_props = copy.deepcopy(props)
    # первый раздел не должен ссылаться на колонтитулы второго
    for tag in ["w:headerReference", "w:footerReference", "w:titlePg", "w:type"]:
        for element in new_props.findall(qn(tag)):
            new_props.remove(element)
    for element in props.findall(qn("w:type")):
        props.remove(element)  # раздел со введением всегда с новой страницы

    if prev is None:
        prev = OxmlElement("w:p")
        p._p.addprevious(prev)
    prev.get_or_add_pPr().append(new_props)


def fix_numbers(doc, profile):
    remove_typed_nums(doc)
    layout = analyze_layout(doc)

    if layout["intro"] is None:
        # введения нет - нумеруем все страницы, кроме первой
        for i, section in enumerate(doc.sections):
            section.footer.is_linked_to_previous = False
            write_number(section.footer, profile)
            section.different_first_page_header_footer = (i == 0)
            if i == 0:
                section.first_page_footer.is_linked_to_previous = False
                clear_footer(section.first_page_footer)
        return

    add_section_break(layout["intro"])
    first = section_number(layout["intro"])
    for i, section in enumerate(doc.sections):
        section.footer.is_linked_to_previous = False
        if i < first:
            clear_footer(section.footer)  # титульный лист и оглавление
        else:
            write_number(section.footer, profile)
            section.different_first_page_header_footer = False
            for element in section._sectPr.findall(qn("w:pgNumType")):
                section._sectPr.remove(element)  # убираем "начать нумерацию заново"


# --- 7: точки в оглавлении ---

def dots_to_tab(p):
    # "Введение.....3" -> "Введение<табуляция>3", чтобы номера встали ровно
    if "\t" in p.text:
        return
    result = split_dots(p.text) or split_space_page(p.text)
    if result is None or p._p.find(qn("w:hyperlink")) is not None:
        return
    runs = p.runs
    if not runs:
        return
    runs[0].text = result[0] + "\t" + result[1]
    for run in runs[1:]:
        run._r.getparent().remove(run._r)


def fix_toc(doc):
    layout = analyze_layout(doc)
    section = doc.sections[0]
    width = section.page_width.cm - section.left_margin.cm - section.right_margin.cm
    for p in layout["toc"]:
        if p.text.strip() == "":
            continue
        dots_to_tab(p)
        # У строки оглавления не должно быть старого отступа:
        # иначе табуляция выходит за край строки и номер страницы переносится вниз.
        p.paragraph_format.left_indent = Cm(0)
        p.paragraph_format.right_indent = Cm(0)
        p.paragraph_format.first_line_indent = Cm(0)
        if not has_leader(p):
            # табуляция по правому краю, а пустое место до неё заполняется точками
            p.paragraph_format.tab_stops.add_tab_stop(Cm(width), WD_TAB_ALIGNMENT.RIGHT, WD_TAB_LEADER.DOTS)


# --- 8: подписи к картинкам ---

def add_captions(doc, profile):
    layout = analyze_layout(doc)
    number = 0
    for p in list(doc.paragraphs):
        if p._p in layout["title_el"] or not has_picture(p):
            continue
        number += 1
        if has_caption(p):
            continue

        # новый абзац "Рис. N" сразу под картинкой
        element = OxmlElement("w:p")
        p._p.addnext(element)
        caption = Paragraph(element, p._parent)
        caption.alignment = WD_ALIGN_PARAGRAPH.CENTER
        caption.paragraph_format.line_spacing_rule = WD_LINE_SPACING.MULTIPLE
        caption.paragraph_format.line_spacing = profile["line_spacing"]
        caption.paragraph_format.first_line_indent = Cm(0)
        run = caption.add_run("Рис. " + str(number))
        run.font.name = profile["font_name"]
        run.font.size = Pt(profile["font_size_pt"])

        if p.text.strip() == "":
            # центрируем только абзац с одной картинкой, абзац с текстом не трогаем
            p.alignment = WD_ALIGN_PARAGRAPH.CENTER
            p.paragraph_format.keep_with_next = True  # картинка не отрывается от подписи


# --- 9: каждая тема с новой страницы и жирным шрифтом ---

def fix_themes(doc):
    layout = analyze_layout(doc)
    for p in layout["themes"]:
        if not new_page(p):
            p.paragraph_format.page_break_before = True
        for run in title_runs(p):
            run.font.bold = True


# --- титульный лист ---

def set_run_style(p, bold, size, upper=False):
    for run in p.runs:
        run.font.bold = bold
        run.font.size = Pt(size)
        if upper:
            run.text = run.text.upper()


def format_title(doc):
    # оформляем титульный лист по зонам:
    # справа - участники (обычным), по центру - название работы и город (жирным), остальное жирным
    layout = analyze_layout(doc)
    center = []
    for p in layout["title"]:
        if p.text.strip() == "":
            continue
        p.paragraph_format.line_spacing_rule = WD_LINE_SPACING.SINGLE
        p.paragraph_format.first_line_indent = Cm(0)
        if p.alignment == WD_ALIGN_PARAGRAPH.RIGHT:
            set_run_style(p, False, 14)
        elif p.alignment == WD_ALIGN_PARAGRAPH.CENTER:
            center.append(p)
        else:
            set_run_style(p, True, 14)
    if len(center) >= 1:
        set_run_style(center[0], True, 16, upper=True)  # название работы
    if len(center) >= 2:
        set_run_style(center[-1], True, 14)  # город и год


def insert_title_page(doc, title_path):
    # вставляем титульный лист из другого файла в начало документа
    validate_docx_file(title_path)
    body = doc.element.body
    new_elements = []
    for element in Document(title_path).element.body:
        if element.tag != qn("w:sectPr"):
            new_elements.append(copy.deepcopy(element))

    # пустые абзацы в начале (место под титульный лист) убираем, чтобы не было пустой страницы
    while len(body) > 1 and body[0].tag == qn("w:p"):
        first = Paragraph(body[0], doc._body)
        if first.text.strip() != "" or has_picture(first) or has_sect(body[0]):
            break
        body.remove(body[0])

    for i, element in enumerate(new_elements):
        body.insert(i, element)

    # старый текст документа начинаем с новой страницы
    first_old = body[len(new_elements)]
    target = None
    if first_old.tag == qn("w:p"):
        target = first_old
    else:
        for node in first_old.iter(qn("w:p")):  # документ начинается с таблицы
            target = node
            break
    if target is not None:
        Paragraph(target, doc._body).paragraph_format.page_break_before = True
    else:
        break_paragraph = OxmlElement("w:p")
        run = OxmlElement("w:r")
        br = OxmlElement("w:br")
        br.set(qn("w:type"), "page")
        run.append(br)
        break_paragraph.append(run)
        first_old.addprevious(break_paragraph)
