import copy

from docx import Document
from docx.enum.style import WD_STYLE_TYPE
from docx.enum.text import WD_ALIGN_PARAGRAPH, WD_LINE_SPACING, WD_TAB_ALIGNMENT, WD_TAB_LEADER
from docx.oxml import OxmlElement
from docx.oxml.ns import qn
from docx.shared import Cm, Pt, RGBColor
from docx.text.paragraph import Paragraph

from checks import (
    JUSTIFY_ROLES, SPACING_ROLES, ZERO_SPACE_ROLES, find_typed_nums, has_caption, has_leader, has_table_caption, has_trailing_num,
    new_page, text_tables, title_runs,
)
from layout import all_paragraphs, analyze_layout, get_role
from word_helpers import (
    get_bold, get_runs, get_size, get_sect, has_picture, has_sect, page_breaks, prev_par, section_number,
    set_font_xml, split_dots, split_space_page, validate_docx_file,
)


def apply_margins(doc, profile):
    for section in doc.sections:
        # лист A4; у альбомных страниц ширина больше высоты, это сохраняем
        if section.page_width is not None and section.page_height is not None \
                and section.page_width > section.page_height:
            section.page_width, section.page_height = Cm(29.7), Cm(21)
        else:
            section.page_width, section.page_height = Cm(21), Cm(29.7)
        section.top_margin = Cm(profile["margin_top_cm"])
        section.bottom_margin = Cm(profile["margin_bottom_cm"])
        section.left_margin = Cm(profile["margin_left_cm"])
        section.right_margin = Cm(profile["margin_right_cm"])


def apply_font(doc, profile):
    # сначала шрифт по умолчанию и в стилях (заголовки берут шрифт из стиля), потом в самом тексте
    font = profile["font_name"]
    defaults = doc.styles.element.find(qn("w:docDefaults"))
    if defaults is not None and defaults.find(qn("w:rPrDefault")) is not None:
        rpr_default = defaults.find(qn("w:rPrDefault"))
        if rpr_default.find(qn("w:rPr")) is None:
            rpr_default.append(OxmlElement("w:rPr"))
        set_font_xml(rpr_default.find(qn("w:rPr")), font)
    for style in doc.styles:
        if style.type in (WD_STYLE_TYPE.PARAGRAPH, WD_STYLE_TYPE.CHARACTER, WD_STYLE_TYPE.TABLE):
            set_font_xml(style.element.get_or_add_rPr(), font)
    layout = analyze_layout(doc)
    for p in all_paragraphs(doc, layout, True):
        for run in get_runs(p):
            set_font_xml(run._r.get_or_add_rPr(), font)


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
        elif role in ("heading", "center", "caption"):
            p.paragraph_format.first_line_indent = Cm(0)


def apply_alignment(doc):
    layout = analyze_layout(doc)
    for p in doc.paragraphs:
        if get_role(p, layout) in JUSTIFY_ROLES:
            p.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY


def apply_paragraph_spaces(doc):
    layout = analyze_layout(doc)
    for p in doc.paragraphs:
        role = get_role(p, layout)
        if role not in ZERO_SPACE_ROLES:
            continue
        p.paragraph_format.space_before = Pt(0)
        p.paragraph_format.space_after = Pt(0)
        if role != "list":
            p.paragraph_format.left_indent = Cm(0)
            p.paragraph_format.right_indent = Cm(0)


def remove_decoration(doc):
    layout = analyze_layout(doc)
    for p in all_paragraphs(doc, layout, True):
        for run in get_runs(p):
            run.font.underline = False
            run.font.highlight_color = None
            run.font.color.rgb = RGBColor(0, 0, 0)


def strip_trailing_num(p):
    nodes = []
    for node in p._p.iter():
        if node.tag in (qn("w:t"), qn("w:br"), qn("w:tab")):
            nodes.append(node)
    nodes.reverse()
    removed = False
    for node in nodes:
        if not removed:
            if node.tag != qn("w:t"):
                return
            node.getparent().remove(node)
            removed = True
        elif node.tag == qn("w:tab"):
            node.getparent().remove(node)
        elif node.tag == qn("w:br") and node.get(qn("w:type")) != "page":
            node.getparent().remove(node)
            return
        else:
            return


def remove_typed_nums(doc):
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
    # удаляем всё, включая блок w:sdt от Word, и оставляем один пустой абзац,
    # иначе номер страницы задвоится
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
    # отдельный раздел нужен, чтобы у него был свой колонтитул
    if p._p.getprevious() is None:
        return
    prev = prev_par(p)
    if prev is not None and has_sect(prev):
        return

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

    first_section = layout["first_section"]
    if first_section is None:
        # введения нет - нумеруем все страницы, кроме первой
        for i, section in enumerate(doc.sections):
            section.footer.is_linked_to_previous = False
            write_number(section.footer, profile)
            section.different_first_page_header_footer = (i == 0)
            if i == 0:
                section.first_page_footer.is_linked_to_previous = False
                clear_footer(section.first_page_footer)
        return

    add_section_break(first_section)
    first = section_number(first_section)
    for i, section in enumerate(doc.sections):
        section.footer.is_linked_to_previous = False
        if i < first:
            clear_footer(section.footer)  # титульный лист и оглавление
        else:
            write_number(section.footer, profile)
            section.different_first_page_header_footer = False
            for element in section._sectPr.findall(qn("w:pgNumType")):
                section._sectPr.remove(element)  # убираем "начать нумерацию заново"


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
            p.paragraph_format.tab_stops.add_tab_stop(Cm(width), WD_TAB_ALIGNMENT.RIGHT, WD_TAB_LEADER.DOTS)


def add_captions(doc, profile):
    layout = analyze_layout(doc)
    number = 0
    for p in list(doc.paragraphs):
        if p._p in layout["title_el"] or not has_picture(p):
            continue
        number += 1
        if has_caption(p):
            continue

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
            p.paragraph_format.keep_with_next = True


def add_table_captions(doc, profile):
    layout = analyze_layout(doc)
    for number, table in enumerate(text_tables(doc, layout), 1):
        if has_table_caption(table):
            continue
        element = OxmlElement("w:p")
        table._tbl.addprevious(element)
        caption = Paragraph(element, table._parent)
        caption.alignment = WD_ALIGN_PARAGRAPH.LEFT
        caption.paragraph_format.line_spacing_rule = WD_LINE_SPACING.MULTIPLE
        caption.paragraph_format.line_spacing = profile["line_spacing"]
        caption.paragraph_format.first_line_indent = Cm(0)
        caption.paragraph_format.keep_with_next = True
        run = caption.add_run("Таблица " + str(number))
        run.font.name = profile["font_name"]
        run.font.size = Pt(profile["font_size_pt"])


def fix_themes(doc):
    layout = analyze_layout(doc)
    for p in layout["themes"]:
        if not new_page(p):
            p.paragraph_format.page_break_before = True
        for run in title_runs(p):
            run.font.bold = True


def set_run_style(p, bold, size, upper=False):
    # bold=None - жирность не трогаем
    for run in p.runs:
        if bold is not None:
            run.font.bold = bold
        run.font.size = Pt(size)
        if upper:
            run.text = run.text.upper()


def title_weight(p):
    # чем "заметнее" строка, тем вероятнее, что это название работы
    size = 0
    bold = False
    for run in p.runs:
        if run.text.strip():
            size = max(size, get_size(run, p) or 0)
            bold = bold or get_bold(run, p)
    return (size, bold, p.text.isupper())


def format_title(doc, profile):
    # Название работы - строка по центру с самым крупным (жирным, заглавным) шрифтом.
    # Выше него - организация, ниже не по центру - участники и руководители,
    # последняя строка по центру - город и год. Остальные строки (например, «Секция ...») не трогаем.
    size = profile["font_size_pt"]
    lines = []
    for p in analyze_layout(doc)["title"]:
        if p.text.strip() != "":
            p.paragraph_format.line_spacing_rule = WD_LINE_SPACING.SINGLE
            p.paragraph_format.first_line_indent = Cm(0)
            lines.append(p)

    centered = [p for p in lines if p.alignment == WD_ALIGN_PARAGRAPH.CENTER]
    city = centered[-1] if len(centered) >= 2 else None
    candidates = [p for p in centered[1:] if p is not city] or centered[:1]
    work_title = None
    for p in candidates:
        if work_title is None or title_weight(p) > title_weight(work_title):
            work_title = p

    above_title = True
    for p in lines:
        if p is work_title:
            above_title = False
            set_run_style(p, True, size + 2, upper=True)
        elif p is city or (above_title and work_title is not None):
            set_run_style(p, True, size)
        elif p.alignment != WD_ALIGN_PARAGRAPH.CENTER and not p.text.strip().endswith(":"):
            set_run_style(p, False, size)
        else:
            set_run_style(p, None, size)


def remove_title_page(doc):
    # удаляем всё от начала документа до последней строки старого титульного листа
    # (вместе с таблицами и картинками на нём)
    title = analyze_layout(doc)["title"]
    if not title:
        return
    last = title[-1]._p
    body = doc.element.body
    while len(body) > 1:
        element = body[0]
        body.remove(element)
        if element is last:
            break


def insert_title_page(doc, title_path):
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
