# Проверки документа. Каждая функция check_... проверяет один критерий
# и возвращает True (всё правильно) или False (критерий нарушен).
# Здесь документ только читается, ничего не меняется.

from docx.oxml.ns import qn
from docx.text.paragraph import Paragraph

from layout import all_paragraphs, analyze_layout, get_role
from profiles import DOCUMENT_TYPES
from punctuation import BIBLIO, apply_edits, find_edits, get_nodes, heading_edit, make_changes, spell_words
from word_helpers import (
    before, break_before, clean, get_bold, get_font, get_indent, get_runs, get_size, get_spacing,
    has_num, has_picture, has_sect, is_caption, page_breaks, prev_par, section_number,
    split_dots, split_space_page, starts_with_break,
)

# у каких абзацев проверяем межстрочный интервал
SPACING_ROLES = ["body", "heading", "list", "center"]


# --- 1-5: поля, шрифт, размер, интервал, отступ ---

def check_margins(doc, profile):
    want = [profile["margin_top_cm"], profile["margin_bottom_cm"],
            profile["margin_left_cm"], profile["margin_right_cm"]]
    for section in doc.sections:
        try:
            got = [section.top_margin.cm, section.bottom_margin.cm,
                   section.left_margin.cm, section.right_margin.cm]
        except (ValueError, TypeError, AttributeError):
            return False
        for real, needed in zip(got, want):
            if abs(real - needed) > 0.05:  # допуск полмиллиметра
                return False
    return True


def check_font(doc, profile, layout):
    for p in all_paragraphs(doc, layout, True):
        for run in get_runs(p):
            if run.text.strip() and get_font(run, p) != profile["font_name"]:
                return False
    return True


def check_size(doc, profile, layout):
    # титульный лист не проверяем: там по правилам бывают крупные буквы
    for p in all_paragraphs(doc, layout, False):
        for run in get_runs(p):
            if run.text.strip():
                size = get_size(run, p)
                if size is None or abs(size - profile["font_size_pt"]) > 0.5:
                    return False
    return True


def check_spacing(doc, profile, layout):
    for p in doc.paragraphs:
        if get_role(p, layout) in SPACING_ROLES:
            spacing = get_spacing(p)
            if spacing is None or abs(spacing - profile["line_spacing"]) > 0.05:
                return False
    return True


def check_indent(doc, profile, layout):
    # красная строка нужна только обычному тексту
    for p in doc.paragraphs:
        if get_role(p, layout) == "body":
            if abs(get_indent(p) - profile["first_line_indent_cm"]) > 0.1:
                return False
    return True


# --- 6: нумерация страниц ---
# Правильно: на титульном листе и в оглавлении номера нет, со введения - есть.

def is_typed_num(p):
    # абзац, в котором руками набран только номер страницы: "7"
    text = p.text.strip()
    if text == "" or has_num(p):
        return False
    return text.isdigit() and text.isascii() and len(text) <= 3


def has_trailing_num(p):
    # номер страницы набран руками в конце абзаца: "...текст.<Enter><табы>7"
    if is_typed_num(p):
        return False
    text = p.text.rstrip()
    if "\n" not in text:
        return False
    last = text.split("\n")[-1].strip()
    return last.isdigit() and last.isascii() and len(last) <= 3


def find_typed_nums(doc, layout):
    result = []
    for p in doc.paragraphs:
        if p._p in layout["title_el"]:
            continue
        if is_typed_num(p) or has_trailing_num(p):
            result.append(p)
    return result


def is_page_code(text):
    # код поля "номер страницы" в Word начинается со слова PAGE
    words = (text or "").split()
    return len(words) > 0 and words[0] == "PAGE"


def has_page_field(sections, i):
    # есть ли номер страницы в нижнем колонтитуле раздела i.
    # Колонтитул "как в предыдущем разделе" - тогда смотрим предыдущий раздел
    while i > 0 and sections[i].footer.is_linked_to_previous:
        i -= 1
    footer = sections[i].footer
    if footer.is_linked_to_previous:
        return False
    for code in footer._element.iter(qn("w:instrText")):
        if is_page_code(code.text):
            return True
    for field in footer._element.iter(qn("w:fldSimple")):
        if is_page_code(field.get(qn("w:instr"))):
            return True
    return False


def check_numbers(doc, layout):
    sections = doc.sections
    intro = layout["intro"]
    if find_typed_nums(doc, layout):
        return False  # номера набраны руками - после нумерации будет два номера

    if intro is None:
        for i in range(len(sections)):
            if not has_page_field(sections, i):
                return False
        return True

    if intro._p.getprevious() is None:
        first = 0  # введение в самом начале документа
    else:
        prev = prev_par(intro)
        if prev is None or not has_sect(prev):
            return False  # перед введением нет разрыва раздела
        first = section_number(intro)

    for i, section in enumerate(sections):
        has_number = has_page_field(sections, i)
        if i < first and has_number:
            return False  # номер стоит на титульном листе или в оглавлении
        if i >= first:
            if not has_number or section.different_first_page_header_footer:
                return False
            if section._sectPr.find(qn("w:pgNumType")) is not None:
                return False  # нумерация начинается заново
    return True


# --- 7: точки в оглавлении ---
# Правильно: между названием и номером стоит табуляция с точками (её делает сам Word),
# а не точки, набранные руками.

def has_leader(p):
    # есть ли у абзаца табуляция с точками
    if 'w:leader="dot"' in p._p.xml:
        return True
    return p.style is not None and 'w:leader="dot"' in p.style.element.xml


def check_toc(doc, layout):
    for p in layout["toc"]:
        text = p.text.strip()
        if text == "":
            continue
        if split_dots(text) is not None:
            return False  # точки набраны руками
        if "\t" not in text and split_space_page(text) is not None:
            return False  # номер страницы отделён пробелом, а не табуляцией
        if "\t" in text and not has_leader(p):
            return False
    return True


# --- 8: подписи к картинкам ---

def caption_below(p):
    # есть ли подпись в ближайшем непустом абзаце под картинкой
    for element in p._p.itersiblings(qn("w:p")):
        next_p = Paragraph(element, p._parent)
        if next_p.text.strip() != "":
            return is_caption(next_p.text)
        if has_picture(next_p):
            return False
    return False


def has_caption(p):
    # подпись может быть строкой внутри абзаца с картинкой или в следующем абзаце
    for line in p.text.split("\n"):
        if is_caption(line):
            return True
    return caption_below(p)


def check_captions(doc, layout):
    for p in doc.paragraphs:
        if p._p in layout["title_el"]:
            continue
        if has_picture(p) and not has_caption(p):
            return False
    return True


# --- 9: каждая тема с новой страницы и жирным шрифтом ---

def new_page(p):
    # начинается ли абзац с новой страницы
    if break_before(p) or starts_with_break(p):
        return True
    for element in reversed(before(p)):
        if has_sect(element) or page_breaks(element):
            return True
        prev = Paragraph(element, p._parent)
        if prev.text.strip() != "" or has_picture(prev):
            return False
    return True  # перед абзацем ничего нет - это первая страница


def title_runs(p):
    # куски текста заголовка. Если после названия в том же абзаце идёт текст,
    # берём только куски до первого переноса строки
    runs = get_runs(p)
    if "\n" not in p.text:
        return runs
    result = []
    for run in runs:
        result.append(run)
        if run._r.find(qn("w:br")) is not None:
            break
    return result


def title_is_bold(p):
    for run in title_runs(p):
        if run.text.strip() and not get_bold(run, p):
            return False
    return True


def check_themes(layout):
    for p in layout["themes"]:
        if not new_page(p) or not title_is_bold(p):
            return False
    return True


# --- 10: пунктуация ---

def punct_work(doc, layout, apply=False):
    # идёт по абзацам и находит правки; если apply=True - сразу вносит их в документ.
    # Возвращает список изменений "было -> стало" для показа пользователю
    changes = []
    biblio = False  # список литературы оформляется по своим правилам, его не трогаем
    prev_dot = None
    for p in doc.paragraphs:
        if p._p in layout["theme_el"] and clean(p.text.split("\n")[0]) in BIBLIO:
            biblio = True
        if biblio or is_typed_num(p):
            continue
        role = get_role(p, layout)
        nodes, text = get_nodes(p)
        long_heading = role == "heading" and len(text.strip()) > 100  # текст, которому дали стиль заголовка
        if role in ("body", "list") or long_heading:
            edits, prev_dot = find_edits(text, role == "list", not has_picture(p), prev_dot)
        elif p._p in layout["theme_el"]:
            edits = heading_edit(text)
            prev_dot = None
        else:
            edits = []
            prev_dot = None
        if apply:
            apply_edits(nodes, edits)
        changes += make_changes(text, edits)
    return changes


def check_punct(doc, layout):
    return punct_work(doc, layout) == []


# --- пояснения к исправлениям ---

def cm_text(value):
    # 2.0 -> "2", 1.5 -> "1,5"
    return ("%g" % round(value, 2)).replace(".", ",")


def margins_description(doc, profile):
    # что именно не так с полями: "левое 2 → 3 см, правое 2 → 1,5 см"
    section = doc.sections[0]
    sides = [
        ("верхнее", section.top_margin, profile["margin_top_cm"]),
        ("нижнее", section.bottom_margin, profile["margin_bottom_cm"]),
        ("левое", section.left_margin, profile["margin_left_cm"]),
        ("правое", section.right_margin, profile["margin_right_cm"]),
    ]
    parts = []
    for name, real, needed in sides:
        if real is None:
            continue
        if abs(real.cm - needed) > 0.05:
            parts.append(name + " " + cm_text(real.cm) + " → " + cm_text(needed) + " см")
    return ", ".join(parts)


# --- предупреждения (их программа не исправляет, а только сообщает) ---

def structure_warnings(doc, profile):
    # есть ли обязательные части работы. Текст за автора программа не придумывает
    document_type = profile.get("document_type", "free")
    rules = profile.get("structure") or DOCUMENT_TYPES.get(document_type, DOCUMENT_TYPES["free"])
    layout = analyze_layout(doc)
    warnings = []

    if rules.get("title_page") and not layout["title"]:
        warnings.append("В документе не найден обязательный титульный лист.")
    if rules.get("toc") and not layout["toc"]:
        warnings.append("В документе не найдено обязательное оглавление.")

    headings = set()
    for p in layout["themes"]:
        if p.text.strip():
            headings.add(clean(p.text))
    for section in rules.get("sections", []):
        if section not in headings:
            warnings.append("Не найден обязательный раздел «" + section.capitalize() + "».")
    return warnings


def spell_warning(doc):
    # слова, которые Word подчеркнул красным; если пометок нет - ничего не пишем
    words = []
    for word in spell_words(doc):
        if word not in words:
            words.append(word)
    if not words:
        return []
    shown = ", ".join("«" + word + "»" for word in words[:10])
    more = ""
    if len(words) > 10:
        more = " и ещё " + str(len(words) - 10)
    return ["Word подчеркнул красной линией слова: " + shown + more + ". Проверьте их правописание вручную."]


def find_warnings(doc):
    layout = analyze_layout(doc)
    warnings = []
    if not layout["title"]:
        warnings.append("Титульный лист не найден. Создайте его на вкладке «Титульный лист» и вставьте в документ.")
    if layout["intro"] is None:
        warnings.append("Раздел «Введение» не найден: номера страниц поставлены на все страницы, кроме первой.")
    return warnings
