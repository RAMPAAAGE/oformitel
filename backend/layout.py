from docx.enum.text import WD_ALIGN_PARAGRAPH

from word_helpers import (
    break_before, clean, has_picture, has_sect, is_caption, is_heading_style, is_list, is_table_caption,
    is_subsection, is_toc_style, is_top_heading, page_breaks, split_toc,
)

KNOWN_TITLES = [
    "оглавление", "содержание", "введение", "заключение",
    "список литературы", "список использованных источников", "список использованной литературы",
    "список источников", "литература",
    "приложение", "приложения",
    # разделы бизнес-плана («Шаг в бизнес»)
    "описание бизнес-идеи", "анализ состояния рынка", "ценностное предложение и анализ целевой аудитории",
    "бизнес-модель", "операционный план", "финансовая модель",
]
TOC_NAMES = ["оглавление", "содержание"]
TITLE_LIMIT = 80  # титульный лист ищем только среди первых 80 абзацев (с пустыми строками)


def find_toc_intro(paragraphs):
    toc = None
    intro = None
    for i, p in enumerate(paragraphs):
        name = clean(p.text)
        if toc is None and name in TOC_NAMES:
            toc = i
        if name == "введение" and not is_toc_style(p) and (toc is None or i > toc):
            intro = i  # берём последнее "Введение": первое может быть строкой оглавления
    return toc, intro


def find_title(paragraphs, toc, intro):
    # титульный лист - всё, что стоит до первого разрыва страницы (или до оглавления)
    end = None
    for i, p in enumerate(paragraphs[:TITLE_LIMIT + 1]):
        if i == toc or i == intro:
            end = i
        elif i > 0 and break_before(p):
            end = i
        elif page_breaks(p._p):
            end = i + 1 if p.text.strip() else i
        elif has_sect(p._p):
            end = i + 1
        if end is not None:
            break
    if end is None:
        return []
    title = paragraphs[:end]
    for p in title:
        if p.text.strip():
            return title
    return []  # на "титульном листе" нет ни одной буквы - значит, это не он


def find_toc_entries(paragraphs, toc, intro):
    end = intro
    if end is None and toc is not None:
        # введения нет - оглавление кончается на первом разрыве страницы, разрыве раздела
        # (его ставит сама программа при нумерации) или на первом заголовке Word
        for i in range(toc + 1, len(paragraphs)):
            p = paragraphs[i]
            if page_breaks(p._p) or break_before(p) or is_heading_style(p):
                end = i
                break
            if has_sect(p._p):
                end = i + 1
                break
    result = []
    for i, p in enumerate(paragraphs):
        if is_toc_style(p) or (toc is not None and end is not None and toc < i < end):
            result.append(p)
    return result


def get_toc_titles(entries):
    titles = []
    for p in entries:
        raw = split_toc(p.text)[0]
        name = clean(raw)
        if is_subsection(raw) or name == "" or name in titles:
            continue
        titles.append(name)
    return titles


def looks_like_heading(paragraphs, i):
    # похоже на заголовок без стиля: короткая строка, перед ней пустая строка,
    # а после неё длинный абзац текста
    p = paragraphs[i]
    text = p.text.strip()
    if text == "" or "\n" in text or len(text) > 70 or len(text.split()) > 8:
        return False
    if text[0].islower() or text[0].isdigit() or text[0] in "-–—•":
        return False
    if text[-1] in ":,;" or is_caption(text) or has_picture(p) or is_list(p):
        return False
    if i == 0 or paragraphs[i - 1].text.strip() != "":
        return False
    for next_p in paragraphs[i + 1:]:
        if next_p.text.strip() != "":
            return len(next_p.text.strip()) >= 100
    return False


def word_beginnings(titles):
    # первые 5 букв каждого длинного слова: "практическая часть" -> ["практ"]
    result = []
    for title in titles:
        for word in title.split():
            if len(word) >= 5:
                result.append(word[:5])
    return result


def shares_word(text, beginnings):
    for word in clean(text).split():
        if len(word) >= 5 and word[:5] in beginnings:
            return True
    return False


def find_themes(paragraphs, title_el, toc_el, toc_titles, intro):
    # заголовки разделов ("темы"): по стилю Word, по известным словам и по оглавлению
    themes = []
    found = []
    for p in paragraphs:
        if p._p in title_el or p._p in toc_el:
            continue
        name = clean(p.text.split("\n")[0])  # заголовок может быть в одном абзаце с текстом
        if name == "":
            continue
        if is_top_heading(p) or name in KNOWN_TITLES or name in toc_titles:
            themes.append(p)
            found.append(name)

    # названия из оглавления, которые в тексте записаны чуть по-другому
    lost = [name for name in toc_titles if name not in found]
    if lost and intro is not None:
        beginnings = word_beginnings(lost)
        added = 0
        for i in range(intro + 1, len(paragraphs)):
            if added >= len(lost):
                break
            p = paragraphs[i]
            if p._p in toc_el or not looks_like_heading(paragraphs, i):
                continue
            if shares_word(p.text, beginnings):
                themes.append(p)
                added += 1

    chosen = [p._p for p in themes]
    return [p for p in paragraphs if p._p in chosen]


def find_first_section(themes, intro, paragraphs, toc, title):
    # раздел, с которого начинается нумерация страниц: обычно «Введение»,
    # а в бизнес-плане его нет - тогда первый раздел после оглавления
    if intro is not None:
        return paragraphs[intro]
    if toc is None and not title:
        return None
    for p in themes:
        if clean(p.text) not in TOC_NAMES:
            return p
    return None


def analyze_layout(doc):
    paragraphs = list(doc.paragraphs)
    toc, intro = find_toc_intro(paragraphs)
    title = find_title(paragraphs, toc, intro)
    toc_entries = find_toc_entries(paragraphs, toc, intro)
    # "_el" - это сами XML-элементы абзацев, по ним удобно проверять "абзац в списке?"
    title_el = [p._p for p in title]
    toc_el = [p._p for p in toc_entries]
    themes = find_themes(paragraphs, title_el, toc_el, get_toc_titles(toc_entries), intro)
    return {
        "first_section": find_first_section(themes, intro, paragraphs, toc, title),
        "title": title,
        "title_el": title_el,
        "toc": toc_entries,
        "toc_el": toc_el,
        "intro": paragraphs[intro] if intro is not None else None,
        "themes": themes,
        "theme_el": [p._p for p in themes],
    }


def get_role(p, layout):
    if p._p in layout["title_el"]:
        return "title"
    has_text = p.text.strip() != ""
    if not has_text and not has_picture(p):
        return "empty"
    if p._p in layout["toc_el"]:
        return "toc"
    if not has_text:
        return "image"
    if p._p in layout["theme_el"] or is_heading_style(p):
        return "heading"
    if is_list(p):
        return "list"
    if len(p.text) < 150 and (is_caption(p.text) or is_table_caption(p.text)):
        return "caption"
    if p.alignment in (WD_ALIGN_PARAGRAPH.CENTER, WD_ALIGN_PARAGRAPH.RIGHT):
        return "center"
    return "body"


def table_paragraphs(table):
    result = []
    for row in table.rows:
        for cell in row.cells:
            result += cell.paragraphs
            for inner in cell.tables:
                result += table_paragraphs(inner)
    return result


def all_paragraphs(doc, layout, with_title):
    result = []
    for p in doc.paragraphs:
        if with_title or p._p not in layout["title_el"]:
            result.append(p)
    for table in doc.tables:
        result += table_paragraphs(table)
    return result
