# Маленькие помощники для работы с документом Word.
#
# Как устроен файл .docx внутри:
#   документ - это XML, абзац - это тег w:p, кусок текста внутри абзаца - тег w:r (run).
#   Один абзац может состоять из нескольких кусков с разным шрифтом,
#   поэтому шрифт и размер мы проверяем у каждого куска.

import os

from docx.oxml.ns import qn
from docx.text.run import Run

MAX_DOCUMENT_SIZE_MB = 5


# --- проверка файла ---

def check_file_size(path, max_mb=MAX_DOCUMENT_SIZE_MB):
    # возвращает (размер в норме?, размер в мегабайтах)
    size_mb = os.path.getsize(path) / (1024 * 1024)
    return size_mb <= max_mb, size_mb


def validate_docx_file(path):
    if not path:
        raise ValueError("Не указан путь к документу.")
    if not os.path.isfile(path):
        raise ValueError("Файл документа не найден.")
    if not path.lower().endswith(".docx"):
        raise ValueError("Можно выбрать только файл с расширением .docx.")
    return True


# --- работа с текстом ---

def clean(text):
    # приводит название раздела к простому виду: "1. Введение:" -> "введение"
    text = text.strip().lower().rstrip(".:").strip()
    parts = text.split(" ", 1)
    if len(parts) == 2 and parts[0].rstrip(".)").replace(".", "").isdigit():
        text = parts[1].strip()
    return text


def is_subsection(text):
    # подраздел - это "1.1 Цели" или "2.3.1 Что-то"
    numbers = text.strip().split(" ")[0].rstrip(".").split(".")
    if len(numbers) < 2:
        return False
    for number in numbers:
        if not number.isdigit():
            return False
    return True


def cut_page_number(text):
    # отрезает номер страницы с конца строки: "Введение 3" -> ("Введение ", "3")
    i = len(text)
    while i > 0 and (text[i - 1].isdigit() or text[i - 1] in "-–"):
        i -= 1
    return text[:i], text[i:]


def split_dots(text):
    # строка оглавления с точками, набранными руками:
    # "Введение.......3" -> ("Введение", "3"); если строка другая - None
    rest, page = cut_page_number(text.strip())
    if page == "":
        return None
    dots = 0
    while rest and rest[-1] in ".… \t":
        if rest[-1] in ".…":
            dots += 1
        rest = rest[:-1]
    if dots < 2 or rest.strip() == "":
        return None
    return rest.strip(), page


def split_space_page(text):
    # строка оглавления, где номер отделён просто пробелом:
    # "Практическая часть 67" -> ("Практическая часть", "67"); если строка другая - None
    rest, page = cut_page_number(text.strip())
    if page == "" or rest == "" or rest[-1] != " " or rest.strip() == "":
        return None
    return rest.strip(), page


def split_toc(text):
    # делит строку оглавления на название и номер страницы
    text = text.strip()
    if "\t" in text:
        title, _, page = text.partition("\t")
        return title.strip(), page.strip()
    result = split_dots(text)
    if result is None:
        return text, ""
    return result


def is_caption(text):
    # подпись к картинке: "Рис. 1", "Рисунок 2" и т.п.
    t = text.strip().lower()
    return t.startswith(("рис.", "рисунок", "рис ", "рис("))


# --- работа с XML абзаца ---

def has_tag(element, tag_name):
    # есть ли внутри элемента хотя бы один тег с таким именем
    for _ in element.iter(qn(tag_name)):
        return True
    return False


def get_runs(p):
    # все куски текста абзаца (в том числе внутри ссылок)
    return [Run(r, p) for r in p._p.iter(qn("w:r"))]


def has_picture(p):
    return has_tag(p._p, "w:drawing") or has_tag(p._p, "w:pict")


def page_breaks(element):
    # ручные разрывы страницы (Ctrl+Enter) внутри элемента
    result = []
    for br in element.iter(qn("w:br")):
        if br.get(qn("w:type")) == "page":
            result.append(br)
    return result


def break_before(p):
    # включена ли у абзаца галочка "с новой страницы"
    settings = p._p.pPr
    if settings is None:
        return False
    flag = settings.find(qn("w:pageBreakBefore"))
    return flag is not None and flag.get(qn("w:val")) not in ("0", "false")


def starts_with_break(p):
    # абзац начинается с ручного разрыва страницы
    for node in p._p.iter():
        if node.tag == qn("w:t") and (node.text or "").strip():
            return False
        if node.tag == qn("w:br") and node.get(qn("w:type")) == "page":
            return True
    return False


def get_sect(element):
    # настройки раздела (sectPr), если абзац заканчивает раздел
    settings = element.find(qn("w:pPr"))
    if settings is None:
        return None
    return settings.find(qn("w:sectPr"))


def has_sect(element):
    return get_sect(element) is not None


def style_name(p):
    # название стиля абзаца маленькими буквами.
    # У некоторых документов стиль указан, но в файле его нет - тогда p.style равен None
    if p.style is None:
        return ""
    return (p.style.name or "").lower()


def has_num(p):
    # абзац - пункт автоматического списка Word
    settings = p._p.pPr
    return settings is not None and settings.find(qn("w:numPr")) is not None


def is_list(p):
    return has_num(p) or style_name(p).startswith("list")


def is_toc_style(p):
    return style_name(p).startswith("toc")


def is_heading_style(p):
    return style_name(p).startswith(("heading", "заголовок"))


def is_top_heading(p):
    return style_name(p) in ("heading 1", "заголовок 1")


def before(p):
    # все абзацы документа перед этим, по порядку
    result = []
    for element in p._p.getroottree().getroot().iter(qn("w:p")):
        if element is p._p:
            break
        result.append(element)
    return result


def section_number(p):
    # номер раздела, в котором лежит абзац (считаем с нуля)
    count = 0
    for element in before(p):
        if has_sect(element):
            count += 1
    return count


def prev_par(p):
    # предыдущий абзац, если прямо перед этим стоит абзац (а не таблица)
    node = p._p.getprevious()
    if node is not None and node.tag == qn("w:p"):
        return node
    return None


# --- чтение оформления ---
# Если у текста не указан шрифт, он берётся из стиля абзаца,
# а если и там нет - из стиля, на котором основан этот стиль, и так далее.

def from_style(p, get_value):
    style = p.style
    while style is not None:
        value = get_value(style)
        if value is not None:
            return value
        style = style.base_style
    return None


def get_font(run, p):
    try:
        return run.font.name or from_style(p, lambda style: style.font.name)
    except (ValueError, TypeError):
        return None


def get_size(run, p):
    try:
        size = run.font.size or from_style(p, lambda style: style.font.size)
    except (ValueError, TypeError):
        return None
    if not size:
        return None
    return size.pt


def get_bold(run, p):
    try:
        if run.font.bold is not None:
            return run.font.bold
        return bool(from_style(p, lambda style: style.font.bold))
    except (ValueError, TypeError):
        return False


def get_spacing(p):
    try:
        value = p.paragraph_format.line_spacing
        if value is None:
            value = from_style(p, lambda style: style.paragraph_format.line_spacing)
    except (ValueError, TypeError):
        return None
    # интервал "1.5" приходит как дробное число, а интервал "ровно 18 пт" - другим типом
    if isinstance(value, float):
        return value
    return None


def get_indent(p):
    try:
        indent = p.paragraph_format.first_line_indent
        if indent is None:
            indent = from_style(p, lambda style: style.paragraph_format.first_line_indent)
    except (ValueError, TypeError):
        return 0.0
    if indent is None:
        return 0.0
    return indent.cm
