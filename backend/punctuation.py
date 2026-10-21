from docx.oxml.ns import qn

W_T = qn("w:t")
W_TAB = qn("w:tab")
W_BR = qn("w:br")
SKIP_TAGS = [qn("w:drawing"), qn("w:pict"), qn("w:txbxContent")]
SPACE_ATTR = "{http://www.w3.org/XML/1998/namespace}space"

ENDS = ".!?…:;"
CLOSERS = "»\")'”"
OPENERS = ',.!?…:;—–-("«“\n\''
BIBLIO = ["список литературы", "список использованных источников", "литература"]

ABBR = ["рис", "табл", "стр", "гг", "напр", "см", "др", "пр", "тыс", "млн", "млрд",
        "руб", "коп", "англ", "нем", "прим", "проф", "доц", "акад", "обл"]
ABBR_FIX = {"тп": "т.п.", "тд": "т.д.", "др": "др.", "пр": "пр."}
ABBR_NO_DOT = ["т.д", "т.п", "т.е", "т.к"]

KOTORY = ["который", "которая", "которое", "которые", "которого", "которому",
          "которым", "которой", "которую", "которых", "которыми", "котором"]
PREPS = ["в", "во", "на", "с", "со", "к", "ко", "о", "об", "обо", "от", "из", "у", "по",
         "для", "за", "над", "под", "при", "про", "через", "без", "до", "между", "среди",
         "около", "вокруг", "после", "перед", "ради", "против", "сквозь"]
QUANT = ["один", "одна", "одно", "одни", "некоторые", "многие", "часть", "все", "всех",
         "каждый", "каждая", "каждое", "оба", "обе", "несколько", "большинство", "любой",
         "любая", "любое", "два", "две", "три", "четыре", "пять", "множество", "половина"]
CONJ = ["и", "а", "но", "да", "или", "либо", "ни"]
# после этих слов запятую перед «что», «если» и т.д. не ставим
SKIP = CONJ + PREPS + ["не", "только", "лишь", "даже", "ли", "же", "именно",
                       "особенно", "как", "так", "тем", "если", "когда", "что", "чтобы"]
POTOMU_SKIP = ["не", "только", "именно", "лишь", "даже"]
STOP_END = CONJ + PREPS


# --- текст абзаца и правки в нём ---
# правка - это кортеж (начало, конец, новый_текст): кусок text[начало:конец] заменяется на новый_текст

def get_nodes(p):
    # кусочки текста абзаца. Перенос строки и табуляция тоже занимают 1 символ
    items = []

    def walk(el):
        for child in el:
            if child.tag in SKIP_TAGS:
                continue
            if child.tag in (W_T, W_TAB, W_BR):
                items.append(child)
            else:
                walk(child)

    walk(p._p)
    nodes = []
    text = ""
    for el in items:
        if el.tag == W_T:
            piece = el.text or ""
        elif el.tag == W_TAB:
            piece = "\t"
        else:
            piece = "\n"
        nodes.append((el, len(text), len(text) + len(piece)))
        text += piece
    return nodes, text


def set_text(el, text):
    el.text = text
    el.set(SPACE_ATTR, "preserve")


def put_text(nodes, pos, new):
    # вставка: кладём в кусок, который стоит слева, чтобы шрифт остался тем же
    near = [n for n in nodes if n[0].tag == W_T and n[1] <= pos <= n[2]]
    if not near:
        return
    left = [n for n in near if n[1] < pos]
    el, a, b = left[-1] if left else near[0]
    old = el.text or ""
    set_text(el, old[:pos - a] + new + old[pos - a:])


def cut_text(nodes, start, end, new):
    first = True
    for el, a, b in nodes:
        if el.tag != W_T or b <= start or a >= end:
            continue
        s = max(start, a) - a
        e = min(end, b) - a
        old = el.text or ""
        set_text(el, old[:s] + (new if first else "") + old[e:])
        first = False


def apply_edits(nodes, edits):
    # с конца в начало, чтобы номера символов впереди не сдвигались
    for start, end, new in sorted(edits, reverse=True):
        if start == end:
            put_text(nodes, start, new)
        else:
            cut_text(nodes, start, end, new)


def show(text, start, end):
    left = max(0, start - 25)
    right = min(len(text), end + 25)
    piece = text[left:right].replace("\n", " ").replace("\t", " ")
    return ("…" if left > 0 else "") + piece + ("…" if right < len(text) else "")


def make_changes(text, edits):
    # правки рядом друг с другом показываем одной строкой «было → стало»
    groups = []
    for edit in edits:
        if groups and edit[0] - groups[-1][-1][1] <= 30:
            groups[-1].append(edit)
        else:
            groups.append([edit])
    result = []
    for group in groups:
        start, end = group[0][0], group[-1][1]
        left = max(0, start - 25)
        right = min(len(text), end + 25)
        middle = ""
        pos = start
        for s, e, new in group:
            middle += text[pos:s] + new
            pos = e
        middle += text[pos:end]
        old_text = text[left:right]
        new_text = text[left:start] + middle + text[end:right]
        result.append({"old": clean_view(old_text, left, right, len(text)),
                       "new": clean_view(new_text, left, right, len(text))})
    return result


def clean_view(piece, left, right, size):
    piece = piece.replace("\n", " ").replace("\t", " ").strip()
    return ("…" if left > 0 else "") + piece + ("…" if right < size else "")


# --- помощники ---

def find_words(text):
    # список (слово, начало, конец)
    words = []
    i = 0
    while i < len(text):
        if text[i].isalnum():
            j = i
            while j < len(text) and (text[j].isalnum() or text[j] == "-"):
                j += 1
            words.append((text[i:j].lower(), i, j))
            i = j
        else:
            i += 1
    return words


def tokens(text):
    # куски текста между пробелами: (кусок, начало, конец)
    result = []
    i = 0
    while i < len(text):
        if text[i].isspace():
            i += 1
            continue
        j = i
        while j < len(text) and not text[j].isspace():
            j += 1
        result.append((text[i:j], i, j))
        i = j
    return result


def prev_char(text, start):
    j = start - 1
    while j >= 0 and text[j] in " \t\xa0":
        j -= 1
    return text[j] if j >= 0 else ""


def comma_pos(text, start):
    # куда ставить запятую перед словом, которое начинается в start
    j = start - 1
    while j >= 0 and text[j] in " \t\xa0":
        j -= 1
    return j + 1


def token_at(text, i):
    a = i
    while a > 0 and not text[a - 1].isspace():
        a -= 1
    b = i
    while b < len(text) and not text[b].isspace():
        b += 1
    return text[a:b]


def is_link(token):
    for part in ["http", "www.", "@", "/"]:
        if part in token:
            return True
    return False


def is_caption(text):
    t = text.strip().lower()
    return t.startswith(("рис.", "рисунок", "рис ", "рис("))


def ends_line(text, i):
    end = text.find(")", i)
    if end == -1:
        return False
    return text[end + 1:].split("\n")[0].strip() == ""


# --- правила: каждое возвращает список правок ---

def double_spaces(text):
    found = []
    i = 0
    while i < len(text) - 1:
        if text[i] == " " and text[i + 1] == " ":
            j = i
            while j < len(text) and text[j] == " ":
                j += 1
            start_of_line = i == 0 or text[i - 1] in "\n\t"
            end_of_line = j == len(text) or text[j] in "\n\t"
            if not start_of_line and not end_of_line:
                found.append((i + 1, j, ""))
            i = j
        else:
            i += 1
    return found


def space_before(text):
    found = []
    for i in range(1, len(text)):
        ch = text[i]
        if ch not in ",;:!?." or text[i - 1] != " ":
            continue
        if ch == "." and (text[i + 1:i + 2] == "." or text[i - 2:i - 1] == "."):
            continue
        a = i
        while a > 0 and text[a - 1] == " ":
            a -= 1
        if a > 0 and text[a - 1] not in "\n\t":
            found.append((a, i, ""))
    return found


def no_space_after(text):
    found = []
    for i in range(1, len(text) - 1):
        ch, nxt = text[i], text[i + 1]
        if ch not in ",;:!?." or is_link(token_at(text, i)):
            continue
        letter_next = nxt.isalpha() or (nxt in "«“\"" and text[i + 2:i + 3].isalpha())
        if not letter_next:
            continue
        if ch == "." and not (nxt.isupper() and text[i - 1].islower()):
            continue
        found.append((i + 1, i + 1, " "))
    return found


def repeated_marks(text):
    found = []
    i = 0
    while i < len(text):
        ch = text[i]
        if ch in ",;:.":
            j = i
            while j < len(text) and text[j] == ch:
                j += 1
            if j - i >= 2 and (ch != "." or j - i == 2):
                found.append((i + 1, j, ""))
            i = j
        else:
            i += 1
    return found


def small_after_dot(text):
    found = []
    for i in range(len(text) - 2):
        if text[i] not in ".!?" or text[i + 1] != " " or not text[i + 2].islower():
            continue
        if text[i] == ".":
            j = i
            while j > 0 and text[j - 1].isalpha():
                j -= 1
            word = text[j:i]
            if i == 0 or text[i - 1].isdigit() or len(word) <= 3 or word.lower() in ABBR:
                continue
            if j > 0 and text[j - 1] == ".":
                continue
        found.append((i + 2, i + 3, text[i + 2].upper()))
    return found


def abbreviations(text):
    found = []
    words = tokens(text)
    for k, (tok, s, e) in enumerate(words):
        lead = len(tok) - len(tok.lstrip("(«“\""))
        core = tok.strip("(«“\"").rstrip(",;:!?)»”\"")
        if core == "":
            continue
        start = s + lead
        end = start + len(core)
        low = core.lower()
        prev = words[k - 1][0].lower() if k > 0 else ""
        want = None
        if prev == "и" and low.replace(".", "") in ABBR_FIX:
            want = ABBR_FIX[low.replace(".", "")]
        elif low in ABBR_NO_DOT:
            want = low + "."
        if want is None or low == want:
            continue
        if want.startswith(low):
            found.append((end, end, want[len(low):]))
        else:
            found.append((start, end, want))
    return found


def space_after_abbr(text):
    found = []
    for i in range(1, len(text) - 1):
        if text[i] != "." or not text[i + 1].isdigit():
            continue
        j = i
        while j > 0 and text[j - 1].isalpha():
            j -= 1
        if text[j:i].lower() in ["рис", "табл", "стр"]:
            found.append((i + 1, i + 1, " "))
    return found


def paren_space(text):
    # "игры(рис. 3)" -> "игры (рис. 3)", "игры.(рис. 3)" -> "игры (рис. 3)"
    found = []
    for key in ["(рис", "(табл"]:
        i = text.find(key)
        while i != -1:
            k = i
            while k > 0 and text[k - 1] == " ":
                k -= 1
            if k > 1 and text[k - 1] == "." and text[k - 2].isalpha() and ends_line(text, i):
                found.append((k - 1, i, " "))
            elif k == i and i > 0 and text[i - 1].isalpha():
                found.append((i, i, " "))
            i = text.find(key, i + 1)
    return found


def missing_commas(text):
    found = []
    words = find_words(text)
    for k, (w, s, e) in enumerate(words):
        before = prev_char(text, s)
        if before == "" or before in OPENERS:
            continue
        prev_word = words[k - 1][0] if k > 0 else ""
        next_word = words[k + 1][0] if k + 1 < len(words) else ""
        gap = text[e:words[k + 1][1]] if k + 1 < len(words) else ""

        if w == "но" and prev_word not in ("а", "и"):
            found.append((comma_pos(text, s), comma_pos(text, s), ","))
        elif w in KOTORY:
            if prev_word in PREPS:
                ps = words[k - 1][1]
                before_prep = prev_char(text, ps)
                word_before = words[k - 2][0] if k > 1 else ""
                if before_prep == "" or before_prep in OPENERS or word_before in QUANT:
                    continue
                found.append((comma_pos(text, ps), comma_pos(text, ps), ","))
            else:
                found.append((comma_pos(text, s), comma_pos(text, s), ","))
        elif w in ("что", "чтобы", "если", "когда"):
            if prev_word not in SKIP and prev_word != "потому":
                found.append((comma_pos(text, s), comma_pos(text, s), ","))
        elif w == "так" and next_word == "как" and gap.strip() == "":
            if prev_word not in SKIP and prev_word != "потому":
                found.append((comma_pos(text, s), comma_pos(text, s), ","))
        elif w == "потому" and next_word == "что" and gap.strip() == "":
            if prev_word in CONJ:
                found.append((e, e, ","))
            elif prev_word not in POTOMU_SKIP:
                found.append((comma_pos(text, s), comma_pos(text, s), ","))
    return found


TEXT_RULES = [double_spaces, space_before, no_space_after, repeated_marks, small_after_dot,
              abbreviations, space_after_abbr, paren_space, missing_commas]


# --- строки абзаца: точка в конце и большая буква в пунктах списка ---

def item_start(line):
    # где начинается текст пункта после "1." или "-"; если строка не пункт, вернёт -1
    i = 0
    while i < len(line) and line[i] in " \t":
        i += 1
    j = i
    while j < len(line) and line[j].isdigit():
        j += 1
    if j > i and j < len(line) and line[j] in ".)":
        j += 1
    elif line[i:i + 1] in ("-", "–", "•"):
        j = i + 1
    else:
        return -1
    if j < len(line) and line[j] not in " \t":
        return -1
    while j < len(line) and line[j] in " \t":
        j += 1
    return j


def end_edit(text, a, b, item):
    line = text[a:b].rstrip(" \t\xa0")
    body = line.strip()
    if len(body) < (5 if item else 20):
        return None
    last_word = body.split()[-1]
    if body.isdigit() or is_link(last_word) or is_caption(body):
        return None
    if last_word.lower().strip(CLOSERS) in STOP_END:
        return None
    tail = line.rstrip(CLOSERS)
    marks = ENDS + "," if item else ENDS
    if tail and tail[-1] in marks:
        return None
    return (a + len(line), a + len(line), ".")


def line_edits(text, is_list, check_end, prev_dot):
    # prev_dot: закончился ли предыдущий пункт точкой (тогда следующий с большой буквы)
    edits = []
    lines = []
    pos = 0
    for line in text.split("\n"):
        if line.strip() != "" and not line.strip().isdigit():
            lines.append((pos, pos + len(line)))
        pos += len(line) + 1

    for k, (a, b) in enumerate(lines):
        line = text[a:b]
        if is_caption(line):
            prev_dot = None
            continue
        s = item_start(line)
        item = is_list or s != -1
        if not item:
            prev_dot = None
            if check_end and k == len(lines) - 1:
                edit = end_edit(text, a, b, False)
                if edit:
                    edits.append(edit)
            continue

        if s == -1:
            s = len(line) - len(line.lstrip())
        first = text[a + s:a + s + 1]
        if prev_dot and first.isalpha() and first.islower():
            edits.append((a + s, a + s + 1, first.upper()))

        edit = end_edit(text, a, b, True) if check_end else None
        if edit:
            edits.append(edit)
            prev_dot = True
        else:
            tail = line.rstrip(" \t\xa0").rstrip(CLOSERS)
            if tail and tail[-1] in ".!?…":
                prev_dot = True
            elif tail and tail[-1] in ";,":
                prev_dot = False
            else:
                prev_dot = None
    return edits, prev_dot


def heading_edit(text):
    t = text.rstrip()
    if "\n" in t or not t.endswith(".") or t.endswith(".."):
        return []
    return [(len(t) - 1, len(t), "")]


def find_edits(text, is_list, check_end, prev_dot):
    edits = []
    for rule in TEXT_RULES:
        edits += rule(text)
    more, prev_dot = line_edits(text, is_list, check_end, prev_dot)
    edits += more
    edits.sort()
    result = []
    last_end = -1
    for edit in edits:
        if edit[0] < last_end or edit in result:
            continue
        dot_again = edit == (edit[0], edit[0], ".") and result and result[-1][1] == edit[0] and result[-1][2].endswith(".")
        if dot_again:
            continue
        result.append(edit)
        last_end = max(last_end, edit[1])
    return result, prev_dot


# --- пометки Word: красные волнистые линии ---

def spell_words(doc):
    words = []
    for p in doc.paragraphs:
        inside = False
        word = ""
        for node in p._p.iter():
            if node.tag == qn("w:proofErr"):
                kind = node.get(qn("w:type"))
                if kind == "spellStart":
                    inside = True
                    word = ""
                elif kind == "spellEnd":
                    if inside and any(ch.isalpha() for ch in word):
                        words.append(word.strip())
                    inside = False
            elif node.tag == W_T and inside:
                word += node.text or ""
    return words
