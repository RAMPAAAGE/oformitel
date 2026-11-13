from docx import Document

from checks import (
    check_alignment, check_captions, check_decoration, check_paragraph_spaces, check_font, check_indent, check_margins, check_numbers, check_punct,
    check_size, check_spacing, check_table_captions, check_themes, check_toc,
    find_warnings, margins_description, punct_work, spell_warning, structure_warnings,
)
from fixes import (
    add_captions, add_table_captions, apply_alignment, apply_paragraph_spaces, remove_decoration, apply_font, apply_indent, apply_margins, apply_size, apply_spacing,
    fix_numbers, fix_themes, fix_toc, format_title, insert_title_page, remove_title_page,
)
from layout import analyze_layout
from profiles import (  # noqa: F401 (их берёт app.py)
    DOCUMENT_TYPES, create_profile, list_profiles, load_profile, profile_for_type, update_profile,
)
from word_helpers import check_file_size, validate_docx_file  # noqa: F401 (их берёт app.py)

CRITERIA_NAMES = {
    1: "формат листа и поля",
    2: "шрифт",
    3: "размер шрифта",
    4: "межстрочный интервал",
    5: "отступ первой строки",
    6: "нумерация страниц",
    7: "точки в оглавлении",
    8: "нумерация картинок",
    9: "тема с новой страницы",
    10: "пунктуация",
    11: "выравнивание по ширине",
    12: "подписи таблиц",
    13: "отступы и интервалы между абзацами",
    14: "подчёркивание и выделение цветом",
}


def check_document(doc, profile):
    # номера непройденных критериев, например [1, 6]
    layout = analyze_layout(doc)
    results = {
        1: check_margins(doc, profile),
        2: check_font(doc, profile, layout),
        3: check_size(doc, profile, layout),
        4: check_spacing(doc, profile, layout),
        5: check_indent(doc, profile, layout),
        6: check_numbers(doc, layout),
        7: check_toc(doc, layout),
        8: check_captions(doc, layout),
        9: check_themes(layout),
        10: check_punct(doc, layout),
        11: check_alignment(doc, layout),
        12: check_table_captions(doc, layout),
        13: check_paragraph_spaces(doc, layout),
        14: check_decoration(doc, layout),
    }
    failed = []
    for number in results:
        if not results[number]:
            failed.append(number)
    return failed


def fix_document(doc, profile, failed):
    # Порядок важен: сначала структура (подписи, номера, разрывы), потом внешний вид
    if 8 in failed:
        add_captions(doc, profile)
    if 12 in failed:
        add_table_captions(doc, profile)
    if 6 in failed:
        fix_numbers(doc, profile)
    if 9 in failed:
        fix_themes(doc)
    if 1 in failed:
        apply_margins(doc, profile)
    if 2 in failed:
        apply_font(doc, profile)
    if 3 in failed:
        apply_size(doc, profile)
    if 4 in failed:
        apply_spacing(doc, profile)
    if 5 in failed:
        apply_indent(doc, profile)
    if 11 in failed:
        apply_alignment(doc)
    if 13 in failed:
        apply_paragraph_spaces(doc)
    if 14 in failed:
        remove_decoration(doc)
    if 7 in failed:
        fix_toc(doc)
    if 2 in failed or 3 in failed or 4 in failed or 5 in failed:
        format_title(doc, profile)  # после смены шрифта и интервалов титульный лист оформляем отдельно


def check_and_format_document(input_path, profile_path, output_path, title_path=None, document_type=None,
                              title_mode="keep"):
    # title_mode: что делать со старым титульным листом, если вставляем новый - "replace" (удалить) или "keep"
    validate_docx_file(input_path)
    profile = profile_for_type(load_profile(profile_path), document_type)

    doc = Document(input_path)
    failed = check_document(doc, profile)
    had_title = len(analyze_layout(doc)["title"]) > 0
    spell = spell_warning(doc)
    details = {}
    if 1 in failed:
        details[1] = margins_description(doc, profile)

    if not failed and not title_path:
        return {
            "is_correct": True,
            "profile_name": profile.get("name", ""),
            "failed": [],
            "title_inserted": False,
            "changes": [],
            "warnings": find_warnings(doc) + structure_warnings(doc, profile) + spell,
        }

    doc = Document(input_path)  # открываем заново, чистую копию
    changes = []
    if 10 in failed:
        # пунктуацию правим до вставки титульного листа,
        # пока старый титульный лист ещё виден программе как титульный
        changes = punct_work(doc, analyze_layout(doc), apply=True)
    if title_path:
        if had_title and title_mode == "replace":
            remove_title_page(doc)
        insert_title_page(doc, title_path)
        # после вставки структура документа изменилась - проверяем заново
        for number in check_document(doc, profile):
            if number not in failed and number != 10:
                failed.append(number)
        failed.sort()

    fix_document(doc, profile, failed)
    doc.save(output_path)

    warnings = find_warnings(doc) + structure_warnings(doc, profile)
    if title_path and had_title and title_mode != "replace":
        warnings.insert(0, "В документе два титульных листа: новый (первая страница) и старый (вторая). "
                           "Лишний удалите в Word или оформите работу заново и выберите «Заменить старый».")
    return {
        "is_correct": False,
        "profile_name": profile.get("name", ""),
        "failed": failed,
        "details": details,
        "title_inserted": bool(title_path),
        "changes": changes,
        "warnings": warnings + spell,
    }
