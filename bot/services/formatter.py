"""
ГОСТ-форматирование .docx документов.

Применяемые правила (ГОСТ 7.32-2017 и требования большинства российских вузов):
- Шрифт: Times New Roman, 14 pt
- Межстрочный интервал: 1.5
- Поля: левое 30 мм, правое 10 мм, верхнее 20 мм, нижнее 20 мм
- Отступ первой строки абзаца: 1.25 см
- Выравнивание основного текста: по ширине
- Заголовки 1 уровня: 14 pt, жирный, КАПСЛОК, по центру, без отступа, с новой страницы
- Заголовки 2 уровня: 14 pt, жирный, по левому краю, без отступа
- Заголовки 3 уровня: 14 pt, жирный, по левому краю, без отступа
- Нумерация страниц: снизу по центру, без точки
"""

import re
from io import BytesIO

from docx import Document
from docx.shared import Pt, Mm, Cm, RGBColor
from docx.enum.text import WD_ALIGN_PARAGRAPH, WD_LINE_SPACING
from docx.oxml import OxmlElement
from docx.oxml.ns import qn


# ---------------------------------------------------------------------------
# Константы ГОСТ
# ---------------------------------------------------------------------------

FONT_NAME = "Times New Roman"
FONT_SIZE = Pt(14)

MARGIN_LEFT   = Mm(30)
MARGIN_RIGHT  = Mm(10)
MARGIN_TOP    = Mm(20)
MARGIN_BOTTOM = Mm(20)

FIRST_LINE_INDENT = Cm(1.25)

HEADING_CONFIG = {
    1: {"bold": True,  "italic": False, "caps": True,  "align": WD_ALIGN_PARAGRAPH.CENTER,
        "space_before": Pt(0),  "space_after": Pt(12), "page_break": True},
    2: {"bold": True,  "italic": False, "caps": False, "align": WD_ALIGN_PARAGRAPH.LEFT,
        "space_before": Pt(12), "space_after": Pt(6),  "page_break": False},
    3: {"bold": True,  "italic": False, "caps": False, "align": WD_ALIGN_PARAGRAPH.LEFT,
        "space_before": Pt(8),  "space_after": Pt(4),  "page_break": False},
}

# Ключевые слова для заголовков 1 уровня без нумерации
UNNUMBERED_HEADINGS = {
    "ВВЕДЕНИЕ", "INTRODUCTION",
    "ЗАКЛЮЧЕНИЕ", "CONCLUSION",
    "ВЫВОДЫ", "SUMMARY",
    "АННОТАЦИЯ", "ABSTRACT",
    "СОДЕРЖАНИЕ", "ОГЛАВЛЕНИЕ",
    "СПИСОК ИСПОЛЬЗОВАННЫХ ИСТОЧНИКОВ",
    "СПИСОК ЛИТЕРАТУРЫ",
    "СПИСОК ИСПОЛЬЗОВАННОЙ ЛИТЕРАТУРЫ",
    "REFERENCES",
    "ПРИЛОЖЕНИЕ", "ПРИЛОЖЕНИЯ",
    "БЛАГОДАРНОСТИ",
}


# ---------------------------------------------------------------------------
# Определение типа параграфа
# ---------------------------------------------------------------------------

def _get_heading_level_by_style(para):
    """Заголовок по стилю Word (Heading 1/2/3)."""
    style_name = para.style.name if para.style else ""
    for level in (1, 2, 3):
        if style_name in (f"Heading {level}", f"Заголовок {level}"):
            return level
    return None


def _get_heading_level_by_content(para):
    """
    Определяет уровень заголовка по содержимому параграфа.
    Работает когда все параграфы написаны стилем Normal/None.
    """
    text = para.text.strip()
    if not text or len(text) < 2:
        return None

    # Проверяем жирность хотя бы одного run
    is_bold = False
    for run in para.runs:
        if run.font.bold:
            is_bold = True
            break

    if not is_bold:
        return None

    # Нумерованные заголовки: "1. ...", "1.1. ...", "1.1.1. ..."
    if re.match(r'^\d+\.\d+\.\d+', text):
        return 3
    if re.match(r'^\d+\.\d+', text):
        return 2
    if re.match(r'^\d+\.?\s+\S', text):
        return 1

    # Ненумерованные заголовки верхнего уровня (всё заглавными)
    text_upper = text.upper()
    if text_upper in UNNUMBERED_HEADINGS:
        return 1
    # Проверка что весь текст заглавными (минимум 4 символа)
    if len(text) >= 4 and text == text_upper and not text[0].isdigit():
        return 1

    return None


def _is_body_text(para) -> bool:
    style_name = para.style.name if para.style else ""
    body_styles = {"Normal", "Body Text", "Основной текст", ""}
    return style_name in body_styles or style_name.startswith("Normal")


def _is_table_caption(para) -> bool:
    text = para.text.strip().upper()
    return text.startswith("ТАБЛИЦА") or text.startswith("ТАБЛИЦЯ")


def _is_figure_caption(para) -> bool:
    text = para.text.strip()
    return (text.lower().startswith("рис.") or
            text.lower().startswith("рисунок") or
            text.lower().startswith("fig."))


def _is_list_item(para) -> bool:
    style_name = para.style.name if para.style else ""
    return ("List" in style_name or
            "Список" in style_name)


# ---------------------------------------------------------------------------
# Применение форматирования
# ---------------------------------------------------------------------------

def _set_font(run, bold=None, italic=None, caps=None):
    run.font.name = FONT_NAME
    run.font.size = FONT_SIZE
    if bold is not None:
        run.font.bold = bold
    if italic is not None:
        run.font.italic = italic
    if caps is not None:
        run.font.all_caps = caps
    run.font.color.rgb = RGBColor(0, 0, 0)


def _apply_heading(para, level: int):
    cfg = HEADING_CONFIG[level]
    fmt = para.paragraph_format
    fmt.alignment        = cfg["align"]
    fmt.first_line_indent = Pt(0)
    fmt.space_before     = cfg["space_before"]
    fmt.space_after      = cfg["space_after"]
    fmt.line_spacing_rule = WD_LINE_SPACING.ONE_POINT_FIVE
    fmt.page_break_before = cfg["page_break"]

    for run in para.runs:
        _set_font(run, bold=cfg["bold"], italic=cfg["italic"], caps=cfg["caps"])


def _apply_body(para):
    fmt = para.paragraph_format
    fmt.alignment         = WD_ALIGN_PARAGRAPH.JUSTIFY
    fmt.first_line_indent = FIRST_LINE_INDENT
    fmt.space_before      = Pt(0)
    fmt.space_after       = Pt(0)
    fmt.line_spacing_rule = WD_LINE_SPACING.ONE_POINT_FIVE
    fmt.page_break_before = False

    for run in para.runs:
        _set_font(run, bold=False, italic=False, caps=False)


def _apply_table_caption(para):
    fmt = para.paragraph_format
    fmt.alignment         = WD_ALIGN_PARAGRAPH.LEFT
    fmt.first_line_indent = Pt(0)
    fmt.space_before      = Pt(12)
    fmt.space_after       = Pt(3)
    fmt.line_spacing_rule = WD_LINE_SPACING.ONE_POINT_FIVE
    for run in para.runs:
        _set_font(run, bold=False, italic=False, caps=False)


def _apply_figure_caption(para):
    fmt = para.paragraph_format
    fmt.alignment         = WD_ALIGN_PARAGRAPH.CENTER
    fmt.first_line_indent = Pt(0)
    fmt.space_before      = Pt(3)
    fmt.space_after       = Pt(12)
    fmt.line_spacing_rule = WD_LINE_SPACING.ONE_POINT_FIVE
    for run in para.runs:
        _set_font(run, bold=False, italic=False, caps=False)


def _apply_list_item(para):
    fmt = para.paragraph_format
    fmt.line_spacing_rule = WD_LINE_SPACING.ONE_POINT_FIVE
    fmt.first_line_indent = Pt(0)
    fmt.left_indent       = Cm(1.25)
    fmt.space_before      = Pt(0)
    fmt.space_after       = Pt(0)
    for run in para.runs:
        _set_font(run, bold=False, italic=False, caps=False)


# ---------------------------------------------------------------------------
# Нумерация страниц
# ---------------------------------------------------------------------------

def _add_page_numbers(doc: Document):
    """Добавляет нумерацию страниц снизу по центру."""
    section = doc.sections[0]
    footer = section.footer
    para = footer.paragraphs[0] if footer.paragraphs else footer.add_paragraph()
    para.clear()
    para.alignment = WD_ALIGN_PARAGRAPH.CENTER

    run = para.add_run()
    run.font.name = FONT_NAME
    run.font.size = FONT_SIZE

    fld_begin = OxmlElement("w:fldChar")
    fld_begin.set(qn("w:fldCharType"), "begin")
    run._r.append(fld_begin)

    instr_run = para.add_run()
    instrText = OxmlElement("w:instrText")
    instrText.set(qn("xml:space"), "preserve")
    instrText.text = " PAGE "
    instr_run._r.append(instrText)

    fld_end_run = para.add_run()
    fld_end = OxmlElement("w:fldChar")
    fld_end.set(qn("w:fldCharType"), "end")
    fld_end_run._r.append(fld_end)


# ---------------------------------------------------------------------------
# Основная функция
# ---------------------------------------------------------------------------

def format_document(docx_bytes: bytes) -> bytes:
    """
    Принимает байты .docx, возвращает байты отформатированного .docx по ГОСТу.
    """
    doc = Document(BytesIO(docx_bytes))

    # 1. Поля страницы
    for section in doc.sections:
        section.left_margin   = MARGIN_LEFT
        section.right_margin  = MARGIN_RIGHT
        section.top_margin    = MARGIN_TOP
        section.bottom_margin = MARGIN_BOTTOM

    # 2. Обход всех параграфов
    for para in doc.paragraphs:
        if not para.text.strip():
            continue

        # Определяем тип: сначала по стилю Word, потом по содержимому
        heading_level = (_get_heading_level_by_style(para) or
                         _get_heading_level_by_content(para))

        if heading_level is not None:
            _apply_heading(para, heading_level)

        elif _is_list_item(para):
            _apply_list_item(para)

        elif _is_table_caption(para):
            _apply_table_caption(para)

        elif _is_figure_caption(para):
            _apply_figure_caption(para)

        elif _is_body_text(para):
            _apply_body(para)

        else:
            # Прочие стили — только шрифт и интервал
            para.paragraph_format.line_spacing_rule = WD_LINE_SPACING.ONE_POINT_FIVE
            for run in para.runs:
                run.font.name = FONT_NAME
                run.font.size = FONT_SIZE

    # 3. Нумерация страниц
    _add_page_numbers(doc)

    # 4. Сохраняем в байты
    buf = BytesIO()
    doc.save(buf)
    buf.seek(0)
    return buf.read()
