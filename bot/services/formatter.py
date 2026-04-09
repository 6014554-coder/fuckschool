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
- Заголовки 3 уровня: 14 pt, курсив, по левому краю, без отступа
- Нумерация страниц: снизу по центру, без точки
"""

from io import BytesIO
from copy import deepcopy

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
    1: {"bold": True,  "italic": False, "caps": True,  "align": WD_ALIGN_PARAGRAPH.CENTER},
    2: {"bold": True,  "italic": False, "caps": False, "align": WD_ALIGN_PARAGRAPH.LEFT},
    3: {"bold": False, "italic": True,  "caps": False, "align": WD_ALIGN_PARAGRAPH.LEFT},
}


# ---------------------------------------------------------------------------
# Вспомогательные функции
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
    # Убираем цвет (возвращаем авто)
    run.font.color.rgb = RGBColor(0, 0, 0)


def _get_heading_level(para):
    """Возвращает уровень заголовка (1-3) или None если не заголовок."""
    style_name = para.style.name if para.style else ""
    for level in (1, 2, 3):
        if style_name in (f"Heading {level}", f"Заголовок {level}"):
            return level
    return None


def _is_body_text(para) -> bool:
    style_name = para.style.name if para.style else ""
    body_styles = {"Normal", "Body Text", "Основной текст", ""}
    return style_name in body_styles or style_name.startswith("Normal")


# ---------------------------------------------------------------------------
# Нумерация страниц
# ---------------------------------------------------------------------------

def _add_page_numbers(doc: Document):
    """Добавляет нумерацию страниц снизу по центру."""
    for section in doc.sections:
        footer = section.footer
        if not footer.paragraphs:
            para = footer.add_paragraph()
        else:
            para = footer.paragraphs[0]

        para.clear()
        para.alignment = WD_ALIGN_PARAGRAPH.CENTER

        run = para.add_run()
        run.font.name = FONT_NAME
        run.font.size = FONT_SIZE

        # XML-поле PAGE
        for tag, fld_type in [("begin", None), ("separate", None), ("end", None)]:
            fld = OxmlElement("w:fldChar")
            fld.set(qn("w:fldCharType"), tag)
            run._r.append(fld)

        # instrText между begin и separate
        instr_run = para.add_run()
        instr = OxmlElement("w:instrText")
        instr.set(qn("xml:space"), "preserve")
        instr.text = " PAGE "
        instr_run._r.insert(0, instr)

        # Убираем старые колонтитулы из run
        # Пересобираем правильно
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

        break  # достаточно первой секции для MVP


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

        heading_level = _get_heading_level(para)

        fmt = para.paragraph_format
        fmt.line_spacing_rule = WD_LINE_SPACING.ONE_POINT_FIVE
        fmt.space_before = Pt(0)
        fmt.space_after  = Pt(0)

        if heading_level is not None:
            cfg = HEADING_CONFIG.get(heading_level, HEADING_CONFIG[1])
            fmt.alignment = cfg["align"]
            fmt.first_line_indent = Pt(0)

            # Заголовок 1 уровня — всегда с новой страницы
            if heading_level == 1:
                fmt.page_break_before = True
            else:
                fmt.page_break_before = False

            for run in para.runs:
                _set_font(
                    run,
                    bold=cfg["bold"],
                    italic=cfg["italic"],
                    caps=cfg["caps"],
                )

        elif _is_body_text(para):
            fmt.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY
            fmt.first_line_indent = FIRST_LINE_INDENT

            for run in para.runs:
                _set_font(run, bold=False, italic=False, caps=False)

        else:
            # Прочие стили (подписи, цитаты и т.д.) — только шрифт
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
