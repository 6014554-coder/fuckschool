"""
Анализатор примерного документа.

Извлекает параметры форматирования из .docx-образца и возвращает
config-словарь, который можно передать в format_document().
"""

import re
from io import BytesIO
from collections import Counter

from docx import Document
from docx.shared import Pt, Mm, Cm
from docx.enum.text import WD_ALIGN_PARAGRAPH, WD_LINE_SPACING
from docx.oxml.ns import qn


# ---------------------------------------------------------------------------
# Константы конвертации
# ---------------------------------------------------------------------------

EMU_PER_PT  = 12700   # 1 пункт = 12700 EMU
EMU_PER_CM  = 360000  # 1 см = 360000 EMU
EMU_PER_MM  = 36000   # 1 мм = 36000 EMU


# ---------------------------------------------------------------------------
# Вспомогательные функции
# ---------------------------------------------------------------------------

def _safe_first_line_indent_emu(para):
    """
    Безопасно читает first_line_indent из XML (в EMU).
    Обходит баг python-docx при дробных значениях twips.
    """
    try:
        pPr = para._p.find(qn("w:pPr"))
        if pPr is None:
            return None
        ind = pPr.find(qn("w:ind"))
        if ind is None:
            return None
        # firstLine или hanging в твипах (1 twip = 635 EMU; 1 pt = 20 twips)
        fl = ind.get(qn("w:firstLine"))
        if fl is not None:
            twips = float(fl)
            return int(twips * 635)
        hang = ind.get(qn("w:hanging"))
        if hang is not None:
            return -int(float(hang) * 635)
        return None
    except Exception:
        return None


def _safe_left_indent_emu(para):
    """Читает left indent из XML (EMU)."""
    try:
        pPr = para._p.find(qn("w:pPr"))
        if pPr is None:
            return None
        ind = pPr.find(qn("w:ind"))
        if ind is None:
            return None
        left = ind.get(qn("w:left"))
        if left is not None:
            return int(float(left) * 635)
        return None
    except Exception:
        return None


def _safe_alignment(para):
    """Читает выравнивание без исключений."""
    try:
        return para.paragraph_format.alignment
    except Exception:
        return None


def _safe_line_spacing(para):
    """Возвращает (rule, value) или None."""
    try:
        fmt = para.paragraph_format
        return fmt.line_spacing_rule, fmt.line_spacing
    except Exception:
        return None, None


def _get_run_font_size_pt(run):
    """Возвращает размер шрифта в пунктах через XML (sz — полуточки)."""
    try:
        rpr = run._r.find(qn("w:rPr"))
        if rpr is not None:
            sz = rpr.find(qn("w:sz"))
            if sz is not None:
                val = sz.get(qn("w:val"))
                if val:
                    return float(val) / 2.0  # полуточки → точки
    except Exception:
        pass
    # Fallback: через python-docx API
    if run.font.size:
        return run.font.size / EMU_PER_PT
    return None


def _get_para_font_name(para):
    """Возвращает имя шрифта из первого run (или из XML)."""
    for run in para.runs:
        name = run.font.name
        if name:
            return name
        # Через XML
        try:
            rpr = run._r.find(qn("w:rPr"))
            if rpr is not None:
                rFonts = rpr.find(qn("w:rFonts"))
                if rFonts is not None:
                    ascii_font = rFonts.get(qn("w:ascii"))
                    if ascii_font:
                        return ascii_font
        except Exception:
            pass
    return None


def _is_bold(para):
    return any(run.font.bold for run in para.runs)


def _is_all_caps_de_facto(para):
    """Проверяет: весь текст — заглавные буквы (шрифт или содержание)."""
    for run in para.runs:
        if run.font.all_caps:
            return True
    text = para.text.strip()
    return len(text) >= 3 and text == text.upper() and any(c.isalpha() for c in text)


def _heading_level_by_style(para):
    style_name = para.style.name if para.style else ""
    for lvl in (1, 2, 3):
        if style_name in (f"Heading {lvl}", f"Заголовок {lvl}"):
            return lvl
    return None


def _heading_level_by_content(para):
    text = para.text.strip()
    if not text or len(text) < 2:
        return None
    if not _is_bold(para):
        return None
    if re.match(r'^\d+\.\d+\.\d+', text):
        return 3
    if re.match(r'^\d+\.\d+', text):
        return 2
    if re.match(r'^\d+\.?\s+\S', text):
        return 1
    text_up = text.upper()
    if len(text) >= 4 and text == text_up and not text[0].isdigit():
        return 1
    return None


def _is_title_page_para(para, idx):
    """
    Эвристически определяет: похоже ли на обложку (первые 15 параграфов,
    текст очень короткий или ALL-CAPS без нумерации).
    """
    if idx < 15:
        return True
    text = para.text.strip()
    # Строки-разделители или пустые после очистки
    if len(text) < 5:
        return True
    return False


# ---------------------------------------------------------------------------
# Основная функция анализа
# ---------------------------------------------------------------------------

def analyze_example(docx_bytes: bytes) -> dict:
    """
    Анализирует образцовый .docx и возвращает config-словарь.

    Структура:
    {
        "font_name":    str,
        "font_size_pt": float,
        "margins": {"left_mm", "right_mm", "top_mm", "bottom_mm"},
        "body": {
            "first_line_indent_cm": float,
            "alignment":            WD_ALIGN_PARAGRAPH,
            "space_before_pt":      float,
            "space_after_pt":       float,
            "line_spacing":         WD_LINE_SPACING,
        },
        "headings": {
            1: {bold, italic, caps, alignment, first_line_indent_cm,
                space_before_pt, space_after_pt, page_break},
            2: {...},
            3: {...},
        }
    }
    """
    doc = Document(BytesIO(docx_bytes))

    # ---- Поля страницы -----------------------------------------------
    margins = {"left_mm": 30.0, "right_mm": 10.0, "top_mm": 20.0, "bottom_mm": 20.0}
    if doc.sections:
        sec = doc.sections[0]
        def _mm(v):
            return round(v / EMU_PER_MM, 1) if v else 0.0
        margins = {
            "left_mm":   _mm(sec.left_margin),
            "right_mm":  _mm(sec.right_margin),
            "top_mm":    _mm(sec.top_margin),
            "bottom_mm": _mm(sec.bottom_margin),
        }

    # ---- Классификация параграфов -----------------------------------
    body_paras    = []
    heading_paras = {1: [], 2: [], 3: []}

    for idx, para in enumerate(doc.paragraphs):
        text = para.text.strip()
        if not text:
            continue
        lvl = (_heading_level_by_style(para) or _heading_level_by_content(para))
        if lvl in (1, 2, 3):
            heading_paras[lvl].append(para)
        elif not _is_title_page_para(para, idx):
            body_paras.append(para)

    # ---- Шрифт -------------------------------------------------------
    font_sizes = []
    font_names = []
    for para in body_paras[:80]:
        for run in para.runs:
            sz = _get_run_font_size_pt(run)
            if sz and 6 <= sz <= 24:
                font_sizes.append(round(sz * 2) / 2)
            nm = _get_para_font_name(para)
            if nm:
                font_names.append(nm)
                break  # достаточно одного имени с параграфа

    font_size_pt = Counter(font_sizes).most_common(1)[0][0] if font_sizes else 14.0
    font_name    = Counter(font_names).most_common(1)[0][0] if font_names else "Times New Roman"

    # ---- Параметры тела (body) --------------------------------------
    indent_vals  = []
    align_vals   = []

    for para in body_paras[:50]:
        emu = _safe_first_line_indent_emu(para)
        if emu is not None and emu >= 0:
            cm = emu / EMU_PER_CM
            if 0 <= cm <= 5:
                indent_vals.append(round(cm * 4) / 4)

        al = _safe_alignment(para)
        if al is not None:
            align_vals.append(al)

    first_line_indent_cm = Counter(indent_vals).most_common(1)[0][0] if indent_vals else 1.25
    body_alignment = Counter(align_vals).most_common(1)[0][0] if align_vals else WD_ALIGN_PARAGRAPH.JUSTIFY

    body_cfg = {
        "first_line_indent_cm": first_line_indent_cm,
        "alignment":            body_alignment,
        "space_before_pt":      0.0,
        "space_after_pt":       0.0,
        "line_spacing":         WD_LINE_SPACING.ONE_POINT_FIVE,
    }

    # ---- Параметры заголовков ----------------------------------------
    def _extract_heading_cfg(paras, level):
        if not paras:
            return None
        para = paras[0]

        bold   = _is_bold(para)
        italic = any(r.font.italic for r in para.runs if r.font.italic is not None)
        caps   = _is_all_caps_de_facto(para)
        align  = _safe_alignment(para) or WD_ALIGN_PARAGRAPH.LEFT

        emu = _safe_first_line_indent_emu(para)
        indent_cm = round(emu / EMU_PER_CM * 4) / 4 if emu and emu > 0 else 0.0

        # space_before/after через XML (w:pPr/w:spacing)
        sp_before = 0.0
        sp_after  = 0.0
        try:
            pPr = para._p.find(qn("w:pPr"))
            if pPr is not None:
                spacing = pPr.find(qn("w:spacing"))
                if spacing is not None:
                    before = spacing.get(qn("w:before"))
                    after  = spacing.get(qn("w:after"))
                    if before:
                        sp_before = round(float(before) / 20.0 * 2) / 2  # twips→pt
                    if after:
                        sp_after  = round(float(after)  / 20.0 * 2) / 2
        except Exception:
            pass

        page_break = False
        try:
            pPr = para._p.find(qn("w:pPr"))
            if pPr is not None:
                pb = pPr.find(qn("w:pageBreakBefore"))
                if pb is not None:
                    val = pb.get(qn("w:val"), "true")
                    page_break = val.lower() not in ("false", "0")
        except Exception:
            pass

        return {
            "bold":                 bold,
            "italic":               italic,
            "caps":                 caps,
            "alignment":            align,
            "first_line_indent_cm": indent_cm,
            "space_before_pt":      sp_before,
            "space_after_pt":       sp_after,
            "page_break":           page_break,
        }

    headings_cfg = {}
    for lvl in (1, 2, 3):
        cfg = _extract_heading_cfg(heading_paras[lvl], lvl)
        if cfg:
            headings_cfg[lvl] = cfg

    return {
        "font_name":    font_name,
        "font_size_pt": font_size_pt,
        "margins":      margins,
        "body":         body_cfg,
        "headings":     headings_cfg,
    }
