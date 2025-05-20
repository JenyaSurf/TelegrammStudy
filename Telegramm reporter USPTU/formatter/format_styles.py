# formatter/format_styles.py

from docx.shared import Cm, Pt
from docx.enum.text import WD_ALIGN_PARAGRAPH

def apply_styles(doc):
    # 1) Параметры страницы A4 + поля
    for section in doc.sections:
        section.page_height = Cm(29.7)
        section.page_width  = Cm(21)
        section.top_margin    = Cm(2)
        section.bottom_margin = Cm(2)
        section.left_margin   = Cm(2)
        section.right_margin  = Cm(1)

    # 2) Шрифт, отступ, интервалы, выравнивание
    for p in doc.paragraphs:
        pf = p.paragraph_format
        pf.first_line_indent = Cm(1.25)
        pf.line_spacing      = 1.5
        pf.alignment         = WD_ALIGN_PARAGRAPH.JUSTIFY

        for run in p.runs:
            run.font.name = "Times New Roman"
            run.font.size = Pt(14)
