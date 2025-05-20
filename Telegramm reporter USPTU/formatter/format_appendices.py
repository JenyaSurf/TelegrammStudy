# formatter/format_appendices.py

import re
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.shared import Pt, Cm

def apply_appendices(doc):
    pars = list(doc.paragraphs)
    for idx, p in enumerate(pars):
        txt = p.text.strip().upper()
        # Ищем "ПРИЛОЖЕНИЕ X"
        if re.match(r'^ПРИЛОЖЕНИЕ [А-Я]\b', txt):
            # Новый раздел – на новой странице
            pf = p.paragraph_format
            pf.page_break_before = True
            pf.alignment = WD_ALIGN_PARAGRAPH.LEFT
            # Большими буквами, без жирности
            for run in p.runs:
                run.font.bold = False
                run.font.size = Pt(14)
            # Следующий параграф – название приложения
            if idx+1 < len(pars):
                cap = pars[idx+1]
                cap.paragraph_format.first_line_indent = Cm(0)
