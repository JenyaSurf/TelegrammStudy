# formatter/format_headers.py

import re
from docx.shared import Pt
from docx.enum.text import WD_ALIGN_PARAGRAPH

HEADER_RULES = [
    (r'^(ВВЕДЕНИЕ|ЗАКЛЮЧЕНИЕ|Введение|введение|заключение|Заключение)\b', True, 14, True,  WD_ALIGN_PARAGRAPH.CENTER,  'intro'),
    (r'^ПРИЛОЖЕНИЕ\s+[А-Я]\b',    True, 14, True,  WD_ALIGN_PARAGRAPH.CENTER,  'appendix'),
    (r'^\d+\s+',                  True, 14, True,  WD_ALIGN_PARAGRAPH.LEFT,    'section'),
    (r'^\d+\.\d+\s+',             True, 14, False, WD_ALIGN_PARAGRAPH.LEFT,    'subsection'),
    (r'^\d+\.\d+\.\d+\s+',        False,14, False, WD_ALIGN_PARAGRAPH.LEFT,    'item'),
    (r'^\d+\.\d+\.\d+\.\d+\s+',   False,14, False, WD_ALIGN_PARAGRAPH.LEFT,    'subitem'),
]

def apply_headers(doc):
    pars = list(doc.paragraphs)
    for idx, p in enumerate(pars):
        txt = p.text.strip()
        for pattern, bold, size, break_before, align, level in HEADER_RULES:
            if re.match(pattern, txt):
                pf = p.paragraph_format
                pf.alignment         = align
                pf.page_break_before = break_before

                # Удаляем пустой параграф перед
                if break_before and idx > 0:
                    prev = pars[idx-1]
                    if not prev.text.strip():
                        prev._p.getparent().remove(prev._p)

                # Для intro/appendix: вставляем ровно один пустой абзац после
                if level in ('intro', 'appendix'):
                    # если следующий не пустой — вставляем пустую строку
                    if idx+1 < len(pars) and pars[idx+1].text.strip():
                        p.insert_paragraph_after('')

                # Точка только у пунктов (item)
                base = p.text.rstrip('.').rstrip()
                p.text = base + ('.' if level == 'item' else '')

                # Шрифт и жирность
                for run in p.runs:
                    run.font.bold = bold
                    run.font.size = Pt(size)
                break
