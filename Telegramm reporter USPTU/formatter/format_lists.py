import re
from docx.shared import Cm, Pt
from docx.oxml.ns import qn

LIST_RULES = [
    (re.compile(r'^–\s+'),               Cm(1.25)),  # уровень 1
    (re.compile(r'^\([а-яё]\)\s+', re.I), Cm(2.5)),   # уровень 2
    (re.compile(r'^\(\d+\)\s+'),         Cm(3.75)),  # уровень 3
]

def _remove_word_numbering(p):
    pPr = p._p.get_or_add_pPr()
    numPr = pPr.find(qn('w:numPr'))
    if numPr is not None:
        pPr.remove(numPr)

def apply_lists(doc):
    paras = list(doc.paragraphs)
    for idx, p in enumerate(paras):
        raw = p.text or ""
        for pattern, indent in LIST_RULES:
            m = pattern.match(raw)
            if not m:
                continue

            # 1) Сбрасываем автоматическую нумерацию Word
            _remove_word_numbering(p)

            # 2) Жёсткие отступы
            fmt = p.paragraph_format
            fmt.left_indent       = indent
            fmt.first_line_indent = Cm(0)

            # 3) Разбираем маркер и тело пункта
            marker = m.group(0)
            core   = raw[len(marker):]
            # Склеиваем переносы строк в пробел
            core = re.sub(r'\s+', ' ', core).strip()
            # Первая буква — заглавная
            if core:
                core = core[0].upper() + core[1:]

            # 4) Признак последнего пункта
            is_last = True
            if idx + 1 < len(paras):
                nxt = paras[idx+1].text or ""
                if any(pr.match(nxt) for pr, _ in LIST_RULES):
                    is_last = False

            punct = '.' if is_last else ';'
            p.text = f"{marker}{core.rstrip(';. ')}{punct}"

            # 5) Настраиваем шрифт для всего параграфа
            for run in p.runs:
                run.font.name = "Times New Roman"
                run.font.size = Pt(14)

            break
