# formatter/format_formulas.py

from docx.shared import Pt
from docx.enum.text import WD_ALIGN_PARAGRAPH

def apply_formulas(doc):
    # Пока python-docx не поддерживает Equation-объекты,
    # будем искать параграфы вида "(X.X)" и проставлять формат:
    for p in doc.paragraphs:
        txt = p.text.strip()
        if txt.startswith("(") and txt.endswith(")"):
            # Выравнять вправо
            pf = p.paragraph_format
            pf.alignment = WD_ALIGN_PARAGRAPH.RIGHT
            # Шрифт формулы
            for run in p.runs:
                run.font.size = Pt(14)
                run.font.name = "Times New Roman"
            # Ищем сразу следующий параграф, начинающийся на "где"
            # и форматируем его как нормальный текст с отступом
            # TODO: добавьте логику для "где …"
