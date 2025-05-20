# formatter/format_toc.py

from docx.oxml import OxmlElement
from docx.oxml.ns import qn

def apply_toc(doc):
    """
    Вставляет поле оглавления в начало документа.
    GPT-стайл: генерирует TOC based on Heading 1..4.
    """
    # Создаём параграф в начале
    p = doc.add_paragraph()
    fld = OxmlElement('w:fldSimple')
    fld.set(qn('w:instr'), 'TOC \\o "1-4" \\h \\z \\u')
    p._p.addprevious(fld)
