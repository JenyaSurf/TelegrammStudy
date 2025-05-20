# formatter/format_paragraphs.py

import re

def apply_paragraphs(doc):
    paragraphs = list(doc.paragraphs)
    prev_empty = False

    for p in paragraphs:
        text = p.text or ""

        # 1) Если это параграф с рисунком — пропускаем полностью
        if p._p.xpath('.//w:drawing'):
            prev_empty = False
            continue

        # 2) Если это заголовок «ВВЕДЕНИЕ», «ЗАКЛЮЧЕНИЕ», «ПРИЛОЖЕНИЕ» или номер раздела — пропускаем точку/удаление
        if re.match(r'^(ВВЕДЕНИЕ|ЗАКЛЮЧЕНИЕ|ПРИЛОЖЕНИЕ|\d+(\.\d+)*)', text.strip()):
            prev_empty = False
            continue

        # 3) Удаляем подряд идущие пустые абзацы
        if not text.strip():
            if prev_empty:
                p._p.getparent().remove(p._p)
                continue
            prev_empty = True
            continue
        prev_empty = False

        # 4) Поправляем лишние знаки в конце: если стоит ';' или ',' — заменяем на '.'
        #    и если нет никакой пунктуации — добавляем '.'
        if not re.search(r'[\.!?\u2026]$', text.strip()):
            # убираем любые ; или , на конце
            cleaned = re.sub(r'[;,]+$', '', text.strip())
            p.text = cleaned + '.'
