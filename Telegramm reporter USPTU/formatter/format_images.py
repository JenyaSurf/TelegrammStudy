# formatter/format_images.py

from docx.enum.text import WD_ALIGN_PARAGRAPH

def apply_images(doc):
    pars = list(doc.paragraphs)
    for idx, p in enumerate(pars):
        # детектируем картинку через тег <w:drawing>
        if p._p.xpath('.//w:drawing'):
            # 1) центрируем сам абзац-картинку
            p.paragraph_format.alignment = WD_ALIGN_PARAGRAPH.CENTER

            # 2) подпись — следующий параграф, если начинается на "Рисунок"
            if idx+1 < len(pars):
                cap = pars[idx+1]
                if cap.text.strip().startswith("Рисунок"):
                    pf2 = cap.paragraph_format
                    pf2.alignment      = WD_ALIGN_PARAGRAPH.CENTER
                    pf2.keep_together  = True
                    pf2.keep_with_next = True

                    # снимаем жирность
                    for run in cap.runs:
                        run.font.bold = False
                    # убираем точку в конце подписи
                    cap.text = cap.text.rstrip('.').rstrip()
