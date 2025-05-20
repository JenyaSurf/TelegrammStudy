import os
from win32com import client

def convert_to_docx(path: str) -> str:
    """
    Конвертирует .doc → .docx (Windows + MS Word должен быть установлен).
    Возвращает путь к новому .docx.
    """
    abs_path = os.path.abspath(path)
    word     = client.Dispatch("Word.Application")
    doc      = word.Documents.Open(abs_path)
    new_path = abs_path + "x"                 # добавляем "x" к имени
    doc.SaveAs(new_path, FileFormat=16)       # 16 = wdFormatDocumentDefault
    doc.Close()
    word.Quit()
    return new_path
