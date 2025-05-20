# formatter/format_main.py

from docx import Document
from .format_styles     import apply_styles
from .format_headers    import apply_headers
from .format_lists      import apply_lists
from .format_paragraphs import apply_paragraphs
from .format_tables     import apply_tables
from .format_images     import apply_images
from .format_formulas   import apply_formulas
from .format_appendices import apply_appendices
from .format_toc        import apply_toc

def process_document(input_path: str) -> str:
    doc = Document(input_path)

    apply_styles(doc)
    apply_headers(doc)
    apply_lists(doc)
    apply_paragraphs(doc)
    apply_tables(doc)
    apply_images(doc)
    apply_formulas(doc)
    apply_appendices(doc)
    apply_toc(doc)

    output_path = input_path.replace(".docx", "_formatted.docx")
    doc.save(output_path)
    return output_path
