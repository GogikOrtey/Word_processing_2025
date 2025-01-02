from docx import Document
from docx.shared import Pt
from docx.enum.text import WD_PARAGRAPH_ALIGNMENT
from docx.oxml.ns import qn
from docx.oxml import OxmlElement


def modify_paragraph_style(paragraph):
    run = paragraph.runs[0]
    run.font.name = 'Times New Roman'
    run.font.bold = True
    run.font.size = Pt(16)
    paragraph.alignment = WD_PARAGRAPH_ALIGNMENT.CENTER

    # Set spacing before and after
    pPr = paragraph._element.get_or_add_pPr()
    spacing = OxmlElement('w:spacing')
    spacing.set(qn('w:before'), '240')  # 12pt before
    spacing.set(qn('w:after'), '60')    # 3pt after
    pPr.append(spacing)

def process_document(file_path):
    doc = Document(file_path)
    for i, paragraph in enumerate(doc.paragraphs):
        if i < 18 or i > 738:  # Skip paragraphs before page 19 and after page 739
            continue
        if paragraph.text.startswith('№'):
            paragraph.style = doc.styles['Heading 1']
            modify_paragraph_style(paragraph)
    doc.save('modified_document_1.docx')

# Path to your document
file_path = '7_1 Сборник бабушкиных стихов.docx'
process_document(file_path)

































