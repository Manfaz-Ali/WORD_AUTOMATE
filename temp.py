from docx import Document

document = Document()

for paragraph in document.paragraphs:
    # Set line spacing after each paragraph
    paragraph.paragraph_format.line_spacing = 0.0
