import docx

# Open the document
doc = docx.Document()

# Add a new paragraph to the end of the document
new_paragraph = doc.add_paragraph('Science and technology have become essential aspects of our lives. Technology was a luxury at a point in time, but now it has become a necessity. It is impossible to survive without electricity, television, music systems, mobile phones, internet connections, etc. We start and end our day with technology. So it is indeed difficult to imagine our life without technology, but it should be used with caution. If we become too dependent on technology, it will end up being harmful to us and our health. Overuse of technology can also become self-destructive, so it is important everyone uses technology only when necessary.')

# Add two line breaks before and after the new paragraph
new_paragraph.text = '\n\n' + new_paragraph.text + '\n\n'

# Maintain the original left alignment and add indentation
paragraph_format = new_paragraph.paragraph_format
paragraph_format.first_line_indent = docx.shared.Inches(1.0)
paragraph_format.left_indent = docx.shared.Inches(0.5)
paragraph_format.alignment = docx.enum.text.WD_ALIGN_PARAGRAPH.LEFT

# Save the modified document
doc.save('modified_document.docx')
