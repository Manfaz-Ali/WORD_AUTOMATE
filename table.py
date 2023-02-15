from docx import Document
doc = Document()
table = doc.add_table(rows=1,cols=3)
row = table.rows[0].cells
doc_name = 'My Document'
row[0].text = doc_name
row[1].text = doc_name
row[2].text = doc_name
table.style = 'Table Grid'
doc.save('Sample1.docx')