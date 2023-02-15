import docx
from docx.shared import Inches,Pt
from docx.enum.table import WD_TABLE_ALIGNMENT
from docx.enum.table import WD_ROW_HEIGHT_RULE
# Open the Word document
doc = docx.Document()
print(help(doc.styles['Normal']))
