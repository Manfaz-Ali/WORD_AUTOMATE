import sys

from PyQt6.QtWidgets import *
from PyQt6.uic import loadUiType
import docx
from docx.shared import Inches,Pt
from docx.enum.table import WD_TABLE_ALIGNMENT
from docx.enum.table import WD_ROW_HEIGHT_RULE
from docx.enum.table import WD_ALIGN_VERTICAL
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.enum.section import WD_SECTION
from docx.oxml import OxmlElement
from docx.oxml.ns import qn
import pandas as pd
from docx.enum.style import WD_STYLE_TYPE
ui, _ = loadUiType("MAIN1.ui")





class MainApp(QMainWindow, ui):
    def __init__(self, parent=None):
        super(MainApp, self).__init__(parent)
        QMainWindow.__init__(self)
        self.doc_name = None
        self.s_grade = None
        self.setupUi(self)
        self.button_actions()

    
    def button_actions(self):
        self.pushButton_paragraph.clicked.connect(self.my_functions_group)
        # Connect the button to a function to perform some action
        self.pushButton.clicked.connect(self.button_clicked)
        self.radioButton1.clicked.connect(self.radio_button_clicked)
        self.radioButton2.clicked.connect(self.radio_button_clicked)
        self.radioButton3.clicked.connect(self.radio_button_clicked)
        self.radioButton4.clicked.connect(self.radio_button_clicked)
        self.radioButton5.clicked.connect(self.radio_button_clicked)
        

    def my_functions_group(self):
        #self.get_security_grade()
        #self.get_document_name()
        self.make_new_doc()

    def radio_button_clicked(self):
        
        if self.radioButton1.isChecked():
            return 1
        elif self.radioButton2.isChecked():
            return 2
        elif self.radioButton3.isChecked():
            return 3
        elif self.radioButton4.isChecked():
            return 4
        else:
            return 5

    
    
    
    

    def button_clicked(self):
        # Perform some action when the button is clicked
        print("Button clicked!")



    def get_document_name(self):
        self.doc_name = self.lineEdit_Document_Name.text()
        self.doc_name = self.doc_name.upper()
        return self.doc_name

    def get_security_grade(self):
        self.s_grade = self.lineEdit_Security_Grade.text()
        self.s_grade = self.s_grade.upper()
        print(self.s_grade)
        return self.s_grade


    def set_pageMargin(self,doc):
        sections = doc.sections
        for section in sections:
            section.top_margin = Inches(0.5)
            section.bottom_margin = Inches(0.5)
            section.left_margin = Inches(1.5)
            section.right_margin = Inches(0.5)

    def set_header(self,doc):
        header = doc.sections[0].header

        print(header.is_linked_to_previous)
        # ui_hdr_doc_nam = self.get_document_name()
        ui_hdr_doc_nam = "my new document for test".upper()
        print(ui_hdr_doc_nam)
        # ui_hdr_s_grd = self.get_security_grade()
        ui_hdr_s_grd = "confidential".upper()
        table = header.add_table(rows=1, cols=3, width=Inches(8.01))
        table.style = 'Table Grid'
        table.alignment = WD_TABLE_ALIGNMENT.CENTER
        # Resize the table to fit the header
        # for all rows
        #--------
        for row in table.rows:
            for cell in row.cells:
                cell.width = docx.shared.Inches(2.67)
        # first cell
        first_cell = table.cell(0, 0)
        first_cell.text = ui_hdr_doc_nam
        
        first_cell.paragraphs[0].paragraph_format.alignment = WD_ALIGN_PARAGRAPH.CENTER
        first_cell.paragraphs[0].style.font.bold = False
        first_cell.vertical_alignment = WD_ALIGN_VERTICAL.CENTER
        self.set_cell_font(first_cell,12)
        
        
        
        # second cell
        second_cell = table.cell(0, 1)
        second_cell.text = ui_hdr_s_grd
        second_cell.paragraphs[0].paragraph_format.alignment = WD_ALIGN_PARAGRAPH.CENTER
        second_cell.vertical_alignment = WD_ALIGN_VERTICAL.CENTER
        second_cell.paragraphs[0].style.font.bold = False
        self.set_cell_font(second_cell,12)
        # third cell
        third_cell = table.cell(0, 2)
        run = third_cell.paragraphs[0].add_run()
        picture = run.add_picture("pic.jpg")
        picture.width = docx.shared.Inches(0.5)
        picture.height = docx.shared.Inches(0.25)
        third_cell.paragraphs[0].paragraph_format.alignment = WD_ALIGN_PARAGRAPH.CENTER
        third_cell.vertical_alignment = WD_ALIGN_VERTICAL.CENTER

        # header from top
        doc.sections[0].header_distance = Inches(0.5)
        #header row height
        for row in table.rows:
            row.height = Inches(0.5)
        
        
        
        
        print("header done")
    
    def set_footer(self,doc):
        # Set the footer of the document
        section = doc.sections[-1]
        footer = section.footer

        # Create a table with one row and three columns
        table = footer.add_table(rows=2, cols=4, width=Inches(8))
        table.alignment = WD_TABLE_ALIGNMENT.CENTER
        table.style = 'Table Grid'
        # Add content to the cells
        cell1 = table.cell(0, 0)
        cell1.text = 'Document No'
        cell1.paragraphs[0].runs[0].bold = False
        cell1.paragraphs[0].alignment = WD_ALIGN_PARAGRAPH.CENTER
        cell1.vertical_alignment = WD_ALIGN_VERTICAL.CENTER
        cell1.width = Inches(2.5)
        self.set_cell_font(cell1,12)

        cell2 = table.cell(0, 1)
        cell2.text = 'Rev No'
        cell2.paragraphs[0].runs[0].bold = False
        cell2.paragraphs[0].alignment = WD_ALIGN_PARAGRAPH.CENTER
        cell2.vertical_alignment = WD_ALIGN_VERTICAL.CENTER
        cell2.width = Inches(1.5)
        self.set_cell_font(cell2,12)

        cell3 = table.cell(0, 2)
        cell3.text = 'Date'
        cell3.paragraphs[0].runs[0].bold = False
        cell3.paragraphs[0].alignment = WD_ALIGN_PARAGRAPH.CENTER
        cell3.vertical_alignment = WD_ALIGN_VERTICAL.CENTER
        cell3.width = Inches(2)
        self.set_cell_font(cell3, 12)

        cell4 = table.cell(0, 3)
        cell4.text = 'Page'
        cell4.paragraphs[0].runs[0].bold = False
        cell4.paragraphs[0].alignment = WD_ALIGN_PARAGRAPH.CENTER
        cell4.vertical_alignment = WD_ALIGN_VERTICAL.CENTER
        cell4.width = Inches(2.5)
        self.set_cell_font(cell4, 12)

        cell5 = table.cell(1, 0)
        doc_ref = 'dfg/opt/1-fgt-7-987'.upper()
        cell5.text = doc_ref
        cell5.paragraphs[0].runs[0].bold = False
        cell5.paragraphs[0].alignment = WD_ALIGN_PARAGRAPH.CENTER
        cell5.vertical_alignment = WD_ALIGN_VERTICAL.CENTER
        cell5.width = Inches(2)
        self.set_cell_font(cell5, 12)

        cell6 = table.cell(1, 1)
        doc_rev = '0'
        cell6.text = doc_rev
        cell6.paragraphs[0].runs[0].bold = False
        cell6.paragraphs[0].alignment = WD_ALIGN_PARAGRAPH.CENTER
        cell6.vertical_alignment = WD_ALIGN_VERTICAL.CENTER
        cell6.width = Inches(1.5)
        self.set_cell_font(cell6, 12)

        cell7 = table.cell(1, 2)
        doc_date = '25-01-2023'.upper()
        cell7.text = doc_date
        cell7.paragraphs[0].runs[0].bold = False
        cell7.paragraphs[0].alignment = WD_ALIGN_PARAGRAPH.CENTER
        cell7.vertical_alignment = WD_ALIGN_VERTICAL.CENTER
        cell7.width = Inches(2)
        self.set_cell_font(cell7, 12)

        cell8 = table.cell(1, 3)
        # Add a run to the third cell to display the page numbers
        # run = cell3.paragraphs[0].add_run()
        # run.text = 'Page '
        # field = OxmlElement('w:fldSimple')
        # field.set(qn('w:instr'), 'PAGE')
        # run._r.append(field)
        # run.add_text(' of ')
        # field = OxmlElement('w:fldSimple')
        # field.set(qn('w:instr'), 'NUMPAGES')
        # run._r.append(field)
        # run = cell3.paragraphs[0].add_run()
        # field = OxmlElement('w:fldSimple')
        # field.set(qn('w:instr'), 'PAGE')
        # run._r.append(field)
        # run.add_text('/')
        # field = OxmlElement('w:fldSimple')
        # field.set(qn('w:instr'), 'NUMPAGES')
        # run._r.append(field)
        # font = run.font
        # font.bold = True
        # font.size = Pt(14)
        run = cell8.paragraphs[0].add_run()
        field = OxmlElement('w:fldSimple')
        field.set(qn('w:instr'), 'PAGE')
        run._r.append(field)
        run.add_text(' of ')
        field = OxmlElement('w:fldSimple')
        field.set(qn('w:instr'), 'NUMPAGES')
        run._r.append(field)
        run.font.bold = False
        run.font.size = Pt(12)
        cell8.paragraphs[0].alignment = WD_ALIGN_PARAGRAPH.CENTER
        cell8.vertical_alignment = WD_ALIGN_VERTICAL.CENTER
        self.set_cell_font(cell8,12)



        paragraph = footer.add_paragraph()
        p = "confidential".upper()
        footer_run = paragraph.add_run(p)
        footer_run.bold = False
        footer_run.font.size = Pt(12)
    
        paragraph.alignment = WD_ALIGN_PARAGRAPH.CENTER
        paragraph_format = paragraph.paragraph_format
        paragraph_format.space_before = Inches(0.2)
        paragraph_format.space_after = Inches(0.2)

        # Set the row height to 1 inch
        # row = table.rows[0]
        # row.height = Inches(0.5)

        # header from top
        doc.sections[0].footer_distance = Inches(0.5)
        # footer row height
        for row in table.rows:
            row.height = Inches(0.5)
        
        
        
        
        print("footer done")


    def add_paragraph(self,doc):
        check = self.radio_button_clicked()
        if check == 1:
            new_paragraph = "Science and technology have become essential aspects of our lives. Technology was a luxury at a point in time, but now it has become a necessity. It is impossible to survive without electricity, television, music systems, mobile phones, internet connections, etc. We start and end our day with technology. So it is indeed difficult to imagine our life without technology, but it should be used with caution. If we become too dependent on technology, it will end up being harmful to us and our health. Overuse of technology can also become self-destructive, so it is important everyone uses technology only when necessary."
            new_paragraph = '\n\n\t' + new_paragraph + '\n\n'
        elif check==2:



        doc.add_paragraph(new_paragraph).style.font.bold = False
        style = doc.styles['Normal']
        font = style.font
        font.name = 'Arial'
        font.size = Pt(12)
        
    
    def set_cell_font(self,cell,size):
        cell.paragraphs[0].runs[0].font.size = Pt(size)

    def add_table(self,doc):
        data = pd.read_csv("data.csv")
        num_rows, num_cols = data.shape
        table = doc.add_table(rows=num_rows+1, cols=num_cols)
        table.style = "Table Grid"
        # Center align the table
        table.alignment = WD_TABLE_ALIGNMENT.CENTER

        # Add the column headings to the first row of the table
        heading_row = table.rows[0]
        for i in range(num_cols):
            heading_cell = heading_row.cells[i]
            heading_cell.text = data.columns[i]
            heading_cell.paragraphs[0].runs[0].bold = True

        # Add the data to the table
        for i in range(num_rows):
            row_data = data.iloc[i]
            row = table.rows[i+1]
            for j in range(num_cols):
                value = str(row_data[j])
                cell = row.cells[j]
                cell.text = value
                if j == 0:
                    # Set the font size and bold the text in the first column
                    cell.paragraphs[0].runs[0].font.size = Pt(12)
                    cell.paragraphs[0].runs[0].bold = True


    

    def make_new_doc(self):
        doc = docx.Document()
        self.set_pageMargin(doc)
        self.set_header(doc)
        self.add_paragraph(doc)
        self.set_footer(doc)
        self.add_table(doc)

        doc.save("example.docx")











def main():
    app = QApplication(sys.argv)
    window = MainApp()
    window.show()
    app.exec()


if __name__ == "__main__":
    main()