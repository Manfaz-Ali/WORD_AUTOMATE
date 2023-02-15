import sys

from PyQt6.QtWidgets import *
from PyQt6.uic import loadUiType
import docx
from docx.shared import Inches,Pt
from docx.enum.table import WD_TABLE_ALIGNMENT
from docx.enum.table import WD_ROW_HEIGHT_RULE
from docx.enum.table import WD_ALIGN_VERTICAL
from docx.enum.text import WD_ALIGN_PARAGRAPH
ui, _ = loadUiType("WORD.ui")





class MainApp(QMainWindow, ui):
    def __init__(self, parent=None):
        super(MainApp, self).__init__(parent)
        QMainWindow.__init__(self)
        self.doc_name = None
        self.s_grade = None
        self.setupUi(self)
        self.button_actions()

    def button_actions(self):
        self.pushButton_HeaderSubmit.clicked.connect(self.my_functions_group)

    def my_functions_group(self):
        self.get_security_grade()
        self.get_document_name()
        self.make_new_doc()


    def get_document_name(self):
        self.doc_name = self.lineEdit_Document_Name.text()
        self.doc_name = self.doc_name.upper()
        return self.doc_name

    def get_security_grade(self):
        self.s_grade = self.lineEdit_Security_Grade.text()
        self.s_grade = self.s_grade.upper()
        print(self.s_grade)
        return self.s_grade






    def make_new_doc(self):
        doc = docx.Document()
        header = doc.sections[0].header


        ui_hdr_doc_nam = self.get_document_name()
        print(ui_hdr_doc_nam)
        ui_hdr_s_grd = self.get_security_grade()
        table = header.add_table(rows=1, cols=3, width=Inches(6.0))
        table.style = 'Table Grid'
        table.alignment = WD_TABLE_ALIGNMENT.CENTER
        # Resize the table to fit the header
        # for all rows
        #--------
        for row in table.rows:
            for cell in row.cells:
                cell.width = docx.shared.Inches(2)
        # first cell
        first_cell = table.cell(0, 0)
        if ui_hdr_doc_nam is None:
            first_cell.text = "Document Name".upper()
        else:
            first_cell.text = ui_hdr_doc_nam
        first_cell.paragraphs[0].paragraph_format.alignment = WD_ALIGN_PARAGRAPH.CENTER
        first_cell.paragraphs[0].style.font.bold = True
        # second cell
        second_cell = table.cell(0, 1)
        if ui_hdr_s_grd is None:
            second_cell.text = "Security Grade".upper()
        else:
            second_cell.text = ui_hdr_s_grd
        second_cell.paragraphs[0].paragraph_format.alignment = WD_ALIGN_PARAGRAPH.CENTER
        second_cell.paragraphs[0].style.font.bold = True
        # third cell
        third_cell = table.cell(0, 2)
        run = third_cell.paragraphs[0].add_run()
        picture = run.add_picture("pic.jpg")
        picture.width = docx.shared.Inches(1)
        picture.height = docx.shared.Inches(0.5)
        third_cell.paragraphs[0].paragraph_format.alignment = WD_ALIGN_PARAGRAPH.CENTER


        # header from top
        doc.sections[0].header_distance = Inches(0.83)
        # header row height
        for row in table.rows:
            row.height = Inches(1)
        #first_cell.vertical_alignment = WD_ALIGN_VERTICAL.TOP
        first_cell.vertical_alignment = WD_ALIGN_VERTICAL.CENTER
        second_cell.vertical_alignment = WD_ALIGN_VERTICAL.CENTER
        third_cell.vertical_alignment = WD_ALIGN_VERTICAL.CENTER
        doc.save("example.docx")











def main():
    app = QApplication(sys.argv)
    window = MainApp()
    window.show()
    app.exec()


if __name__ == "__main__":
    main()