from PyQt6.QtCore import Qt
from PyQt6.QtWidgets import QApplication, QMainWindow
from PyQt6.uic import loadUiType

# Load the UI file
Ui_MainWindow, _ = loadUiType("WORD1.ui")

class MainWindow(QMainWindow, Ui_MainWindow):
    def __init__(self):
        super().__init__()
        self.setupUi(self)
        
        # Connect the radio buttons to their respective functions
        self.radioButton1.clicked.connect(self.radio_button_1_clicked)
        self.radioButton2.clicked.connect(self.radio_button_2_clicked)
        self.radioButton3.clicked.connect(self.radio_button_3_clicked)
        self.radioButton4.clicked.connect(self.radio_button_4_clicked)
        self.radioButton5.clicked.connect(self.radio_button_5_clicked)
        self.radioButton6.clicked.connect(self.radio_button_6_clicked)
    
    def radio_button_1_clicked(self):
        # Do something when radio button 1 is clicked
        if self.radioButton1.isChecked():
            print("Radio button 1 is checked")
    
    def radio_button_2_clicked(self):
        # Do something when radio button 2 is clicked
        if self.radioButton2.isChecked():
            print("Radio button 2 is checked")
    
    def radio_button_3_clicked(self):
        # Do something when radio button 3 is clicked
        if self.radioButton3.isChecked():
            print("Radio button 3 is checked")
    
    def radio_button_4_clicked(self):
        # Do something when radio button 4 is clicked
        if self.radioButton4.isChecked():
            print("Radio button 4 is checked")
    
    def radio_button_5_clicked(self):
        # Do something when radio button 5 is clicked
        if self.radioButton5.isChecked():
            print("Radio button 5 is checked")
    
    def radio_button_6_clicked(self):
        # Do something when radio button 6 is clicked
        if self.radioButton6.isChecked():
            print("Radio button 6 is checked")

if __name__ == '__main__':
    app = QApplication([])
    window = MainWindow()
    window.show()
    app.exec()
