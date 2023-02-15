from PyQt6.QtWidgets import QApplication, QMainWindow, QFileDialog
from PyQt6.uic import loadUi

class MyWindow(QMainWindow):
    def __init__(self):
        super().__init__()
        # Load the UI file
        loadUi('my_ui_file.ui', self)
        # Connect the "Browse" button to the browse_file function
        self.browse_button.clicked.connect(self.browse_file)
    
    def browse_file(self):
        # Open the file browsing dialog
        file_path, _ = QFileDialog.getOpenFileName(self, "Open CSV File", "", "CSV Files (*.csv);;All Files (*)")
        # Set the text of the file_path_textedit to the selected file path
        self.file_path_textedit.setText(file_path)
    
if __name__ == '__main__':
    app = QApplication([])
    window = MyWindow()
    window.show()
    app.exec()
