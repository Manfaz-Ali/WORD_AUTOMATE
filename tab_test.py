from PyQt5.QtWidgets import QApplication, QMainWindow
from PyQt5.uic import loadUi

class MyWindow(QMainWindow):
    def __init__(self):
        super().__init__()

        # Load the UI file
        loadUi('my_ui_file.ui', self)

        # Connect the button to a function to perform some action
        self.pushButton.clicked.connect(self.button_clicked)

    def button_clicked(self):
        # Perform some action when the button is clicked
        print("Button clicked!")

if __name__ == '__main__':
    # Create a QApplication instance
    app = QApplication([])

    # Create an instance of the window and show it
    window = MyWindow()
    window.show()

    # Run the event loop
    app.exec_()
