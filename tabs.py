from PyQt6.QtWidgets import QApplication, QMainWindow, QTabWidget, QVBoxLayout
from PyQt6.uic import loadUi

class MyWindow(QMainWindow):
    def __init__(self):
        super().__init__()
        # Load the UI file
        loadUi('tabs.ui', self)
        # Access the tabs
        self.tab1 # Replace "tab1" with the actual name of your first tab
        self.tab2 # Replace "tab2" with the actual name of your second tab
        self.tab3 # Replace "tab3" with the actual name of your third tab
        self.tab4 # Replace "tab4" with the actual name of your fourth tab
        self.tab5 # Replace "tab5" with the actual name of your fifth tab
        self.tab6 # Replace "tab6" with the actual name of your sixth tab
        self.tab7 # Replace "tab6" with the actual name of your sixth tab
        # Add the tabs to the tab widget
        self.tabWidget = QTabWidget(self)
        self.tabWidget.addTab(self.tab1, "Tab 1")
        self.tabWidget.addTab(self.tab2, "Tab 2")
        self.tabWidget.addTab(self.tab3, "Tab 3")
        self.tabWidget.addTab(self.tab4, "Tab 4")
        self.tabWidget.addTab(self.tab5, "Tab 5")
        self.tabWidget.addTab(self.tab6, "Tab 6")
        self.tabWidget.addTab(self.tab7, "Tab 7")
        # Add the tab widget to a layout
        layout = QVBoxLayout()
        layout.addWidget(self.tabWidget)
        # Set the layout for the main window
        self.setLayout(layout)
    
if __name__ == '__main__':
    app = QApplication([])
    window = MyWindow()
    window.show()
    app.exec()
