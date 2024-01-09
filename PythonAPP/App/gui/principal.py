from PyQt6 import uic
from PyQt6.QtWidgets import QMessageBox
from SortByPriority import Window

class Principal():
    def __init__(self):
        self.principal = uic.loadUi("App/gui/main.ui")
        self.principal.show()
        self.principal.choose_btn.clicked.connect(Window.Consola)
