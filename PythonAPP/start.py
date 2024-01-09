from PyQt6.QtWidgets import QApplication
from gui.principal import Principal
class Start():
    def __init__(self):
        self.app = QApplication([])
        self.principal = Principal()
        self.app.exec()


        