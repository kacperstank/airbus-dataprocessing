from PyQt5.QtWidgets import QMainWindow


class Vista_xlsxio (QMainWindow):
    def __init__(self):
        super().__init__()
        self.initUI()

    def initUI(self):
        self.setWindowTitle("XLSXIO")
        self.setGeometry(100, 100, 600, 400)