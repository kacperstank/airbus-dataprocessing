from PyQt5.QtCore import QThread, pyqtSignal
from data_generator import DataGenerator

class FileGenerationThread(QThread):
    """
    Thread to generate an Excel file with dummy data.
    """
    file_generated = pyqtSignal()  # Signal emitted when file generation is complete

    def __init__(self, filename="data.xlsx"):
        super().__init__()
        self.filename = filename

    def run(self):
        data_gen = DataGenerator(self.filename)
        data_gen.create_excel_file()
        self.file_generated.emit()  # Emit the signal when done