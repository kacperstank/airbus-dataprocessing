from PyQt5.QtCore import QThread, pyqtSignal
from data_processor import DataProcessor

class DataProcessingThread(QThread):
    """
    Thread to process data in an Excel file.
    """
    processing_complete = pyqtSignal(dict, float) # Signal emitted when data processing is complete

    def __init__(self, filename, method="openpyxl"):
        super().__init__()
        self.filename = filename
        self.method = method

    def run(self):
        processor = DataProcessor(self.filename)

        if self.method == "openpyxl":
            counts, time_taken = processor.count_occurrences(processor.extract_unique_values())
        elif self.method == "pandas":
            counts, time_taken = processor.process_with_pandas()
        elif self.method == "xlsxio":
            counts, time_taken = processor.process_with_xlsxio()
        else:
            counts, time_taken = {}, 0  # Placeholder if we add more methods

        # Emit both results and time taken
        self.processing_complete.emit(counts, time_taken)