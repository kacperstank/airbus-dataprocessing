from PyQt5.QtCore import Qt
from PyQt5.QtWidgets import QMainWindow, QPushButton, QTextEdit, QVBoxLayout, QHBoxLayout, QWidget

from data_processing_thread import DataProcessingThread
from file_generation_thread import FileGenerationThread


class DataProcessorGUI(QMainWindow):
    """
    GUI for the data processor, providing file generation, data processing,
    and result display.
    """

    def __init__(self):
        super().__init__()
        self.initUI()

    def initUI(self):
        """
        Initialize GUI components and layout.
        """
        self.setWindowTitle("Data Processor")
        self.setGeometry(100, 100, 600, 400)

        # Main layout
        main_layout = QVBoxLayout()

        # Button to generate the Excel file
        self.generate_file_btn = QPushButton("Generate Excel File", self)
        self.generate_file_btn.clicked.connect(self.start_file_generation)
        self.generate_file_btn.setFixedWidth(200)
        main_layout.addWidget(self.generate_file_btn, alignment=Qt.AlignHCenter)

        # Horizontal layout for processing buttons
        processing_buttons_layout = QHBoxLayout()

        # Button to process data with openpyxl
        self.process_openpyxl_btn = QPushButton("Process with openpyxl", self)
        self.process_openpyxl_btn.clicked.connect(self.process_data_openpyxl)
        self.process_openpyxl_btn.setFixedWidth(180)
        self.process_openpyxl_btn.setEnabled(False)  # Initially disabled

        # Button to process data with pandas
        self.process_pandas_btn = QPushButton("Process with pandas", self)
        self.process_pandas_btn.clicked.connect(self.process_data_pandas)
        self.process_pandas_btn.setFixedWidth(180)
        self.process_pandas_btn.setEnabled(False)  # Initially disabled

        # Placeholder button for xlsxio
        self.process_xlsxio_btn = QPushButton("Process with xlsxio", self)
        self.process_xlsxio_btn.clicked.connect(self.process_data_xlsxio)
        self.process_xlsxio_btn.setFixedWidth(180)
        self.process_xlsxio_btn.setEnabled(False)  # Initially disabled

        # Add processing buttons to layout
        processing_buttons_layout.addWidget(self.process_openpyxl_btn)
        processing_buttons_layout.addWidget(self.process_pandas_btn)
        processing_buttons_layout.addWidget(self.process_xlsxio_btn)

        # Text area to display results
        self.result_display = QTextEdit(self)
        self.result_display.setReadOnly(True)

        # Add layouts to main layout
        main_layout.addLayout(processing_buttons_layout)
        main_layout.addWidget(self.result_display)

        container = QWidget()
        container.setLayout(main_layout)
        self.setCentralWidget(container)

    def start_file_generation(self):
        """
        Initiates file generation in a separate thread to avoid UI blocking.
        """
        self.result_display.append("Generating Excel file...")
        self.generate_file_btn.setEnabled(False)

        # Start the file generation thread
        self.file_thread = FileGenerationThread("data.xlsx")
        # "When file_thread finishes creating the file, it will emit file_generated, and when that happens,
        # call on_file_generated:"
        self.file_thread.file_generated.connect(self.on_file_generated)
        self.file_thread.start() # "(Actually start the thread here)"

    def on_file_generated(self):
        """
        Handles enabling processing buttons post file generation.
        """
        self.result_display.append("Excel file generated successfully.")
        # Disable the generate button after file creation
        self.generate_file_btn.setEnabled(False) #? is this redundant?

        # Enable processing buttons now that the file is generated
        self.process_openpyxl_btn.setEnabled(True)
        self.process_pandas_btn.setEnabled(True)
        self.process_xlsxio_btn.setEnabled(True)

    def process_data_openpyxl(self):
        """
        Initiates data processing with openpyxl.
        """
        self.result_display.append("Processing with openpyxl...")
        self.process_openpyxl_btn.setEnabled(False)
        self.process_pandas_btn.setEnabled(False)
        self.process_xlsxio_btn.setEnabled(False)

        self.data_thread = DataProcessingThread("data.xlsx", "openpyxl")
        self.data_thread.processing_complete.connect(self.display_results)
        self.data_thread.start()

    def process_data_pandas(self):
        """
        Initiates data processing with pandas.
        """
        self.result_display.append("Processing with pandas...")
        self.process_openpyxl_btn.setEnabled(False)
        self.process_pandas_btn.setEnabled(False)
        self.process_xlsxio_btn.setEnabled(False)

        self.data_thread = DataProcessingThread("data.xlsx", "pandas")
        self.data_thread.processing_complete.connect(self.display_results)
        self.data_thread.start()

    def process_data_xlsxio(self):
        """
        Placeholder for xlsxio processing initiation.
        """
        self.result_display.append("xlsxio processing is not implemented.")

    def display_results(self, counts, time_taken):
        """
        Displays data processing results and processing time.

        Parameters:
            counts (dict): Counts of unique values.
            time_taken (float): Time taken for processing in seconds.
        """
        self.result_display.append("\nProcessing complete. Results:")
        self.result_display.append("Unique Value - Count (only values with a count > 1):\n")

        for value, count in counts.items():
            if count > 1:
                self.result_display.append(f"{value}: {count}")

        self.result_display.append(f"\nTime taken: {time_taken:.2f} seconds")  # Display time taken

        # Re-enable buttons after processing
        self.process_openpyxl_btn.setEnabled(True)
        self.process_pandas_btn.setEnabled(True)
        self.process_xlsxio_btn.setEnabled(True)