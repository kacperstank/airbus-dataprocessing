import sys
from PyQt5.QtWidgets import QApplication

from data_processor_gui import DataProcessorGUI

if __name__ == "__main__":
    # Create the application instance
    app = QApplication(sys.argv)

    # Create the main window instance
    mainWindow = DataProcessorGUI()

    # Show the main window
    mainWindow.show()

    # Run the application event loop
    sys.exit(app.exec_())