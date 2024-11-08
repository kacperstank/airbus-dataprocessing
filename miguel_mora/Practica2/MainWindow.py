from PyQt5.QtWidgets import QMainWindow, QWidget, QPushButton, QVBoxLayout, QMessageBox, QDesktopWidget
from PyQt5.QtCore import Qt
from Procesar_Panda import ProcesarPanda as Panda
from Procesar_Xlsxio import Procesar_Xlsxio as Xlsxio
from view.Vista import Vista
from view.Vista_xlsxio import Vista_xlsxio

class MainWindow(QMainWindow):
    def __init__(self):
        super().__init__()
        self.initUI()

    def initUI(self):
        # Configuración de la ventana principal
        self.setWindowTitle("Práctica 2, procesar XLSX PANDAS/XLSXIO")
        self.setGeometry(100, 100, 600, 400)
        self.center_window()  # Centramos la ventana en la pantalla

        # Widget central y layout principal
        central_widget = QWidget(self)
        self.setCentralWidget(central_widget)

        layout = QVBoxLayout()
        layout.setAlignment(Qt.AlignCenter)  # Centrar el layout

        # Crear los botones
        self.btn_json = QPushButton("Pandas", self)
        self.btn_json.clicked.connect(self.on_click_Panda)

        self.btn_xlsx = QPushButton("XLSXIO", self)
        self.btn_xlsx.clicked.connect(self.on_click_Xlsxio)

        # Añadir los botones al layout
        layout.addWidget(self.btn_json)
        layout.addWidget(self.btn_xlsx)

        # Asignar layout al widget central
        central_widget.setLayout(layout)

    def center_window(self):
        # Centra la ventana en la pantalla
        frame_geometry = self.frameGeometry()
        screen_center = QDesktopWidget().availableGeometry().center()
        frame_geometry.moveCenter(screen_center)
        self.move(frame_geometry.topLeft())

    def on_click_Panda(self):
        try:
            panda = Panda()
            data = panda.repetidos()
            self.ventana_panda = Vista(data)
            self.ventana_panda.show()
            self.hide()
        except Exception as e:
            QMessageBox.critical(self, "Error", f"Se produjo un error al cargar la ventana de procesado de datos (Panda):\n{str(e)}")

    def on_click_Xlsxio(self):
        try:
            xlsxio = Xlsxio()
            data = xlsxio.repetidos()
            self.ventana_xlsxio = Vista(data)
            self.ventana_xlsxio.show()
            self.hide()
        except Exception as e:
            QMessageBox.critical(self, "Error", f"Se produjo un error al cargar la ventana de procesado de datos (Xlsxio):\n{str(e)}")
