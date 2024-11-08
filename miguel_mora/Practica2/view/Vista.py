from PyQt5.QtCore import Qt
from PyQt5.QtWidgets import QMainWindow, QDesktopWidget, QTableWidgetItem, QTableWidget, QWidget, QVBoxLayout, \
    QHBoxLayout, QLabel, QPushButton, QMessageBox


class Vista(QMainWindow):
    def __init__(self, data):
        super().__init__()
        self.data = data  # Guardamos los datos para usarlos en initUI
        self.initUI()

    def initUI(self):
        self.setWindowTitle("Datos procesados")
        self.setGeometry(100, 100, 600, 800)
        self.center_window()  # Centramos la ventana en la pantalla

        # Widget central y layout
        central_widget = QWidget()
        layout = QVBoxLayout()

        # Crear y añadir el título
        title_label = QLabel("Datos Procesados")
        title_label.setAlignment(Qt.AlignCenter)  # Centrar el título
        title_label.setStyleSheet("font-size: 20px; font-weight: bold; color: #1a75ff;")  # Estilo del título
        layout.addWidget(title_label)

        # Crear la tabla
        self.tableWidget = QTableWidget()

        # Verificar que self.data no sea None o vacío
        if self.data is None or len(self.data) == 0:
            self.tableWidget.setRowCount(1)
            self.tableWidget.setColumnCount(1)
            self.tableWidget.setHorizontalHeaderLabels(["Error"])
            self.tableWidget.setItem(0, 0, QTableWidgetItem("No hay datos disponibles."))
        else:
            self.tableWidget.setRowCount(len(self.data))
            self.tableWidget.setColumnCount(2)
            self.tableWidget.setHorizontalHeaderLabels(["Valor", "Frecuencia"])

            # Rellenar la tabla con los datos
            for row, (valor, frecuencia) in enumerate(self.data.items()):
                self.tableWidget.setItem(row, 0, QTableWidgetItem(str(valor)))
                self.tableWidget.setItem(row, 1, QTableWidgetItem(str(frecuencia)))

        # Configurar barra de desplazamiento
        self.tableWidget.setVerticalScrollBarPolicy(Qt.ScrollBarAlwaysOn)
        self.tableWidget.setHorizontalScrollBarPolicy(Qt.ScrollBarAlwaysOff)
        self.tableWidget.resizeColumnsToContents()

        # Ajustar el tamaño de la tabla
        self.tableWidget.setFixedWidth(223)
        self.tableWidget.setFixedHeight(450)

        # Crear un layout horizontal para centrar la tabla
        table_layout = QHBoxLayout()
        table_layout.addWidget(self.tableWidget, alignment=Qt.AlignCenter)  # Centrar la tabla
        layout.addLayout(table_layout)

        # Añadir el botón al layout
        self.btn_xlsx = QPushButton("Menu", self)
        self.btn_xlsx.clicked.connect(self.return_to_menu)
        layout.addWidget(self.btn_xlsx, alignment=Qt.AlignCenter)  # Añadir el botón al centro

        # Añadir el layout al widget central
        central_widget.setLayout(layout)
        self.setCentralWidget(central_widget)

    def center_window(self):
        # Centra la ventana en la pantalla
        screen_center = QDesktopWidget().availableGeometry().center()
        self.move(screen_center - self.rect().center())

    def return_to_menu(self):
        try:
            from MainWindow import MainWindow  
            self.ventana = MainWindow()  # Instancia la ventana principal del menú
            self.ventana.show()  # Muestra la ventana del menú
            self.close()  # Cierra la ventana actual (en lugar de hide para liberar recursos)
        except Exception as e:
            QMessageBox.critical(self, "Error", f"Error al cargar Menu:\n{str(e)}")
