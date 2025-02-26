from PyQt5.QtWidgets import QApplication, QMainWindow, QPushButton, QFileDialog, QLabel, QVBoxLayout, QWidget, QGridLayout
from PyQt5.QtGui import QIcon, QFont
from PyQt5.QtCore import Qt, QSize

class MainWindow(QMainWindow):
    def __init__(self):
        super().__init__()
        self.setWindowTitle("Extractor de Estados de Cuenta")
        self.setGeometry(200, 200, 600, 500)
        self.setFixedSize(600, 500)
        self.setWindowIcon(QIcon("heza_logo.jpg"))

        # Estilos CSS mejorados
        self.setStyleSheet("""
            QMainWindow {
                background-color: #f0f0f0;
            }
            QLabel {
                font-size: 18px;
                font-weight: bold;
                color: #333333;
                margin-bottom: 20px;
                text-align: center;
            }
            QPushButton {
                background-color: #4CAF50;
                color: white;
                font-size: 14px;
                font-weight: bold;
                padding: 10px;
                border-radius: 8px;
                border: none;
                margin: 5px;
                text-align: left;
                min-width: 150px;
                min-height: 20px;
                icon-size: 10px;
            }
            QPushButton:hover {
                background-color: #45a049;
            }
            QPushButton:disabled {
                background-color: #cccccc;
                color: #666666;
            }
        """)

        # Fuente personalizada
        font = QFont("Arial", 12)

        # Etiqueta principal
        self.label = QLabel("Selecciona el banco y carga un estado de cuenta en PDF", self)
        self.label.setFont(font)
        self.label.setAlignment(Qt.AlignCenter)

        # Botones para seleccionar el banco
        self.buttons = [
            ("BBVA", "bbva_icon.png"),
            ("Banco del Bajío", "banco_bajio_icon.png"),
            ("Scotiabank", "scotiabank_icon.png"),
            ("Banorte", "banorte_icon.png"),
            ("BBASE", "bbase_icon.png"),
        ]

        # Crear botones dinámicamente
        self.bank_buttons = []
        for name, icon in self.buttons:
            btn = QPushButton(name, self)
            btn.setIcon(QIcon(icon))
            btn.setFont(font)
            btn.setIconSize(QSize(24, 24))  # Ajustar tamaño del icono
            btn.setMinimumSize(180, 50)  # Tamaño mínimo para que sean uniformes
            btn.clicked.connect(lambda _, b=name: self.select_bank(b))
            self.bank_buttons.append(btn)

        # Layout en Grid (2 columnas)
        bank_layout = QGridLayout()
        bank_layout.setAlignment(Qt.AlignCenter)
        bank_layout.setSpacing(10)

        # Insertar botones en el grid (2 en cada fila)
        for i, btn in enumerate(self.bank_buttons):
            row = i // 2
            col = i % 2
            bank_layout.addWidget(btn, row, col)

        # Botón para seleccionar archivo
        self.btn_select_file = QPushButton("Seleccionar PDF", self)
        self.btn_select_file.setFont(font)
        self.btn_select_file.setMinimumSize(200, 50)
        self.btn_select_file.setEnabled(False)  # Deshabilitado hasta que se seleccione un banco
        self.btn_select_file.clicked.connect(self.load_pdf)

        # Botón para exportar a Excel
        self.btn_export_excel = QPushButton("Exportar a Excel", self)
        self.btn_export_excel.setFont(font)
        self.btn_export_excel.setMinimumSize(200, 50)
        self.btn_export_excel.setEnabled(False)
        self.btn_export_excel.clicked.connect(self.export_to_excel)

        # Diseño de la ventana
        layout = QVBoxLayout()
        layout.addWidget(self.label)
        layout.addLayout(bank_layout)  # Agregar los botones en grid
        layout.addWidget(self.btn_select_file, alignment=Qt.AlignCenter)
        layout.addWidget(self.btn_export_excel, alignment=Qt.AlignCenter)
        layout.setSpacing(15)
        layout.setContentsMargins(30, 30, 30, 30)

        container = QWidget()
        container.setLayout(layout)
        self.setCentralWidget(container)

        self.file_path = None
        self.selected_bank = None

    def select_bank(self, bank):
        self.selected_bank = bank
        self.label.setText(f"Banco seleccionado: {bank}\nCarga un estado de cuenta en PDF")
        self.btn_select_file.setEnabled(True)

    def load_pdf(self):
        options = QFileDialog.Options()
        file_path, _ = QFileDialog.getOpenFileName(self, "Seleccionar PDF", "", "Archivos PDF (*.pdf);;Todos los archivos (*)", options=options)
        if file_path:
            self.file_path = file_path
            self.label.setText(f"Archivo seleccionado:\n{file_path}")
            self.btn_export_excel.setEnabled(True)

    def export_to_excel(self):
        if self.file_path and self.selected_bank:
            from extractor import PDFExtractor
            extractor = PDFExtractor(self.file_path, self.selected_bank)
            data = extractor.extract_data()
            save_path, _ = QFileDialog.getSaveFileName(
                self, "Guardar como", f"estado_de_cuenta_{self.selected_bank}.xlsx",
                "Archivos de Excel (*.xlsx);;Todos los archivos (*)"
            )
            if save_path:
                try:
                    extractor.save_to_excel(data, save_path)
                    self.label.setText(f"Datos guardados en:\n{save_path}")
                except PermissionError:
                    self.label.setText("Error: No se pudo guardar el archivo. Verifica los permisos o cierra el archivo si está abierto.")
                except Exception as e:
                    self.label.setText(f"Error inesperado: {str(e)}")

if __name__ == "__main__":
    import sys
    app = QApplication(sys.argv)
    window = MainWindow()
    window.show()
    sys.exit(app.exec_())
