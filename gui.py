from PyQt5.QtWidgets import QApplication, QMainWindow, QPushButton, QFileDialog, QLabel, QVBoxLayout, QWidget, QHBoxLayout
from PyQt5.QtGui import QIcon, QFont
from PyQt5.QtCore import Qt

class MainWindow(QMainWindow):
    def __init__(self):
        super().__init__()
        self.setWindowTitle("Extractor de Estados de Cuenta")
        self.setGeometry(200, 200, 600, 450)
        self.setFixedSize(600, 450)
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
                background-color: #2196F3;  /* Azul profesional */
                color: white;
                font-size: 14px;  /* Tamaño de letra ajustado */
                font-weight: bold;
                padding: 12px -4px;
                border-radius: 8px;
                border: 2px solid #1976D2;  /* Borde para mejor contraste */
                margin: 5px;
                text-align: center;
                min-width: 140px;
            }
            QPushButton:hover {
                background-color: #1976D2;  /* Azul más oscuro al pasar el mouse */
            }
            QPushButton:disabled {
                background-color: #BBDEFB;  /* Azul claro para deshabilitados */
                color: #757575;
                border: 2px solid #90CAF9;
            }
        """)

        # Fuente personalizada
        font = QFont()
        font.setFamily("Arial")
        font.setPointSize(12)

        # Etiqueta principal
        self.label = QLabel("Selecciona el banco y carga un estado de cuenta en PDF", self)
        self.label.setFont(font)
        self.label.setAlignment(Qt.AlignCenter)

        # Botones para seleccionar el banco
        self.btn_bbva = QPushButton("BBVA", self)
        self.btn_bbva.setIcon(QIcon("bbva_icon.png"))  # Icono para BBVA
        self.btn_bbva.setFont(font)
        self.btn_bbva.clicked.connect(lambda: self.select_bank("BBVA"))

        self.btn_banco_bajio = QPushButton("Banco del Bajío", self)
        self.btn_banco_bajio.setIcon(QIcon("banco_bajio_icon.png"))  # Icono para Banco del Bajío
        self.btn_banco_bajio.setFont(font)
        self.btn_banco_bajio.clicked.connect(lambda: self.select_bank("Banco del Bajío"))

        self.btn_scotiabank = QPushButton("Scotiabank", self)
        self.btn_scotiabank.setIcon(QIcon("scotiabank_icon.png"))  # Icono para Scotiabank
        self.btn_scotiabank.setFont(font)
        self.btn_scotiabank.clicked.connect(lambda: self.select_bank("Scotiabank"))

        self.btn_banorte = QPushButton("Banorte", self)
        self.btn_banorte.setIcon(QIcon("banorte_icon.png"))  # Icono para Banorte
        self.btn_banorte.setFont(font)
        self.btn_banorte.clicked.connect(lambda: self.select_bank("Banorte"))

        # Layout para los botones de bancos
        bank_layout = QHBoxLayout()
        bank_layout.addWidget(self.btn_bbva)
        bank_layout.addWidget(self.btn_banco_bajio)
        bank_layout.addWidget(self.btn_scotiabank)
        bank_layout.addWidget(self.btn_banorte)

        # Botón para seleccionar archivo
        self.btn_select_file = QPushButton("Seleccionar PDF", self)
        self.btn_select_file.setFont(font)
        self.btn_select_file.clicked.connect(self.load_pdf)
        self.btn_select_file.setEnabled(False)  # Deshabilitado hasta que se seleccione un banco

        # Botón para exportar a Excel
        self.btn_export_excel = QPushButton("Exportar a Excel", self)
        self.btn_export_excel.setFont(font)
        self.btn_export_excel.setEnabled(False)
        self.btn_export_excel.clicked.connect(self.export_to_excel)

        # Diseño de la ventana
        layout = QVBoxLayout()
        layout.addWidget(self.label)
        layout.addLayout(bank_layout)
        layout.addWidget(self.btn_select_file)
        layout.addWidget(self.btn_export_excel)
        layout.setSpacing(20)
        layout.setContentsMargins(30, 30, 30, 30)

        container = QWidget()
        container.setLayout(layout)
        self.setCentralWidget(container)

        self.file_path = None
        self.selected_bank = None

        # Habilitar arrastrar y soltar
        self.setAcceptDrops(True)

    def select_bank(self, bank):
        self.selected_bank = bank
        self.label.setText(f"Banco seleccionado: {bank}\nCarga un estado de cuenta en PDF")
        self.btn_select_file.setEnabled(True)

    def dragEnterEvent(self, event):
        if event.mimeData().hasUrls():
            event.acceptProposedAction()

    def dropEvent(self, event):
        for url in event.mimeData().urls():
            file_path = url.toLocalFile()
            if file_path.endswith('.pdf'):
                self.file_path = file_path
                self.label.setText(f"Archivo seleccionado:\n{file_path}")
                self.btn_export_excel.setEnabled(True)
                break

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

            # Nombre por defecto del archivo
            default_file_name = f"estado_de_cuenta_{self.selected_bank}.xlsx"
            
            # Configurar el diálogo de guardado con el nombre por defecto
            save_path, _ = QFileDialog.getSaveFileName(
                self, 
                "Guardar como", 
                default_file_name,  # Nombre por defecto
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