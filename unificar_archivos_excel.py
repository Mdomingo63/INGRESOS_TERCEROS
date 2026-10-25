import os
import sys
import pandas as pd
from PySide6.QtWidgets import (
    QApplication, QWidget, QVBoxLayout, QPushButton, 
    QTextEdit, QLabel, QProgressBar, QMessageBox
)
from PySide6.QtGui import QFont
from PySide6.QtCore import Qt

# Ruta especificada
FOLDER_PATH = r"C:/Users/Domingo Sánchez/Escritorio/DOCUMENTOS TRABAJOS/CONTABILIDAD/DATOS EDISON/INGRESOS_TERCEROS"
OUTPUT_FILE_NAME = "RESUMEN_TERCERO_DGII_EDISON.xlsx"

class UnificadorDGIIApp(QWidget):
    def __init__(self):
        super().__init__()
        self.initUI()

    def initUI(self):
        self.setWindowTitle("Unificador de Archivos DGII - Edison")
        self.resize(650, 420)

        layout = QVBoxLayout()

        # Título
        title = QLabel("Unificador de Ingresos Según Terceros (DGII)")
        title.setFont(QFont("Segoe UI", 14, QFont.Weight.Bold))
        title.setAlignment(Qt.AlignmentFlag.AlignCenter)
        layout.addWidget(title)

        # Ruta
        path_label = QLabel(f"<b>Carpeta de Trabajo:</b><br>{FOLDER_PATH}")
        path_label.setWordWrap(True)
        layout.addWidget(path_label)

        # Botón de Procesar
        self.btn_run = QPushButton("Unificar Archivos de Excel")
        self.btn_run.setFont(QFont("Segoe UI", 11, QFont.Weight.Bold))
        self.btn_run.setStyleSheet("""
            QPushButton {
                background-color: #27ae60;
                color: white;
                padding: 12px;
                border-radius: 6px;
            }
            QPushButton:hover {
                background-color: #219150;
            }
        """)
        self.btn_run.clicked.connect(self.procesar_archivos)
        layout.addWidget(self.btn_run)

        # Área de Log
        self.log_area = QTextEdit()
        self.log_area.setReadOnly(True)
        self.log_area.setFont(QFont("Consolas", 10))
        layout.addWidget(self.log_area)

        # Barra de progreso
        self.progress = QProgressBar()
        self.progress.setValue(0)
        layout.addWidget(self.progress)

        self.setLayout(layout)

    def log(self, mensaje):
        self.log_area.append(mensaje)

    def procesar_archivos(self):
        self.log_area.clear()
        self.log("▶ Iniciando proceso de unificación...")

        if not os.path.exists(FOLDER_PATH):
            QMessageBox.critical(self, "Error de Ruta", f"La carpeta no existe:\n{FOLDER_PATH}")
            self.log("❌ Error: La carpeta especificada no existe.")
            return

        # Buscar archivos en la carpeta
        archivos = [f for f in os.listdir(FOLDER_PATH) if f.lower().endswith(('.xls', '.xlsx')) and not f.startswith('~$') and f != OUTPUT_FILE_NAME]

        if not archivos:
            QMessageBox.warning(self, "Sin Archivos", "No se encontraron archivos de Excel (.xls / .xlsx) en la carpeta.")
            self.log("⚠️ No hay archivos para consolidar.")
            return

        self.log(f"📁 Archivos encontrados ({len(archivos)}):")
        for f in archivos:
            self.log(f"  - {f}")

        dataframes = []
        self.progress.setMaximum(len(archivos))

        for idx, archivo in enumerate(archivos, 1):
            file_path = os.path.join(FOLDER_PATH, archivo)
            try:
                # Intentar leer primero como tabla HTML (formato nativo de exportación de la DGII)
                try:
                    dfs = pd.read_html(file_path)
                    df = dfs[0]
                except Exception:
                    # Si no es HTML, leer como Excel normal
                    df = pd.read_excel(file_path)

                dataframes.append(df)
                self.log(f"✔ Cargado correctamente ({len(df)} filas): {archivo}")
            except Exception as e:
                self.log(f"❌ Error al leer {archivo}: {str(e)}")

            self.progress.setValue(idx)

        if not dataframes:
            QMessageBox.critical(self, "Error", "No se pudo leer ninguno de los archivos.")
            return

        # Concatenar todos los archivos
        df_consolidado = pd.concat(dataframes, ignore_index=True)

        # Guardar archivo consolidado
        output_path = os.path.join(FOLDER_PATH, OUTPUT_FILE_NAME)
        try:
            with pd.ExcelWriter(output_path, engine='openpyxl') as writer:
                df_consolidado.to_excel(writer, index=False, sheet_name="RESUMEN_TERCEROS")

            self.log(f"\n🎉 ¡Proceso completado exitosamente!")
            self.log(f"📊 Total de registros consolidados: {len(df_consolidado)}")
            self.log(f"💾 Archivo guardado en:\n{output_path}")

            QMessageBox.information(
                self, "Proceso Exitoso", 
                f"Se han unificado {len(archivos)} archivos ({len(df_consolidado)} filas totales).\n\nGuardado en:\n{output_path}"
            )
        except Exception as e:
            QMessageBox.critical(self, "Error al Guardar", f"No se pudo guardar el archivo final:\n{str(e)}")
            self.log(f"❌ Error al guardar archivo consolidado: {str(e)}")


if __name__ == "__main__":
    app = QApplication(sys.argv)
    window = UnificadorDGIIApp()
    window.show()
    sys.exit(app.exec())