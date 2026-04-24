#!/usr/bin/env python3
"""
Homologación ISSFA - Aplicación PyQt6
Login único → Seleccionar Excel → Analizar → Vista previa → Aplicar
"""
import os
import sys
import shutil
from datetime import datetime
from pathlib import Path
from PyQt6.QtWidgets import (
    QApplication, QMainWindow, QWidget, QVBoxLayout, QHBoxLayout,
    QLabel, QLineEdit, QPushButton, QFileDialog, QTableWidget,
    QTableWidgetItem, QHeaderView, QMessageBox, QProgressBar,
    QStatusBar, QCheckBox, QComboBox, QSpinBox, QGroupBox,
    QScrollArea, QSplitter, QFrame, QDialog, QTabWidget
)
from PyQt6.QtCore import Qt, QThread, pyqtSignal, QTimer
from PyQt6.QtGui import QFont, QColor

from database import OracleConnection
from homology import HomologacionEngine
from config import get_template_excel_path


class LoginDialog(QDialog):
    """Diálogo de login - solo una vez."""
    
    def __init__(self, parent=None):
        super().__init__(parent)
        self.setWindowTitle("🔐 Login Oracle ISSFA")
        self.setModal(True)
        self.setFixedSize(400, 200)
        
        layout = QVBoxLayout()
        
        # Título
        title = QLabel("Conexión Oracle RAC")
        title.setFont(QFont("Arial", 14, QFont.Weight.Bold))
        title.setAlignment(Qt.AlignmentFlag.AlignCenter)
        layout.addWidget(title)
        
        layout.addSpacing(10)
        
        # Usuario
        user_layout = QHBoxLayout()
        user_layout.addWidget(QLabel("Usuario:"))
        self.user_edit = QLineEdit()
        self.user_edit.setPlaceholderText("Usuario Oracle")
        user_layout.addWidget(self.user_edit)
        layout.addLayout(user_layout)
        
        # Contraseña
        pass_layout = QHBoxLayout()
        pass_layout.addWidget(QLabel("Clave:"))
        self.pass_edit = QLineEdit()
        self.pass_edit.setEchoMode(QLineEdit.EchoMode.Password)
        self.pass_edit.setPlaceholderText("Contraseña Oracle")
        pass_layout.addWidget(self.pass_edit)
        layout.addLayout(pass_layout)
        
        # Botón conectar
        self.connect_btn = QPushButton("🔗 Conectar")
        self.connect_btn.setFixedHeight(40)
        self.connect_btn.clicked.connect(self.try_connect)
        layout.addWidget(self.connect_btn)
        
        layout.addStretch()
        
        # Estado
        self.status_label = QLabel("")
        self.status_label.setAlignment(Qt.AlignmentFlag.AlignCenter)
        layout.addWidget(self.status_label)
        
        self.setLayout(layout)
        
        # Presionar Enter
        self.pass_edit.returnPressed.connect(self.try_connect)
        self.user_edit.setFocus()
    
    def try_connect(self):
        user = self.user_edit.text().strip()
        password = self.pass_edit.text()
        
        if not user or not password:
            self.status_label.setText("⚠️ Ingrese usuario y clave")
            return
        
        self.status_label.setText("⏳ Conectando...")
        self.connect_btn.setEnabled(False)
        QApplication.processEvents()
        
        try:
            self.db = OracleConnection(user, password)
            if self.db.connect():
                self.status_label.setText("✅ Conexión exitosa")
                QTimer.singleShot(500, lambda: self.accept())
            else:
                self.status_label.setText("❌ No se pudo conectar")
                self.connect_btn.setEnabled(True)
        except Exception as e:
            self.status_label.setText(f"❌ Error: {str(e)[:50]}")
            self.connect_btn.setEnabled(True)


class WorkerThread(QThread):
    """Thread para operaciones de base de datos."""
    progress = pyqtSignal(str)
    finished = pyqtSignal(bool, str)
    stats_ready = pyqtSignal(dict)
    results_ready = pyqtSignal(dict)
    
    def __init__(self, engine: HomologacionEngine, action: str):
        super().__init__()
        self.engine = engine
        self.action = action
    
    def run(self):
        try:
            if self.action == "analizar":
                self.progress.emit("Analizando items...")
                stats = self.engine.analizar()
                self.stats_ready.emit(stats)
                self.finished.emit(True, "Análisis completado")
            elif self.action == "aplicar":
                self.progress.emit("Aplicando cambios...")
                self.engine.crear_backup()
                results = self.engine.aplicar_cambios()
                self.results_ready.emit(results)
                self.finished.emit(True, "Cambios aplicados")
            elif self.action == "excel":
                self.progress.emit("Cargando Excel...")
                count = self.engine.load_excel(self.excel_path)
                self.finished.emit(True, f"Cargados {count} registros")
        except Exception as e:
            self.finished.emit(False, str(e))


class MainWindow(QMainWindow):
    """Ventana principal de homologación."""
    
    def __init__(self):
        super().__init__()
        self.db = None
        self.engine = None
        self.worker = None
        self.init_ui()
    
    def init_ui(self):
        self.setWindowTitle("🔄 Homologación ISSFA - Fase 01")
        self.setGeometry(100, 100, 1400, 800)
        
        # Widget central
        central = QWidget()
        self.setCentralWidget(central)
        main_layout = QVBoxLayout(central)
        
        # Toolbar
        toolbar = QFrame()
        toolbar.setStyleSheet("QFrame { background: #f0f0f0; padding: 5px; }")
        toolbar_layout = QHBoxLayout(toolbar)
        
        # Botón Descargar plantilla
        self.btn_template = QPushButton("📥 Descargar plantilla Excel")
        self.btn_template.clicked.connect(self.download_template)
        toolbar_layout.addWidget(self.btn_template)
        
        toolbar_layout.addSpacing(10)
        
        # Botón Excel
        self.btn_excel = QPushButton("📁 Seleccionar Excel")
        self.btn_excel.clicked.connect(self.select_excel)
        toolbar_layout.addWidget(self.btn_excel)
        
        # ID_ITISF
        toolbar_layout.addWidget(QLabel("ID_ITISF:"))
        self.spin_id = QSpinBox()
        self.spin_id.setRange(1, 9999)
        self.spin_id.setValue(1)
        self.spin_id.setFixedWidth(80)
        toolbar_layout.addWidget(self.spin_id)
        
        # Threshold
        toolbar_layout.addWidget(QLabel("Threshold %:"))
        self.spin_threshold = QSpinBox()
        self.spin_threshold.setRange(50, 100)
        self.spin_threshold.setValue(88)
        self.spin_threshold.setFixedWidth(60)
        toolbar_layout.addWidget(self.spin_threshold)
        
        # Botón Analizar
        self.btn_analizar = QPushButton("🔍 Analizar")
        self.btn_analizar.clicked.connect(self.analizar)
        toolbar_layout.addWidget(self.btn_analizar)
        
        toolbar_layout.addStretch()
        
        # Botón Aplicar
        self.btn_aplicar = QPushButton("✅ Aplicar Cambios")
        self.btn_aplicar.clicked.connect(self.aplicar)
        self.btn_aplicar.setStyleSheet("QPushButton { background: #4CAF50; color: white; font-weight: bold; }")
        toolbar_layout.addWidget(self.btn_aplicar)
        
        # Botón CSV
        self.btn_csv = QPushButton("📊 Generar CSV")
        self.btn_csv.clicked.connect(self.generar_csv)
        toolbar_layout.addWidget(self.btn_csv)
        
        # Botón Panic
        self.btn_panic = QPushButton("🚨 Restaurar Backup")
        self.btn_panic.setStyleSheet("QPushButton { background: #f44336; color: white; }")
        self.btn_panic.clicked.connect(self.restaurar_backup)
        toolbar_layout.addWidget(self.btn_panic)
        
        main_layout.addWidget(toolbar)
        
        # Estadísticas
        stats_frame = QFrame()
        stats_frame.setStyleSheet("QFrame { background: #e3f2fd; padding: 10px; }")
        stats_layout = QHBoxLayout(stats_frame)
        
        self.lbl_total = QLabel("Total: 0")
        self.lbl_update = QLabel("UPDATE: 0")
        self.lbl_insert = QLabel("INSERT: 0")
        self.lbl_bloqueado = QLabel("BLOQUEADO: 0")
        self.lbl_confianza = QLabel("Alta confianza: 0")
        
        for lbl in [self.lbl_total, self.lbl_update, self.lbl_insert, 
                    self.lbl_bloqueado, self.lbl_confianza]:
            lbl.setFont(QFont("Arial", 10, QFont.Weight.Bold))
            stats_layout.addWidget(lbl)
        
        main_layout.addWidget(stats_frame)
        
        # Tabla de resultados
        self.table = QTableWidget()
        self.table.setColumnCount(14)
        self.table.setHorizontalHeaderLabels([
            'Aplicar', 'Fila', 'Acción', 'Estado', 'Decisión',
            'Código Actual', 'Desc. Actual Oracle', 'Desc. Actual Excel',
            'Código Nuevo', 'Desc. Nueva',
            'Score Oracle/Excel', 'Score Actual/Nueva',
            'Existe', 'Tipo'
        ])
        self.table.horizontalHeader().setSectionResizeMode(QHeaderView.ResizeMode.Stretch)
        self.table.setAlternatingRowColors(True)
        self.table.setStyleSheet("""
            QTableWidget { font-size: 11px; }
            QTableWidget::item { padding: 4px; }
        """)
        main_layout.addWidget(self.table)
        
        # Barra de progreso
        self.progress_bar = QProgressBar()
        self.progress_bar.setVisible(False)
        main_layout.addWidget(self.progress_bar)
        
        # Status bar
        self.statusBar().showMessage("Listo")
        
        # Archivo Excel seleccionado
        self.excel_path = None
        self.lbl_excel = QLabel("Ningún archivo seleccionado")
    
    def download_template(self):
        """Permite descargar una copia de la plantilla oficial."""
        try:
            source_path = get_template_excel_path()
            destination_path, _ = QFileDialog.getSaveFileName(
                self,
                "Guardar plantilla Excel",
                "plantilla_homologacion_items_issfa.xlsx",
                "Excel (*.xlsx)"
            )
            if not destination_path:
                return
            if not destination_path.lower().endswith(".xlsx"):
                destination_path += ".xlsx"
            shutil.copyfile(source_path, destination_path)
            QMessageBox.information(
                self,
                "Plantilla descargada",
                "La plantilla fue descargada correctamente.\n\n"
                "Llene el Excel y luego cárguelo con 'Seleccionar Excel'."
            )
        except Exception as e:
            QMessageBox.critical(self, "Error", str(e))
    
    def select_excel(self):
        file_path, _ = QFileDialog.getOpenFileName(
            self, "Seleccionar Excel", "",
            "Archivos Excel (*.xlsx *.xls);;Todos los archivos (*.*)"
        )
        if file_path:
            self.excel_path = file_path
            self.lbl_excel.setText(f"📁 {os.path.basename(file_path)}")
            self.statusBar().showMessage(f"Excel: {file_path}")
            
            # Crear engine
            user = getattr(self, 'oracle_user', '')
            password = getattr(self, 'oracle_password', '')
            if self.db:
                self.engine = HomologacionEngine(
                    self.db,
                    id_itisf=self.spin_id.value(),
                    threshold=self.spin_threshold.value()
                )
                
                # Cargar Excel en thread
                self.worker = WorkerThread(self.engine, "excel")
                self.worker.excel_path = file_path
                self.worker.progress.connect(lambda m: self.statusBar().showMessage(m))
                self.worker.finished.connect(self.on_excel_loaded)
                self.progress_bar.setVisible(True)
                self.worker.start()
    
    def on_excel_loaded(self, success, message):
        self.progress_bar.setVisible(False)
        if success:
            self.statusBar().showMessage(message)
            self.btn_analizar.setEnabled(True)
        else:
            QMessageBox.critical(self, "Error", message)
    
    def analizar(self):
        if not self.engine:
            QMessageBox.warning(self, "Aviso", "Seleccione primero un archivo Excel")
            return
        
        self.engine.id_itisf = self.spin_id.value()
        self.engine.threshold = self.spin_threshold.value()
        
        self.worker = WorkerThread(self.engine, "analizar")
        self.worker.stats_ready.connect(self.mostrar_stats)
        self.worker.progress.connect(lambda m: self.statusBar().showMessage(m))
        self.worker.finished.connect(self.mostrar_tabla)
        self.progress_bar.setVisible(True)
        self.worker.start()
    
    def mostrar_stats(self, stats):
        self.lbl_total.setText(f"Total: {stats['total']}")
        self.lbl_update.setText(f"UPDATE: {stats['update']}")
        self.lbl_insert.setText(f"INSERT: {stats['insert']}")
        self.lbl_bloqueado.setText(f"BLOQUEADO: {stats['bloqueado']}")
        self.lbl_confianza.setText(f"Alta confianza: {stats['alta_confianza']}")
    
    def mostrar_tabla(self, success, message):
        self.progress_bar.setVisible(False)
        self.statusBar().showMessage(message)
        
        if not success or not self.engine:
            return
        
        self.table.setRowCount(0)
        self.table.setRowCount(len(self.engine.items))
        
        for row, item in enumerate(self.engine.items):
            # Checkbox aplicar
            checkbox = QCheckBox()
            checkbox.setChecked(item.aplicar)
            checkbox.stateChanged.connect(lambda s, i=row: self.toggle_aplicar(i, s))
            self.table.setCellWidget(row, 0, checkbox)
            
            # Datos
            datos = [
                str(item.fila_excel),
                item.accion,
                item.status,
                item.decision,
                item.codigo_actual,
                item.descripcion_actual_oracle[:50] if item.descripcion_actual_oracle else "",
                item.descripcion_actual_excel[:50],
                item.codigo_nuevo,
                item.descripcion_nueva[:50],
                f"{item.score_oracle_excel:.1f}%",
                f"{item.score_actual_nueva:.1f}%",
                "S" if item.existe_actual else "N",
                item.tipo,
            ]
            
            for col, valor in enumerate(datos, 1):
                cell = QTableWidgetItem(valor)
                
                # Colorear según acción
                if item.accion == "BLOQUEADO":
                    cell.setBackground(QColor("#ffcccc"))
                elif item.aplicar and item.decision == "ALTA_CONFIANZA":
                    cell.setBackground(QColor("#ccffcc"))
                elif item.aplicar:
                    cell.setBackground(QColor("#ffffcc"))
                
                self.table.setItem(row, col, cell)
    
    def toggle_aplicar(self, row, state):
        if self.engine and row < len(self.engine.items):
            self.engine.items[row].aplicar = state > 0
    
    def aplicar(self):
        if not self.engine:
            QMessageBox.warning(self, "Aviso", "Primero analice el archivo")
            return
        
        # Contar por tipo de acción
        updates = [i for i in self.engine.items if i.aplicar and i.accion == "UPDATE"]
        inserts = [i for i in self.engine.items if i.aplicar and i.accion == "INSERT"]
        bloqueados = [i for i in self.engine.items if i.accion == "BLOQUEADO"]
        desmarcados = [i for i in self.engine.items if not i.aplicar and i.accion not in ["BLOQUEADO", ""]]
        
        total_aplicar = len(updates) + len(inserts)
        if total_aplicar == 0:
            QMessageBox.warning(self, "Aviso", "No hay filas seleccionadas para aplicar")
            return
        
        # Construir mensaje de confirmación técnica
        msg = "=" * 60 + "\n"
        msg += "RESUMEN DE ACCIONES EN BASE DE DATOS\n"
        msg += "=" * 60 + "\n\n"
        msg += "Tabla afectada:\n"
        msg += "  • SIS.ITEMS_ISSFA_DETALLE\n\n"
        msg += "Acciones:\n"
        msg += f"  • UPDATE: {len(updates)} registro(s)\n"
        msg += f"  • INSERT: {len(inserts)} registro(s)\n"
        msg += f"  • BLOQUEADOS: {len(bloqueados)} registro(s)\n"
        msg += f"  • DESMARCADOS: {len(desmarcados)} registro(s)\n\n"
        msg += "NO se tocará:\n"
        msg += "  • SIS.ITEMS_ISSFA_CABECERA\n"
        msg += "  • SIS.EQUIVALENCIAS_ITEMS_ISSFA (salvo validación)\n\n"
        
        # Detalle de Updates
        if updates:
            msg += "Detalle UPDATE:\n"
            for item in updates[:10]:  # Solo primeros 10
                msg += f"  {item.codigo_actual} → {item.codigo_nuevo}\n"
            if len(updates) > 10:
                msg += f"  ... y {len(updates) - 10} más\n"
            msg += "\n"
        
        # Detalle de Inserts
        if inserts:
            msg += "Detalle INSERT:\n"
            for item in inserts[:10]:  # Solo primeros 10
                msg += f"  {item.codigo_nuevo} ({item.tipo})\n"
            if len(inserts) > 10:
                msg += f"  ... y {len(inserts) - 10} más\n"
            msg += "\n"
        
        msg += "-" * 60 + "\n"
        msg += "La operación se ejecutará dentro de una transacción.\n"
        msg += "Si ocurre un error, se hará rollback.\n"
        msg += "-" * 60 + "\n\n"
        msg += "¿Desea continuar?"
        
        reply = QMessageBox.question(
            self, "⚠️ Confirmar Cambios",
            msg,
            QMessageBox.StandardButton.Yes | QMessageBox.StandardButton.No
        )
        
        if reply == QMessageBox.StandardButton.No:
            self.statusBar().showMessage("Operación cancelada por el usuario")
            return
        
        self.worker = WorkerThread(self.engine, "aplicar")
        self.worker.results_ready.connect(self.on_aplicar_result)
        self.worker.progress.connect(lambda m: self.statusBar().showMessage(m))
        self.worker.finished.connect(lambda s, m: self.statusBar().showMessage(m))
        self.progress_bar.setVisible(True)
        self.worker.start()
    
    def on_aplicar_result(self, results):
        msg = f"Updates: {results['updates']}, Inserts: {results['inserts']}"
        if results['errores']:
            QMessageBox.critical(self, "Error", f"{msg}\n\nErrores:\n" + "\n".join(results['errores']))
        else:
            QMessageBox.information(self, "Éxito", msg)
    
    def generar_csv(self):
        if not self.engine:
            QMessageBox.warning(self, "Aviso", "Primero analice el archivo")
            return
        
        file_path, _ = QFileDialog.getSaveFileName(
            self, "Guardar CSV", 
            f"auditoria_{datetime.now().strftime('%Y%m%d_%H%M%S')}.csv",
            "CSV (*.csv)"
        )
        
        if file_path:
            self.engine.generar_csv_auditoria(file_path)
            QMessageBox.information(self, "CSV", f"Guardado en:\n{file_path}")
    
    def restaurar_backup(self):
        if not self.engine:
            QMessageBox.warning(self, "Aviso", "No hay backup disponible")
            return
        
        reply = QMessageBox.question(
            self, "Restaurar Backup",
            "¿Restaurar datos al estado anterior?",
            QMessageBox.StandardButton.Yes | QMessageBox.StandardButton.No
        )
        
        if reply == QMessageBox.StandardButton.Yes:
            if self.engine.restaurar_backup():
                QMessageBox.information(self, "Restaurado", "Backup restaurado correctamente")
            else:
                QMessageBox.critical(self, "Error", "No se pudo restaurar el backup")


def main():
    # Configurar timezone para Oracle
    os.environ["JAVA_TOOL_OPTIONS"] = "-Doracle.jdbc.timezoneAsRegion=false -Duser.timezone=UTC"
    
    app = QApplication(sys.argv)
    
    # Login inicial
    login = LoginDialog()
    if login.exec() != QDialog.DialogCode.Accepted:
        return
    
    # Crear ventana principal con conexión
    main_win = MainWindow()
    main_win.db = login.db
    main_win.oracle_user = login.user_edit.text()
    main_win.oracle_password = login.pass_edit.text()
    main_win.show()
    
    sys.exit(app.exec())


if __name__ == "__main__":
    main()
