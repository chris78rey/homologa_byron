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
    QScrollArea, QSplitter, QFrame, QDialog, QTabWidget,
    QAbstractItemView, QTextEdit, QDialogButtonBox
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
        self.lbl_decision = QLabel("DECISIÓN: 0")
        self.lbl_confianza = QLabel("Alta confianza: 0")
        
        for lbl in [self.lbl_total, self.lbl_update, self.lbl_insert, 
                    self.lbl_decision, self.lbl_confianza]:
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
        
        # ======================================================
        # NUEVO: Configuración visual de tabla para textos largos
        # ======================================================
        
        self.table.setWordWrap(True)
        self.table.setTextElideMode(Qt.TextElideMode.ElideNone)
        self.table.setSelectionBehavior(QAbstractItemView.SelectionBehavior.SelectRows)
        self.table.setSelectionMode(QAbstractItemView.SelectionMode.SingleSelection)
        
        # Permite que el usuario estire columnas manualmente
        header = self.table.horizontalHeader()
        header.setSectionResizeMode(QHeaderView.ResizeMode.Interactive)
        header.setStretchLastSection(False)
        
        self.table.setHorizontalScrollBarPolicy(Qt.ScrollBarPolicy.ScrollBarAsNeeded)
        self.table.setVerticalScrollBarPolicy(Qt.ScrollBarPolicy.ScrollBarAsNeeded)
        self.table.verticalHeader().setDefaultSectionSize(60)
        
        self.table.setAlternatingRowColors(True)
        self.table.setStyleSheet("""
            QTableWidget { font-size: 11px; }
            QTableWidget::item { padding: 4px; }
        """)
        
        # Doble clic para ver detalle completo
        self.table.cellDoubleClicked.connect(self.show_row_detail)
        
        # ======================================================
        # FIN NUEVO
        # ======================================================
        
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
        self.lbl_update.setText(f"AUTO: {stats['update_auto']}")
        self.lbl_insert.setText(f"INSERT: {stats['insert']}")
        self.lbl_decision.setText(f"MIGR: {stats['migrar']}")
        self.lbl_confianza.setText(f"Alta: {stats['alta_confianza']}")
    
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
            
            # Determinar opciones y acción actual
            opciones = item.get_opciones_disponibles()
            accion_actual = item.accion_final if item.accion_final else item.accion
            
            # Combo para acción final (columna 2)
            combo = QComboBox()
            combo.addItems(opciones)
            if accion_actual in opciones:
                combo.setCurrentText(accion_actual)
            combo.currentTextChanged.connect(lambda text, r=row: self.on_accion_final_changed(r, text))
            self.table.setCellWidget(row, 2, combo)
            
            # Datos: Fila, Estado, Acción sugerida
            datos = [
                str(item.fila_excel),
                item.status,
                item.accion,
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
                
                cell.setToolTip(str(valor))
                cell.setTextAlignment(Qt.AlignmentFlag.AlignTop | Qt.AlignmentFlag.AlignLeft)
                
                # Colorear según estado
                if item.status == "NO_APLICABLE_TECNICO":
                    cell.setBackground(QColor("#ffcccc"))
                elif item.status == "REQUIERE_DECISION_HUMANA":
                    cell.setBackground(QColor("#fff3cd"))
                elif item.status == "APLICABLE_AUTOMATICO":
                    cell.setBackground(QColor("#ccffcc"))
                
                self.table.setItem(row, col, cell)
        
        # Ajustar altura de filas
        self.table.resizeRowsToContents()
        for row_index in range(self.table.rowCount()):
            if self.table.rowHeight(row_index) > 140:
                self.table.setRowHeight(row_index, 140)
    
    def on_accion_final_changed(self, row, text):
        """Cuando el usuario cambia la acción final en el combo."""
        if self.engine and row < len(self.engine.items):
            item = self.engine.items[row]
            item.accion_final = text
            # Si选择了 algo distinto a OMITIR, marcar como aplicar
            item.aplicar = text != "OMITIR"
            # Actualizar checkbox
            checkbox = self.table.cellWidget(row, 0)
            if checkbox:
                checkbox.setChecked(item.aplicar)
    
    def toggle_aplicar(self, row, state):
        if self.engine and row < len(self.engine.items):
            self.engine.items[row].aplicar = state > 0
    
    def aplicar(self):
        if not self.engine:
            QMessageBox.warning(self, "Aviso", "Primero analice el archivo")
            return
        
        # Contar por acción final
        items_para_aplicar = [i for i in self.engine.items if i.aplicar]
        
        # Agrupar por tipo de acción final
        counts = {}
        for item in items_para_aplicar:
            accion = item.accion_final if item.accion_final else item.accion
            counts[accion] = counts.get(accion, 0) + 1
        
        if not items_para_aplicar:
            QMessageBox.warning(self, "Aviso", "No hay filas seleccionadas para aplicar")
            return
        
        # Construir mensaje de confirmación técnica
        msg = "=" * 60 + "\n"
        msg += "RESUMEN DE ACCIONES EN BASE DE DATOS\n"
        msg += "=" * 60 + "\n\n"
        msg += "Tabla afectada:\n"
        msg += "  • SIS.ITEMS_ISSFA_DETALLE\n"
        msg += "  • SIS.EQUIVALENCIAS_ITEMS_ISSFA (si aplica UPDATE_CON_EQUIVALENCIAS)\n\n"
        msg += "Acciones por tipo:\n"
        for accion, count in sorted(counts.items()):
            msg += f"  • {accion}: {count} registro(s)\n"
        msg += "\nNO se tocará:\n"
        msg += "  • SIS.ITEMS_ISSFA_CABECERA\n\n"
        
        # Detalle por acción
        for accion in sorted(counts.keys()):
            items_accion = [i for i in items_para_aplicar if (i.accion_final or i.accion) == accion][:5]
            if items_accion:
                msg += f"Detalle {accion}:\n"
                for item in items_accion:
                    if accion == "INSERT":
                        msg += f"  + {item.codigo_nuevo} ({item.tipo})\n"
                    else:
                        msg += f"  {item.codigo_actual} → {item.codigo_nuevo}\n"
                if len([i for i in items_para_aplicar if (i.accion_final or i.accion) == accion]) > 5:
                    msg += f"  ... y más\n"
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
    
    def show_row_detail(self, row: int, column: int):
        """Muestra detalle completo de una fila seleccionada."""
        if row < 0 or row >= len(self.engine.items):
            return
        
        item = self.engine.items[row]
        opciones = item.get_opciones_disponibles()
        
        lines = []
        lines.append("=" * 60)
        lines.append("DETALLE COMPLETO DE LA FILA")
        lines.append("=" * 60)
        lines.append("")
        lines.append(f"Fila Excel: {item.fila_excel}")
        lines.append(f"Estado: {item.status}")
        lines.append(f"Acción sugerida: {item.accion}")
        lines.append(f"Acción final: {item.accion_final or '(pendiente)'}")
        lines.append(f"Decisión: {item.decision}")
        lines.append("")
        if item.motivo_riesgo:
            lines.append(f"⚠️ Motivo de riesgo: {item.motivo_riesgo}")
            lines.append("")
        lines.append("Opciones disponibles:")
        for opt in opciones:
            lines.append(f"  • {opt}")
        lines.append("")
        lines.append(f"Aplicar: {'Sí' if item.aplicar else 'No'}")
        lines.append("")
        lines.append(f"Código actual: {item.codigo_actual}")
        lines.append(f"Existe en Oracle: {'Sí' if item.existe_actual else 'No'}")
        lines.append("")
        lines.append("Descripción actual Excel:")
        lines.append(str(item.descripcion_actual_excel))
        lines.append("")
        lines.append("Descripción actual Oracle:")
        lines.append(str(item.descripcion_actual_oracle))
        lines.append("")
        lines.append(f"Código nuevo: {item.codigo_nuevo}")
        lines.append(f"Existe en Oracle: {'Sí' if item.existe_nuevo else 'No'}")
        lines.append("")
        lines.append("Descripción nueva:")
        lines.append(str(item.descripcion_nueva))
        lines.append("")
        lines.append(f"Tipo: {item.tipo}")
        lines.append(f"Score Oracle/Excel: {item.score_oracle_excel:.1f}%")
        lines.append(f"Score Actual/Nueva: {item.score_actual_nueva:.1f}%")
        lines.append("")
        lines.append(f"Tiene equivalencias: {'Sí' if item.tiene_equivalencias else 'No'}")
        
        dialog = QDialog(self)
        dialog.setWindowTitle("🔍 Detalle de la Fila")
        dialog.resize(900, 650)
        
        text = QTextEdit()
        text.setReadOnly(True)
        text.setPlainText("\n".join(lines))
        
        buttons = QDialogButtonBox(QDialogButtonBox.StandardButton.Close)
        buttons.rejected.connect(dialog.reject)
        
        layout = QVBoxLayout()
        layout.addWidget(text)
        layout.addWidget(buttons)
        
        dialog.setLayout(layout)
        dialog.exec()
        
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
