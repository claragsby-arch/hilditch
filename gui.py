import sys
import os
from pathlib import Path
from PySide6.QtWidgets import (
    QApplication, QMainWindow, QWidget, QVBoxLayout, QHBoxLayout,
    QPushButton, QLabel, QLineEdit, QTextEdit, QFileDialog, QProgressBar,
    QCheckBox
)
from PySide6.QtCore import Qt, QThread, Signal
from datetime import datetime, timedelta
import time
from PySide6.QtGui import QFont
import pandas as pd
from openpyxl import load_workbook
from translate_prepare import preparer_et_traduire_excel, import_format_csv


class ProcessWorker(QThread):
    """Worker thread for processing Excel files without freezing the GUI"""
    progress = Signal(str)
    progress_update = Signal(int, str)  # percentage, time_remaining
    finished = Signal(bool)
    
    def __init__(self, file_path, numero_vente, create_xlsx=True, include_provenance=True):
        super().__init__()
        self.file_path = file_path
        self.numero_vente = numero_vente
        self.create_xlsx = create_xlsx
        self.include_provenance = include_provenance
        self.start_time = None
    
    def run(self):
        try:
            self.start_time = time.time()
            self.progress.emit("Démarrage du traitement...\n")
            self.progress_update.emit(0, "Calcul en cours...")
            
            # Préparation et traduction avec callback de progression
            self.progress.emit("Préparation et traduction du fichier Excel...\n")
            fichier_sortie_excel = f"{self.numero_vente}_excel.xlsx"
            
            def progress_callback(message):
                self.progress.emit(message + "\n")
                # Extraire le pourcentage si présent dans le message
                if "%" in message:
                    try:
                        # Chercher un nombre suivi de %
                        import re
                        match = re.search(r'(\d+)%', message)
                        if match:
                            percentage = int(match.group(1))
                            self._update_progress(percentage)
                    except:
                        pass
            
            result_df = preparer_et_traduire_excel(
                self.file_path,
                fichier_sortie_excel,
                progress_callback,
                save_xlsx=self.create_xlsx,
                include_provenance=self.include_provenance
            )
            
            # Conversion en CSV (100%)
            if self.create_xlsx:
                self.progress.emit("Conversion en CSV...\n")
                fichier_sortie_csv = f"{self.numero_vente}_CSV.csv"
                import_format_csv(fichier_sortie_excel, fichier_sortie_csv, progress_callback)
            else:
                self.progress.emit("Conversion en CSV sans création XLSX...\n")
                fichier_sortie_csv = f"{self.numero_vente}_CSV.csv"
                import_format_csv(nom_fichier_csv=fichier_sortie_csv, progress_callback=progress_callback, df=result_df)
            self._update_progress(100)
            
            self.progress.emit(f"\n✓ Terminé avec succès !\n")
            self.progress.emit(f"  - Excel: {fichier_sortie_excel if self.create_xlsx else 'non créé'}\n")
            self.progress.emit(f"  - CSV: {fichier_sortie_csv}\n")
            
            self.progress.emit(f"\n✓ Terminé avec succès !\n")
            self.progress.emit(f"  - Excel: {fichier_sortie_excel}\n")
            self.progress.emit(f"  - CSV: {fichier_sortie_csv}\n")
            self.finished.emit(True)
        except Exception as e:
            self.progress.emit(f"\n✗ Erreur : {str(e)}\n")
            self.finished.emit(False)
    
    def _update_progress(self, percentage):
        """Update progress with time estimation"""
        if self.start_time:
            elapsed = time.time() - self.start_time
            if percentage > 0 and percentage < 100:
                total_estimated = elapsed * (100 / percentage)
                remaining = total_estimated - elapsed
                time_str = self._format_time(remaining)
                self.progress_update.emit(percentage, time_str)
            elif percentage == 100:
                total_time = self._format_time(elapsed)
                self.progress_update.emit(100, f"Temps total: {total_time}")
    
    def _format_time(self, seconds):
        """Format seconds to readable time"""
        if seconds < 60:
            return f"{int(seconds)}s"
        elif seconds < 3600:
            minutes = int(seconds / 60)
            secs = int(seconds % 60)
            return f"{minutes}m {secs}s"
        else:
            hours = int(seconds / 3600)
            minutes = int((seconds % 3600) / 60)
            return f"{hours}h {minutes}m"


class MainWindow(QMainWindow):
    def __init__(self):
        super().__init__()
        self.worker = None
        self.init_ui()

    def init_ui(self):
        """Initialize the user interface"""
        self.setWindowTitle("Optimisation Hilditch Interface")
        self.setGeometry(100, 100, 1000, 700)

        # Create central widget
        central_widget = QWidget()
        self.setCentralWidget(central_widget)

        # Create main layout
        main_layout = QVBoxLayout()

        # Title
        title = QLabel("Optimisation Hilditch - Excel Processing")
        title_font = QFont()
        title_font.setPointSize(14)
        title_font.setBold(True)
        title.setFont(title_font)
        main_layout.addWidget(title)

        # File selection section
        file_layout = QHBoxLayout()
        file_label = QLabel("Excel File:")
        self.file_input = QLineEdit()
        self.file_input.setReadOnly(True)
        browse_btn = QPushButton("Browse")
        browse_btn.clicked.connect(self.browse_file)
        file_layout.addWidget(file_label)
        file_layout.addWidget(self.file_input)
        file_layout.addWidget(browse_btn)
        main_layout.addLayout(file_layout)

        # Numero de vente section
        numero_layout = QHBoxLayout()
        numero_label = QLabel("Numéro de vente:")
        self.numero_input = QLineEdit()
        self.numero_input.setText("9525")
        numero_layout.addWidget(numero_label)
        numero_layout.addWidget(self.numero_input)
        main_layout.addLayout(numero_layout)

        # Option de génération XLSX
        xlsx_option_layout = QHBoxLayout()
        self.xlsx_checkbox = QCheckBox("Créer le fichier XLSX")
        self.xlsx_checkbox.setChecked(False)
        xlsx_option_layout.addWidget(self.xlsx_checkbox)
        xlsx_option_layout.addStretch()
        main_layout.addLayout(xlsx_option_layout)

        # Option colonne Provenance
        provenance_option_layout = QHBoxLayout()
        self.provenance_checkbox = QCheckBox("Inclure la colonne Provenance")
        self.provenance_checkbox.setChecked(False)
        provenance_option_layout.addWidget(self.provenance_checkbox)
        provenance_option_layout.addStretch()
        main_layout.addLayout(provenance_option_layout)

        # Processing options
        options_label = QLabel("Actions:")
        options_font = QFont()
        options_font.setBold(True)
        options_label.setFont(options_font)
        main_layout.addWidget(options_label)

        # Buttons
        buttons_layout = QHBoxLayout()
        
        prepare_btn = QPushButton("Préparer & Traduire")
        prepare_btn.clicked.connect(self.prepare_data)
        
        buttons_layout.addWidget(prepare_btn)
        main_layout.addLayout(buttons_layout)

        # Progress section
        progress_label = QLabel("Progression:")
        progress_label.setFont(options_font)
        main_layout.addWidget(progress_label)

        # Progress bar
        self.progress_bar = QProgressBar()
        self.progress_bar.setMinimum(0)
        self.progress_bar.setMaximum(100)
        self.progress_bar.setValue(0)
        main_layout.addWidget(self.progress_bar)

        # Time estimation
        time_layout = QHBoxLayout()
        time_label = QLabel("Temps restant:")
        self.time_label = QLabel("")
        time_layout.addWidget(time_label)
        time_layout.addWidget(self.time_label)
        time_layout.addStretch()
        main_layout.addLayout(time_layout)

        # Output section
        output_label = QLabel("Output:")
        output_label.setFont(options_font)
        main_layout.addWidget(output_label)

        self.output_text = QTextEdit()
        self.output_text.setReadOnly(True)
        main_layout.addWidget(self.output_text)

        # Set central widget layout
        central_widget.setLayout(main_layout)

    def prepare_data(self):
        """Prepare and translate data"""
        file_path = self.file_input.text()
        numero_vente = self.numero_input.text()
        
        if not file_path:
            self.output_text.append("❌ Erreur : Aucun fichier sélectionné !\n")
            return
        
        if not numero_vente.strip():
            self.output_text.append("❌ Erreur : Veuillez entrer un numéro de vente !\n")
            return
        
        # Reset progress bar
        self.progress_bar.setValue(0)
        self.time_label.setText("")
        self.output_text.clear()
        self.output_text.append("Démarrage du traitement...\n")
        
        # Create and start worker thread
        self.worker = ProcessWorker(
            file_path,
            numero_vente,
            self.xlsx_checkbox.isChecked(),
            self.provenance_checkbox.isChecked()
        )
        self.worker.progress.connect(self.output_text.append)
        self.worker.progress_update.connect(self.update_progress)
        self.worker.finished.connect(self.on_processing_finished)
        self.worker.start()
    
    def update_progress(self, percentage, time_remaining):
        """Update progress bar and time label"""
        self.progress_bar.setValue(percentage)
        self.time_label.setText(time_remaining)

    def on_processing_finished(self, success):
        """Handle processing completion"""
        if not success:
            self.output_text.append("\n⚠️ Le traitement s'est terminé avec des erreurs.\n")
    
    def browse_file(self):
        """Open file dialog to select Excel file"""
        file_path, _ = QFileDialog.getOpenFileName(
            self, "Select Excel File", "", "Excel Files (*.xlsx *.xls);;All Files (*)"
        )
        if file_path:
            self.file_input.setText(file_path)
            self.output_text.append(f"📂 Fichier sélectionné : {file_path}\n")


if __name__ == "__main__":
    app = QApplication(sys.argv)
    window = MainWindow()
    window.show()
    sys.exit(app.exec())
