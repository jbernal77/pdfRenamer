from PyQt5.QtGui import QIcon
import os
import sys
import re
import csv
import pdfplumber
from datetime import datetime
from PyQt5.QtWidgets import (QApplication, QFileDialog, QMessageBox, 
                             QMainWindow, QPushButton, QVBoxLayout, 
                             QWidget, QLabel, QTextEdit, QProgressBar)
from PyQt5.QtCore import Qt, QThread, pyqtSignal


def sanitize_filename(name):
    return re.sub(r'[<>:"/\\|?*]', '-', name).strip()


def strip_after_label(text, cutoffs):
    for cutoff in cutoffs:
        if cutoff in text:
            return text.split(cutoff)[0].strip()
    return text.strip()


class PDFRenamerThread(QThread):
    progress_signal = pyqtSignal(int, int)  # current, total
    log_signal = pyqtSignal(str)
    file_signal = pyqtSignal(str, str, str)  # old_name, new_name, status
    finished_signal = pyqtSignal(str)  # log file path
    error_signal = pyqtSignal(str)

    def __init__(self, folder):
        super().__init__()
        self.folder = folder

    def run(self):
        try:
            renamed_files = []
            cutoffs = ["Type/Type:", "Coverage/Couverture:", "Product/Produit:", "Demo/Group Cible:"]
            
            pdf_files = [f for f in os.listdir(self.folder) if f.lower().endswith('.pdf')]
            total = len(pdf_files)
            
            if total == 0:
                self.log_signal.emit("No PDF files found in the selected folder.")
                self.finished_signal.emit("")
                return
                
            self.log_signal.emit(f"Found {total} PDF files to process.")
            
            for idx, filename in enumerate(pdf_files):
                self.progress_signal.emit(idx + 1, total)
                self.log_signal.emit(f"Processing: {filename}")
                
                filepath = os.path.join(self.folder, filename)
                try:
                    with pdfplumber.open(filepath) as pdf:
                        text = pdf.pages[0].extract_text()

                    if not text:
                        raise ValueError("No text found on page 1")

                    lines = text.splitlines()
                    title = "NoTitle"

                    for i, line in enumerate(lines):
                        if "Title/Titre:" in line:
                            title_raw = strip_after_label(line.split("Title/Titre:")[1].strip(), cutoffs)

                            if i + 1 < len(lines):
                                next_line = strip_after_label(lines[i + 1].strip(), cutoffs)
                                if next_line and not any(next_line.startswith(label) for label in cutoffs):
                                    title_raw += " " + next_line

                            if i + 2 < len(lines):
                                next_next_line = strip_after_label(lines[i + 2].strip(), cutoffs)
                                if next_next_line and not any(next_next_line.startswith(label) for label in cutoffs):
                                    title_raw += " " + next_next_line

                            title = title_raw.strip()
                            break

                    number_match = re.search(r'Number/Numéro:\s*(\d+)', text)
                    number = number_match.group(1).strip() if number_match else "NoNumber"

                    safe_title = sanitize_filename(title)
                    new_name = f"{safe_title} - {number}.pdf"
                    new_path = os.path.join(self.folder, new_name)
                    
                    self.log_signal.emit(f"  Renaming to: {new_name}")
                    os.rename(filepath, new_path)
                    renamed_files.append((filename, new_name, "Success"))
                    self.file_signal.emit(filename, new_name, "Success")
                    
                except Exception as e:
                    error_msg = f"ERROR: {str(e)}"
                    self.log_signal.emit(f"  {error_msg}")
                    renamed_files.append((filename, "", error_msg))
                    self.file_signal.emit(filename, "", error_msg)

            # Write log to CSV
            log_path = os.path.join(self.folder, f"rename_log_{datetime.now().strftime('%Y%m%d_%H%M%S')}.csv")
            with open(log_path, mode="w", newline="", encoding="utf-8") as csvfile:
                writer = csv.writer(csvfile)
                writer.writerow(["Original Filename", "New Filename", "Status"])
                for row in renamed_files:
                    writer.writerow(row)

            self.finished_signal.emit(log_path)
            
        except Exception as e:
            self.error_signal.emit(f"An error occurred: {str(e)}")


class PDFRenamerApp(QMainWindow):
    def __init__(self):
        super().__init__()

        # --- set window‑specific icon ---
        ico_path = os.path.join(getattr(sys, '_MEIPASS', os.path.dirname(__file__)),
                                'pdf_renamer_icon.ico')
        self.setWindowIcon(QIcon(ico_path))
        # --------------------------------

        self.initUI()
        
    def initUI(self):
        self.setWindowTitle('PDF Renamer Tool')
        self.setGeometry(300, 300, 600, 500)
        
        # Central widget and layout
        central_widget = QWidget()
        self.setCentralWidget(central_widget)
        layout = QVBoxLayout(central_widget)
        
        # Select folder button
        self.select_btn = QPushButton('Select Folder with PDFs', self)
        self.select_btn.clicked.connect(self.select_folder)
        layout.addWidget(self.select_btn)
        
        # Folder label
        self.folder_label = QLabel('No folder selected', self)
        layout.addWidget(self.folder_label)
        
        # Progress bar
        self.progress_bar = QProgressBar(self)
        self.progress_bar.setVisible(False)
        layout.addWidget(self.progress_bar)
        
        # Log area
        layout.addWidget(QLabel('Log:', self))
        self.log_text = QTextEdit(self)
        self.log_text.setReadOnly(True)
        layout.addWidget(self.log_text)
        
        # Status label
        self.status_label = QLabel('Ready', self)
        layout.addWidget(self.status_label)
        
        self.show()
        
    def select_folder(self):
        folder = QFileDialog.getExistingDirectory(self, 'Select Folder with PDFs')
        if folder:
            self.folder_label.setText(folder)
            self.log_text.clear()
            self.status_label.setText('Ready')
            self.progress_bar.setVisible(False)
            
            reply = QMessageBox.question(self, 'Start Processing', 
                                         f'Start processing PDFs in:\n{folder}?',
                                         QMessageBox.Yes | QMessageBox.No, QMessageBox.Yes)
            
            if reply == QMessageBox.Yes:
                self.process_folder(folder)
    
    def process_folder(self, folder):
        self.progress_bar.setVisible(True)
        self.progress_bar.setValue(0)
        self.select_btn.setEnabled(False)
        self.status_label.setText('Processing...')
        
        # Create and start worker thread
        self.worker = PDFRenamerThread(folder)
        self.worker.progress_signal.connect(self.update_progress)
        self.worker.log_signal.connect(self.add_log)
        self.worker.file_signal.connect(self.handle_file_result)
        self.worker.finished_signal.connect(self.process_finished)
        self.worker.error_signal.connect(self.handle_error)
        self.worker.start()
    
    def update_progress(self, current, total):
        self.progress_bar.setMaximum(total)
        self.progress_bar.setValue(current)
    
    def add_log(self, message):
        self.log_text.append(message)
        # Auto-scroll to bottom
        self.log_text.verticalScrollBar().setValue(self.log_text.verticalScrollBar().maximum())
    
    def handle_file_result(self, old_name, new_name, status):
        pass  # We're using the log_signal for displaying this info
    
    def process_finished(self, log_path):
        self.select_btn.setEnabled(True)
        if log_path:
            self.status_label.setText('Completed!')
            self.add_log(f"\nCompleted! Log saved to: {log_path}")
            QMessageBox.information(self, 'Completed', 
                                   f'PDF renaming process completed.\nLog saved to:\n{log_path}')
        else:
            self.status_label.setText('Ready')
    
    def handle_error(self, error_message):
        self.select_btn.setEnabled(True)
        self.status_label.setText('Error')
        self.add_log(f"ERROR: {error_message}")
        QMessageBox.critical(self, 'Error', error_message)


def excepthook(exc_type, exc_value, exc_traceback):
    """Handle uncaught exceptions to prevent silent failures"""
    print(f"Uncaught exception: {exc_value}", file=sys.stderr)
    # Write to error log
    try:
        with open(f"error_log_{datetime.now().strftime('%Y%m%d_%H%M%S')}.txt", "w") as f:
            import traceback
            traceback.print_exception(exc_type, exc_value, exc_traceback, file=f)
    except:
        pass
    # Show error dialog
    if QApplication.instance():
        QMessageBox.critical(None, "Error", f"An unexpected error occurred:\n{exc_value}")
    # Call the default handler
    sys.__excepthook__(exc_type, exc_value, exc_traceback)


if __name__ == '__main__':
    # Set the exception hook to catch unhandled exceptions
    sys.excepthook = excepthook
    
    app = QApplication(sys.argv)
    
# --- set task‑bar / title‑bar icon ---
ico_path = os.path.join(getattr(sys, '_MEIPASS', os.path.dirname(__file__)),
                        'pdf_renamer_icon.ico')
app.setWindowIcon(QIcon(ico_path))
# -------------------------------------
ex = PDFRenamerApp()
sys.exit(app.exec_())
