# pdf_renamer_tool_v2.4.py
# Version 2.4 of the PDF Renamer Tool with type selection and file count display

from PyQt5.QtGui import QIcon  # For application and window icons
import os
import sys
import re
import csv
import pdfplumber  # For extracting text from PDFs
from datetime import datetime
from PyQt5.QtWidgets import (
    QApplication, QFileDialog, QMessageBox,
    QMainWindow, QPushButton, QVBoxLayout,
    QWidget, QLabel, QTextEdit, QProgressBar,
    QComboBox  # Dropdown for selecting PDF type
)
from PyQt5.QtCore import Qt, QThread, pyqtSignal
# ---- Telemetry setup ----------------------------------------
APPINSIGHTS_CONN_STRING = "__REPLACE_ME__"

# Initialize global variables
pdf_counter = None
option_counter = None

def setup_telemetry():
    global pdf_counter, option_counter
    
    print("=== TELEMETRY DEBUG START ===", file=sys.stderr)
    
    # Let's examine the connection string in detail
    placeholder = "__REPLACE_ME__"
    
    print(f"Connection string repr: {repr(APPINSIGHTS_CONN_STRING)}", file=sys.stderr)
    print(f"Connection string length: {len(APPINSIGHTS_CONN_STRING)}", file=sys.stderr)
    print(f"Placeholder repr: {repr(placeholder)}", file=sys.stderr)
    print(f"Placeholder length: {len(placeholder)}", file=sys.stderr)
    
    # Character-by-character comparison for first 20 chars
    print("First 20 characters comparison:", file=sys.stderr)
    for i in range(min(20, len(APPINSIGHTS_CONN_STRING))):
        char = APPINSIGHTS_CONN_STRING[i]
        print(f"  [{i}]: '{char}' (ord: {ord(char)})", file=sys.stderr)
    
    # Check exact equality
    is_equal = (APPINSIGHTS_CONN_STRING == placeholder)
    contains_placeholder = (placeholder in APPINSIGHTS_CONN_STRING)
    
    print(f"Equals check: {is_equal}", file=sys.stderr)
    print(f"Contains check: {contains_placeholder}", file=sys.stderr)
    
    # If it contains the placeholder, let's see where
    if contains_placeholder:
        index = APPINSIGHTS_CONN_STRING.find(placeholder)
        print(f"Placeholder found at index: {index}", file=sys.stderr)
        if index >= 0:
            before = APPINSIGHTS_CONN_STRING[:index]
            after = APPINSIGHTS_CONN_STRING[index + len(placeholder):]
            print(f"Before placeholder: {repr(before)}", file=sys.stderr)
            print(f"After placeholder: {repr(after)}", file=sys.stderr)
    
    try:
        # Use a simpler, more explicit check
        if len(APPINSIGHTS_CONN_STRING) < 50:  # Real connection string should be much longer
            print("❌ Connection string too short - likely still placeholder", file=sys.stderr)
            return False
            
        if not APPINSIGHTS_CONN_STRING.startswith("InstrumentationKey="):
            print("❌ Connection string doesn't start with InstrumentationKey=", file=sys.stderr)
            return False
            
        print("✓ Connection string passes basic validation", file=sys.stderr)
        
        # Try importing required modules
        print("Importing Azure modules...", file=sys.stderr)
        try:
            from azure.monitor.opentelemetry import configure_azure_monitor
            print("✓ azure.monitor.opentelemetry imported successfully", file=sys.stderr)
        except ImportError as e:
            print(f"❌ Failed to import azure.monitor.opentelemetry: {e}", file=sys.stderr)
            return False
            
        try:
            from opentelemetry import metrics
            print("✓ opentelemetry.metrics imported successfully", file=sys.stderr)
        except ImportError as e:
            print(f"❌ Failed to import opentelemetry.metrics: {e}", file=sys.stderr)
            return False

        # Configure Azure Monitor
        print("Configuring Azure Monitor...", file=sys.stderr)
        configure_azure_monitor(connection_string=APPINSIGHTS_CONN_STRING)
        print("✓ Azure Monitor configured successfully", file=sys.stderr)

        # Create meter and counters
        print("Creating telemetry meter...", file=sys.stderr)
        _meter = metrics.get_meter("pdf-renamer", "2.4.0")
        print("✓ Meter created successfully", file=sys.stderr)
        
        print("Creating counters...", file=sys.stderr)
        pdf_counter = _meter.create_counter(
            name="pdfs_renamed",
            unit="1",
            description="Total PDFs renamed per batch",
        )
        option_counter = _meter.create_counter(
            name="rename_option",
            unit="1",
            description="Option label per batch",
        )
        
        print("✓ Telemetry counters created successfully", file=sys.stderr)
        print("=== TELEMETRY DEBUG END ===", file=sys.stderr)
        return True
        
    except ImportError as e:
        print(f"❌ Import error: {e}", file=sys.stderr)
        print("=== TELEMETRY DEBUG END ===", file=sys.stderr)
        return False
    except Exception as e:
        print(f"❌ Unexpected error: {e}", file=sys.stderr)
        print(f"❌ Error type: {type(e).__name__}", file=sys.stderr)
        print("=== TELEMETRY DEBUG END ===", file=sys.stderr)
        return False

# Initialize telemetry
telemetry_enabled = setup_telemetry()

# Temporary test - remove after confirming telemetry works
if telemetry_enabled:
    print("🎉 Telemetry is enabled and ready", file=sys.stderr)
else:
    print("💥 Telemetry is disabled", file=sys.stderr)
# -------------------------------------------------------------

def sanitize_filename(name):
    """
    Replace invalid filename characters with hyphens to ensure filesystem compatibility.
    """
    return re.sub(r'[<>:"/\\|?*]', '-', name).strip()


def strip_after_label(text, cutoffs):
    """
    Given a text line and cutoff labels, return the portion before any cutoff appears.
    E.g. strip_after_label("Title/Titre: ABC Demo/Group Cible: XYZ", ["Demo/Group Cible:"]) -> "Title/Titre: ABC"
    """
    for cutoff in cutoffs:
        if cutoff in text:
            return text.split(cutoff)[0].strip()
    return text.strip()


class PDFRenamerThread(QThread):
    # Signals to communicate progress, logs, and results back to the UI thread
    progress_signal = pyqtSignal(int, int)       # current index, total count
    log_signal = pyqtSignal(str)                 # text log messages
    file_signal = pyqtSignal(str, str, str)      # original name, new name, status
    finished_signal = pyqtSignal(str)            # path to generated log file
    error_signal = pyqtSignal(str)               # error messages

    def __init__(self, folder, prefix):
        super().__init__()
        self.folder = folder       # Folder containing PDFs to rename
        self.prefix = prefix       # Prefix to add based on user selection

    def run(self):
        """
        Process each PDF in the folder:
        - Extract title and number from first page
        - Sanitize title for filenames
        - Prepend the chosen prefix
        - Rename file and log result
        """
        try:
            renamed_files = []
            # Labels that indicate where to stop reading extra title lines
            cutoffs = ["Type/Type:", "Coverage/Couverture:",
                       "Product/Produit:", "Demo/Group Cible:", "Advertiser/Annonceur:", "Contact Name/Nom du contact:", 
                       "Revenue Type/Type de revenu:", "Sec. Demo/Cible secondaire:"]
            # List all PDF files in folder
            pdf_files = [f for f in os.listdir(self.folder) if f.lower().endswith('.pdf')]
            total = len(pdf_files)

            # If no PDFs, notify and exit
            if total == 0:
                self.log_signal.emit("No PDF files found in the selected folder.")
                self.finished_signal.emit("")
                return

            self.log_signal.emit(f"Found {total} PDF files to process.")

            # Loop through each PDF and rename
            for idx, filename in enumerate(pdf_files):
                self.progress_signal.emit(idx + 1, total)
                self.log_signal.emit(f"Processing: {filename}")
                filepath = os.path.join(self.folder, filename)
                try:
                    # Extract text from first page
                    with pdfplumber.open(filepath) as pdf:
                        text = pdf.pages[0].extract_text()

                    if not text:
                        raise ValueError("No text found on page 1")

                    lines = text.splitlines()
                    title = "NoTitle"

                    # Find the line with "Title/Titre:" and capture subsequent lines
                    for i, line in enumerate(lines):
                        if "Title/Titre:" in line:
                            title_raw = strip_after_label(
                                line.split("Title/Titre:")[1].strip(), cutoffs)
                            # Append up to two following lines if they are part of title
                            for j in (i+1, i+2):
                                if j < len(lines):
                                    next_line = strip_after_label(lines[j].strip(), cutoffs)
                                    if next_line and not any(next_line.startswith(label) for label in cutoffs):
                                        if not (title_raw.endswith('-') and not title_raw.endswith(' -')):
                                            title_raw += ' '
                                        title_raw += next_line
                            title = title_raw.strip()
                            break

                    # Extract the Number/Numéro field
                    number_match = re.search(r'Number/Numéro:\s*(\d+)', text)
                    number = number_match.group(1).strip() if number_match else "NoNumber"

                    # Create safe filename and include user-selected prefix
                    safe_title = sanitize_filename(title)
                    new_name = f"{self.prefix}{safe_title} - {number}.pdf"
                    new_path = os.path.join(self.folder, new_name)

                    # Perform the file rename
                    self.log_signal.emit(f"  Renaming to: {new_name}")
                    os.rename(filepath, new_path)
                    renamed_files.append((filename, new_name, "Success"))
                    self.file_signal.emit(filename, new_name, "Success")

                except Exception as e:
                    # Log any errors encountered during processing
                    error_msg = f"ERROR: {str(e)}"
                    self.log_signal.emit(f"  {error_msg}")
                    renamed_files.append((filename, "", error_msg))
                    self.file_signal.emit(filename, "", error_msg)

            # After processing all files, write a CSV log
            log_path = os.path.join(
                self.folder, f"rename_log_{datetime.now().strftime('%Y%m%d_%H%M%S')}.csv")
            with open(log_path, mode="w", newline="", encoding="utf-8") as csvfile:
                writer = csv.writer(csvfile)
                writer.writerow(["Original Filename", "New Filename", "Status"])
                for row in renamed_files:
                    writer.writerow(row)
            
            # ---- Telemetry counters ---------------------------------
            if pdf_counter and option_counter:
                pdf_counter.add(total)
                option_counter.add(total, {"option": (self.prefix or "ORIGINAL").split('-')[0].strip()})
            # ----------------------------------------------------------
            
            # Signal completion with the log file path
            self.finished_signal.emit(log_path)

        except Exception as e:
            # Catch any top-level errors
            self.error_signal.emit(f"An error occurred: {str(e)}")


class PDFRenamerApp(QMainWindow):
    """
    Main application window:
    - Dropdown for PDF type
    - Button to select folder
    - Progress bar, log area, and status label
    """
    def __init__(self):
        super().__init__()
        # Load and set the window/taskbar icon
        ico_path = os.path.join(
            getattr(sys, '_MEIPASS', os.path.dirname(__file__)),
            'pdf_renamer_icon.ico'
        )
        self.setWindowIcon(QIcon(ico_path))
        self.initUI()

    def initUI(self):
        # Window properties
        self.setWindowTitle('PDF Renamer Tool')
        self.setGeometry(300, 300, 600, 550)

        central_widget = QWidget()
        self.setCentralWidget(central_widget)
        layout = QVBoxLayout(central_widget)

        # 1) PDF type selection label + dropdown
        self.type_label = QLabel('PDF Type:', self)
        layout.addWidget(self.type_label)
        self.type_combo = QComboBox(self)
        self.type_combo.addItems(['Original', 'NA Notice', 'Renegotiation', 'Post'])
        layout.addWidget(self.type_combo)

        # 2) Button to open folder chooser
        self.select_btn = QPushButton('Select Folder with PDFs', self)
        self.select_btn.clicked.connect(self.select_folder)
        layout.addWidget(self.select_btn)

        # 3) Display selected folder path
        self.folder_label = QLabel('No folder selected', self)
        layout.addWidget(self.folder_label)

        # 4) Progress bar (hidden until processing starts)
        self.progress_bar = QProgressBar(self)
        self.progress_bar.setVisible(False)
        layout.addWidget(self.progress_bar)

        # 5) Logging area to show real-time messages
        layout.addWidget(QLabel('Log:', self))
        self.log_text = QTextEdit(self)
        self.log_text.setReadOnly(True)
        layout.addWidget(self.log_text)

        # 6) Status label for overall state
        self.status_label = QLabel('Ready', self)
        layout.addWidget(self.status_label)

        self.show()

    def select_folder(self):
        """
        Open a folder dialog, count PDFs, confirm, and start processing with prefix.
        """
        folder = QFileDialog.getExistingDirectory(self, 'Select Folder with PDFs')
        if not folder:
            return
        # Update UI for selected folder
        self.folder_label.setText(folder)
        self.log_text.clear()
        self.status_label.setText('Ready')
        self.progress_bar.setVisible(False)

        # Count PDF files in chosen folder
        pdf_files = [f for f in os.listdir(folder) if f.lower().endswith('.pdf')]
        count = len(pdf_files)

        # Confirm with user; show count
        reply = QMessageBox.question(
            self, 'Start Processing',
            f'Start processing {count} PDFs in:\n{folder}?',
            QMessageBox.Yes | QMessageBox.No, QMessageBox.Yes
        )

        if reply == QMessageBox.Yes:
            # Determine prefix from dropdown selection
            sel = self.type_combo.currentText()
            if sel == 'NA Notice':
                prefix = 'NA - '
            elif sel == 'Renegotiation':
                prefix = 'RENEG - '
            elif sel == 'Post':
                prefix = 'POST - '    
            else:
                prefix = ''
            self.process_folder(folder, prefix)

    def process_folder(self, folder, prefix):
        """
        Initialize progress bar and start the background renaming thread.
        """
        self.progress_bar.setVisible(True)
        self.progress_bar.setValue(0)
        self.select_btn.setEnabled(False)
        self.status_label.setText('Processing...')

        # Start worker thread with chosen folder and prefix
        self.worker = PDFRenamerThread(folder, prefix)
        self.worker.progress_signal.connect(self.update_progress)
        self.worker.log_signal.connect(self.add_log)
        self.worker.file_signal.connect(self.handle_file_result)
        self.worker.finished_signal.connect(self.process_finished)
        self.worker.error_signal.connect(self.handle_error)
        self.worker.start()

    def update_progress(self, current, total):
        # Update the progress bar values
        self.progress_bar.setMaximum(total)
        self.progress_bar.setValue(current)

    def add_log(self, message):
        # Append log message and scroll to bottom
        self.log_text.append(message)
        self.log_text.verticalScrollBar().setValue(
            self.log_text.verticalScrollBar().maximum()
        )

    def handle_file_result(self, old_name, new_name, status):
        # Individual file results are already logged via log_signal
        pass

    def process_finished(self, log_path):
        # Re-enable UI and notify user of completion
        self.select_btn.setEnabled(True)
        if log_path:
            self.status_label.setText('Completed!')
            self.add_log(f"\nCompleted! Log saved to: {log_path}")
            QMessageBox.information(
                self, 'Completed',
                f'PDF renaming process completed.\nLog saved to:\n{log_path}'
            )
        else:
            self.status_label.setText('Ready')

    def handle_error(self, error_message):
        # Handle any unexpected errors
        self.select_btn.setEnabled(True)
        self.status_label.setText('Error')
        self.log_text.append(f"ERROR: {error_message}")
        QMessageBox.critical(self, 'Error', error_message)


def excepthook(exc_type, exc_value, exc_traceback):
    """
    Global exception hook to catch and log uncaught exceptions,
    and show an error dialog rather than crashing silently.
    """
    print(f"Uncaught exception: {exc_value}", file=sys.stderr)
    try:
        with open(f"error_log_{datetime.now().strftime('%Y%m%d_%H%M%S')}.txt", "w") as f:
            import traceback
            traceback.print_exception(
                exc_type, exc_value, exc_traceback, file=f
            )
    except Exception:
        pass
    if QApplication.instance():
        QMessageBox.critical(
            None, "Error", f"An unexpected error occurred:\n{exc_value}"
        )
    sys.__excepthook__(exc_type, exc_value, exc_traceback)


if __name__ == '__main__':
    # Install exception hook and start the GUI event loop
    sys.excepthook = excepthook
    app = QApplication(sys.argv)
    ico_path = os.path.join(
        getattr(sys, '_MEIPASS', os.path.dirname(__file__)),
        'pdf_renamer_icon.ico'
    )
    app.setWindowIcon(QIcon(ico_path))
    ex = PDFRenamerApp()
    sys.exit(app.exec_())
