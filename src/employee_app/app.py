import sys
import pandas as pd
from PySide6.QtWidgets import (
    QApplication, QWidget, QVBoxLayout, QLineEdit, QPushButton,
    QLabel, QFileDialog, QMessageBox
)
from .generator import generate_employee_data
from .exporter import export_to_excel

class EmployeeApp(QWidget):
    def __init__(self):
        super().__init__()
        self.setWindowTitle("Employee Data Generator")
        self.resize(400, 300)
        self.folder_path = ""
        self.data = pd.DataFrame()
        self.init_ui()

    def init_ui(self):
        layout = QVBoxLayout()

        self.num_input = QLineEdit()
        self.num_input.setPlaceholderText("Enter number of employees")
        layout.addWidget(self.num_input)

        self.select_btn = QPushButton("Select Folder")
        self.select_btn.clicked.connect(self.select_folder)
        layout.addWidget(self.select_btn)

        self.generate_btn = QPushButton("Generate Data")
        self.generate_btn.clicked.connect(self.generate_data)
        layout.addWidget(self.generate_btn)

        self.export_btn = QPushButton("Export to Excel")
        self.export_btn.clicked.connect(self.export_excel)
        layout.addWidget(self.export_btn)

        self.message_label = QLabel("")
        layout.addWidget(self.message_label)

        self.setLayout(layout)

    def select_folder(self):
        folder = QFileDialog.getExistingDirectory(self, "Select Folder")
        if folder:
            self.folder_path = folder
            self.message_label.setText(f"Selected folder: {folder}")

    def generate_data(self):
        try:
            n = int(self.num_input.text())
            self.data = generate_employee_data(n)
            self.message_label.setText(f"Generated {len(self.data)} employees successfully!")
        except ValueError:
            QMessageBox.warning(self, "Invalid Input", "Please enter a valid number.")

    def export_excel(self):
        if self.data.empty:
            QMessageBox.warning(self, "No Data", "Generate data first!")
            return
        if not self.folder_path:
            QMessageBox.warning(self, "No Folder", "Select a folder first!")
            return

        file_path = export_to_excel(self.data, self.folder_path)
        self.message_label.setText("File generated successfully!")
        QMessageBox.information(self, "Export Complete", f"Saved at:\n{file_path}")

def run_app():
    app = QApplication(sys.argv)
    window = EmployeeApp()
    window.show()
    sys.exit(app.exec())