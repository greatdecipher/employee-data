import sys
import pandas as pd
from PySide6.QtWidgets import (
    QApplication, QWidget, QVBoxLayout, QHBoxLayout, QLineEdit, QPushButton,
    QLabel, QFileDialog, QMessageBox, QFrame, QSpacerItem, QSizePolicy,
    QGroupBox, QProgressBar
)
from PySide6.QtCore import Qt
from PySide6.QtGui import QFont, QPalette, QIcon
from .generator import generate_employee_data
from .exporter import export_to_excel

class EmployeeApp(QWidget):
    def __init__(self):
        super().__init__()
        self.setWindowTitle("🏢 Employee Data Generator - Professional Edition")
        self.setMinimumSize(600, 650)
        self.resize(700, 750)
        self.folder_path = ""
        self.data = pd.DataFrame()
        self.setup_styling()
        self.init_ui()

    def setup_styling(self):
        """Apply modern styling to the application."""
        self.setStyleSheet("""
            QWidget {
                background-color: #fafbfc;
                font-family: 'Segoe UI', 'Helvetica Neue', Arial, sans-serif;
                font-size: 13px;
            }
            
            QGroupBox {
                font-weight: 600;
                border: 1px solid #e1e8ed;
                border-radius: 12px;
                margin-top: 1ex;
                padding-top: 20px;
                background-color: #ffffff;
                color: #2c3e50;
            }
            
            QGroupBox::title {
                subcontrol-origin: margin;
                left: 15px;
                padding: 0 15px 0 15px;
                color: #34495e;
                font-size: 15px;
                font-weight: 600;
            }
            
            QLineEdit {
                border: 2px solid #e1e8ed;
                border-radius: 8px;
                padding: 12px 15px;
                font-size: 14px;
                background-color: #ffffff;
                color: #2c3e50;
                selection-background-color: #3498db;
                selection-color: white;
            }
            
            QLineEdit:focus {
                border-color: #3498db;
                background-color: #ffffff;
            }
            
            QLineEdit:hover {
                border-color: #bdc3c7;
            }
            
            QPushButton {
                background-color: #3498db;
                color: white;
                border: none;
                border-radius: 8px;
                padding: 14px 20px;
                font-size: 14px;
                font-weight: 600;
                min-height: 20px;
                text-align: center;
            }
            
            QPushButton:hover {
                background-color: #2980b9;
            }
            
            QPushButton:pressed {
                background-color: #21618c;
            }
            
            QPushButton:disabled {
                background-color: #ecf0f1;
                color: #95a5a6;
                border: 1px solid #e1e8ed;
            }
            
            QPushButton#primaryButton {
                background-color: #27ae60;
                color: #ffffff;
                font-size: 15px;
                padding: 16px 24px;
                font-weight: 600;
            }
            
            QPushButton#primaryButton:hover {
                background-color: #229954;
                color: #ffffff;
            }
            
            QPushButton#primaryButton:pressed {
                background-color: #1e8449;
                color: #ffffff;
            }
            
            QPushButton#secondaryButton {
                background-color: #e74c3c;
                color: #ffffff;
                font-size: 15px;
                padding: 16px 24px;
                font-weight: 600;
            }
            
            QPushButton#secondaryButton:hover {
                background-color: #c0392b;
                color: #ffffff;
            }
            
            QPushButton#secondaryButton:pressed {
                background-color: #a93226;
                color: #ffffff;
            }
            
            QLabel {
                color: #2c3e50;
                font-size: 13px;
                background-color: transparent;
            }
            
            QLabel#titleLabel {
                font-size: 28px;
                font-weight: 700;
                color: #2c3e50;
                margin: 15px 0;
                background-color: transparent;
            }
            
            QLabel#stepLabel {
                font-size: 16px;
                font-weight: 600;
                color: #34495e;
                margin: 8px 0;
                background-color: transparent;
            }
            
            QLabel#infoLabel {
                color: #7f8c8d;
                font-style: italic;
                font-size: 13px;
                background-color: transparent;
            }
            
            QProgressBar {
                border: 1px solid #e1e8ed;
                border-radius: 6px;
                text-align: center;
                background-color: #f8f9fa;
                color: #2c3e50;
                font-weight: 600;
                min-height: 28px;
            }
            
            QProgressBar::chunk {
                background-color: #3498db;
                border-radius: 4px;
            }
            
            QProgressBar[value="0"] {
                color: #95a5a6;
            }
            
            QProgressBar[value="1"] {
                color: #f39c12;
            }
            
            QProgressBar[value="2"] {
                color: #e67e22;
            }
            
            QProgressBar[value="3"] {
                color: #27ae60;
            }
        """)

    def center_on_screen(self):
        """Move the window to the center of the primary screen's available geometry.

        Uses Qt's screen geometry to compute the center point and repositions
        the window so it appears centered when shown.
        """
        try:
            screen = QApplication.primaryScreen()
            if screen is None:
                return
            screen_geom = screen.availableGeometry()
            frame_geom = self.frameGeometry()
            center_point = screen_geom.center()
            frame_geom.moveCenter(center_point)
            self.move(frame_geom.topLeft())
        except Exception:
            # If anything goes wrong (headless, WSL without display, etc.),
            # just skip centering silently.
            pass

    def center_dialog_on_window(self, dialog):
        """Center a dialog on the main window.
        
        Args:
            dialog: The QDialog or QMessageBox to center
        """
        try:
            # Get the geometry of the main window
            main_geom = self.frameGeometry()
            
            # Calculate the center point of the main window
            center_point = main_geom.center()
            
            # Move the dialog so its center aligns with the main window's center
            dialog_geom = dialog.frameGeometry()
            dialog_geom.moveCenter(center_point)
            dialog.move(dialog_geom.topLeft())
        except Exception:
            # If anything goes wrong, just skip centering
            pass

    def init_ui(self):
        main_layout = QVBoxLayout()
        main_layout.setSpacing(20)
        main_layout.setContentsMargins(30, 30, 30, 30)

        # Title
        title_label = QLabel("🏢 Employee Data Generator")
        title_label.setObjectName("titleLabel")
        title_label.setAlignment(Qt.AlignCenter)
        main_layout.addWidget(title_label)

        # Subtitle
        subtitle_label = QLabel("Generate realistic employee data and export to Excel")
        subtitle_label.setObjectName("infoLabel")
        subtitle_label.setAlignment(Qt.AlignCenter)
        main_layout.addWidget(subtitle_label)

        # Progress indicator
        self.progress_bar = QProgressBar()
        self.progress_bar.setMaximum(3)
        self.progress_bar.setValue(0)
        self.progress_bar.setFormat("Step %v of %m")
        main_layout.addWidget(self.progress_bar)

        # Step 1: Number of employees
        step1_group = QGroupBox("Step 1: Specify Number of Employees")
        step1_layout = QVBoxLayout()
        
        step1_info = QLabel("Enter how many employee records you want to generate:")
        step1_info.setObjectName("infoLabel")
        step1_layout.addWidget(step1_info)
        
        self.num_input = QLineEdit()
        self.num_input.setPlaceholderText("e.g., 50")
        self.num_input.textChanged.connect(self.validate_input)
        step1_layout.addWidget(self.num_input)
        
        step1_group.setLayout(step1_layout)
        main_layout.addWidget(step1_group)

        # Step 2: Folder selection
        step2_group = QGroupBox("Step 2: Choose Output Folder")
        step2_layout = QVBoxLayout()
        
        step2_info = QLabel("Select where you want to save the Excel file:")
        step2_info.setObjectName("infoLabel")
        step2_layout.addWidget(step2_info)
        
        folder_layout = QHBoxLayout()
        self.folder_label = QLabel("No folder selected")
        self.folder_label.setObjectName("infoLabel")
        folder_layout.addWidget(self.folder_label)
        
        self.select_btn = QPushButton("📁 Browse...")
        self.select_btn.clicked.connect(self.select_folder)
        folder_layout.addWidget(self.select_btn)
        
        step2_layout.addLayout(folder_layout)
        step2_group.setLayout(step2_layout)
        main_layout.addWidget(step2_group)

        # Step 3: Generate and export
        step3_group = QGroupBox("Step 3: Generate and Export Data")
        step3_layout = QVBoxLayout()
        
        step3_info = QLabel("Generate employee data and export to Excel:")
        step3_info.setObjectName("infoLabel")
        step3_layout.addWidget(step3_info)
        
        button_layout = QHBoxLayout()
        
        self.generate_btn = QPushButton("🎲 Generate Data")
        self.generate_btn.setObjectName("primaryButton")
        self.generate_btn.clicked.connect(self.generate_data)
        self.generate_btn.setEnabled(False)
        button_layout.addWidget(self.generate_btn)

        self.export_btn = QPushButton("📊 Export to Excel")
        self.export_btn.setObjectName("secondaryButton")
        self.export_btn.clicked.connect(self.export_excel)
        self.export_btn.setEnabled(False)
        button_layout.addWidget(self.export_btn)
        
        step3_layout.addLayout(button_layout)
        step3_group.setLayout(step3_layout)
        main_layout.addWidget(step3_group)

        # Status message
        self.message_label = QLabel("👋 Welcome! Start by entering the number of employees above.")
        self.message_label.setAlignment(Qt.AlignCenter)
        self.message_label.setWordWrap(True)
        main_layout.addWidget(self.message_label)

        # Add stretch to push everything up
        main_layout.addStretch()

        self.setLayout(main_layout)

    def validate_input(self):
        """Validate user input and update UI state accordingly."""
        text = self.num_input.text().strip()
        
        if text == "":
            self.generate_btn.setEnabled(False)
            self.message_label.setText("👋 Welcome! Start by entering the number of employees above.")
            self.progress_bar.setValue(0)
            return
            
        try:
            num = int(text)
            if num <= 0:
                self.generate_btn.setEnabled(False)
                self.message_label.setText("❌ Please enter a positive number.")
                self.progress_bar.setValue(0)
            elif num > 10000:
                self.generate_btn.setEnabled(False)
                self.message_label.setText("⚠️ Maximum 10,000 employees allowed.")
                self.progress_bar.setValue(0)
            else:
                self.generate_btn.setEnabled(True)
                self.message_label.setText(f"✅ Ready to generate {num} employees. Click 'Generate Data' below!")
                self.progress_bar.setValue(1)
        except ValueError:
            self.generate_btn.setEnabled(False)
            self.message_label.setText("❌ Please enter a valid number.")
            self.progress_bar.setValue(0)

    def select_folder(self):
        folder = QFileDialog.getExistingDirectory(self, "Select Output Folder")
        if folder:
            self.folder_path = folder
            # Shorten the path for display if it's too long
            display_path = folder
            if len(display_path) > 50:
                display_path = "..." + display_path[-47:]
            
            self.folder_label.setText(f"📁 {display_path}")
            self.message_label.setText(f"✅ Folder selected! {self.get_next_step_message()}")
            self.update_export_button_state()

    def get_next_step_message(self):
        """Get appropriate next step message based on current state."""
        if not self.data.empty and self.folder_path:
            return "Ready to export!"
        elif not self.data.empty:
            return "Now select a folder to export the data."
        elif self.num_input.text().strip() and self.folder_path:
            return "Now generate the employee data."
        else:
            return ""

    def update_export_button_state(self):
        """Enable export button only when data exists and folder is selected."""
        self.export_btn.setEnabled(not self.data.empty and bool(self.folder_path))
        if not self.data.empty and self.folder_path:
            self.progress_bar.setValue(3)

    def generate_data(self):
        try:
            n = int(self.num_input.text())
            if n <= 0:
                QMessageBox.warning(self, "Invalid Input", "Please enter a positive number.")
                return
            if n > 10000:
                QMessageBox.warning(self, "Too Many Records", "Maximum 10,000 employees allowed for performance reasons.")
                return
                
            # Show loading state
            self.generate_btn.setText("🔄 Generating...")
            self.generate_btn.setEnabled(False)
            QApplication.processEvents()  # Update UI
            
            self.data = generate_employee_data(n)
            
            # Reset button state
            self.generate_btn.setText("🎲 Generate Data")
            self.generate_btn.setEnabled(True)
            
            self.message_label.setText(f"🎉 Successfully generated {len(self.data)} employee records! {self.get_next_step_message()}")
            self.progress_bar.setValue(2)
            self.update_export_button_state()
            
        except ValueError:
            QMessageBox.warning(self, "Invalid Input", "Please enter a valid number.")
        except Exception as e:
            QMessageBox.critical(self, "Error", f"An error occurred while generating data:\n{str(e)}")
            self.generate_btn.setText("🎲 Generate Data")
            self.generate_btn.setEnabled(True)

    def export_excel(self):
        if self.data.empty:
            QMessageBox.warning(self, "No Data", "Please generate employee data first!")
            return
        if not self.folder_path:
            QMessageBox.warning(self, "No Folder Selected", "Please select an output folder first!")
            return

        try:
            # Show loading state
            self.export_btn.setText("💾 Exporting...")
            self.export_btn.setEnabled(False)
            QApplication.processEvents()  # Update UI
            
            file_path = export_to_excel(self.data, self.folder_path)
            
            # Reset button state
            self.export_btn.setText("📊 Export to Excel")
            self.export_btn.setEnabled(True)
            
            self.message_label.setText("🎉 Excel file created successfully!")
            self.progress_bar.setValue(3)
            
            # Show success dialog with file location
            msg = QMessageBox(self)
            msg.setIcon(QMessageBox.Information)
            msg.setWindowTitle("Export Complete")
            msg.setText("Employee data exported successfully!")
            msg.setInformativeText(f"File saved at:\n{file_path}")
            msg.addButton("Open Folder", QMessageBox.ActionRole)
            msg.addButton("OK", QMessageBox.AcceptRole)
            
            # Center the dialog on the main window
            self.center_dialog_on_window(msg)
            
            result = msg.exec()
            if result == 0:  # Open Folder clicked
                import subprocess
                import os
                if os.name == 'nt':  # Windows
                    subprocess.run(['explorer', str(file_path.parent)])
                elif os.name == 'posix':  # macOS/Linux
                    subprocess.run(['xdg-open', str(file_path.parent)])
                    
        except Exception as e:
            QMessageBox.critical(self, "Export Error", f"An error occurred while exporting:\n{str(e)}")
            self.export_btn.setText("📊 Export to Excel")
            self.export_btn.setEnabled(True)

def run_app():
    app = QApplication(sys.argv)
    window = EmployeeApp()
    # Center the window on screen where possible, then show.
    window.center_on_screen()
    window.show()
    sys.exit(app.exec())