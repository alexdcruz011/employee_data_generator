import sys
import pandas as pd
from datetime import (datetime, date, timedelta)
from PySide6.QtWidgets import (
    QApplication, QWidget, QVBoxLayout,
    QLabel, QLineEdit, QPushButton, QFileDialog
)
import random
import names
import os

class EmployeeDataWidget(QWidget):
    """Class EmployeeDataWidget"""
    def __init__(self):
        super().__init__()
        self.setWindowTitle("Employee Data Generator")
        self.resize(400, 200)
        self.layout = QVBoxLayout()

        # Input for number of rows
        self.row_input = QLineEdit()
        self.row_input.setPlaceholderText("Enter number of records")
        self.layout.addWidget(self.row_input)

        # Folder selection
        self.folder_button = QPushButton("Select Folder")
        self.folder_button.clicked.connect(self.select_folder)
        self.layout.addWidget(self.folder_button)

        # Generate data
        self.generate_button = QPushButton("Generate Data")
        self.generate_button.clicked.connect(self.generate_data)
        self.layout.addWidget(self.generate_button)

        # Export to Excel
        self.export_button = QPushButton("Export to Excel")
        self.export_button.setStyleSheet("background-color: blue; font-weight: bold; color: white;")
        self.export_button.clicked.connect(self.export_data)
        self.layout.addWidget(self.export_button)

        # Timestamp label
        self.status_label = QLabel(f"Selected folder: {os.getcwd()}")
        self.layout.addWidget(self.status_label)

        self.setLayout(self.layout)
        self.data = None
        self.folder_path = ""

    def select_folder(self):
        self.folder_path = QFileDialog.getExistingDirectory(self, "Select Folder")
        self.status_label.setText(f"Selected folder: {self.folder_path}")

    def generate_data(self):
        try:
            self.status_label.setText('Generating Data...')
            num_rows = int(self.row_input.text())
            self.data = pd.DataFrame({
                "emp_id": [i+1 for i in range(num_rows)],
                "full_name": [names.get_full_name() for _ in range(num_rows)],
                "department": self.random_department(num_rows),
                "salary": [random.randint(25000,120000) for _ in range(num_rows)],
                "hire_date": self.hire_date_generator(num_rows)
            }).head(num_rows)
            timestamp = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
            self.status_label.setText(f"Generated Data: {timestamp}")
        except ValueError:
            self.status_label.setText("Error: Number of rows must be a number greater than 0")

    def export_data(self):
        if self.data is not None and self.folder_path:
            file_path = f"{self.folder_path}/employees.xlsx"
            self.data.to_excel(file_path, index=False, sheet_name='Employees')
            self.status_label.setText(f"File Generated. Exported to Excel at {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}")
        else:
            self.status_label.setText("Error: No data or folder selected")

    def hire_date_generator(self, num_rows):
        min_date = date(2020,1,1)
        max_date = date.today()
        delta_days = max_date - min_date
        return [(min_date + timedelta(days=random.randint(0, delta_days.days))) for _ in range(num_rows)]
    
    def random_department(self, num_rows):
        departments = ["IT", "HR", "Operations","Administration" , "Finance"]
        return [random.choice(departments) for _ in range(num_rows)]

if __name__ == "__main__":
    app = QApplication(sys.argv)
    widget = EmployeeDataWidget()
    widget.show()
    sys.exit(app.exec())