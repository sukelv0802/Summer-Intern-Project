import sys
import time
import serial
import pandas as pd
import openpyxl
from openpyxl.styles import PatternFill
from PyQt5.QtWidgets import (QApplication, QMainWindow, QTreeWidget, QTreeWidgetItem, 
                             QPushButton, QVBoxLayout, QHBoxLayout, QWidget, QLabel, 
                             QLineEdit, QMessageBox, QHeaderView, QSplitter, QStyle,
                             QStyleFactory, QComboBox, QFileDialog, QDateTimeEdit,
                             QHBoxLayout, QVBoxLayout, QPushButton, QLabel, QLineEdit)
from PyQt5.QtCore import QTimer, Qt, QSettings, QDateTime
from PyQt5.QtGui import QFont, QColor

class MainWindow(QMainWindow):
    def __init__(self):
        super().__init__()
        self.setWindowTitle("Serial Data Logger")
        self.setGeometry(100, 100, 1000, 600)

        # Settings
        self.settings = QSettings("Test", "SerialDataLogger")
        self.load_settings()

        # Main layout
        main_widget = QWidget()
        main_layout = QHBoxLayout(main_widget)
        
        # Splitter for resizable panels
        splitter = QSplitter(Qt.Horizontal)
        
        # Options bar
        options_bar = QWidget()
        options_layout = QVBoxLayout(options_bar)
        options_layout.setAlignment(Qt.AlignTop)
        
        threshold_label = QLabel("Threshold")
        threshold_label.setFont(QFont("Arial", 12, QFont.Bold))
        self.threshold_entry = QLineEdit()
        self.threshold_entry.setPlaceholderText("Enter threshold value")
        confirm_button = QPushButton("Set Threshold")
        confirm_button.clicked.connect(self.set_threshold)
        
        # COM Port selection
        com_label = QLabel("COM Port")
        com_label.setFont(QFont("Arial", 12, QFont.Bold))
        self.com_combo = QComboBox()
        self.com_combo.addItems([f"COM{i}" for i in range(1, 21)])
        self.com_combo.setCurrentText(self.settings.value("com_port", "COM10"))
        
        # Theme selection
        theme_label = QLabel("Theme")
        theme_label.setFont(QFont("Arial", 12, QFont.Bold))
        self.theme_combo = QComboBox()
        self.theme_combo.addItems(QStyleFactory.keys())
        self.theme_combo.setCurrentText(self.settings.value("theme", "Fusion"))
        self.theme_combo.currentTextChanged.connect(self.change_theme)
        
        options_layout.addWidget(threshold_label)
        options_layout.addWidget(self.threshold_entry)
        options_layout.addWidget(confirm_button)
        options_layout.addWidget(com_label)
        options_layout.addWidget(self.com_combo)
        options_layout.addWidget(theme_label)
        options_layout.addWidget(self.theme_combo)
        options_layout.addStretch()
        
        # Tree widget
        self.tree = QTreeWidget()
        self.tree.setHeaderLabels(["Timestamp", "Data"])
        self.tree.header().setSectionResizeMode(QHeaderView.Stretch)
        self.tree.setAlternatingRowColors(True)
        self.tree.setSortingEnabled(True)
        
        # Filter layout
        filter_layout = QVBoxLayout()

        # Timestamp range filter
        timestamp_layout = QHBoxLayout()
        self.start_time = QDateTimeEdit(QDateTime.currentDateTime().addDays(-1))
        self.end_time = QDateTimeEdit(QDateTime.currentDateTime())
        timestamp_layout.addWidget(QLabel("From:"))
        timestamp_layout.addWidget(self.start_time)
        timestamp_layout.addWidget(QLabel("To:"))
        timestamp_layout.addWidget(self.end_time)

        # Value range filter
        value_layout = QHBoxLayout()
        self.min_value = QLineEdit()
        self.max_value = QLineEdit()
        value_layout.addWidget(QLabel("Min Value:"))
        value_layout.addWidget(self.min_value)
        value_layout.addWidget(QLabel("Max Value:"))
        value_layout.addWidget(self.max_value)

        # Apply filter button
        apply_filter_button = QPushButton("Apply Filter")
        apply_filter_button.clicked.connect(self.apply_filter)
        filter_layout.addLayout(timestamp_layout)
        filter_layout.addLayout(value_layout)
        filter_layout.addWidget(apply_filter_button)
        
        # Button layout
        button_layout = QHBoxLayout()
        self.start_button = QPushButton("Start")
        self.stop_button = QPushButton("Stop")
        export_button = QPushButton("Export to Excel")
        clear_button = QPushButton("Clear Data")
        
        self.start_button.clicked.connect(self.start_update)
        self.stop_button.clicked.connect(self.stop_update)
        export_button.clicked.connect(self.export_to_excel)
        clear_button.clicked.connect(self.clear_data)
        
        button_layout.addWidget(self.start_button)
        button_layout.addWidget(self.stop_button)
        button_layout.addWidget(export_button)
        button_layout.addWidget(clear_button)
        
        # Combine layouts
        right_widget = QWidget()
        right_layout = QVBoxLayout(right_widget)
        right_layout.addWidget(self.tree)
        right_layout.addLayout(button_layout)
        right_layout.addLayout(filter_layout)
        # Add a clear filters button in the __init__ method:
        clear_filter_button = QPushButton("Clear Filters")
        clear_filter_button.clicked.connect(self.clear_filters)
        filter_layout.addWidget(clear_filter_button)
        
        # Add widgets to splitter
        splitter.addWidget(options_bar)
        splitter.addWidget(right_widget)
        splitter.setStretchFactor(1, 1) 
        
        main_layout.addWidget(splitter)
        
        self.setCentralWidget(main_widget)
        
        # Serial connection
        self.baudrate = 115200
        self.serialConnection = None
        
        # Other variables
        self.threshold_value = None
        self.update_flag = False
        self.timer = QTimer()
        self.timer.timeout.connect(self.update_text)

        self.change_theme(self.theme_combo.currentText())

    def load_settings(self):
        self.move(self.settings.value("pos", self.pos()))
        self.resize(self.settings.value("size", self.size()))

    def closeEvent(self, event):
        self.settings.setValue("pos", self.pos())
        self.settings.setValue("size", self.size())
        self.settings.setValue("com_port", self.com_combo.currentText())
        self.settings.setValue("theme", self.theme_combo.currentText())
        super().closeEvent(event)

    def change_theme(self, theme):
        QApplication.setStyle(QStyleFactory.create(theme))

    def set_threshold(self):
        try:
            self.threshold_value = int(self.threshold_entry.text())
            QMessageBox.information(self, "Success", f"Threshold set to {self.threshold_value}")
        except ValueError:
            QMessageBox.warning(self, "Invalid Input", "Please enter a valid integer value for the threshold.")

    def start_update(self):
        if not self.update_flag:
            try:
                com_port = self.com_combo.currentText()
                self.serialConnection = serial.Serial(com_port, self.baudrate)
                self.update_flag = True
                self.timer.start(500) # CHANGE IF NEEDS TO BE FASTER/SLOWER
                self.start_button.setEnabled(False)
                self.stop_button.setEnabled(True)
            except serial.SerialException as e:
                QMessageBox.critical(self, "Error", f"Failed to open {com_port}: {str(e)}")

    def stop_update(self):
        self.update_flag = False
        self.timer.stop()
        if self.serialConnection:
            self.serialConnection.close()
            self.serialConnection = None
        self.start_button.setEnabled(True)
        self.stop_button.setEnabled(False)

    def update_text(self):
        if self.update_flag and self.serialConnection:
            try:
                data = self.serialConnection.readline()
                if data == b"EOF":
                    self.stop_update()
                else:
                    value = data.decode('utf-8').strip()
                    timestamp = time.strftime("%Y-%m-%d %H:%M:%S", time.localtime())
                    try:
                        numeric_value = float(value.split(':')[-1].strip())
                        item = QTreeWidgetItem([timestamp, value])
                        item.setData(2, Qt.UserRole, numeric_value)
                        if self.threshold_value is not None and numeric_value > self.threshold_value:
                            item.setBackground(0, QColor(255, 255, 0, 100))
                            item.setBackground(1, QColor(255, 255, 0, 100))
                        self.tree.addTopLevelItem(item)
                        self.tree.scrollToBottom()
                    except ValueError:
                        self.tree.addTopLevelItem(QTreeWidgetItem([timestamp, value]))
            except serial.SerialException as e:
                self.stop_update()
                QMessageBox.critical(self, "Error", f"Serial communication error: {str(e)}")

    def export_to_excel(self):
        if self.tree.topLevelItemCount() == 0:
            QMessageBox.warning(self, "No Data", "There is no data to export.")
            return

        try:
            data = []
            for i in range(self.tree.topLevelItemCount()):
                item = self.tree.topLevelItem(i)
                data.append([item.text(0), item.text(1)])

            filename, _ = QFileDialog.getSaveFileName(self, "Save Excel File", "", "Excel Files (*.xlsx)")
            if not filename: # If the file dialog is cancelled
                return 

            workbook = openpyxl.Workbook()
            sheet = workbook.active
            sheet.title = "Serial Data"

            sheet['A1'] = "Timestamp"
            sheet['B1'] = "Data"

            # Write data and apply conditional formatting
            yellow_fill = PatternFill(start_color="FFFF00", end_color="FFFF00", fill_type="solid")
            for row, (timestamp, value) in enumerate(data, start=2):
                sheet.cell(row=row, column=1, value=timestamp)
                sheet.cell(row=row, column=2, value=value)
                
                try:
                    numeric_value = int(value.split(':')[-1].strip())
                    if self.threshold_value is not None and numeric_value > self.threshold_value:
                        sheet.cell(row=row, column=1).fill = yellow_fill
                        sheet.cell(row=row, column=2).fill = yellow_fill
                except ValueError:
                    pass 

            # Auto-adjust column widths
            for column in sheet.columns:
                max_length = 0
                column_letter = openpyxl.utils.get_column_letter(column[0].column)
                for cell in column:
                    try:
                        if len(str(cell.value)) > max_length:
                            max_length = len(cell.value)
                    except:
                        pass
                adjusted_width = (max_length + 2)
                sheet.column_dimensions[column_letter].width = adjusted_width

            workbook.save(filename)
            QMessageBox.information(self, "Success", f"Data has been successfully exported to {filename}")
        except Exception as e:
            QMessageBox.critical(self, "Error", f"Failed to export data: {str(e)}")
            
    def clear_data(self):
        self.tree.clear()
        
    def apply_filter(self):
        start_time = self.start_time.dateTime().toString("yyyy-MM-dd HH:mm:ss")
        end_time = self.end_time.dateTime().toString("yyyy-MM-dd HH:mm:ss")
    
        try:
            min_value = float(self.min_value.text()) if self.min_value.text() else float('-inf')
            max_value = float(self.max_value.text()) if self.max_value.text() else float('inf')
        except ValueError:
            QMessageBox.warning(self, "Invalid Input", "Please enter valid numeric values for min and max.")
            return

        for i in range(self.tree.topLevelItemCount()):
            item = self.tree.topLevelItem(i)
            timestamp = item.text(0)
            value_str = item.text(1).split(':')[-1].strip()
            
            try:
                value = float(value_str)
            except ValueError:
                # If value can't be converted to float, hide the item
                item.setHidden(True)
                continue

            # Check if item is within the specified ranges
            if start_time <= timestamp <= end_time and min_value <= value <= max_value:
                item.setHidden(False)
            else:
                item.setHidden(True)

    # Add a method to clear filters
    def clear_filters(self):
        for i in range(self.tree.topLevelItemCount()):
            self.tree.topLevelItem(i).setHidden(False)
        self.start_time.setDateTime(QDateTime.currentDateTime().addDays(-1))
        self.end_time.setDateTime(QDateTime.currentDateTime())
        self.min_value.clear()
        self.max_value.clear()

if __name__ == "__main__":
    app = QApplication(sys.argv)
    window = MainWindow()
    window.show()
    sys.exit(app.exec_())