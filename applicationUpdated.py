import sys
import time
import serial
import pandas as pd
import openpyxl
from openpyxl.styles import PatternFill
from openpyxl.chart import LineChart, Reference
from PyQt5.QtWidgets import (QApplication, QMainWindow, QTreeWidget, QTreeWidgetItem, 
                             QPushButton, QVBoxLayout, QHBoxLayout, QWidget, QLabel, 
                             QLineEdit, QMessageBox, QHeaderView, QSplitter, QCheckBox,
                             QStyleFactory, QComboBox, QFileDialog, QDateTimeEdit)
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
        self.tree.setHeaderLabels(["Timestamp", "Mux", "Channel", "Temperature", "Voltage"])
        self.tree.header().setSectionResizeMode(QHeaderView.Interactive)
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
        
        
        self.filter_checkbox = QCheckBox("Show only above threshold")
        self.filter_checkbox.stateChanged.connect(self.apply_filter)
        options_layout.addWidget(self.filter_checkbox)
        
        # Combine layouts
        right_widget = QWidget()
        right_layout = QVBoxLayout(right_widget)
        right_layout.addWidget(self.tree)
        right_layout.addLayout(button_layout)
        right_layout.addLayout(filter_layout)
        
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
        self.last_temp = None
        self.last_pot = None

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
            self.apply_filter()  # Reapply filter after changing threshold
        except ValueError:
            QMessageBox.warning(self, "Invalid Input", "Please enter a valid integer value for the threshold.")

    def start_update(self):
        if not self.update_flag:
            try:
                com_port = self.com_combo.currentText()
                self.serialConnection = serial.Serial(com_port, self.baudrate)
                print(f"Connected to {com_port} at {self.baudrate} baud.") # Debug message                
                self.serialConnection.flush()
                self.serialConnection.write("RESUME\r".encode())
                self.update_flag = True
                self.timer.start(50) # CHANGE IF NEEDS TO BE FASTER/SLOWER, no faster than freq in main
                self.start_button.setEnabled(False)
                self.stop_button.setEnabled(True)
                self.serialConnection.reset_input_buffer()
            except serial.SerialException as e:
                QMessageBox.critical(self, "Error", f"Failed to open {com_port}: {str(e)}")
                print(f"Failed to open {com_port}: {str:e}") # Debug message

    def stop_update(self):
        self.update_flag = False
        self.timer.stop()
        if self.serialConnection:
            # Send PAUSE command
            self.serialConnection.flush()
            self.serialConnection.write("PAUSE\r".encode())
            print("Pause command sent")
            try:
                # Wait for a response with a timeout
                self.serialConnection.timeout = 0.1  # Set a timeout for reading (e.g., 2 seconds)
                response = ''
                while True:
                    response = self.serialConnection.readline().strip()
                    print(f"Response received: '{response.decode()}")
                    if 'Pause confirmed' in response.decode():
                        # print(f"Response received: '{response.decode()}")
                        break
                    elif 'Mux:' in response.decode() and 'Channel:' in response.decode() and 'Temperature:' in response.decode() and 'Voltage:' in response.decode():
                        parts = response.decode().split()
                        mux = parts[1]
                        channel = parts[3]
                        temperature = parts[5]
                        voltage = parts[7]
                        timestamp = time.strftime("%Y-%m-%d %H:%M:%S", time.localtime())
                        item = QTreeWidgetItem([
                            timestamp,
                            mux,
                            channel,
                            f"{temperature}°C",
                            f"{voltage}"
                        ])
                        
                        if self.threshold_value is not None and float(voltage) * 65535 / 3.3 < self.threshold_value:
                            item.setBackground(1, QColor(255, 255, 0, 100))
                            
                        self.tree.addTopLevelItem(item)
                        self.tree.scrollToBottom()
                        self.apply_filter()  # Apply filter after adding new item                        
            except Exception as e:
                print(f"Error receiving confirmation: {e}")
            self.serialConnection.close()
            self.serialConnection = None
        self.start_button.setEnabled(True)
        self.stop_button.setEnabled(False)

    def update_text(self):
        if self.update_flag and self.serialConnection:
            try:
                data = self.serialConnection.readline()
                print(f"Data received: {data}") # Debug message
                if data == b"EOF":
                    self.stop_update()
                else:
                    value = data.decode('utf-8').strip()
                    print(f"Decoded value: {value}") # Debug message
                    timestamp = time.strftime("%Y-%m-%d %H:%M:%S", time.localtime())
                    try:
                        if "Mux:" in value and "Channel:" in value and "Temperature:" in value and "Voltage:" in value:
                            parts = value.split()
                            mux = parts[1]
                            channel = parts[3]
                            temperature = parts[5]
                            voltage = parts[7]
                            
                            item = QTreeWidgetItem([
                                timestamp,
                                mux,
                                channel,
                                f"{temperature}°C",
                                f"{voltage}"
                            ])
                            
                            if self.threshold_value is not None and float(voltage) * 65535 / 3.3 < self.threshold_value:
                                item.setBackground(1, QColor(255, 255, 0, 100))
                            
                            self.tree.addTopLevelItem(item)
                            self.tree.scrollToBottom()
                            self.apply_filter()  # Apply filter after adding new item
                    except ValueError:
                        self.tree.addTopLevelItem(QTreeWidgetItem([timestamp, value, "", "", ""]))
                        self.apply_filter()  # Apply filter even for invalid data
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
                data.append([item.text(0), item.text(1), item.text(2)])

            filename, _ = QFileDialog.getSaveFileName(self, "Save Excel File", "", "Excel Files (*.xlsx)")
            if not filename: # If the file dialog is cancelled
                return 

            workbook = openpyxl.Workbook()
            sheet = workbook.active
            sheet.title = "Serial Data"
            sheet['A1'] = "Timestamp"
            sheet['B1'] = "Potentiometer"
            sheet['C1'] = "Temperature"

            # Write data and apply conditional formatting
            yellow_fill = PatternFill(start_color="FFFF00", end_color="FFFF00", fill_type="solid")
            for row, (timestamp, pot_value, temp_value) in enumerate(data, start=2):
                sheet.cell(row=row, column=1, value=timestamp)
                sheet.cell(row=row, column=2, value=pot_value)
                sheet.cell(row=row, column=3, value=temp_value)
                
                try:
                    numeric_value = float(pot_value)
                    if self.threshold_value is not None and numeric_value > self.threshold_value:
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
        self.filter_checkbox.setChecked(False)
        
    def apply_filter(self):
        for i in range(self.tree.topLevelItemCount()):
            item = self.tree.topLevelItem(i)
            if self.filter_checkbox.isChecked():
                try:
                    pot_value = float(item.text(1))
                    item.setHidden(not (self.threshold_value is not None and pot_value > self.threshold_value))
                except ValueError:
                    item.setHidden(True)
            else:
                item.setHidden(False)

if __name__ == "__main__":
    app = QApplication(sys.argv)
    window = MainWindow()
    window.show()
    sys.exit(app.exec_())