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
                             QStyleFactory, QComboBox, QFileDialog, QDateTimeEdit,
                             QStatusBar, QListWidget)
from PyQt5.QtCore import QTimer, Qt, QSettings, QDateTime
from PyQt5.QtGui import QFont, QColor
import pyqtgraph as pg

class MainWindow(QMainWindow):
    def __init__(self):
        super().__init__()
        self.setWindowTitle("Serial Data Logger")
        self.setGeometry(100, 100, 1200, 800)

        # Settings
        self.settings = QSettings("YourCompany", "SerialDataLogger")
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
        self.theme_combo.addItems(["Light", "Dark"])
        self.theme_combo.setCurrentText(self.settings.value("theme", "Light"))
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
        filter_layout = QHBoxLayout()
        self.filter_checkbox = QCheckBox("Show only above threshold")
        self.filter_checkbox.stateChanged.connect(self.apply_filter)
        filter_layout.addWidget(self.filter_checkbox)
        
        # Button layout
        button_layout = QHBoxLayout()
        self.start_button = QPushButton("Start")
        self.resume_button = QPushButton("Resume")
        self.stop_button = QPushButton("Stop")
        export_button = QPushButton("Export to Excel")
        clear_button = QPushButton("Clear Data")
        
        self.start_button.clicked.connect(self.start_update)
        self.resume_button.clicked.connect(self.resume_update)
        self.stop_button.clicked.connect(self.stop_update)
        export_button.clicked.connect(self.export_to_excel)
        clear_button.clicked.connect(self.clear_data)
        
        button_layout.addWidget(self.start_button)
        button_layout.addWidget(self.resume_button)
        button_layout.addWidget(self.stop_button)
        button_layout.addWidget(export_button)
        button_layout.addWidget(clear_button)
        
        # Combine layouts
        right_widget = QWidget()
        right_layout = QVBoxLayout(right_widget)
        right_layout.addWidget(self.tree)
        right_layout.addLayout(button_layout)
        right_layout.addLayout(filter_layout)
        
        # Add a plot widget
        self.plot_widget = pg.PlotWidget()
        self.plot_widget.setBackground('w')
        self.plot_widget.setTitle("Real-time Voltage Plot")
        self.plot_widget.setLabel('left', "Voltage")
        self.plot_widget.setLabel('bottom', "Time", units='s')
        self.plot_data = {}
        self.plot_widget.addLegend()
    
        # Mux selection
        mux_label = QLabel("Select Mux")
        mux_label.setFont(QFont("Arial", 12, QFont.Bold))
        self.mux_combo = QComboBox()
        self.mux_combo.addItems([f"Mux {i}" for i in range(1, 9)])  # Assuming 8 muxes
        self.mux_combo.currentIndexChanged.connect(self.update_plot)
        options_layout.addWidget(mux_label)
        options_layout.addWidget(self.mux_combo)

        # Channel selection
        channel_label = QLabel("Select Channels")
        channel_label.setFont(QFont("Arial", 12, QFont.Bold))
        self.channel_list = QListWidget()
        self.channel_list.setSelectionMode(QListWidget.MultiSelection)
        for i in range(1, 33):
            self.channel_list.addItem(f"Channel {i}")
        self.channel_list.itemSelectionChanged.connect(self.update_plot)
        options_layout.addWidget(channel_label)
        options_layout.addWidget(self.channel_list)

        # Enable zooming and panning
        self.plot_widget.setMouseEnabled(x=True, y=True)
        self.plot_widget.enableAutoRange()

        # Add the plot widget to the right layout
        right_layout.addWidget(self.plot_widget)

        # Add widgets to splitter
        splitter.addWidget(options_bar)
        splitter.addWidget(right_widget)
        splitter.setStretchFactor(1, 1) 
        
        main_layout.addWidget(splitter)
        
        self.setCentralWidget(main_widget)
        
        # Status bar
        self.statusBar = QStatusBar()
        self.setStatusBar(self.statusBar)
        
        # Serial connection
        self.baudrate = 115200
        self.serialConnection = None
        
        # Other variables
        self.threshold_value = None
        self.update_flag = False
        self.timer = QTimer()
        self.timer.timeout.connect(self.update_text)
        self.start_time = None

        self.change_theme(self.theme_combo.currentText())
        
        self.plot_data = {}
        self.current_mux = 1

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
        if theme == "Dark":
            self.setStyleSheet("""
                QWidget { background-color: #2D2D2D; color: #FFFFFF; }
                QTreeWidget { alternate-background-color: #3D3D3D; }
                QPushButton { background-color: #4D4D4D; border: 1px solid #5D5D5D; }
                QLineEdit, QComboBox { background-color: #3D3D3D; border: 1px solid #5D5D5D; }
            """)
            self.plot_widget.setBackground('k')
        else:
            self.setStyleSheet("")
            self.plot_widget.setBackground('w')

    def set_threshold(self):
        try:
            self.threshold_value = int(self.threshold_entry.text())
            QMessageBox.information(self, "Success", f"Threshold set to {self.threshold_value}")
            self.apply_filter()
            self.statusBar.showMessage(f"Threshold set to {self.threshold_value}")
        except ValueError:
            QMessageBox.warning(self, "Invalid Input", "Please enter a valid integer value for the threshold.")

    def start_update(self):
        if not self.update_flag:
            try:
                com_port = self.com_combo.currentText()
                self.serialConnection = serial.Serial(com_port, self.baudrate)
                self.serialConnection.flush()
                self.serialConnection.write("START\r".encode())
                self.update_flag = True
                self.timer.start(50)
                self.start_button.setEnabled(False)
                self.resume_button.setEnabled(False)
                self.stop_button.setEnabled(True)
                self.serialConnection.reset_input_buffer()
                self.start_time = time.time()
                self.statusBar.showMessage(f"Connected to {com_port} at {self.baudrate} baud")
            except serial.SerialException as e:
                QMessageBox.critical(self, "Error", f"Failed to open {com_port}: {str(e)}")

    def resume_update(self):
        if not self.update_flag:
            try:
                com_port = self.com_combo.currentText()
                self.serialConnection = serial.Serial(com_port, self.baudrate)
                self.serialConnection.flush()
                self.serialConnection.write("RESUME\r".encode())
                self.update_flag = True
                self.timer.start(50)
                self.start_button.setEnabled(False)
                self.resume_button.setEnabled(False)
                self.stop_button.setEnabled(True)
                self.serialConnection.reset_input_buffer()
                self.statusBar.showMessage(f"Resumed connection to {com_port}")
            except serial.SerialException as e:
                QMessageBox.critical(self, "Error", f"Failed to open {com_port}: {str(e)}")

    def stop_update(self):
        self.update_flag = False
        self.timer.stop()
        if self.serialConnection:
            self.serialConnection.flush()
            self.serialConnection.write("PAUSE\r".encode())
            try:
                self.serialConnection.timeout = 0.1
                while True:
                    response = self.serialConnection.readline().strip()
                    if 'Pause confirmed' in response.decode():
                        break
                    elif 'Mux:' in response.decode() and 'Channel:' in response.decode() and 'Temperature:' in response.decode() and 'Voltage:' in response.decode():
                        self.process_data(response.decode())
            except Exception as e:
                print(f"Error receiving confirmation: {e}")
            self.serialConnection.close()
            self.serialConnection = None
        self.start_button.setEnabled(True)
        self.resume_button.setEnabled(True)
        self.stop_button.setEnabled(False)
        self.statusBar.showMessage("Connection paused")

    def update_text(self):
        if self.update_flag and self.serialConnection:
            try:
                data = self.serialConnection.readline()
                if data == b"EOF":
                    self.stop_update()
                else:
                    value = data.decode('utf-8').strip()
                    self.process_data(value)
            except serial.SerialException as e:
                self.stop_update()
                QMessageBox.critical(self, "Error", f"Serial communication error: {str(e)}")

    def process_data(self, value):
        timestamp = time.strftime("%Y-%m-%d %H:%M:%S", time.localtime())
        try:
            if "Mux:" in value and "Channel:" in value and "Temperature:" in value and "Voltage:" in value:
                parts = value.split()
                mux = int(parts[1])
                channel = int(parts[3])
                temperature = float(parts[5])
                voltage = float(parts[7])
                
                item = QTreeWidgetItem([
                    timestamp,
                    str(mux),
                    str(channel),
                    f"{temperature:.2f}°C",
                    f"{voltage:.4f}"
                ])
                
                if self.threshold_value is not None and voltage * 65535 / 3.3 > self.threshold_value:
                    item.setBackground(1, QColor(255, 255, 0, 100))
                
                self.tree.addTopLevelItem(item)
                self.tree.scrollToBottom()
                self.apply_filter()

                # Store data for plotting
                if mux not in self.plot_data:
                    self.plot_data[mux] = {}
                if channel not in self.plot_data[mux]:
                    self.plot_data[mux][channel] = {'x': [], 'y': []}
                
                current_time = time.time() - self.start_time
                self.plot_data[mux][channel]['x'].append(current_time)
                self.plot_data[mux][channel]['y'].append(voltage)

                self.update_plot()

        except ValueError:
            self.tree.addTopLevelItem(QTreeWidgetItem([timestamp, value, "", "", ""]))
            self.apply_filter()

    def update_plot(self):
        self.plot_widget.clear()
        selected_mux = int(self.mux_combo.currentText().split()[1])
        selected_channels = [int(item.text().split()[1]) for item in self.channel_list.selectedItems()]

        if selected_mux in self.plot_data:
            for channel in self.plot_data[selected_mux]:
                if not selected_channels or channel in selected_channels:
                    self.plot_widget.plot(
                        self.plot_data[selected_mux][channel]['x'],
                        self.plot_data[selected_mux][channel]['y'],
                        pen=(channel * 20) % 256,
                        name=f'Channel {channel}'
                    )

        self.plot_widget.addLegend()

    def export_to_excel(self):
        data = []
        for i in range(self.tree.topLevelItemCount()):
            item = self.tree.topLevelItem(i)
            data.append([item.text(j) for j in range(5)])
        
        if not data:
            QMessageBox.warning(self, "No Data", "There is no data to export.")
            return

        try:
            df = pd.DataFrame(data, columns=["Timestamp", "Mux", "Channel", "Temperature", "Voltage"])

            filename, _ = QFileDialog.getSaveFileName(self, "Save Excel File", "", "Excel Files (*.xlsx)")
            if not filename:
                return 

            df.to_excel(filename, index=False)

            QMessageBox.information(self, "Success", f"Data has been successfully exported to {filename}")
        except Exception as e:
            QMessageBox.critical(self, "Error", f"Failed to export data: {str(e)}")
            
    def clear_data(self):
        self.tree.clear()
        self.filter_checkbox.setChecked(False)
        self.plot_widget.clear()
        self.plot_data = {}
        self.start_time = None
        
    def apply_filter(self):
        for i in range(self.tree.topLevelItemCount()):
            item = self.tree.topLevelItem(i)
            if self.filter_checkbox.isChecked():
                try:
                    voltage = float(item.text(4))
                    item.setHidden(not (self.threshold_value is not None and voltage * 65535 / 3.3 > self.threshold_value))
                except ValueError:
                    item.setHidden(True)
            else:
                item.setHidden(False)

if __name__ == "__main__":
    app = QApplication(sys.argv)
    window = MainWindow()
    window.show()
    sys.exit(app.exec_())