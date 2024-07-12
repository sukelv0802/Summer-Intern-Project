import sys
import serial
import serial.tools.list_ports
import pandas as pd
import time
from PyQt5 import QtWidgets, QtCore, QtGui
from PyQt5.QtWidgets import QMainWindow, QApplication, QLabel, QPushButton, QVBoxLayout, QHBoxLayout, QComboBox, QLineEdit, QTableWidget, QTableWidgetItem, QTabWidget, QPlainTextEdit, QFileDialog, QHeaderView
from PyQt5.QtCore import QThread, pyqtSignal
import matplotlib.pyplot as plt
from matplotlib.backends.backend_qt5agg import FigureCanvasQTAgg as FigureCanvas

class SerialThread(QThread):
    data_received = pyqtSignal(str)
    
    def __init__(self, port, threshold):
        super().__init__()
        self.port = port
        self.threshold = threshold
        self.running = True
    
    def run(self):
        try:
            with serial.Serial(self.port, 115200, timeout=1) as ser:
                while self.running:
                    line = ser.readline().decode('utf-8').strip()
                    if line:
                        print(f"Thread received: {line}")
                        self.data_received.emit(line)
        except Exception as e:
            print(f"Serial thread error: {str(e)}")
    
    def stop(self):
        self.running = False

class MainWindow(QMainWindow):
    def __init__(self):
        super().__init__()
        self.initUI()
        self.current_temperature = None
        self.cycle_data = {}
        self.threshold = 0.1
                
    def initUI(self):
        self.setWindowTitle('Solder Joint Monitoring')
        self.setGeometry(100, 100, 1500, 900)
        
        self.main_widget = QtWidgets.QWidget(self)
        self.setCentralWidget(self.main_widget)
        
        self.layout = QVBoxLayout(self.main_widget)
        
        self.top_layout = QHBoxLayout()
        
        self.port_label = QLabel('Select COM Port:')
        self.top_layout.addWidget(self.port_label)
        
        self.combobox = QComboBox()
        self.top_layout.addWidget(self.combobox)
        self.refresh_ports()
        
        self.refresh_button = QPushButton('Refresh')
        self.refresh_button.clicked.connect(self.refresh_ports)
        self.top_layout.addWidget(self.refresh_button)
        
        self.threshold_label = QLabel('Set Threshold:')
        self.top_layout.addWidget(self.threshold_label)
        
        self.threshold_input = QLineEdit('0.1')
        self.top_layout.addWidget(self.threshold_input)
        
        self.set_threshold_button = QPushButton('Set Threshold')
        self.set_threshold_button.clicked.connect(self.set_threshold)
        self.top_layout.addWidget(self.set_threshold_button)
        
        self.start_button = QPushButton('Start')
        self.start_button.clicked.connect(self.start_monitoring)
        self.top_layout.addWidget(self.start_button)
        
        self.stop_button = QPushButton('Stop')
        self.stop_button.setEnabled(False)
        self.stop_button.clicked.connect(self.stop_monitoring)
        self.top_layout.addWidget(self.stop_button)
        
        self.layout.addLayout(self.top_layout)
        
        self.tabs = QTabWidget()
        
        self.data_tab = QtWidgets.QWidget()
        self.graph_tab = QtWidgets.QWidget()
        self.log_tab = QtWidgets.QWidget()
        
        self.tabs.addTab(self.data_tab, 'Data')
        self.tabs.addTab(self.graph_tab, 'Graph')
        self.tabs.addTab(self.log_tab, 'Log')
        
        self.layout.addWidget(self.tabs)
        
        self.data_layout = QVBoxLayout(self.data_tab)
        
        self.table_widget = QTableWidget()
        self.table_widget.setStyleSheet("""
            QTableWidget {
                background-color: white;
                gridline-color: #d0d0d0;
            }
            QHeaderView::section {
                background-color: #f0f0f0;
                padding: 4px;
                border: 1px solid #d0d0d0;
                font-weight: bold;
            }
            QTableWidget::item {
                padding: 4px;
            }
        """)
        self.data_layout.addWidget(self.table_widget)
        
        self.export_button = QPushButton('Export to Excel')
        self.export_button.clicked.connect(self.export_to_excel)
        self.data_layout.addWidget(self.export_button)
        
        self.filter_button = QPushButton('Filter Cracked Joints')
        self.filter_button.clicked.connect(self.filter_data)
        self.data_layout.addWidget(self.filter_button)
        
        self.graph_layout = QVBoxLayout(self.graph_tab)
        
        self.figure, self.ax = plt.subplots()
        self.canvas = FigureCanvas(self.figure)
        self.graph_layout.addWidget(self.canvas)
        
        # Create a dock widget for the log terminal
        self.log_dock_widget = QtWidgets.QDockWidget("Log Terminal", self)
        self.log_dock_widget.setAllowedAreas(QtCore.Qt.BottomDockWidgetArea)

        # Add the log terminal to the dock widget
        self.log_terminal = QPlainTextEdit()
        self.log_terminal.setReadOnly(True)
        self.log_dock_widget.setWidget(self.log_terminal)

        # Add the dock widget to the main window
        self.addDockWidget(QtCore.Qt.BottomDockWidgetArea, self.log_dock_widget)
        
        # Add a "Show Log Terminal" button
        self.show_log_button = QPushButton('Show Log Terminal')
        self.show_log_button.clicked.connect(self.show_log_terminal)
        self.top_layout.addWidget(self.show_log_button)
        
        self.serial_thread = None

    
    def show_log_terminal(self):
        self.log_dock_widget.show()
    
    def set_threshold(self):
        try:
            self.threshold = float(self.threshold_input.text())
            self.log(f'Threshold set to {self.threshold}')
            self.update_table()  
        except ValueError:
            self.log('Invalid threshold value')
        
    def refresh_ports(self):
        self.combobox.clear()
        ports = serial.tools.list_ports.comports()
        for port in ports:
            self.combobox.addItem(port.device)
    
    def start_monitoring(self):
        port = self.combobox.currentText()
        self.threshold = float(self.threshold_input.text())
        if port:
            self.serial_thread = SerialThread(port, self.threshold)
            self.serial_thread.data_received.connect(self.update_data)
            self.serial_thread.start()
            self.start_button.setEnabled(False)
            self.stop_button.setEnabled(True)
            self.log('Monitoring started on port ' + port)

    def stop_monitoring(self):
        if self.serial_thread:
            self.serial_thread.stop()
            self.serial_thread.wait()
            self.serial_thread = None
            self.start_button.setEnabled(True)
            self.stop_button.setEnabled(False)
            self.log('Monitoring stopped')
    
    def update_data(self, line):
        try:
            print(f"Raw data: {line}")
            self.log(f"Raw data received: {line}")

            if "Cycles Number:" in line:
                self.current_cycle = int(line.split(":")[1].strip())
                self.cycle_data[self.current_cycle] = {"Temperature": None}
                for i in range(1, 33):
                    self.cycle_data[self.current_cycle][f"Channel {i}"] = None
            elif "Temperature:" in line:
                temp = float(line.split(":")[1].strip().split("°")[0])
                self.cycle_data[self.current_cycle]["Temperature"] = temp
            elif "Mux" in line and "Channel" in line and "Voltage:" in line:
                parts = line.split()
                channel = int(parts[3])
                voltage = float(parts[5])
                self.cycle_data[self.current_cycle][f"Channel {channel}"] = voltage

            if "------------------------" in line:
                self.update_table()
                self.update_graph()
        except Exception as e:
            print(f"Error processing data: {str(e)}")
            self.log(f"Error processing data: {str(e)}")
    
    def update_table(self):
        cycles = list(self.cycle_data.keys())
        self.table_widget.setColumnCount(len(cycles) + 1)  # +1 for the channel labels
        self.table_widget.setRowCount(33)  # 32 channels + 1 for temperature

        # Set headers
        headers = ["Channel"] + [f"Cycle {cycle}" for cycle in cycles]
        self.table_widget.setHorizontalHeaderLabels(headers)

        # Set row labels
        row_labels = ["Temperature"] + [f"Channel {i}" for i in range(1, 33)]
        self.table_widget.setVerticalHeaderLabels(row_labels)

        # Populate data
        for col, cycle in enumerate(cycles, start=1):
            # Set temperature
            temp_item = QTableWidgetItem(f"{self.cycle_data[cycle]['Temperature']:.2f}")
            temp_item.setTextAlignment(QtCore.Qt.AlignCenter)
            self.table_widget.setItem(0, col, temp_item)

            # Set channel voltages
            for row in range(1, 33):
                channel_data = self.cycle_data[cycle].get(f"Channel {row}")
                if channel_data is not None:
                    item = QTableWidgetItem(f"{channel_data:.3f}")
                    item.setTextAlignment(QtCore.Qt.AlignCenter)
                    if channel_data < self.threshold:
                        item.setBackground(QtGui.QColor(255, 200, 200))
                    self.table_widget.setItem(row, col, item)

        # Adjust column widths
        self.table_widget.resizeColumnsToContents()

        # Freeze the first column
        self.table_widget.setColumnWidth(0, 100)
        self.table_widget.horizontalHeader().setSectionResizeMode(0, QHeaderView.Fixed)

        # Allow horizontal scrolling for cycles
        self.table_widget.horizontalHeader().setSectionResizeMode(QHeaderView.Interactive)
        self.table_widget.horizontalHeader().setStretchLastSection(False)

        # Set row height
        for row in range(self.table_widget.rowCount()):
            self.table_widget.setRowHeight(row, 25)

        self.table_widget.scrollToBottom()
    
    def update_graph(self):
        self.ax.clear()
        cycles = list(self.cycle_data.keys())
        temperatures = [self.cycle_data[cycle]['Temperature'] for cycle in cycles]
        self.ax.plot(cycles, temperatures, label='Temperature')
        
        for channel in range(1, 33):
            voltages = [self.cycle_data[cycle].get(f"Channel {channel}", 0) for cycle in cycles]
            self.ax.plot(cycles, voltages, label=f'Channel {channel}')
        
        self.ax.legend()
        self.ax.set_xlabel('Cycle')
        self.ax.set_ylabel('Temperature (°C) / Voltage (V)')
        self.canvas.draw()
    
    def export_to_excel(self):
        file_name, _ = QFileDialog.getSaveFileName(self, "Save Excel File", "", "Excel Files (*.xlsx)")
        if file_name:
            df = pd.DataFrame(self.data, columns=['Mux', 'Channel', 'Voltage', 'Temperature'])
            df.to_excel(file_name, index=False)
            self.log('Data exported to ' + file_name)
    
    def filter_data(self):
        self.table_widget.setRowCount(0)
        filtered_data = [d for d in self.data if d[2] < self.threshold]
        self.table_widget.setRowCount(len(filtered_data))
        for row, (mux, channel, voltage, temperature) in enumerate(filtered_data):
            self.table_widget.setItem(row, 0, QTableWidgetItem(str(mux)))
            self.table_widget.setItem(row, 1, QTableWidgetItem(str(channel)))
            self.table_widget.setItem(row, 2, QTableWidgetItem(str(voltage)))
            self.table_widget.setItem(row, 3, QTableWidgetItem(str(temperature)))
            for col in range(4):
                self.table_widget.item(row, col).setBackground(QtGui.QColor(255, 0, 0))
        self.log('Filtered data to show only cracked solder joints')
    
    def log(self, message):
        self.log_terminal.appendPlainText(message)

if __name__ == '__main__':
    app = QApplication(sys.argv)
    main_win = MainWindow()
    main_win.show()
    sys.exit(app.exec_())
