import sys
from PyQt5.QtWidgets import (QApplication, QMainWindow, QLineEdit, QPushButton, 
                             QVBoxLayout, QHBoxLayout, QWidget, QTextEdit, QLabel, 
                             QSpacerItem, QSizePolicy, QFileDialog)
from threading import Thread
import serial
from PyQt5.QtCore import pyqtSignal, QObject, QThread, Qt
from PyQt5.QtGui import QColor
import pandas as pd
import time

class SerialWorker(QObject):
    data_received = pyqtSignal(str)

    def __init__(self, port, baudrate, output_path):
        super().__init__()
        self.port = port
        self.baudrate = baudrate
        self.output_path = output_path
        self.running = False
        self.exit = False
        self.serial_connection = None

    def work(self):
        try:
            self.serial_connection = serial.Serial(self.port, self.baudrate, timeout=0.1)
            # self.running = True
            # with open(self.output_path, "a") as destination_file:
            while not self.exit:
                if self.running:
                    if self.serial_connection.in_waiting > 0:
                        data = self.serial_connection.readline().decode('utf-8').strip()
                        if data:
                            self.data_received.emit(data)  # Emit signal with the data
                else:
                    time.sleep(0.1)  # Add a small delay to release CPU when paused
        except serial.SerialException as e:
            print(f"Error opening serial port {self.port}: {e}")
        finally:
            if self.serial_connection and self.serial_connection.is_open:
                self.serial_connection.close()

    def stop(self):
        self.running = False
        if hasattr(self, 'serial_connection') and self.serial_connection.is_open:
            self.serial_connection.reset_input_buffer()  # Clear the input buffer
    
    def exit_work(self):
        self.exit = True

class Window(QMainWindow):
    def __init__(self):
        super().__init__()
        self.setGeometry(300, 300, 800, 600)
        self.setWindowTitle("PyQt5Test")
        self.init_ui()

        # Initialize threshold value
        self.threshold_value = 70000

        self.serial_worker = None
        self.serial_thread = None

    def init_ui(self):
        # Create layouts
        left_layout = QVBoxLayout()
        right_layout = QVBoxLayout()
        main_layout = QHBoxLayout()

        # Create textbox
        threshold_label = QLabel('Threshold Value:')
        self.threshold_input = QLineEdit()

        # Create buttons
        set_threshold_button = QPushButton('Set Threshold Value')
        start_button = QPushButton('Start Reading')
        stop_button = QPushButton('Stop Reading')
        clear_button = QPushButton('Clear')
        # Create a blank space
        vertical_spacer = QSpacerItem(20, 40, QSizePolicy.Minimum, QSizePolicy.Expanding)
        export_button = QPushButton('Export to Excel')

        # Create text area for data display
        self.data_display = QTextEdit()

        # Connect button signals to slots
        # self.threshold_input.textChanged.connect(self.read_threshold_value)
        set_threshold_button.clicked.connect(self.read_threshold_value)
        start_button.clicked.connect(self.start_reading)
        stop_button.clicked.connect(self.stop_reading)
        clear_button.clicked.connect(self.clear)
        export_button.clicked.connect(self.export_to_excel)

        # Add buttons to left layout
        left_layout.addWidget(threshold_label)
        left_layout.addWidget(self.threshold_input)
        left_layout.addWidget(set_threshold_button)
        left_layout.addWidget(start_button)
        left_layout.addWidget(stop_button)
        left_layout.addWidget(clear_button)

        # Add spacer to 'clear' and 'export'
        left_layout.addSpacerItem(vertical_spacer)
        left_layout.addWidget(export_button)

        left_layout.addStretch()

        # Add text area to right layout
        right_layout.addWidget(self.data_display)

        # Set main layout
        main_layout.addLayout(left_layout, 1)  # The second argument is stretch factor
        main_layout.addLayout(right_layout, 3)  # The second argument is stretch factor

        # Set central widget
        central_widget = QWidget()
        central_widget.setLayout(main_layout)
        self.setCentralWidget(central_widget)

    # Read the threshold value from the threshold textbox
    def read_threshold_value(self):
        threshold_text = self.threshold_input.text()
        try:
            self.threshold_value = int(threshold_text)
        except:
            self.threshold_value = 70000
        print('Threshold Value Set!')

    def start_reading(self):
        if self.serial_thread is None:  # Use isRunning() for QThread
            self.serial_worker = SerialWorker("COM9", 115200, "C:/Users/Chenyang.Hu/OneDrive - Gentex Corporation/Desktop/analog_data.txt")
            self.serial_thread = QThread()
            self.serial_worker.moveToThread(self.serial_thread)
            self.serial_worker.data_received.connect(self.update_data)
            self.serial_thread.started.connect(self.serial_worker.work)
            self.serial_thread.finished.connect(self.serial_thread.deleteLater)  # Ensure proper thread cleanup
            self.serial_thread.start()
        else:
            if self.serial_worker.serial_connection and self.serial_worker.serial_connection.is_open:
                self.serial_worker.serial_connection.reset_input_buffer()  # Flush the buffer before starting to read again
        self.serial_worker.running = True

    def stop_reading(self):
        if self.serial_worker:
            self.serial_worker.running = False

    def clear(self):
        self.data_display.clear()

    def export_to_excel(self):
        # Get the text from the QTextEdit widget
        text = self.data_display.toPlainText()
        # Split the text by newlines to get a list of data
        data = text.split('\n')
        # Convert the list to a pandas DataFrame
        df = pd.DataFrame(data, columns=['Data'])
        # Open a QFileDialog to ask the user for a file path and name
        options = QFileDialog.Options()
        options |= QFileDialog.DontUseNativeDialog
        filename, _ = QFileDialog.getSaveFileName(self, "Save to Excel", "", "Excel Files (*.xlsx)", options=options)

        # If a file path is provided, export the DataFrame to the specified Excel file
        if filename:
            if not filename.endswith('.xlsx'):
                filename += '.xlsx'
            df.to_excel(filename, index=False)

    def update_data(self, data):
        data_value = int(data)
        if data_value > self.threshold_value:
            # Highlight the data to red if data above threshold
            self.data_display.setTextColor(QColor(Qt.red))
            self.data_display.append(data)
            # Reset color to black for next data
            self.data_display.setTextColor(QColor(Qt.black))
        else:
            self.data_display.append(data)

    def closeEvent(self, event):
        # Signal the worker to stop running
        if self.serial_worker:
            self.serial_worker.exit_work()  # Tell the worker to finish its loop
            self.serial_worker.running = False
        if self.serial_thread:
            self.serial_thread.quit()
            self.serial_thread.wait(1000)
        event.accept()
        print('Testing Terminates')

if __name__ == '__main__':
    app = QApplication(sys.argv)
    main_window = Window()
    main_window.show()
    sys.exit(app.exec_())