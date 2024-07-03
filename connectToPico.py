import serial

def read_from_serial(port, baudrate, output_path):
    # Attempt to open the serial connection
    try:
        serial_connection = serial.Serial(port, baudrate)
    except serial.SerialException as e:
        print(f"Failed to connect on {port}: {str(e)}")
        return

    # Open the destination file in append mode
    with open(output_path, "ab") as destination_file:
        try:
            while True:
                data = serial_connection.readline().decode('utf-8').strip()
                if data == "EOF":
                    break
                print(data)  # For debugging, to see the data in the terminal
                destination_file.write((data + "\n").encode('utf-8'))
                destination_file.flush()  # Flush after each write to ensure real-time updates
        except KeyboardInterrupt:
            print("Interrupted by the user")
        finally:
            serial_connection.close()
            print("File written and serial connection closed.")

# Configure the serial connection
port = "COM9"
baudrate = 115200
output_path = "C:/Users/Chenyang.Hu/OneDrive - Gentex Corporation/Desktop/analog_data.txt"

# Call the function to start the process
read_from_serial(port, baudrate, output_path)