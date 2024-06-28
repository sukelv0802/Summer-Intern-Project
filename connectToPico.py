import serial

# To configure the serial connection
baudrate = 115200
serialConnection = serial.Serial("COM10", baudrate)

destinationFile = open("/Users/Kelvin.Su/connect/testFile.txt", "w")

try:
    while True:
        data = serialConnection.read(128)
        if data == b"EOF":
            break
        print(data)
        destinationFile.write(data.decode('utf-8')) 
        destinationFile.flush()  
        
finally:
    destinationFile.close()
    serialConnection.close()