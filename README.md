Summer Intern Project led by Jim Dratz

Team Members:
1. Kelvin Su
2. Chenyang Hu

----------------------Introduction----------------------
The python script main.py is designed for use with the Raspberry Pi Pico microcontroller to manage data acquisition from multiple sensors through a series of multiplexers. It utilizes UART for serial communication with a host computer, I2C for interfacing with an MCP23017 I/O expander that controls 8 ADG732 multiplexers, and ADC for reading sensor data. The script supports dynamic control via stdin and stdout, allowing operations like pausing and resuming data collection.

----------------------Prerequisites----------------------
- Raspberry Pi Pico
- MCP23017 I/O Expander
- ADG732 Multiplexers (up to 8, each with 32 channels)
- Analog sensors compatible with the multiplexer setup
- Necessary wires and resistors for connections
- Breadboard or similar setup for prototyping
- IDE for developing Python code to Raspberry Pi Pico
- (Optional) Serial terminal software (e.g. Putty, Tera Term)

----------------------Installation----------------------
- Connect each multiplexer's control pins to the specified GPIO pins on the Pico as per the script settings.    (e.g. Connect the MCP23017 to the Raspberry Pi Pico via I2C, SDA to Pin2, SCL to Pin3)
- The Raspberry Pi Pico is used for prividing power for ADG732 multiplexers, and the constant voltage source is  used for providing power for testing PCBs.
- CS (Chip Selecing) lines use 8 separate lines, while EN (Enable) line and WR (Write) line are only 1 for each.
