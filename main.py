import machine
from machine import I2C, Pin, ADC
import time
import uos
import select
import sys

# need this UART to send data to local computer
uart = machine.UART(0, baudrate=115200)
uart.init(115200, bits=8, parity=None, stop=1, tx=Pin(0), rx=Pin(1))
uos.dupterm(uart)

# Set up the poll object
poll_obj = select.poll()
poll_obj.register(sys.stdin, select.POLLIN)

# I2C setup
i2c = I2C(1, scl=Pin(3), sda=Pin(2))

# MCP23017 Constants
MCP23017_ADDR = 0x20
IODIRA = 0x00
GPIOA = 0x12

# Mux constants
mux_num = 8
mux_channels = 32

# ADC setup
adc = ADC(Pin(27))  # Assuming GPIO 27 is ADC capable

mux_en_pins = [Pin(i, Pin.OUT) for i in range(8, 8 + mux_num)] # These are connected to each CS pin on muxes
wr_pin = Pin(20, Pin.OUT) # This is connected to the WR pins
en_pin = Pin(21, Pin.OUT) # This is connected to the EN pins
gnd_pin = Pin(22, Pin.OUT)
temp_pin = machine.ADC(4)

# Configure GPIOA pins as outputs
def setup_mcp23017():
    i2c.writeto_mem(MCP23017_ADDR, IODIRA, b'\x00')
    for pins in mux_en_pins:
        pins.value(1)
    wr_pin.value(1)
    en_pin.value(1)
    gnd_pin.value(0)

# Enable mux
def enable_mux(mux_index):
    mux_en_pins[mux_index].value(0)

# Disable mux
def disable_all_muxes():
    for pin in mux_en_pins:
        pin.value(1)

# Reset mux
def reset_mux():
    i2c.writeto_mem(MCP23017_ADDR, GPIOA, b'\x00')

# Select the multiplexer channel
def select_channel(channel):
    if 0 <= channel <= 31:
        wr_pin.value(0)
        i2c.writeto_mem(MCP23017_ADDR, GPIOA, bytes([channel]))
        wr_pin.value(1)
    # else:
    #     print("Invalid Channel")

def discharge_input():
    gnd_pin.value(1)
    time.sleep(0.05)
    gnd_pin.value(0)

def read_adc():
    return adc.read_u16()

# Read the temperature
def adc_to_temp(adc_value):
    voltage = (adc_value / 65535) * 3.3
    return 27 - (voltage - 0.706) / 0.001721

def find_potentiometer():
    potentiometers = []
    for mux_index in range(mux_num):
        enable_mux(mux_index)
        for channel in range(mux_channels):
            check_for_pause()
            temp_adc_value = temp_pin.read_u16()
            temp = adc_to_temp(temp_adc_value)

            # Deals with the last channel of each mux
            if channel == mux_channels - 1:
                select_channel(channel)
                discharge_input()
                en_pin.value(0)
                adc_value = read_adc()
                voltage = (adc_value / 65535) * 3.3
                data = f"Mux: {mux_index + 1}  Channel: {channel + 1}  Temperature: {temp:.5f}  Voltage: {voltage:.4f}"
                # print(data)
                time.sleep(0.1)  # Allow the multiplexer to settle and ADC to stabilize
                sys.stdout.write(data.encode() + b'\r\n')
                if adc_value > threshold:  # Define a suitable threshold based on your setup
                    potentiometers.append((mux_index + 1, channel + 1))
                en_pin.value(1)
                # Just randomly select a channel after reading the data of last channel
                # Avoid the selecting stops at the last channel before jumping out of the loop
                # 1 - 31 channel should all be fine, here is 31st channel
                select_channel(channel - 1)
                break

            select_channel(channel)
            discharge_input()
            en_pin.value(0)
            adc_value = read_adc()
            voltage = (adc_value / 65535) * 3.3
            data = f"Mux: {mux_index + 1}  Channel: {channel + 1}  Temperature: {temp:.5f}  Voltage: {voltage:.4f}"
            # print(data)
            time.sleep(0.1)  # Allow the multiplexer to settle and ADC to stabilize
            sys.stdout.write(data.encode() + b'\r\n')
            if adc_value > threshold:  # Define a suitable threshold based on your setup
                potentiometers.append((mux_index + 1, channel + 1))
            en_pin.value(1)
        disable_all_muxes()
    return potentiometers


def check_for_pause():
    pull_results = poll_obj.poll(1) # '1' is how long it will wait for message before looping again (in milliseconds)
    if pull_results:
        PC_command = sys.stdin.readline().strip()
        if PC_command == 'PAUSE':
            sys.stdout.write("Pause confirmed\r")
            while True:
                PC_command = sys.stdin.readline().strip()
                if PC_command == 'RESUME':
                    # sys.stdout.write("Resume confirmed\r")
                    break
                else:
                    continue


# Main execution
setup_mcp23017()
reset_mux()
threshold = 15000  # Example threshold, adjust based on potentiometer's expected ADC output
while True:
    # check_for_reset()
    pot_channels = find_potentiometer()
    # if pot_channels:
    #     for mux, channel in pot_channels:
    #         print(f"Potentiometer found on Mux {mux}, Channel {channel}")
    # else:
    #     print("Potentiometer not found on any channel")
    time.sleep(1)