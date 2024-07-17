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
i2c = I2C(0, scl=Pin(17), sda=Pin(16))

# MCP23017 Constants
MCP23017_ADDR = 0x20
IODIRA = 0x00
GPIOA = 0x12
IODIRB = 0x01
GPIOB = 0x13

# Mux constants
NUM_MUXES = 4
CHANNELS_PER_MUX = 32

# ADC setup
adc = ADC(Pin(26)) 
temp_pin = machine.ADC(4)

# GPIO setup
CS_PINS = [Pin(i, Pin.OUT) for i in range(2, 2 + NUM_MUXES)]  # These are connected to each CS pin on muxes
WR_PIN = Pin(20, Pin.OUT) 
EN_PIN = Pin(21, Pin.OUT)  
GND_PIN = Pin(22, Pin.OUT) 

# Configure GPIOA pins as outputs
def setup_mcp23017():
    i2c.writeto_mem(MCP23017_ADDR, IODIRA, b'\x00')
    for cs_pin in CS_PINS:
        cs_pin.value(1)
    WR_PIN.value(1)
    EN_PIN.value(1)
    GND_PIN.value(0)

# Enable mux
def enable_mux(mux_index):
    CS_PINS[mux_index].value(0)

# Disable mux
def disable_all_muxes(mux_index):
    CS_PINS[mux_index].value(1)
    select_channel(0xFF)

# Reset mux
def reset_mux():
    i2c.writeto_mem(MCP23017_ADDR, GPIOA, b'\x00')

def select_channel(channel):
    if 0 <= channel < CHANNELS_PER_MUX:
        WR_PIN.value(0)
        i2c.writeto_mem(MCP23017_ADDR, GPIOA, bytes([channel]))
        # print("Just wrote channel") #DEBUGGING
        # time.sleep(1)
        WR_PIN.value(1)
        # print("Just latched address") #DEBUGGING
        # time.sleep(1)
    # else:
    #     print("Invalid Channel")


def discharge_input():
    GND_PIN.value(1)
    time.sleep(0.05)
    GND_PIN.value(0)

def read_adc():
    num_readings = 3
    total = 0
    for _ in range(num_readings):
        total += adc.read_u16()
        time.sleep(0.05)
    return total / num_readings

def adc_to_temp(adc_value):
    voltage = (adc_value / 65535) * 3.3
    return 27 - (voltage - 0.706) / 0.001721

def close_channel(channel):
    select_channel(channel)
    time.sleep(0.01)
    select_channel(0xFF) 
    time.sleep(0.01)
    
def reset_mux_completely():
    # Cycle through and close all channels
    for channel in range(CHANNELS_PER_MUX):
        close_channel(channel)
    
    # Reset the MCP23017
    reset_mux()
    time.sleep(0.1)

def find_potentiometer():
    potentiometers = []
    for mux_index in range(NUM_MUXES):
        enable_mux(mux_index)
        time.sleep(0.05)  # Allow time for mux to stabilize after enabling
        
        for channel in range(CHANNELS_PER_MUX):
            check_for_pause()
            
            # Read temperature
            temp_adc_value = temp_pin.read_u16()
            temp = adc_to_temp(temp_adc_value)
            
            # Open all channels, then select the current one
            select_channel(0xFF)
            time.sleep(0.02)
            select_channel(channel)
            time.sleep(0.02)
            
            discharge_input()
            EN_PIN.value(0)
            time.sleep(0.1)  # Increased delay to allow for settling
            
            adc_value = read_adc()
            voltage = (adc_value / 65535) * 3.3
            
            data = f"Mux: {mux_index + 1}  Channel: {channel + 1}  Temperature: {temp:.5f}  Voltage: {voltage:.4f}"
            sys.stdout.write(data.encode() + b'\r\n')
            
            if adc_value > threshold:        
                potentiometers.append((mux_index + 1, channel + 1))
            
            EN_PIN.value(1)
            time.sleep(0.05)
            
            # Forcibly close the current channel after reading
            close_channel(channel)
            time.sleep(0.02)  # Additional delay after closing
        
        # After finishing with a mux, reset it completely
        reset_mux_completely()
        disable_all_muxes(mux_index)
        time.sleep(0.1)  # Allow time for mux to stabilize after disabling
    
    return potentiometers

def check_for_pause():
    pull_results = poll_obj.poll(1)
    if pull_results:
        PC_command = sys.stdin.readline().strip()
        if PC_command == 'PAUSE':
            sys.stdout.write("Pause confirmed\r")
            while True:
                PC_command = sys.stdin.readline().strip()
                if PC_command == 'RESUME':
                    break
                else:
                    continue
        
def check_for_reset():
    pull_results = poll_obj.poll(1)
    if pull_results:
        PC_command = sys.stdin.readline().strip()
        if PC_command == 'RESET':
            reset_mux()
            enable_mux(0)
            select_channel(0)
            sys.stdout.write("Reset to Mux1, Channel1\r")
            while True:
                PC_command = sys.stdin.readline().strip()
                if PC_command == 'RESUME':
                    break
                else:
                    continue

# Main execution
setup_mcp23017()
reset_mux_completely()
threshold = 15000  # Example threshold, adjust based on potentiometer's expected ADC output
while True:
    pot_channels = find_potentiometer()
    time.sleep(5)