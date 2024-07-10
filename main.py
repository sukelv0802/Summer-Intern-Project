import machine
from machine import Pin, ADC, I2C
import time
import uos

# UART setup
uart = machine.UART(0, baudrate=115200)
uart.init(115200, bits=8, parity=None, stop=1, tx=Pin(0), rx=Pin(1))
uos.dupterm(uart)

i2c = I2C(0, scl=Pin(17), sda=Pin(16))

# MCP23017 Constants
MCP23017_ADDR = 0x20
IODIRA = 0x00
GPIOA = 0x12
IODIRB = 0x01
GPIOB = 0x13

# ADC setup for reading voltage and temperature
adc = ADC(Pin(26))  
sensor_temp = machine.ADC(4)

##########CONFIGURATION: CAN CHANGE DEPENDING ON WHAT MUX IS BEING USED################## 
NUM_MUXES = 1  
CHANNELS_PER_MUX = 32

# GPIO pins for CS pins (Starts from Pin 2)
CS_PINS = [Pin(i, Pin.OUT) for i in range(2, 2 + NUM_MUXES)] 

# GPIO pins for WR and EN
WR_PIN = Pin(20, Pin.OUT)
EN_PIN = Pin(21, Pin.OUT)
############################

def setup_mcp23017():
    i2c.writeto_mem(MCP23017_ADDR, IODIRA, b'\x00')
    for cs_pin in CS_PINS:
        cs_pin.value(1)
    WR_PIN.value(1)
    EN_PIN.value(1)

def set_mux_channel(mux, channel):
    if 0 <= mux < NUM_MUXES and 0 <= channel < CHANNELS_PER_MUX:
        for cs_pin in CS_PINS:
            cs_pin.value(1)

        address = channel & 0x1F
        WR_PIN.value(0)
        i2c.writeto_mem(MCP23017_ADDR, GPIOA, bytes([address]))
        CS_PINS[mux].value(0)
        WR_PIN.value(1) # Set WR high to latch the address
        CS_PINS[mux].value(1)
    else:
        print("Invalid Mux or Channel")

def read_voltage(mux, channel):
    set_mux_channel(mux, channel)
    EN_PIN.value(0)
    adc_value = adc.read_u16()
    voltage = (adc_value / 65535) * 3.3
    print(f"Mux {mux+1}, Channel {channel+1} Voltage: {voltage:.3f} V")
    time.sleep(0.01)
    EN_PIN.value(1)
    # time.sleep(1)
    return voltage

def adc_to_temp(adc_value):
    voltage = (adc_value / 65535) * 3.3
    return 27 - (voltage - 0.706) / 0.001721

# Main Program
setup_mcp23017()

cycles = 1
while True:
    temp_adc_value = sensor_temp.read_u16()
    temp = adc_to_temp(temp_adc_value)
    print(f'Cycles Number: {cycles}')
    cycles += 1
    print(f'Temperature: {temp:.2f}°C')
    
    for mux in range(NUM_MUXES):
        for channel in range(CHANNELS_PER_MUX):
            voltage = read_voltage(mux, channel)
            # print(f"Mux {mux+1}, Channel {channel+1} Voltage: {voltage:.3f} V")
    
    print("------------------------")
    time.sleep(60)