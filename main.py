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

def setup_mcp23017():
    i2c.writeto_mem(MCP23017_ADDR, IODIRA, b'\x00')
    i2c.writeto_mem(MCP23017_ADDR, IODIRB, b'\x00')
    i2c.writeto_mem(MCP23017_ADDR, GPIOB, b'\x03')

def set_mux_channel(channel):
    if 0 <= channel <= 31:
        address = channel & 0x1F
        i2c.writeto_mem(MCP23017_ADDR, GPIOB, b'\x02')
        i2c.writeto_mem(MCP23017_ADDR, GPIOA, bytes([address]))
        i2c.writeto_mem(MCP23017_ADDR, GPIOB, b'\x03') # Set WR high to latch the address
    else:
        print("Invalid Channel")

def read_voltage(channel):
    set_mux_channel(channel)
    i2c.writeto_mem(MCP23017_ADDR, GPIOB, b'\x01')
    # time.sleep(1)
    adc_value = adc.read_u16()
    voltage = (adc_value / 65535) * 3.3
    i2c.writeto_mem(MCP23017_ADDR, GPIOB, b'\x03')
    print(f"Channel {channel+1} Voltage: {voltage:.3f} V")
    if voltage >= 1.5:
        time.sleep(5)
    else:
        time.sleep(1)
    return voltage

def adc_to_temp(adc_value):
    voltage = (adc_value / 65535) * 3.3
    return 27 - (voltage - 0.706) / 0.001721

# Main Program
setup_mcp23017()

while True:
    temp_adc_value = sensor_temp.read_u16()
    temp = adc_to_temp(temp_adc_value)
    print(f'Temperature: {temp}°C')
    
    for channel in range(32):
        voltage = read_voltage(channel)
        # print(f"Channel {channel+1} Voltage: {voltage:.3f} V")
    
    print("------------------------")
    time.sleep(1)