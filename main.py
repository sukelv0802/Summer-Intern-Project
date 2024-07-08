import machine
from machine import Pin, I2C
import time
import uos

# UART setup (unchanged)
uart = machine.UART(0, baudrate=115200)
uart.init(115200, bits=8, parity=None, stop=1, tx=Pin(0), rx=Pin(1))
uos.dupterm(uart)

i2c = I2C(0, scl=Pin(17), sda=Pin(16))

# MCP23017 Constants
MCP23017_ADDR = 0x20
IODIRA = 0x00
IODIRB = 0x01
GPIOA = 0x12
GPIOB = 0x13

led = Pin(19, Pin.OUT)

def setup_mcp23017():
    i2c.writeto_mem(MCP23017_ADDR, IODIRA, b'\x00')
    i2c.writeto_mem(MCP23017_ADDR, IODIRB, b'\x00')
    i2c.writeto_mem(MCP23017_ADDR, GPIOB, b'\x03')

def set_mux_channel(channel):
    if 0 <= channel <= 31:
        address = channel & 0x1F
        i2c.writeto_mem(MCP23017_ADDR, GPIOB, b'\x02')
        i2c.writeto_mem(MCP23017_ADDR, GPIOA, bytes([address]))
        i2c.writeto_mem(MCP23017_ADDR, GPIOB, b'\x03') # latch the address by setting the WR to high
    else:
        print("Invalid Channel")

def light_led(channel, duration=0.5):
    set_mux_channel(channel)
    i2c.writeto_mem(MCP23017_ADDR, GPIOB, b'\x01') # Set EN low to enable the selected channel
    led.value(1)
    print(f"Lit LED on channel: {channel+1}")
    time.sleep(duration)
    led.value(0)
    i2c.writeto_mem(MCP23017_ADDR, GPIOB, b'\x03') # Set EN high to disable all channels


# Main Program
setup_mcp23017()
led.value(0)
while True:
    for channel in range(32):
        light_led(channel)
        time.sleep(0.5)