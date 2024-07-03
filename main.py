import machine
from machine import Pin, ADC, I2C
import time
import uos

# need this UART to send data to local computer
uart = machine.UART(0, baudrate=115200)
uart.init(115200, bits=8, parity=None, stop=1, tx=Pin(0), rx=Pin(1))
uos.dupterm(uart)

i2c = I2C(0, scl=Pin(17), sda=Pin(16))

# MCP23017 Constants
MCP23017_ADDR = 0x20
IODIRA = 0x00
GPIOA = 0x12

# GPIO for controlling the common output
led = Pin(19, Pin.OUT)

# Configure GPIOA pins as outputs
def setup_mcp23017():
    i2c.writeto_mem(MCP23017_ADDR, IODIRA, b'\x00')
    
def set_mux_channel(channel):
    if 0 <= channel <= 31:
        i2c.writeto_mem(MCP23017_ADDR, GPIOA, bytes([channel]))
    else:
        print("Invalid Channel")
        
def light_led(channel, duration=0.1):
    set_mux_channel(channel)
    led.value(1)
    time.sleep(duration)
    led.value(0)
        
def cycle_through_leds():
    for channel in range(32):
        light_led(channel)
        print(f"Lit LED on channel: {channel}")
        time.sleep(1)
        
# Main Program
setup_mcp23017()
led.value(0)
while True:
    cycle_through_leds()

        
    
    

    


