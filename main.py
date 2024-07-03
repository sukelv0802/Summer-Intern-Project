from machine import I2C, Pin, ADC
import time

# I2C setup
i2c = I2C(1, scl=Pin(3), sda=Pin(2))

# MCP23017 Constants
MCP23017_ADDR = 0x20
IODIRA = 0x00
GPIOA = 0x12

# ADC setup
adc = ADC(Pin(27))  # Assuming GPIO 27 is ADC capable

# Configure GPIOA pins as outputs
def setup_mcp23017():
    i2c.writeto_mem(MCP23017_ADDR, IODIRA, b'\x00')

# Select the multiplexer channel
def select_channel(channel):
    if 0 <= channel <= 31:
        i2c.writeto_mem(MCP23017_ADDR, GPIOA, bytes([channel]))
    else:
        print("Invalid Channel")

def read_adc():
    return adc.read_u16()

def find_potentiometer():
    potentiometers = []
    for channel in range(32):
        select_channel(channel)
        time.sleep(0.5)  # Allow the multiplexer to settle and ADC to stabilize
        adc_value = read_adc()
        print(f"Channel {channel}: ADC Value = {adc_value}")
        if adc_value > threshold:  # Define a suitable threshold based on your setup
            potentiometers.append(channel)
    return potentiometers

# Main execution
setup_mcp23017()
threshold = 15000  # Example threshold, adjust based on potentiometer's expected ADC output
pot_channel = find_potentiometer()
if pot_channel:
    print(f"Potentiometer found on channel {', '.join(map(str, pot_channel))}")
else:
    print("Potentiometer not found on any channel")