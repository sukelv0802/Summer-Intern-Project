from machine import Pin, I2C

class MCP23017:
    def __init__(self, i2c, address=0x20):
        self.i2c = i2c
        self.address = address
        self.i2c.writeto_mem(self.address, 0x00, b'\x00')  # Set all A and B pins as outputs

    def output(self, pin, value):
        # Assume pin is from 0 to 7 (GPIOA)
        gpio = 0x12 # GPIOA output register
        current_value = self.i2c.readfrom_mem(self.address, gpio, 1)[0]
        if value:
            current_value |= (1 << pin)
        else:
            current_value &= ~(1 << pin)
        self.i2c.writeto_mem(self.address, gpio, bytes([current_value]))        