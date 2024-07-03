import machine
from machine import Pin, ADC
import time
from time import sleep
import uos

uart = machine.UART(0, baudrate = 115200)
uart.init(115200, bits = 8, parity = None, stop = 1, tx = Pin(0), rx = Pin(1))
uos.dupterm(uart)

pot = ADC(Pin(26))

while True:  
  pot_value = pot.read_u16() # read value, 0-65535 across voltage range 0.0v - 3.3v
  print(pot_value)
  time.sleep(0.5)