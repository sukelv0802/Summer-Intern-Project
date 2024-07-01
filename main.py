import machine
from machine import Pin, ADC
import time
import uos

# need this UART to send data to local computer
uart = machine.UART(0, baudrate=115200)
uart.init(115200, bits=8, parity=None, stop=1, tx=Pin(0), rx=Pin(1))
uos.dupterm(uart)

sensor_temp = machine.ADC(4)

def adc_to_temp(adc_value):
    voltage = (adc_value / 65535) * 3.3
    return 27 - (voltage - 0.706) / 0.001721

while True:
    adc_value = sensor_temp.read_u16()
    temp = adc_to_temp(adc_value)
    print('Temperature Value: ', temp)
    time.sleep(1)
