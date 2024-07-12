import machine
import sys

led = machine.Pin("LED", machine.Pin.OUT)
while True:
    print("Going to sleep")
    machine.lightsleep(5000)
    print("Toggling light")
    led.toggle()