Summer Intern Project led by Jim Dratz

Team Members:
1. Kelvin Su
2. Chenyang Hu

----------------------**Introduction**----------------------
- Only two python scripts will be used in this project, ***main.py*** and ***applicationUpdated.py***
- The python script ***main.py*** is designed for use with the Raspberry Pi Pico microcontroller to manage data acquisition from multiple sensors through a series of multiplexers. It utilizes UART for serial communication with a host computer, I2C for interfacing with an MCP23017 I/O expander that controls 8 ADG732 multiplexers, and ADC for reading sensor data. The script supports dynamic control via stdin and stdout, allowing operations like pausing and resuming data collection.
- The Python script ***applicationUpdated.py*** serves as a GUI-based control and monitoring system for a Raspberry Pi Pico-based data acquisition setup. It allows users to start, pause, and resume data collection, set operational parameters like thresholds and cycle periods, and visualize the acquired data in real-time. Additionally, the application supports exporting collected data to Excel for further analysis.

----------------------**Prerequisites**----------------------
- PC with a USB port
- Raspberry Pi Pico
- MCP23017 I/O Expander
- ADG732 Multiplexers (up to 8, each with 32 channels)
- Analog sensors compatible with the multiplexer setup
- Necessary wires and resistors for connections
- Breadboard or similar setup for prototyping
- IDE for developing Python code to Raspberry Pi Pico
- Libraries used in the program (e.g. ***PyQt5***, ***pyserial***, ***openpyxl***, ***numpy***)
- (Optional) Serial terminal software (e.g. ***Putty***, ***Tera Term***)

----------------------**Installation**----------------------
- Connect each multiplexer's control pins to the specified GPIO pins on the Pico as per the script settings.    (e.g. Connect the MCP23017 to the Raspberry Pi Pico via I2C, SDA to Pin2, SCL to Pin3)
- The Raspberry Pi Pico is used for providing power for ADG732 multiplexers, and the constant voltage source is  used for providing power for testing PCBs.
- **CS** (Chip Selecing) lines use 8 separate lines, while **EN** (Enable) line and **WR** (Write) line are only 1 for each.
- Ensure the connection is at same the baud rate. (typically 115200)
- Ensure the Raspberry Pi Pico is connected via USB and that the correct COM port is selected in the application.

----------------------**Additional Notes**----------------------
1. Modify the script parameters such as **channel_period** in ***main.py*** based on the specific timing and performance requirements of your sensors and multiplexers. The unit of **channel_period** is *s*. (e.g. **channel_period** = 0.1 means a frequecny of 10Hz)
2. The **channel_period_value** in ***applicationUpdated.py*** is **NOT** the actual frequecny. However, it should be corresponding to the **channel_period** in ***main.py***, typically, half of it for convenience. Also, the unit for **channel_period_value** is *ms*. (e.g. if **channel_period** is 0.1 in ***main.py***, then **channel_period_value** should be 50 in ***applicationUpdated.py***)
3. When connecting the Raspberry Pi Pico to the PC, make sure Pico is disconnected. Otherwise, thread blocking might occur.
4. Power off the constant voltage source when connecting the Raspberry Pi Pico to the PC.
5. The threshold voltage for Raspberry Pi Pico is 3.3V, so the constant voltage source can be no larger than 4V

----------------------**Operational Controls**----------------------
- **Start**: Begins the data acquisition process and reset the reading to mux1 channel1. But actually you should only use this button when launching the project.
- **Resume**: Continues data acquisition from where it was paused.
- **Stop**: Temporarily halts the data acquisition, retaining the current state.
- **Clear**: Clear all the data displaying on the UI.
- **Set threshold**: Once set, the voltage below this value will be highlighted. Initially defaulted as *None*, so nothing will be highlighted if no value is set.
- **Set cycle period**: Default value is 60s.
- **COM ports**: Varies from different PCs. Not necessarily COM9.
- **Export to Excel**: Export collecting data to an Excel and preserved those highlighted parts.
- **Show below threshold**: Only show data below the set threshold. That is the highlighted ones.


----------------------**Potential Problems**----------------------
1. The **check_for_pause()** function in ***main.py*** is not robust, changing them might cause unexpected errors or crashes. 
2. **star_update**, **resume_update** and **stop_update** functions in ***applicationUpdate.py*** are not robust, changing them might cause unexpected errors or crashes.
3. **START** command is used for reset the selection from mux1 channel1. But in the UI window, you should press **STOP** button before pressing **START** button to function properly, or it just behaves like **RESUME** button. (If **START** still doesn't behave properly after pressing **STOP**, try this process again)
4. When the program keeps receiving data for a long time (several hours) without stop, there might will be a short stuck when pressing **Stop** button.

----------------------**Future Plans**----------------------
- Add code robustness, especially the serial communication part. 
- Add an input bar in UI, which can enable users to select at which multiplexer and which channel to start, not only mux1 channel1. 
- Add time robustness. Though some problems don't happen at first, they happen after some time.
