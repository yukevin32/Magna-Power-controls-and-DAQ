# Import pySerial, which encapsulates the serial port access
# Import time: Time access and conversions to allow for pausing
import serial, time
# Create serial connection object with default baudrate for MagnaLOADs
conn = serial.Serial(port='COM3', baudrate=19200)

# Add a start delay here in seconds
time.sleep(20000)

# Send SCPI command requesting the product to identify itself
conn.write('*IDN?\n'.encode())
# Receive the product's response and display it in the terminal
print(conn.readline())

# Send SCPI command to configure the MagnaDC for local control
conn.write('CONF:SETPT 0\n'.encode())
# Send SCPI command to set the DC output current to 0 Adc before enabling DC input
conn.write('CURR 0\n'.encode())
# Send SCPI command to enable the MagnaDC power supply output
conn.write('OUTP:START\n'.encode())
# Send SCPI command to set the DC input current to desired Adc
conn.write('CURR 580\n'.encode())

# Send SCPI command to measure the voltage and display it after 5 min
time.sleep(300)
conn.write('MEAS:VOLT?\n'.encode())
print(f"The voltage at 5 minutes is {conn.readline()}")

# Enter desired run time here
time.sleep(28800)
# Send SCPI command to measure the voltage before stopping the program
conn.write('MEAS:VOLT?\n'.encode())
print(f"The ending voltage is {conn.readline()}")

# Send SCPI command to disable the DC output
conn.write('OUTP:STOP\n'.encode())

# Close the communication channel to the product
conn.close()