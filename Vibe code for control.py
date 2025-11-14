import serial
import time
import datetime as dt
import matplotlib.pyplot as plt

# --- Configuration for Magna-Power TS Series ---
# Serial port settings
SERIAL_PORT = 'COM3'
BAUDRATE = 19200

# Measurement settings
MEASUREMENT_INTERVAL_SECONDS = 5  # 1 minute
TOTAL_MEASUREMENTS = 5  # Take X number of measurements (X * 1 minute total running time)
SET_CURRENT = 5.0  # The constant current to source (Constant Current mode)
SET_VOLTAGE = 9.0  # The maximum compliance voltage allowed (safeguard)

# --- Data Storage ---
# Lists to store the measurement time (x-axis) and voltage (y-axis)
time_stamps = []
voltages = []


# --- Helper Function ---
def send_command(connection, command):
    """Sends an SCPI command and prints it to the console."""
    command_str = command + '\n'
    connection.write(command_str.encode())
    print(f"Sent: {command_str.strip()}")


def query_device(connection, query):
    """Sends an SCPI query, prints it, and returns the response."""
    # Ensure the command ends with a question mark and newline
    query_str = query + '\n'
    connection.write(query_str.encode())
    # Read response. Timeout is handled by the serial object config.
    response = connection.readline().decode().strip()
    return response


# --- Main Logic ---
conn = None  # Initialize connection variable

try:
    print(f"Attempting to connect to {SERIAL_PORT}...")
    # 1. Create serial connection object
    conn = serial.Serial(port=SERIAL_PORT, baudrate=BAUDRATE, timeout=5)

    # Add a 5-second start delay
    time.sleep(5)

    # 2. Identify the product
    identity = query_device(conn, '*IDN?')
    print(f"Device Identity: {identity}")

    # 3. Configure and Enable Power Supply (TS Series Setup)
    send_command(conn, 'CONF:SETPT 0')  # Local control (setting 0)
    send_command(conn, 'CURR 0')  # Set current to 0 Adc initially
    send_command(conn, f'VOLT {SET_VOLTAGE}')  # Set Voltage limit (compliance)
    send_command(conn, 'OUTP:START')  # Enable DC output

    # Set the test current for the stability test
    send_command(conn, f'CURR {SET_CURRENT}')
    print(f"TS Power Supply set to source {SET_CURRENT} Adc (CC Mode, V-limit: {SET_VOLTAGE}V).")

    # 4. Data Logging Loop
    print("-" * 30)
    print(f"Starting {TOTAL_MEASUREMENTS} measurements (1 minute intervals)...")

    # Wait 20 seconds for the unit to settle at the new current setpoint
    print("Initial 20 second stabilization wait...")
    time.sleep(20)

    for i in range(TOTAL_MEASUREMENTS):
        print(f"\n--- Measurement {i + 1}/{TOTAL_MEASUREMENTS} ---")

        # A. Query Voltage Measurement
        # This measures the actual voltage output while sourcing SET_CURRENT
        voltage_response = query_device(conn, 'MEAS:VOLT?')

        try:
            # B. Parse and Store Data
            measured_voltage = float(voltage_response)
            current_time = dt.datetime.now()

            time_stamps.append(current_time)
            voltages.append(measured_voltage)

            print(f"LOGGED: Time: {current_time.strftime('%H:%M:%S')}, Measured Voltage: {measured_voltage:.3f} V")

        except ValueError:
            print(f"ERROR: Could not parse voltage response: '{voltage_response}'")
            continue

        # C. Wait for the next measurement interval (60 seconds)
        if i < TOTAL_MEASUREMENTS - 1:
            print(f"Waiting for {MEASUREMENT_INTERVAL_SECONDS} seconds...")
            time.sleep(MEASUREMENT_INTERVAL_SECONDS)

    print("\nData logging complete.")

    # 5. --- Matplotlib Plotting ---
    if voltages:
        print("Generating plot...")
        fig, ax = plt.subplots(figsize=(10, 6))

        # Plot the data
        ax.plot(time_stamps, voltages, marker='o', linestyle='-', color='#10b981', linewidth=2, markersize=6)

        # Add labels and title
        ax.set_title(f'Magna-Power TS Voltage Stability (CC={SET_CURRENT}A)', fontsize=16, fontweight='bold')
        ax.set_xlabel('Time', fontsize=12)
        ax.set_ylabel('Measured Output Voltage (V)', fontsize=12)

        # Add a grid and format the background
        ax.grid(True, linestyle='--', alpha=0.7)
        fig.patch.set_facecolor('#f3f4f6')
        ax.set_facecolor('white')

        # Improve x-axis readability for time stamps
        fig.autofmt_xdate(rotation=45)
        plt.tight_layout()
        plt.show()
    else:
        print("No valid voltage data collected to plot.")


except serial.SerialException as e:
    print(f"\nFATAL ERROR: Could not open or communicate with the serial port: {e}")
    print("Please check that the port is correct and the device is connected.")

except Exception as e:
    print(f"\nAn unexpected error occurred: {e}")

finally:
    # 6. Safety and Cleanup
    print("-" * 30)
    if conn and conn.is_open:
        try:
            # Send SCPI command to set current to zero and then disable the DC output
            send_command(conn, 'CURR 0')
            send_command(conn, 'OUTP:STOP')
            print("DC Output safely stopped and current set to 0.")
        except Exception as e:
            print(f"Error during shutdown: {e}")

        # Close the communication channel
        conn.close()
        print("Serial connection closed.")
    print("Program finished.")