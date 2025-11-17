import serial
import time
import datetime as dt
import openpyxl
import pandas as pd  # Import pandas for Excel export

# --- Configuration for Magna-Power TS Series ---
# Serial port settings
SERIAL_PORT = 'COM4'
BAUDRATE = 19200

# Measurement settings
MEASUREMENT_INTERVAL_SECONDS = 5  # Set measurement interval in seconds
TOTAL_MEASUREMENTS = 5  # set number of measurements (multiply measurement interval by total measurements to get total running time)
SET_CURRENT = 5.0  # The constant current to source (Constant Current mode)
SET_VOLTAGE = 9.0  # The maximum compliance voltage allowed (safeguard)
EXCEL_FILENAME = r'C:\Users\kevin\OneDrive\Desktop\saratoga energy\Magna Power\magnatest1.xlsx' # Set file name & location <filename>.xlsx

# --- Data Storage ---
# Lists to store the measurement data (elapsed_times now stores HOURS)
elapsed_times = []
voltages = []
currents = []


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
test_start_time = None

try:
    print(f"Attempting to connect to {SERIAL_PORT}...")
    # 1. Create serial connection object
    conn = serial.Serial(port=SERIAL_PORT, baudrate=BAUDRATE, timeout=5)

    # Add a start delay in seconds
    time.sleep(5)

    # 2. Identify the product
    identity = query_device(conn, '*IDN?')
    print(f"Device Identity: {identity}")

    # Record the start time for elapsed time calculation
    test_start_time = time.time()

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
    print(f"Starting {TOTAL_MEASUREMENTS} measurements ({MEASUREMENT_INTERVAL_SECONDS} second intervals)...")

    # Wait 20 seconds for the unit to settle at the new current setpoint
    print("Initial 20 second stabilization wait...")
    time.sleep(20)

    for i in range(TOTAL_MEASUREMENTS):
        # Calculate elapsed time in seconds (used for console display)
        elapsed_time_seconds = time.time() - test_start_time
        # Calculate elapsed time in hours (used for Excel logging)
        elapsed_time_hours = elapsed_time_seconds / 3600.0

        # Console output shows time in seconds for clarity during the run
        print(f"\n--- Measurement {i + 1}/{TOTAL_MEASUREMENTS} (Elapsed: {elapsed_time_seconds:.1f}s) ---")

        # A. Query Measurements
        voltage_response = query_device(conn, 'MEAS:VOLT?')
        current_response = query_device(conn, 'MEAS:CURR?')

        try:
            # B. Parse and Store Data
            measured_voltage = float(voltage_response)
            measured_current = float(current_response)

            elapsed_times.append(elapsed_time_hours)  # Storing in Hours
            voltages.append(measured_voltage)
            currents.append(measured_current)

            print(f"LOGGED: V: {measured_voltage:.3f} V, I: {measured_current:.3f} A")

        except ValueError:
            print(f"ERROR: Could not parse one or more responses (V: '{voltage_response}', I: '{current_response}')")
            continue

        # C. Wait for the next measurement interval
        if i < TOTAL_MEASUREMENTS - 1:
            print(f"Waiting for {MEASUREMENT_INTERVAL_SECONDS} seconds...")
            time.sleep(MEASUREMENT_INTERVAL_SECONDS)

    print("\nData logging complete.")

    # 5. --- Excel Export ---
    if elapsed_times:
        print(f"Creating Excel file: {EXCEL_FILENAME}...")

        # Create a dictionary for the DataFrame
        data = {
            'Elapsed Time (h)': elapsed_times,  # Column header is now (h)
            'Measured Voltage (V)': voltages,
            'Measured Current (A)': currents
        }

        # Create the DataFrame
        df = pd.DataFrame(data)

        # Save to Excel
        df.to_excel(EXCEL_FILENAME, index=False)

        print(f"Successfully saved {len(df)} data points to {EXCEL_FILENAME}")
    else:
        print("No valid data collected for export.")


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