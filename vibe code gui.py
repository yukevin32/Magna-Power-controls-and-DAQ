import tkinter as tk
from tkinter import ttk, scrolledtext, messagebox, filedialog
import threading
import queue
import time
import serial
import pandas as pd
import openpyxl  # Required by pandas for .xlsx files


# =============================================================================
# === WORKER FUNCTION (YOUR ORIGINAL LOGIC) ===
# =============================================================================

def send_command(connection, command, log_queue):
    """Sends an SCPI command and logs it to the GUI."""
    command_str = command + '\n'
    connection.write(command_str.encode())
    log_queue.put(f"Sent: {command_str.strip()}")


def query_device(connection, query, log_queue):
    """Sends an SCPI query, logs it, and returns the response."""
    query_str = query + '\n'
    connection.write(query_str.encode())
    response = connection.readline().decode().strip()
    return response


def run_measurement_task(config, log_queue, stop_event):
    """
    This is the main logic from your script, adapted to run in a thread.
    """

    # --- Data Storage ---
    elapsed_times = []
    voltages = []
    currents = []

    conn = None
    test_start_time = None

    try:
        log_queue.put(f"Attempting to connect to {config['port']}...")

        # 1. Create serial connection
        conn = serial.Serial(
            port=config['port'],
            baudrate=config['baudrate'],
            timeout=5
        )

        # 2. Wait for connection to establish (User-defined Start Delay)
        # --- MODIFIED: Use configured start_delay ---
        start_delay = config['start_delay']
        log_queue.put(f"Waiting {start_delay}s for connection to stabilize (start delay)...")
        if stop_event.wait(timeout=start_delay):
            log_queue.put("Stop requested during initial wait.")
            raise InterruptedException("Test stopped by user.")
        # --- END MODIFICATION ---

        # 3. Identify the product
        identity = query_device(conn, '*IDN?', log_queue)
        log_queue.put(f"Device Identity: {identity}")

        # Record start time
        test_start_time = time.time()

        # 4. Configure and Enable Power Supply
        send_command(conn, 'CONF:SETPT 0', log_queue)  # Local control
        send_command(conn, 'CURR 0', log_queue)
        send_command(conn, f"VOLT {config['voltage']}", log_queue)
        send_command(conn, 'OUTP:START', log_queue)
        send_command(conn, f"CURR {config['current']}", log_queue)
        log_queue.put(f"Power Supply set: {config['current']} Adc (V-limit: {config['voltage']}V).")

        # 5. Data Logging Loop
        log_queue.put("-" * 30)
        log_queue.put(f"Starting {config['measurements']} measurements...")

        # Initial stabilization wait
        log_queue.put("Initial 20 second stabilization wait...")
        if stop_event.wait(timeout=20):
            log_queue.put("Stop requested during stabilization.")
            raise InterruptedException("Test stopped by user.")

        for i in range(config['measurements']):
            # --- Check for Stop Signal ---
            if stop_event.is_set():
                log_queue.put("Stop requested, halting measurement loop.")
                break

            elapsed_time_seconds = time.time() - test_start_time
            elapsed_time_hours = elapsed_time_seconds / 3600.0

            log_queue.put(
                f"\n--- Measurement {i + 1}/{config['measurements']} (Elapsed: {elapsed_time_seconds:.1f}s) ---")

            # A. Query Measurements
            voltage_response = query_device(conn, 'MEAS:VOLT?', log_queue)
            current_response = query_device(conn, 'MEAS:CURR?', log_queue)

            try:
                # B. Parse and Store Data
                measured_voltage = float(voltage_response)
                measured_current = float(current_response)

                elapsed_times.append(elapsed_time_hours)
                voltages.append(measured_voltage)
                currents.append(measured_current)

                log_queue.put(f"LOGGED: V: {measured_voltage:.3f} V, I: {measured_current:.3f} A")

            except ValueError:
                log_queue.put(f"ERROR: Could not parse responses (V: '{voltage_response}', I: '{current_response}')")
                continue

            # C. Wait for the next interval
            if i < config['measurements'] - 1:
                log_queue.put(f"Waiting for {config['interval']} seconds...")
                if stop_event.wait(timeout=config['interval']):
                    log_queue.put("Stop requested during interval.")
                    break

        log_queue.put("\nData logging complete.")

        # 6. --- Excel Export ---
        if elapsed_times:
            log_queue.put(f"Creating Excel file: {config['filename']}...")
            data = {
                'Elapsed Time (h)': elapsed_times,
                'Measured Voltage (V)': voltages,
                'Measured Current (A)': currents
            }
            df = pd.DataFrame(data)
            df.to_excel(config['filename'], index=False)
            log_queue.put(f"Successfully saved {len(df)} data points.")
        else:
            log_queue.put("No valid data collected for export.")

    except serial.SerialException as e:
        log_queue.put(f"\nFATAL ERROR: Could not communicate with serial port: {e}")
        messagebox.showerror("Serial Port Error", str(e))
    except InterruptedException:
        log_queue.put("Test was manually stopped.")
    except Exception as e:
        log_queue.put(f"\nAn unexpected error occurred: {e}")
        messagebox.showerror("Unexpected Error", str(e))
    finally:
        # 7. Safety and Cleanup
        log_queue.put("-" * 30)
        if conn and conn.is_open:
            try:
                send_command(conn, 'CURR 0', log_queue)
                send_command(conn, 'OUTP:STOP', log_queue)
                log_queue.put("DC Output safely stopped and current set to 0.")
            except Exception as e:
                log_queue.put(f"Error during shutdown: {e}")
            conn.close()
            log_queue.put("Serial connection closed.")

        log_queue.put("Program finished.")
        log_queue.put(None)

    # Custom exception for handling user-requested stops


class InterruptedException(Exception):
    pass


# =============================================================================
# === TKINTER GUI APPLICATION CLASS ===
# =============================================================================

class MagnaPowerApp:
    def __init__(self, root):
        self.root = root
        self.root.title("Magna-Power TS Controller")
        self.root.geometry("650x650")  # --- MODIFIED: Made window taller

        self.log_queue = None
        self.worker_thread = None
        self.stop_event = threading.Event()

        self.create_widgets()

    def create_widgets(self):
        config_frame = ttk.LabelFrame(self.root, text="Configuration", padding=10)
        config_frame.pack(fill='x', padx=10, pady=5)

        config_frame.columnconfigure(1, weight=1)
        config_frame.columnconfigure(2, weight=0)

        # StringVars to hold the entry data
        self.port_var = tk.StringVar(value='COM4')
        self.baud_var = tk.StringVar(value='19200')
        self.start_delay_var = tk.StringVar(value='5')  # --- ADDED ---
        self.volt_var = tk.StringVar(value='9.0')
        self.curr_var = tk.StringVar(value='5.0')
        self.interval_var = tk.StringVar(value='5')
        self.total_var = tk.StringVar(value='5')
        self.file_var = tk.StringVar(
            value=r'C:\Users\kevin\OneDrive\Desktop\saratoga energy\Magna Power\magnatest1.xlsx')

        # Create Labels and Entries
        self.entries = {}

        # --- MODIFIED ROW INDICES ---
        self.entries['port'] = self.create_entry_row(config_frame, "Serial Port:", self.port_var, 0)
        self.entries['baud'] = self.create_entry_row(config_frame, "Baudrate:", self.baud_var, 1)
        self.entries['start_delay'] = self.create_entry_row(config_frame, "Start Delay (s):", self.start_delay_var,
                                                            2)  # --- ADDED ---
        self.entries['volt'] = self.create_entry_row(config_frame, "Set Voltage (V):", self.volt_var, 3)
        self.entries['curr'] = self.create_entry_row(config_frame, "Set Current (A):", self.curr_var, 4)
        self.entries['interval'] = self.create_entry_row(config_frame, "Interval (s):", self.interval_var, 5)
        self.entries['total'] = self.create_entry_row(config_frame, "Total Measurements:", self.total_var, 6)

        # --- MODIFIED ROW INDICES ---
        ttk.Label(config_frame, text="Output File:").grid(row=7, column=0, padx=5, pady=5, sticky='w')
        self.entries['file'] = ttk.Entry(config_frame, textvariable=self.file_var, width=60)
        self.entries['file'].grid(row=7, column=1, padx=5, pady=5, sticky='ew')

        self.browse_button = ttk.Button(config_frame, text="Browse...", command=self.browse_file)
        self.browse_button.grid(row=7, column=2, padx=5, pady=5)

        # --- Control Frame ---
        control_frame = ttk.Frame(self.root, padding=10)
        control_frame.pack(fill='x')

        self.start_button = ttk.Button(control_frame, text="Start Test", command=self.start_test)
        self.start_button.pack(side='left', padx=5, expand=True, fill='x')

        self.stop_button = ttk.Button(control_frame, text="Stop Test", command=self.stop_test, state=tk.DISABLED)
        self.stop_button.pack(side='left', padx=5, expand=True, fill='x')

        # --- Log Frame ---
        log_frame = ttk.LabelFrame(self.root, text="Log", padding=10)
        log_frame.pack(fill='both', expand=True, padx=10, pady=10)

        self.log_area = scrolledtext.ScrolledText(log_frame, wrap=tk.WORD, height=15)
        self.log_area.pack(fill='both', expand=True)
        self.log_area.configure(state=tk.DISABLED)

    def create_entry_row(self, parent, label, var, row):
        """Helper to create a Label-Entry row."""
        ttk.Label(parent, text=label).grid(row=row, column=0, padx=5, pady=5, sticky='w')
        entry = ttk.Entry(parent, textvariable=var, width=20)
        entry.grid(row=row, column=1, padx=5, pady=5, sticky='ew')
        return entry

    def browse_file(self):
        """Open a 'save as' dialog to pick the output file."""
        filename = filedialog.asksaveasfilename(
            defaultextension=".xlsx",
            filetypes=[("Excel files", "*.xlsx"), ("All files", "*.*")],
            initialfile=self.file_var.get()
        )
        if filename:
            self.file_var.set(filename)

    def log(self, message):
        """Inserts a message into the log area and scrolls to the end."""
        self.log_area.configure(state=tk.NORMAL)
        self.log_area.insert(tk.END, message + '\n')
        self.log_area.see(tk.END)
        self.log_area.configure(state=tk.DISABLED)

    def check_log_queue(self):
        """Polls the queue for messages from the worker thread."""
        try:
            while True:
                message = self.log_queue.get_nowait()
                if message is None:
                    self.on_task_complete()
                    return
                else:
                    self.log(str(message))

        except queue.Empty:
            if self.worker_thread and self.worker_thread.is_alive():
                self.root.after(100, self.check_log_queue)
            else:
                self.on_task_complete()

    def start_test(self):
        """Gathers config, validates, and starts the worker thread."""
        self.log_area.configure(state=tk.NORMAL)
        self.log_area.delete('1.0', tk.END)
        self.log_area.configure(state=tk.DISABLED)

        # 1. Gather config and validate
        try:
            config = {
                'port': self.port_var.get(),
                'baudrate': int(self.baud_var.get()),
                'start_delay': int(self.start_delay_var.get()),  # --- ADDED ---
                'voltage': float(self.volt_var.get()),
                'current': float(self.curr_var.get()),
                'interval': int(self.interval_var.get()),
                'measurements': int(self.total_var.get()),
                'filename': self.file_var.get()
            }
        except ValueError as e:
            messagebox.showerror("Validation Error", f"Invalid input: {e}\nPlease check your numbers.")
            return

        if not config['filename']:
            messagebox.showerror("Validation Error", "Output file name cannot be empty.")
            return

        # 2. Update GUI state
        self.set_ui_running(True)

        # 3. Create queue and stop event
        self.log_queue = queue.Queue()
        self.stop_event.clear()

        # 4. Create and start the worker thread
        self.worker_thread = threading.Thread(
            target=run_measurement_task,
            args=(config, self.log_queue, self.stop_event)
        )
        self.worker_thread.start()

        # 5. Start polling the queue
        self.root.after(100, self.check_log_queue)

    def stop_test(self):
        """Sets the stop_event to signal the thread to stop."""
        if self.worker_thread and self.worker_thread.is_alive():
            self.log("Sending stop request...")
            self.stop_event.set()
            self.stop_button.configure(state=tk.DISABLED, text="Stopping...")

    def on_task_complete(self):
        """Resets the UI after the task is finished or stopped."""
        self.log("Task monitoring complete.")
        self.set_ui_running(False)
        self.worker_thread = None
        self.log_queue = None

    def set_ui_running(self, is_running):
        """Disables/Enables UI elements based on task status."""
        if is_running:
            self.start_button.configure(state=tk.DISABLED)
            self.stop_button.configure(state=tk.NORMAL, text="Stop Test")
            self.browse_button.configure(state=tk.DISABLED)
            for entry in self.entries.values():
                entry.configure(state=tk.DISABLED)
        else:
            self.start_button.configure(state=tk.NORMAL)
            self.stop_button.configure(state=tk.DISABLED, text="Stop Test")
            self.browse_button.configure(state=tk.NORMAL)
            for entry in self.entries.values():
                entry.configure(state=tk.NORMAL)


# =============================================================================
# === MAIN EXECUTION ===
# =============================================================================

if __name__ == "__main__":
    try:
        import serial
        import pandas
        import openpyxl
    except ImportError as e:
        print(f"Missing required library: {e.name}")
        print("Please install it using: pip install pyserial pandas openpyxl")
        exit()

    root = tk.Tk()
    app = MagnaPowerApp(root)
    root.mainloop()