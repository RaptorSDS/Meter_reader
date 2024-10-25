import sys
import serial
import time
import tkinter as tk
from tkinter import ttk, scrolledtext, messagebox, filedialog
from smllib import SmlStreamReader, const
from collections import deque
import threading
import logging
from serial.tools import list_ports
import queue
import os
from datetime import datetime

class Constants:
    DEFAULT_BAUDRATE = 9600
    DEFAULT_DATABITS = 8
    DEFAULT_UPDATE_INTERVAL = 1
    BUFFER_SIZE_LIMIT = 1024 * 1024  # 1MB
    MAX_HISTORY_SIZE = 60
    HIGHLIGHTED_OBIS = ["1.8.0", "2.8.0", "C.1.0", "0.2.0"]
    LOG_FORMAT = '%(asctime)s - %(levelname)s - %(message)s'
    LOG_FILE = 'smartmeter.log'
    UI_UPDATE_INTERVAL = 100  # ms
    SERIAL_TIMEOUT = 1
    MAX_RECONNECT_ATTEMPTS = 3

class SmartMeterReader:
    def __init__(self, root):
        self.root = root
        self.root.title("Smart Meter Ausleser")
        self.root.minsize(800, 600)
        
        # Initialize core components
        self.setup_logging()
        self.init_variables()
        self.init_ui()
        
        # Start UI update loop
        self.update_ui_loop()

    def setup_logging(self):
        """Initializes the logging system"""
        log_dir = 'logs'
        if not os.path.exists(log_dir):
            os.makedirs(log_dir)
        
        log_path = os.path.join(log_dir, f'smartmeter_{datetime.now().strftime("%Y%m%d_%H%M%S")}.log')
        
        logging.basicConfig(
            level=logging.DEBUG,
            format=Constants.LOG_FORMAT,
            handlers=[
                logging.FileHandler(log_path),
                logging.StreamHandler(sys.stdout)
            ]
        )
        self.logger = logging.getLogger(__name__)
        self.logger.info("Smart Meter Reader gestartet")

    def init_variables(self):
        """Initializes all instance variables"""
        self.serial_connection = None
        self.sml_reader = SmlStreamReader()
        self.raw_buffer = bytearray()
        self.measurement_history = deque(maxlen=Constants.MAX_HISTORY_SIZE)
        self.update_interval = Constants.DEFAULT_UPDATE_INTERVAL
        self.frame_count = 0
        
        # Threading components
        self.serial_lock = threading.Lock()
        self.stop_event = threading.Event()
        self.read_thread = None
        self.data_queue = queue.Queue()
        
        # UI update flags
        self.last_ui_update = 0
        self.needs_ui_update = False

    def init_ui(self):
        """Initializes the complete UI"""
        # Create main frame with two columns
        main_frame = ttk.Frame(self.root)
        main_frame.pack(fill=tk.BOTH, expand=True, padx=5, pady=5)
        
        # Left column
        left_frame = ttk.Frame(main_frame)
        left_frame.pack(side=tk.LEFT, fill=tk.BOTH, padx=5)
        
        # Connection settings
        self.create_connection_group(left_frame)
        
        # Button frame for Info and Save
        self.create_button_frame(left_frame)
        
        # Debug view
        self.create_debug_group(left_frame)
        
        # Right column - Table
        right_frame = ttk.Frame(main_frame)
        right_frame.pack(side=tk.LEFT, fill=tk.BOTH, expand=True, padx=5)
        
        self.create_table_group(right_frame)
        
        # Status bar
        self.create_status_bar()

    def create_connection_group(self, parent):
        """Creates the connection settings group"""
        group = ttk.LabelFrame(parent, text="Verbindungseinstellungen")
        group.pack(fill=tk.X, pady=5)
        
        # Port selection with refresh button
        port_frame = ttk.Frame(group)
        port_frame.pack(fill=tk.X, padx=5, pady=2)
        
        ttk.Label(port_frame, text="Port:").pack(side=tk.LEFT)
        self.port_combo = ttk.Combobox(port_frame)
        self.port_combo.pack(side=tk.LEFT, fill=tk.X, expand=True, padx=(5, 0))
        
        refresh_button = ttk.Button(port_frame, text="↻", width=3, command=self.refresh_ports)
        refresh_button.pack(side=tk.LEFT, padx=(5, 0))
        
        # Refresh ports initially
        self.refresh_ports()
        
        # Baud rate selection
        ttk.Label(group, text="Baudrate:").pack(padx=5, pady=2)
        self.baud_rate_combo = ttk.Combobox(group, values=["9600", "19200", "38400", "57600", "115200"])
        self.baud_rate_combo.set(str(Constants.DEFAULT_BAUDRATE))
        self.baud_rate_combo.pack(fill=tk.X, padx=5, pady=2)
        
        # Data bits
        ttk.Label(group, text="Datenbits:").pack(padx=5, pady=2)
        self.data_bits_combo = ttk.Combobox(group, values=["6", "7", "8"])
        self.data_bits_combo.set(str(Constants.DEFAULT_DATABITS))
        self.data_bits_combo.pack(fill=tk.X, padx=5, pady=2)
        
        # Parity
        ttk.Label(group, text="Parität:").pack(padx=5, pady=2)
        self.parity_combo = ttk.Combobox(group, values=["None", "Odd", "Even"])
        self.parity_combo.set("None")
        self.parity_combo.pack(fill=tk.X, padx=5, pady=2)
        
        # Update Interval
        ttk.Label(group, text="Aktualisierungsintervall:").pack(padx=5, pady=2)
        self.update_interval_combo = ttk.Combobox(group, values=["1s", "5s", "15s"])
        self.update_interval_combo.set("1s")
        self.update_interval_combo.bind('<<ComboboxSelected>>', lambda e: self.set_update_interval())
        self.update_interval_combo.pack(fill=tk.X, padx=5, pady=2)
        
        # Connect/Disconnect buttons
        ttk.Button(group, text="Verbinden", command=self.connect_serial).pack(fill=tk.X, padx=5, pady=2)
        ttk.Button(group, text="Trennen", command=self.disconnect_serial).pack(fill=tk.X, padx=5, pady=2)

    def create_button_frame(self, parent):
        """Creates the frame containing Info and Save buttons"""
        button_frame = ttk.Frame(parent)
        button_frame.pack(fill=tk.X, pady=5)
        
        info_button = ttk.Button(button_frame, text="?", width=3, command=self.show_info)
        info_button.pack(side=tk.RIGHT, padx=2)
        
        save_button = ttk.Button(button_frame, text="Daten speichern", command=self.save_last_minute_data)
        save_button.pack(side=tk.RIGHT, padx=2)

    def create_debug_group(self, parent):
        """Creates the debug view"""
        group = ttk.LabelFrame(parent, text="Debug-Anzeige")
        group.pack(fill=tk.BOTH, expand=True, pady=5)
        
        # Add clear button
        clear_button = ttk.Button(group, text="Debug leeren", command=self.clear_debug)
        clear_button.pack(fill=tk.X, padx=5, pady=(5, 0))
        
        self.debug_output = scrolledtext.ScrolledText(group, wrap=tk.WORD)
        self.debug_output.pack(fill=tk.BOTH, expand=True, padx=5, pady=5)

    def create_table_group(self, parent):
        """Creates the table for OBIS values"""
        group = ttk.LabelFrame(parent, text="OBIS-Werte")
        group.pack(fill=tk.BOTH, expand=True, pady=5)
        
        # Create Treeview
        columns = ('obis', 'value', 'unit')
        self.tree = ttk.Treeview(group, columns=columns, show='headings')
        
        # Define headings
        self.tree.heading('obis', text='OBIS-Code')
        self.tree.heading('value', text='Wert')
        self.tree.heading('unit', text='Einheit')
        
        # Configure column widths
        self.tree.column('obis', width=100)
        self.tree.column('value', width=100)
        self.tree.column('unit', width=100)
        
        # Add scrollbar
        scrollbar = ttk.Scrollbar(group, orient=tk.VERTICAL, command=self.tree.yview)
        self.tree.configure(yscrollcommand=scrollbar.set)
        
        # Pack everything
        self.tree.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        scrollbar.pack(side=tk.RIGHT, fill=tk.Y)
        
        # Configure row tags
        self.tree.tag_configure('highlighted', background='yellow')

    def create_status_bar(self):
        """Creates the status bar"""
        self.status_var = tk.StringVar(value="Bereit")
        status_label = ttk.Label(self.root, textvariable=self.status_var, relief=tk.SUNKEN, anchor=tk.W)
        status_label.pack(side=tk.BOTTOM, fill=tk.X)

    def refresh_ports(self):
        """Refreshes the list of available serial ports"""
        available_ports = [port.device for port in list_ports.comports()]
        self.port_combo['values'] = available_ports or ["Keine Ports gefunden"]
        if available_ports:
            self.port_combo.set(available_ports[0])
        else:
            self.port_combo.set("Keine Ports gefunden")

    def set_update_interval(self):
        """Updates the data display interval"""
        interval_text = self.update_interval_combo.get()
        self.update_interval = int(interval_text.replace("s", ""))
        self.logger.debug(f"Update-Intervall auf {self.update_interval}s gesetzt")

    def show_info(self):
        """Displays information about the program"""
        info_text = (
            "Smart Meter Ausleser\n\n"
            "Entwickelt von Tobias Baumann\n"
            "GitHub: https://github.com/RaptorSDS/\n"
            "Hilfe durch Claude AI 3.5 / ChatGPT 4\n\n"
            "Dieses Programm steht unter der GNU General Public License (GPL).\n\n"
            "Verwendete Bibliotheken:\n"
            "- smllib von spaceman_spiff (GNU License)\n"
            "- Tkinter für die Benutzeroberfläche\n\n"
            "Copyright (C) 2024 Tobias Baumann"
        )
        messagebox.showinfo("Info", info_text)

    def connect_serial(self):
        """Connects to the selected serial port"""
        try:
            port = self.port_combo.get()
            if port == "Keine Ports gefunden":
                raise ValueError("Kein COM-Port verfügbar")
            
            # Get connection parameters
            baud_rate = int(self.baud_rate_combo.get())
            data_bits = int(self.data_bits_combo.get())
            parity_map = {"None": serial.PARITY_NONE, "Odd": serial.PARITY_ODD, "Even": serial.PARITY_EVEN}
            parity = parity_map[self.parity_combo.get()]
            
            # Close existing connection if any
            self.disconnect_serial()
            
            # Create new connection
            self.serial_connection = serial.Serial(
                port=port,
                baudrate=baud_rate,
                bytesize=data_bits,
                parity=parity,
                timeout=Constants.SERIAL_TIMEOUT
            )
            
            if self.serial_connection.is_open:
                self.logger.info(f"Verbunden mit {port}")
                self.update_debug(f"Verbunden mit {port}\n")
                self.status_var.set(f'Verbunden mit {port}')
                
                # Start reading thread
                self.stop_event.clear()
                self.read_thread = threading.Thread(target=self.read_serial_loop, daemon=True)
                self.read_thread.start()
        
        except Exception as e:
            error_msg = f"Fehler beim Verbinden: {str(e)}"
            self.logger.error(error_msg)
            messagebox.showerror("Verbindungsfehler", error_msg)
            self.update_debug(f"Verbindungsfehler: {str(e)}\n")

    def disconnect_serial(self):
        """Disconnects the serial connection"""
        try:
            # Stop reading thread
            if self.read_thread and self.read_thread.is_alive():
                self.stop_event.set()
                self.read_thread.join(timeout=1.0)
            
            # Close serial connection
            if self.serial_connection and self.serial_connection.is_open:
                with self.serial_lock:
                    self.serial_connection.close()
                self.update_debug("Verbindung getrennt.\n")
                self.status_var.set('Verbindung getrennt')
                self.logger.info("Verbindung getrennt")
        
        except Exception as e:
            error_msg = f"Fehler beim Trennen der Verbindung: {str(e)}"
            self.logger.error(error_msg)
            self.update_debug(f"{error_msg}\n")

   def read_serial_loop(self):
        """Main loop for reading serial data"""
        while not self.stop_event.is_set():
            try:
                with self.serial_lock:
                    if not self.serial_connection or not self.serial_connection.is_open:
                        break
                    
                    data = self.serial_connection.read_all()
                    if data:
                        # Manage buffer size
                        if len(self.raw_buffer) > Constants.BUFFER_SIZE_LIMIT:
                            self.raw_buffer = self.raw_buffer[-Constants.BUFFER_SIZE_LIMIT:]
                        
                        self.raw_buffer.extend(data)
                        self.sml_reader.add(self.raw_buffer)
                        sml_frame = self.sml_reader.get_frame()
                        
                        if sml_frame:
                            self.raw_buffer.clear()
                            parsed_msgs = sml_frame.parse_frame()
                            if parsed_msgs:
                                self.frame_count += 1
                                obis_values = parsed_msgs[1].message_body.val_list
                                self.data_queue.put(('frame_received', self.frame_count))
                                self.measurement_history.append(obis_values)
                                
                                if self.frame_count % self.update_interval == 0:
                                    self.data_queue.put(('update_display', obis_values))
                
                time.sleep(0.1)  # Prevent CPU overload
            
            except Exception as e:
                error_msg = f"Fehler beim Lesen der Daten: {str(e)}"
                self.logger.error(error_msg)
                self.data_queue.put(('error', error_msg))
                break

    def update_ui_loop(self):
        """Processes queued UI updates"""
        try:
            while not self.data_queue.empty():
                msg_type, data = self.data_queue.get_nowait()
                
                if msg_type == 'frame_received':
                    self.update_debug(f"Frame empfangen: {data}\n")
                
                elif msg_type == 'update_display':
                    self.display_obis_values(data)
                
                elif msg_type == 'error':
                    self.update_debug(f"{data}\n")
        
        except queue.Empty:
            pass
        
        finally:
            # Schedule next update
            self.root.after(Constants.UI_UPDATE_INTERVAL, self.update_ui_loop)

    def display_obis_values(self, obis_values):
        """Displays OBIS values in the table"""
        try:
            # Clear existing items
            for item in self.tree.get_children():
                self.tree.delete(item)
            
            for entry in obis_values:
                try:
                    obis_short = entry.obis.obis_short
                    unit = const.UNITS.get(entry.unit, "")
                    
                    scaler = entry.scaler if entry.scaler is not None else 0
                    scaled_value = entry.value * (10 ** scaler)
                    
                    item = self.tree.insert('', tk.END, values=(obis_short, f"{scaled_value}", unit))
                    
                    # Highlight specific OBIS codes
                    if obis_short in Constants.HIGHLIGHTED_OBIS:
                        self.tree.item(item, tags=('highlighted',))
                
                except Exception as e:
                    self.logger.error(f"Fehler beim Anzeigen von OBIS-Wert: {str(e)}")
                    self.update_debug(f"Fehler beim Anzeigen von OBIS-Wert: {str(e)}\n")
        
        except Exception as e:
            self.logger.error(f"Fehler beim Aktualisieren der Anzeige: {str(e)}")
            self.update_debug(f"Fehler beim Aktualisieren der Anzeige: {str(e)}\n")

    def update_debug(self, message):
        """Thread-safe debug output update"""
        self.root.after(0, self._update_debug, message)

    def _update_debug(self, message):
        """Internal method for updating debug output"""
        self.debug_output.insert(tk.END, message)
        self.debug_output.see(tk.END)

    def clear_debug(self):
        """Clears the debug output"""
        self.debug_output.delete(1.0, tk.END)

    def save_last_minute_data(self):
        """Saves the last 60 seconds of data to a file"""
        try:
            if not self.measurement_history:
                messagebox.showwarning("Keine Daten", "Keine Daten zum Speichern verfügbar!")
                return

            file_path = filedialog.asksaveasfilename(
                defaultextension=".txt",
                filetypes=[("Textdateien", "*.txt"), ("CSV-Dateien", "*.csv"), ("Alle Dateien", "*.*")],
                initialdir="data"
            )
            
            if file_path:
                # Ensure directory exists
                os.makedirs(os.path.dirname(file_path), exist_ok=True)
                
                # Show progress dialog
                progress_window = tk.Toplevel(self.root)
                progress_window.title("Speichere Daten...")
                progress_window.geometry("300x150")
                progress_window.transient(self.root)
                progress_window.grab_set()
                
                progress_label = ttk.Label(progress_window, text="Speichere Messdaten...")
                progress_label.pack(pady=10)
                
                progress_bar = ttk.Progressbar(progress_window, mode='determinate')
                progress_bar.pack(fill=tk.X, padx=20, pady=10)
                
                # Start saving in separate thread
                save_thread = threading.Thread(
                    target=self._save_data_thread,
                    args=(file_path, progress_window, progress_bar),
                    daemon=True
                )
                save_thread.start()
        
        except Exception as e:
            error_msg = f"Fehler beim Speichern: {str(e)}"
            self.logger.error(error_msg)
            messagebox.showerror("Fehler", error_msg)

    def _save_data_thread(self, file_path, progress_window, progress_bar):
        """Thread for saving data"""
        try:
            total_entries = len(self.measurement_history)
            progress_bar['maximum'] = total_entries
            
            with open(file_path, "w", encoding='utf-8') as file:
                file.write("Smart Meter Daten der letzten 60 Sekunden:\n\n")
                
                for i, obis_values in enumerate(self.measurement_history):
                    file.write(f"Zeitpunkt {i + 1}:\n")
                    for entry in obis_values:
                        unit = const.UNITS.get(entry.unit, "")
                        scaler = entry.scaler if entry.scaler is not None else 0
                        scaled_value = entry.value * (10 ** scaler)
                        file.write(f"{entry.obis.obis_short}: {scaled_value} {unit}\n")
                    file.write("\n")
                    
                    # Update progress
                    self.root.after(0, progress_bar.step, 1)
            
            self.logger.info(f"Daten erfolgreich gespeichert in: {file_path}")
            self.root.after(0, self._show_save_success, progress_window)
        
        except Exception as e:
            error_msg = f"Fehler beim Speichern: {str(e)}"
            self.logger.error(error_msg)
            self.root.after(0, self._show_save_error, progress_window, error_msg)

    def _show_save_success(self, progress_window):
        """Shows success message and closes progress window"""
        progress_window.destroy()
        messagebox.showinfo("Gespeichert", "Daten erfolgreich gespeichert!")

    def _show_save_error(self, progress_window, error_msg):
        """Shows error message and closes progress window"""
        progress_window.destroy()
        messagebox.showerror("Fehler", error_msg)

    def on_closing(self):
        """Handle application closure"""
        try:
            self.logger.info("Beende Anwendung...")
            self.disconnect_serial()
            self.root.destroy()
        except Exception as e:
            self.logger.error(f"Fehler beim Beenden: {str(e)}")
            self.root.destroy()

if __name__ == "__main__":
    try:
        root = tk.Tk()
        app = SmartMeterReader(root)
        root.protocol("WM_DELETE_WINDOW", app.on_closing)
        root.mainloop()
    except Exception as e:
        logging.error(f"Kritischer Fehler: {str(e)}")
        messagebox.showerror("Kritischer Fehler", f"Ein kritischer Fehler ist aufgetreten:\n{str(e)}")
