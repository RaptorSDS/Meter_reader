import sys
import serial
import time
import tkinter as tk
from tkinter import ttk, scrolledtext, messagebox, filedialog
from smllib import SmlStreamReader, const
from collections import deque

class SmartMeterReader:
    def __init__(self, root):
        self.root = root
        self.root.title("Smart Meter Ausleser")
        self.root.minsize(800, 600)
        
        self.serial_connection = None
        self.read_timer = None
        self.sml_reader = SmlStreamReader()
        self.raw_buffer = bytearray()
        self.measurement_history = deque(maxlen=60)
        self.update_interval = 1  # Default update interval in seconds
        self.frame_count = 0
        
        self.init_ui()

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
        button_frame = ttk.Frame(left_frame)
        button_frame.pack(fill=tk.X, pady=5)
        
        info_button = ttk.Button(button_frame, text="?", width=3, command=self.show_info)
        info_button.pack(side=tk.RIGHT, padx=2)
        
        save_button = ttk.Button(button_frame, text="Daten speichern", command=self.save_last_minute_data)
        save_button.pack(side=tk.RIGHT, padx=2)
        
        # Debug view
        self.create_debug_group(left_frame)
        
        # Right column - Table
        right_frame = ttk.Frame(main_frame)
        right_frame.pack(side=tk.LEFT, fill=tk.BOTH, expand=True, padx=5)
        
        self.create_table_group(right_frame)
        
        # Status bar
        self.status_var = tk.StringVar(value="Bereit")
        status_label = ttk.Label(self.root, textvariable=self.status_var, relief=tk.SUNKEN, anchor=tk.W)
        status_label.pack(side=tk.BOTTOM, fill=tk.X)

    def create_connection_group(self, parent):
        """Creates the connection settings group"""
        group = ttk.LabelFrame(parent, text="Verbindungseinstellungen")
        group.pack(fill=tk.X, pady=5)
        
        # Port selection
        ttk.Label(group, text="Port:").pack(padx=5, pady=2)
        self.port_combo = ttk.Combobox(group, values=self.get_serial_ports())
        self.port_combo.pack(fill=tk.X, padx=5, pady=2)
        
        # Baud rate selection
        ttk.Label(group, text="Baudrate:").pack(padx=5, pady=2)
        self.baud_rate_combo = ttk.Combobox(group, values=["9600", "19200", "38400", "57600", "115200"])
        self.baud_rate_combo.set("9600")
        self.baud_rate_combo.pack(fill=tk.X, padx=5, pady=2)
        
        # Data bits
        ttk.Label(group, text="Datenbits:").pack(padx=5, pady=2)
        self.data_bits_combo = ttk.Combobox(group, values=["6", "7", "8"])
        self.data_bits_combo.set("8")
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

    def create_debug_group(self, parent):
        """Creates the debug view"""
        group = ttk.LabelFrame(parent, text="Debug-Anzeige")
        group.pack(fill=tk.BOTH, expand=True, pady=5)
        
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

    def get_serial_ports(self):
        """Returns a list of available serial ports"""
        ports = [f"COM{i}" for i in range(1, 256)]
        available_ports = []
        for port in ports:
            try:
                s = serial.Serial(port)
                s.close()
                available_ports.append(port)
            except (OSError, serial.SerialException):
                pass
        return available_ports or ["Keine Ports gefunden"]

    def set_update_interval(self):
        """Updates the data display interval"""
        interval_text = self.update_interval_combo.get()
        self.update_interval = int(interval_text.replace("s", ""))

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
                
            baud_rate = int(self.baud_rate_combo.get())
            data_bits = int(self.data_bits_combo.get())
            parity_map = {"None": serial.PARITY_NONE, "Odd": serial.PARITY_ODD, "Even": serial.PARITY_EVEN}
            parity = parity_map[self.parity_combo.get()]

            if self.serial_connection and self.serial_connection.is_open:
                self.serial_connection.close()

            self.serial_connection = serial.Serial(
                port=port,
                baudrate=baud_rate,
                bytesize=data_bits,
                parity=parity,
                timeout=1
            )
            
            if self.serial_connection.is_open:
                self.debug_output.insert(tk.END, f"Verbunden mit {port}\n")
                self.debug_output.see(tk.END)
                self.root.after(1000, self.read_serial_data)
                self.status_var.set(f'Verbunden mit {port}')
        except Exception as e:
            messagebox.showerror("Verbindungsfehler", f"Fehler beim Verbinden: {str(e)}")
            self.debug_output.insert(tk.END, f"Verbindungsfehler: {str(e)}\n")
            self.debug_output.see(tk.END)

    def disconnect_serial(self):
        """Disconnects the serial connection"""
        try:
            if self.serial_connection and self.serial_connection.is_open:
                self.serial_connection.close()
                self.debug_output.insert(tk.END, "Verbindung getrennt.\n")
                self.debug_output.see(tk.END)
                self.status_var.set('Verbindung getrennt')
        except Exception as e:
            self.debug_output.insert(tk.END, f"Fehler beim Trennen der Verbindung: {str(e)}\n")
            self.debug_output.see(tk.END)

    def read_serial_data(self):
        """Reads data from serial and updates the table with OBIS values"""
        if not (self.serial_connection and self.serial_connection.is_open):
            return

        try:
            data = self.serial_connection.read_all()
            if data:
                self.raw_buffer.extend(data)
                
                self.sml_reader.add(self.raw_buffer)
                sml_frame = self.sml_reader.get_frame()
                
                if sml_frame:
                    self.raw_buffer.clear()
                    parsed_msgs = sml_frame.parse_frame()
                    if parsed_msgs:
                        self.frame_count += 1
                        obis_values = parsed_msgs[1].message_body.val_list
                        self.debug_output.insert(tk.END, f"Frame empfangen: {self.frame_count}\n")
                        self.debug_output.see(tk.END)
                        self.measurement_history.append(obis_values)
                        
                        if self.frame_count % self.update_interval == 0:
                            self.display_obis_values(obis_values)
            
            # Schedule next read
            self.root.after(1000, self.read_serial_data)
        except Exception as e:
            self.debug_output.insert(tk.END, f"Fehler beim Lesen der Daten: {str(e)}\n")
            self.debug_output.see(tk.END)

    def display_obis_values(self, obis_values):
        """Displays OBIS values in the table"""
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
                if obis_short in ["1.8.0", "2.8.0", "C.1.0", "0.2.0"]:
                    self.tree.tag_configure('highlighted', background='yellow')
                    self.tree.item(item, tags=('highlighted',))
                    
            except Exception as e:
                self.debug_output.insert(tk.END, f"Fehler beim Anzeigen von OBIS-Wert: {str(e)}\n")
                self.debug_output.see(tk.END)
              def save_last_minute_data(self):
        """Saves the last 60 seconds of data to a file"""
        try:
            if not self.measurement_history:
                messagebox.showwarning("Keine Daten", "Keine Daten zum Speichern verfügbar!")
                return

            file_path = filedialog.asksaveasfilename(
                defaultextension=".txt",
                filetypes=[("Textdateien", "*.txt"), ("Alle Dateien", "*.*")]
            )
            
            if file_path:
                with open(file_path, "w", encoding='utf-8') as file:
                    file.write("Smart Meter Daten der letzten 60 Sekunden:\n\n")
                    for timestamp, obis_values in enumerate(self.measurement_history):
                        file.write(f"Zeitpunkt {timestamp + 1}:\n")
                        for entry in obis_values:
                            unit = const.UNITS.get(entry.unit, "")
                            scaler = entry.scaler if entry.scaler is not None else 0
                            scaled_value = entry.value * (10 ** scaler)
                            file.write(f"{entry.obis.obis_short}: {scaled_value} {unit}\n")
                        file.write("\n")
                messagebox.showinfo("Gespeichert", "Daten erfolgreich gespeichert!")
        except Exception as e:
            messagebox.showerror("Fehler", f"Fehler beim Speichern: {str(e)}")

    def on_closing(self):
        """Handle application closure"""
        self.disconnect_serial()
        self.root.destroy()

if __name__ == "__main__":
    root = tk.Tk()
    app = SmartMeterReader(root)
    root.protocol("WM_DELETE_WINDOW", app.on_closing)
    root.mainloop()
