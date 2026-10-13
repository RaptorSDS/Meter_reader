
import serial
import time
import tkinter as tk
from tkinter import ttk, scrolledtext, messagebox, filedialog
from smllib import SmlStreamReader, const
from collections import deque
from serial.tools import list_ports
import re  # Für Eingabevalidierung

# Zeitraum in Sekunden, der beim Speichern geschrieben wird
SAVE_WINDOW_SECONDS = 60
# Maximale Anzahl Zeilen im Debug-Fenster (ältere Zeilen werden entfernt)
MAX_DEBUG_LINES = 1500
# OBIS-Codes, die in der Tabelle hervorgehoben werden
HIGHLIGHT_OBIS = ["1.8.0", "2.8.0", "C.1.0", "0.2.0"]


class SmartMeterReader:
    def __init__(self, root):
        self.root = root
        self.root.title("Smart Meter Ausleser")
        self.root.minsize(800, 600)

        self.serial_connection = None
        self.read_active = False
        self.read_after_id = None  # ID des geplanten Lese-Aufrufs (für after_cancel)
        self.sml_reader = SmlStreamReader()
        # Speichert Tupel (Zeitstempel, obis_values); ältere Einträge als
        # SAVE_WINDOW_SECONDS werden beim Hinzufügen entfernt
        self.measurement_history = deque()
        self.update_interval = 1  # in Sekunden
        self.last_update_time = time.time()  # Zeitbasierte Aktualisierung

        # Statusleiste vor dem Hauptbereich anlegen, damit sie beim
        # Verkleinern des Fensters nicht abgeschnitten wird
        self.status_var = tk.StringVar(value="Bereit")
        self.status_bar = ttk.Label(self.root, textvariable=self.status_var, relief="sunken", anchor="w")
        self.status_bar.pack(side="bottom", fill="x")

        self.init_ui()

    def init_ui(self):
        main_frame = ttk.Frame(self.root)
        main_frame.pack(expand=True, fill="both", padx=5, pady=5)

        left_frame = ttk.Frame(main_frame)
        left_frame.pack(side="left", fill="y", padx=(0, 5))

        self.create_connection_group(left_frame)

        button_frame = ttk.Frame(left_frame)
        button_frame.pack(fill="x", pady=5)

        info_button = ttk.Button(button_frame, text="?", width=3, command=self.show_info)
        info_button.pack(side="right", padx=2)

        # TODO: Export als CSV ist geplant, derzeit wird nur TXT gespeichert
        save_button = ttk.Button(button_frame, text="Daten speichern (TXT)", command=self.save_last_minute_data)
        save_button.pack(side="right", padx=2)

        self.create_debug_group(left_frame)
        self.create_table_group(main_frame)

    def create_connection_group(self, parent):
        group = ttk.LabelFrame(parent, text="Verbindungseinstellungen")
        group.pack(fill="x", pady=5)

        ttk.Label(group, text="Port:").pack(padx=5, pady=2)
        self.port_combo = ttk.Combobox(group, values=self.get_serial_ports())
        self.port_combo.pack(fill="x", padx=5, pady=2)

        # Button zum Aktualisieren der Port-Liste
        refresh_button = ttk.Button(group, text="Ports aktualisieren", command=self.refresh_ports)
        refresh_button.pack(fill="x", padx=5, pady=2)

        ttk.Label(group, text="Baudrate:").pack(padx=5, pady=2)
        self.baud_rate_combo = ttk.Combobox(group, values=["9600", "19200", "38400", "57600", "115200"])
        self.baud_rate_combo.set("9600")
        self.baud_rate_combo.pack(fill="x", padx=5, pady=2)

        ttk.Label(group, text="Datenbits:").pack(padx=5, pady=2)
        self.data_bits_combo = ttk.Combobox(group, values=["6", "7", "8"])
        self.data_bits_combo.set("8")
        self.data_bits_combo.pack(fill="x", padx=5, pady=2)

        ttk.Label(group, text="Parität:").pack(padx=5, pady=2)
        self.parity_combo = ttk.Combobox(group, values=["None", "Odd", "Even"])
        self.parity_combo.set("None")
        self.parity_combo.pack(fill="x", padx=5, pady=2)

        # TODO: Stoppbits sind noch nicht einstellbar, pySerial nutzt fest 1 Stoppbit
        # (für SML-Zähler üblich: 9600 8N1)
        ttk.Label(group, text="Stoppbits: 1 (fest)").pack(padx=5, pady=2)

        ttk.Label(group, text="Aktualisierungsintervall:").pack(padx=5, pady=2)
        self.update_interval_combo = ttk.Combobox(group, values=["1s", "5s", "15s"])
        self.update_interval_combo.set("1s")
        self.update_interval_combo.bind('<<ComboboxSelected>>', lambda e: self.set_update_interval())
        self.update_interval_combo.pack(fill="x", padx=5, pady=2)

        self.connect_button = ttk.Button(group, text="Verbinden", command=self.connect_serial)
        self.connect_button.pack(fill="x", padx=5, pady=2)
        self.disconnect_button = ttk.Button(group, text="Trennen", command=self.disconnect_serial, state="disabled")
        self.disconnect_button.pack(fill="x", padx=5, pady=2)

    def create_debug_group(self, parent):
        group = ttk.LabelFrame(parent, text="Debug-Anzeige")
        group.pack(fill="both", expand=True, pady=5)

        self.debug_output = scrolledtext.ScrolledText(group, wrap=tk.WORD, height=10)
        self.debug_output.pack(fill="both", expand=True, padx=5, pady=5)

    def create_table_group(self, parent):
        group = ttk.LabelFrame(parent, text="OBIS-Werte")
        group.pack(side="left", fill="both", expand=True)

        columns = ('obis', 'value', 'unit')
        self.table = ttk.Treeview(group, columns=columns, show='headings')
        self.table.heading('obis', text='OBIS-Code')
        self.table.heading('value', text='Wert')
        self.table.heading('unit', text='Einheit')
        self.table.tag_configure('highlight', background='yellow')

        scrollbar = ttk.Scrollbar(group, orient="vertical", command=self.table.yview)
        self.table.configure(yscrollcommand=scrollbar.set)
        self.table.pack(side="left", fill="both", expand=True)
        scrollbar.pack(side="right", fill="y")

    def get_serial_ports(self):
        # Nutzt list_ports aus pySerial für eine zuverlässige Ermittlung
        ports = [port.device for port in list_ports.comports()]
        return ports if ports else ["Keine Ports gefunden"]

    def refresh_ports(self):
        """Aktualisiert die Liste der verfügbaren seriellen Ports"""
        ports = self.get_serial_ports()
        self.port_combo['values'] = ports
        self.log_debug("Port-Liste aktualisiert")

        # Wenn der aktuell ausgewählte Port nicht mehr in der Liste ist, ersten Port auswählen
        if self.port_combo.get() not in ports and ports:
            self.port_combo.set(ports[0])

    def set_update_interval(self):
        interval_text = self.update_interval_combo.get()
        try:
            # Entfernt 's' und konvertiert zu Integer
            if re.match(r'^\d+s$', interval_text):
                self.update_interval = int(interval_text.replace("s", ""))
            else:
                raise ValueError("Ungültiges Intervallformat")
        except ValueError:
            self.log_debug(f"Ungültiges Intervallformat: {interval_text}, setze auf 1s")
            self.update_interval = 1  # Fallback
            self.update_interval_combo.set("1s")

    def show_info(self):
        info_text = (
            "Smart Meter Ausleser\n\n"
            "Entwickelt von Tobias Baumann\n"
            "GitHub: https://github.com/RaptorSDS/\n"
            "Hilfe durch Claude AI 3.5 / ChatGPT 3o_high\n\n"
            "Dieses Programm steht unter der GNU General Public License (GPL).\n\n"
            "Verwendete Bibliotheken:\n"
            "- smllib von spaceman_spiff (GNU License)\n"
            "- Tkinter für die Benutzeroberfläche\n\n"
            "Hinweise:\n"
            "- Stoppbits sind fest auf 1 eingestellt\n"
            "- Gespeichert wird als TXT, CSV-Export ist geplant\n\n"
            "Copyright (C) 2024 Tobias Baumann"
        )
        messagebox.showinfo("Info", info_text)

    def log_debug(self, message):
        """Schreibt Debug-Informationen in das Debug-Fenster."""
        timestamp = time.strftime("%H:%M:%S", time.localtime())
        self.debug_output.insert(tk.END, f"[{timestamp}] {message}\n")

        # Älteste Zeilen entfernen, damit das Fenster nicht unbegrenzt wächst
        line_count = int(self.debug_output.index('end-1c').split('.')[0]) - 1
        if line_count > MAX_DEBUG_LINES:
            self.debug_output.delete('1.0', f'{line_count - MAX_DEBUG_LINES + 1}.0')

        self.debug_output.see(tk.END)

    def set_connected_state(self, connected):
        """Aktiviert bzw. deaktiviert die Verbinden-/Trennen-Buttons"""
        self.connect_button.configure(state="disabled" if connected else "normal")
        self.disconnect_button.configure(state="normal" if connected else "disabled")

    def validate_connection_params(self):
        """Validiert die Verbindungsparameter"""
        try:
            port = self.port_combo.get()
            if port == "Keine Ports gefunden" or not port:
                raise ValueError("Kein COM-Port ausgewählt oder verfügbar")

            baud_rate_str = self.baud_rate_combo.get()
            if not re.match(r'^\d+$', baud_rate_str):
                raise ValueError(f"Ungültige Baudrate: {baud_rate_str}")
            baud_rate = int(baud_rate_str)

            data_bits_str = self.data_bits_combo.get()
            if not re.match(r'^[678]$', data_bits_str):
                raise ValueError(f"Ungültige Datenbits: {data_bits_str}")
            data_bits = int(data_bits_str)

            parity = self.parity_combo.get()
            if parity not in ["None", "Odd", "Even"]:
                raise ValueError(f"Ungültige Parität: {parity}")

            return port, baud_rate, data_bits, parity
        except ValueError as e:
            messagebox.showerror("Validierungsfehler", str(e))
            self.log_debug(f"Validierungsfehler: {str(e)}")
            return None

    def connect_serial(self):
        # Verbindungsparameter validieren
        params = self.validate_connection_params()
        if not params:
            return

        port, baud_rate, data_bits, parity_str = params
        parity_map = {"None": serial.PARITY_NONE, "Odd": serial.PARITY_ODD, "Even": serial.PARITY_EVEN}
        parity = parity_map.get(parity_str, serial.PARITY_NONE)

        # Bestehende Verbindung schließen, falls vorhanden
        self.disconnect_serial()

        try:
            # Serielle Verbindung öffnen (wird in disconnect_serial wieder geschlossen)
            self.serial_connection = serial.Serial(
                port=port,
                baudrate=baud_rate,
                bytesize=data_bits,
                parity=parity,
                timeout=0.1
            )

            if self.serial_connection.is_open:
                self.read_active = True
                self.log_debug(f"Verbunden mit {port} bei {baud_rate} Baud")
                self.status_var.set(f"Verbunden mit {port}")
                self.set_connected_state(True)
                self.last_update_time = time.time()  # Timer zurücksetzen
                self.schedule_read()
            else:
                self.log_debug(f"Verbindung zu {port} konnte nicht geöffnet werden")
        except serial.SerialException as e:
            messagebox.showerror("Verbindungsfehler", f"Serieller Port-Fehler: {str(e)}")
            self.log_debug(f"Serieller Port-Fehler: {str(e)}")
        except Exception as e:
            messagebox.showerror("Verbindungsfehler", f"Unerwarteter Fehler: {str(e)}")
            self.log_debug(f"Unerwarteter Fehler beim Verbinden: {str(e)}")

    def disconnect_serial(self):
        """Trennt die serielle Verbindung und gibt Ressourcen frei"""
        self.read_active = False

        # Geplanten Lese-Aufruf abbrechen, damit keine zweite Leseschleife entsteht
        if self.read_after_id is not None:
            self.root.after_cancel(self.read_after_id)
            self.read_after_id = None

        connection = self.serial_connection
        self.serial_connection = None
        self.set_connected_state(False)

        if connection is None:
            return

        try:
            if connection.is_open:
                connection.close()
                self.log_debug("Verbindung getrennt.")
                self.status_var.set("Verbindung getrennt")
        except Exception as e:
            self.log_debug(f"Fehler beim Trennen der Verbindung: {str(e)}")

    def schedule_read(self):
        """Plant den nächsten Lesezyklus ein"""
        self.read_after_id = self.root.after(100, self.read_serial_data)

    def read_serial_data(self):
        self.read_after_id = None
        if not self.read_active:
            return

        if not self.serial_connection or not self.serial_connection.is_open:
            self.log_debug("Serielle Verbindung nicht mehr verfügbar")
            self.disconnect_serial()
            return

        try:
            if self.serial_connection.in_waiting:
                data = self.serial_connection.read(self.serial_connection.in_waiting)
                if data:
                    # Der SmlStreamReader puffert die Daten selbst, daher nur neue Bytes übergeben
                    self.sml_reader.add(data)
                    self.process_frames()
        except serial.SerialException as e:
            self.log_debug(f"Serielle Verbindung unterbrochen: {str(e)}")
            self.disconnect_serial()
            return
        except Exception as e:
            self.log_debug(f"Fehler beim Lesen der Daten: {str(e)}")

        if self.read_active:
            self.schedule_read()

    def process_frames(self):
        """Verarbeitet alle vollständigen SML-Frames im Puffer"""
        while True:
            try:
                sml_frame = self.sml_reader.get_frame()
            except Exception as e:
                # z.B. CRC-Fehler: der fehlerhafte Frame wurde bereits aus dem Puffer entfernt
                self.log_debug(f"Fehlerhafter SML-Frame verworfen: {str(e)}")
                continue

            if sml_frame is None:
                return

            try:
                parsed_msgs = sml_frame.parse_frame()

                # Validierung der SML-Daten
                if not parsed_msgs:
                    self.log_debug("Warnung: Leere SML-Nachricht empfangen")
                elif len(parsed_msgs) <= 1:
                    self.log_debug("Warnung: Unvollständige SML-Nachricht empfangen")
                else:
                    # Gültige Nachricht verarbeiten
                    obis_values = parsed_msgs[1].message_body.val_list
                    current_time = time.time()
                    self.add_to_history(current_time, obis_values)
                    self.log_debug(f"Frame empfangen, {len(obis_values)} OBIS Werte.")

                    # Aktualisiere die Anzeige, wenn das Zeitintervall erreicht wurde
                    if current_time - self.last_update_time >= self.update_interval:
                        self.display_obis_values(obis_values)
                        self.last_update_time = current_time
            except Exception as e:
                self.log_debug(f"Fehler beim Parsen des SML-Frames: {str(e)}")

    def add_to_history(self, timestamp, obis_values):
        """Speichert Messwerte und entfernt Einträge, die älter als SAVE_WINDOW_SECONDS sind"""
        self.measurement_history.append((timestamp, obis_values))
        while self.measurement_history and timestamp - self.measurement_history[0][0] > SAVE_WINDOW_SECONDS:
            self.measurement_history.popleft()

    @staticmethod
    def format_value(entry):
        """Liefert den skalierten Wert eines OBIS-Eintrags als Text"""
        value = entry.get_value()  # berücksichtigt Scaler, Texte und None
        if value is None:
            return ""
        if isinstance(value, float):
            return f"{value:.2f}"
        return str(value)

    def display_obis_values(self, obis_values):
        if not obis_values:
            self.log_debug("Warnung: Keine OBIS-Werte zum Anzeigen")
            return

        # Vorhandene Zeilen werden anhand des OBIS-Codes aktualisiert statt neu angelegt,
        # damit Scroll-Position und Auswahl erhalten bleiben
        seen = set()
        for entry in obis_values:
            try:
                # Validierung der OBIS-Einträge
                if not hasattr(entry, 'obis') or not hasattr(entry.obis, 'obis_short'):
                    self.log_debug(f"Ungültiger OBIS-Eintrag: {entry}")
                    continue

                obis_short = entry.obis.obis_short
                unit = const.UNITS.get(entry.unit, "")
                values = (obis_short, self.format_value(entry), unit)
                tags = ('highlight',) if obis_short in HIGHLIGHT_OBIS else ()

                if self.table.exists(obis_short):
                    self.table.item(obis_short, values=values, tags=tags)
                else:
                    self.table.insert('', 'end', iid=obis_short, values=values, tags=tags)
                seen.add(obis_short)
            except Exception as e:
                self.log_debug(f"Fehler beim Anzeigen von OBIS-Wert: {str(e)}")

        # Zeilen entfernen, die im aktuellen Frame nicht mehr vorkommen
        for item in self.table.get_children():
            if item not in seen:
                self.table.delete(item)

    def save_last_minute_data(self):
        try:
            # Nur Messwerte der letzten SAVE_WINDOW_SECONDS Sekunden speichern
            cutoff = time.time() - SAVE_WINDOW_SECONDS
            history = [(t, values) for t, values in self.measurement_history if t >= cutoff]
            if not history:
                messagebox.showwarning("Keine Daten", "Keine Daten zum Speichern verfügbar!")
                return

            # TODO: CSV-Export ist geplant, derzeit wird nur als TXT gespeichert
            file_path = filedialog.asksaveasfilename(
                title="Daten speichern (TXT)",
                defaultextension=".txt",
                filetypes=[("Textdateien", "*.txt"), ("Alle Dateien", "*.*")]
            )

            if not file_path:
                return  # Benutzer hat Abbrechen gedrückt

            with open(file_path, "w", encoding="utf-8") as file:
                file.write(f"Smart Meter Daten der letzten {SAVE_WINDOW_SECONDS} Sekunden:\n\n")
                for idx, (time_val, obis_values) in enumerate(history):
                    time_str = time.strftime("%Y-%m-%d %H:%M:%S", time.localtime(time_val))
                    file.write(f"Zeitpunkt {idx + 1} ({time_str}):\n")

                    if not obis_values:
                        file.write("  Keine OBIS-Werte verfügbar\n\n")
                        continue

                    for entry in obis_values:
                        try:
                            if not hasattr(entry, 'obis') or not hasattr(entry.obis, 'obis_short'):
                                continue

                            unit = const.UNITS.get(entry.unit, "")
                            file.write(f"  {entry.obis.obis_short}: {entry.get_value()} {unit}\n")
                        except Exception as e:
                            file.write(f"  Fehler beim Schreiben eines OBIS-Wertes: {str(e)}\n")
                    file.write("\n")
            messagebox.showinfo("Gespeichert", "Daten erfolgreich gespeichert!")
        except PermissionError:
            messagebox.showerror("Zugriffsfehler", "Keine Berechtigung zum Schreiben der Datei!")
        except Exception as e:
            messagebox.showerror("Fehler", f"Fehler beim Speichern: {str(e)}")

    def on_closing(self):
        """Wird aufgerufen, wenn das Fenster geschlossen wird"""
        try:
            self.disconnect_serial()
        finally:
            self.root.destroy()

def main():
    root = tk.Tk()
    app = SmartMeterReader(root)
    root.protocol("WM_DELETE_WINDOW", app.on_closing)
    root.mainloop()

if __name__ == "__main__":
    main()
