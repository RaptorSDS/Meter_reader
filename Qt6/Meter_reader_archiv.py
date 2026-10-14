import sys
import serial
import time
from PyQt6.QtWidgets import (QApplication, QMainWindow, QWidget, QVBoxLayout, QHBoxLayout, 
                            QLabel, QComboBox, QPushButton, QTextEdit, QTreeWidget, QTreeWidgetItem,
                            QGroupBox, QScrollArea, QFileDialog, QMessageBox, QStatusBar)
from PyQt6.QtCore import QTimer, Qt
from smllib import SmlStreamReader, const
from collections import deque
from serial.tools import list_ports
import re  # Für Eingabevalidierung

class SmartMeterReader(QMainWindow):
    def __init__(self):
        super().__init__()
        self.setWindowTitle("Smart Meter Ausleser")
        self.setMinimumSize(800, 600)
        
        self.serial_connection = None
        self.read_active = False
        self.sml_reader = SmlStreamReader()
        self.raw_buffer = bytearray()
        # Speichert Tupel (Zeitstempel, obis_values) für die Historie
        self.measurement_history = deque(maxlen=60)
        self.update_interval = 1  # in Sekunden
        self.last_update_time = time.time()  # Zeitbasierte Aktualisierung
        
        self.init_ui()
        
        # Timer für das Lesen der seriellen Daten
        self.read_timer = QTimer()
        self.read_timer.timeout.connect(self.read_serial_data)
        self.read_timer.setInterval(100)  # 100ms
    
    def init_ui(self):
        # Hauptwidget und Layout
        central_widget = QWidget()
        self.setCentralWidget(central_widget)
        main_layout = QHBoxLayout(central_widget)
        
        # Linkes Panel
        left_panel = QWidget()
        left_layout = QVBoxLayout(left_panel)
        left_panel.setMaximumWidth(300)
        
        # Verbindungseinstellungen
        self.create_connection_group(left_layout)
        
        # Button-Leiste
        button_layout = QHBoxLayout()
        save_button = QPushButton("Daten speichern")
        save_button.clicked.connect(self.save_last_minute_data)
        info_button = QPushButton("?")
        info_button.setMaximumWidth(30)
        info_button.clicked.connect(self.show_info)
        
        button_layout.addWidget(save_button)
        button_layout.addWidget(info_button)
        left_layout.addLayout(button_layout)
        
        # Debug-Anzeige
        self.create_debug_group(left_layout)
        
        # Rechtes Panel mit OBIS-Tabelle
        self.create_table_group()
        
        # Layouts zusammenfügen
        main_layout.addWidget(left_panel)
        main_layout.addWidget(self.table_group)
        
        # Statusleiste
        self.statusBar = QStatusBar()
        self.setStatusBar(self.statusBar)
        self.statusBar.showMessage("Bereit")
    
    def create_connection_group(self, parent_layout):
        group_box = QGroupBox("Verbindungseinstellungen")
        layout = QVBoxLayout(group_box)
        
        # Port
        layout.addWidget(QLabel("Port:"))
        self.port_combo = QComboBox()
        self.port_combo.addItems(self.get_serial_ports())
        layout.addWidget(self.port_combo)
        
        # Refresh-Button für Ports
        refresh_button = QPushButton("Ports aktualisieren")
        refresh_button.clicked.connect(self.refresh_ports)
        layout.addWidget(refresh_button)
        
        # Baudrate
        layout.addWidget(QLabel("Baudrate:"))
        self.baud_rate_combo = QComboBox()
        self.baud_rate_combo.addItems(["9600", "19200", "38400", "57600", "115200"])
        self.baud_rate_combo.setCurrentText("9600")
        layout.addWidget(self.baud_rate_combo)
        
        # Datenbits
        layout.addWidget(QLabel("Datenbits:"))
        self.data_bits_combo = QComboBox()
        self.data_bits_combo.addItems(["6", "7", "8"])
        self.data_bits_combo.setCurrentText("8")
        layout.addWidget(self.data_bits_combo)
        
        # Parität
        layout.addWidget(QLabel("Parität:"))
        self.parity_combo = QComboBox()
        self.parity_combo.addItems(["None", "Odd", "Even"])
        self.parity_combo.setCurrentText("None")
        layout.addWidget(self.parity_combo)
        
        # Aktualisierungsintervall
        layout.addWidget(QLabel("Aktualisierungsintervall:"))
        self.update_interval_combo = QComboBox()
        self.update_interval_combo.addItems(["1s", "5s", "15s"])
        self.update_interval_combo.setCurrentText("1s")
        self.update_interval_combo.currentTextChanged.connect(self.set_update_interval)
        layout.addWidget(self.update_interval_combo)
        
        # Verbinden/Trennen Buttons
        connect_button = QPushButton("Verbinden")
        connect_button.clicked.connect(self.connect_serial)
        layout.addWidget(connect_button)
        
        disconnect_button = QPushButton("Trennen")
        disconnect_button.clicked.connect(self.disconnect_serial)
        layout.addWidget(disconnect_button)
        
        parent_layout.addWidget(group_box)
    
    def create_debug_group(self, parent_layout):
        group_box = QGroupBox("Debug-Anzeige")
        layout = QVBoxLayout(group_box)
        
        self.debug_output = QTextEdit()
        self.debug_output.setReadOnly(True)
        layout.addWidget(self.debug_output)
        
        parent_layout.addWidget(group_box, 1)  # 1 = stretch factor
    
    def create_table_group(self):
        self.table_group = QGroupBox("OBIS-Werte")
        layout = QVBoxLayout(self.table_group)
        
        self.table = QTreeWidget()
        self.table.setHeaderLabels(["OBIS-Code", "Wert", "Einheit"])
        self.table.setColumnWidth(0, 150)
        self.table.setColumnWidth(1, 150)
        
        layout.addWidget(self.table)
    
    def get_serial_ports(self):
        # Nutzt list_ports aus pySerial für eine zuverlässige Ermittlung
        ports = [port.device for port in list_ports.comports()]
        return ports if ports else ["Keine Ports gefunden"]
    
    def refresh_ports(self):
        """Aktualisiert die Liste der verfügbaren seriellen Ports"""
        ports = self.get_serial_ports()
        self.port_combo.clear()
        self.port_combo.addItems(ports)
        self.log_debug("Port-Liste aktualisiert")
    
    def set_update_interval(self):
        interval_text = self.update_interval_combo.currentText()
        try:
            # Entfernt 's' und konvertiert zu Integer
            if re.match(r'^\d+s$', interval_text):
                self.update_interval = int(interval_text.replace("s", ""))
            else:
                raise ValueError("Ungültiges Intervallformat")
        except ValueError:
            self.log_debug(f"Ungültiges Intervallformat: {interval_text}, setze auf 1s")
            self.update_interval = 1  # Fallback
            self.update_interval_combo.setCurrentText("1s")
        
    def show_info(self):
        info_text = (
            "Smart Meter Ausleser\n\n"
            "Entwickelt von Tobias Baumann\n"
            "GitHub: https://github.com/RaptorSDS/\n"
            "Hilfe durch Claude AI 3.5 / ChatGPT 3o_high\n\n"
            "Dieses Programm steht unter der GNU General Public License (GPL).\n\n"
            "Verwendete Bibliotheken:\n"
            "- smllib von spaceman_spiff (GNU License)\n"
            "- PyQt6 für die Benutzeroberfläche\n\n"
            "Copyright (C) 2024 Tobias Baumann"
        )
        QMessageBox.information(self, "Info", info_text)
    
    def log_debug(self, message):
        """Schreibt Debug-Informationen in das Debug-Fenster."""
        timestamp = time.strftime("%H:%M:%S", time.localtime())
        self.debug_output.append(f"[{timestamp}] {message}")
        # Scrolle zum Ende
        scrollbar = self.debug_output.verticalScrollBar()
        scrollbar.setValue(scrollbar.maximum())
    
    def validate_connection_params(self):
        """Validiert die Verbindungsparameter"""
        try:
            port = self.port_combo.currentText()
            if port == "Keine Ports gefunden" or not port:
                raise ValueError("Kein COM-Port ausgewählt oder verfügbar")
            
            baud_rate_str = self.baud_rate_combo.currentText()
            if not re.match(r'^\d+$', baud_rate_str):
                raise ValueError(f"Ungültige Baudrate: {baud_rate_str}")
            baud_rate = int(baud_rate_str)
            
            data_bits_str = self.data_bits_combo.currentText()
            if not re.match(r'^[678]$', data_bits_str):
                raise ValueError(f"Ungültige Datenbits: {data_bits_str}")
            data_bits = int(data_bits_str)
            
            parity = self.parity_combo.currentText()
            if parity not in ["None", "Odd", "Even"]:
                raise ValueError(f"Ungültige Parität: {parity}")
            
            return port, baud_rate, data_bits, parity
        except ValueError as e:
            QMessageBox.critical(self, "Validierungsfehler", str(e))
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
            # Verbindung mit Context Manager öffnen
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
                self.statusBar.showMessage(f"Verbunden mit {port}")
                self.last_update_time = time.time()  # Timer zurücksetzen
                self.read_timer.start()
            else:
                self.log_debug(f"Verbindung zu {port} konnte nicht geöffnet werden")
        except serial.SerialException as e:
            QMessageBox.critical(self, "Verbindungsfehler", f"Serieller Port-Fehler: {str(e)}")
            self.log_debug(f"Serieller Port-Fehler: {str(e)}")
        except Exception as e:
            QMessageBox.critical(self, "Verbindungsfehler", f"Unerwarteter Fehler: {str(e)}")
            self.log_debug(f"Unerwarteter Fehler beim Verbinden: {str(e)}")
    
    def disconnect_serial(self):
        """Trennt die serielle Verbindung und gibt Ressourcen frei"""
        try:
            self.read_active = False
            self.read_timer.stop()
            if self.serial_connection and self.serial_connection.is_open:
                self.serial_connection.close()
                self.log_debug("Verbindung getrennt.")
                self.statusBar.showMessage("Verbindung getrennt")
                self.serial_connection = None
        except Exception as e:
            self.log_debug(f"Fehler beim Trennen der Verbindung: {str(e)}")
    
    def read_serial_data(self):
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
                    self.raw_buffer.extend(data)
                    self.sml_reader.add(self.raw_buffer)
                    sml_frame = self.sml_reader.get_frame()
                    
                    if sml_frame:
                        self.raw_buffer.clear()
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
                                # Speichern mit Zeitstempel
                                self.measurement_history.append((current_time, obis_values))
                                self.log_debug(f"Frame empfangen, {len(obis_values)} OBIS Werte.")
                                
                                # Aktualisiere die Anzeige, wenn das Zeitintervall erreicht wurde
                                if current_time - self.last_update_time >= self.update_interval:
                                    self.display_obis_values(obis_values)
                                    self.last_update_time = current_time
                        except Exception as e:
                            self.log_debug(f"Fehler beim Parsen des SML-Frames: {str(e)}")
        
        except serial.SerialException as e:
            self.log_debug(f"Serielle Verbindung unterbrochen: {str(e)}")
            self.disconnect_serial()
        except Exception as e:
            self.log_debug(f"Fehler beim Lesen der Daten: {str(e)}")
    
    def display_obis_values(self, obis_values):
        # Lösche alle bisherigen Einträge
        self.table.clear()
        
        if not obis_values:
            self.log_debug("Warnung: Keine OBIS-Werte zum Anzeigen")
            return
            
        for entry in obis_values:
            try:
                # Validierung der OBIS-Einträge
                if not hasattr(entry, 'obis') or not hasattr(entry.obis, 'obis_short'):
                    self.log_debug(f"Ungültiger OBIS-Eintrag: {entry}")
                    continue
                    
                obis_short = entry.obis.obis_short
                unit = const.UNITS.get(entry.unit, "")
                scaler = entry.scaler if entry.scaler is not None else 0
                
                # Validierung des Wertes
                if not hasattr(entry, 'value'):
                    self.log_debug(f"OBIS-Eintrag ohne Wert: {obis_short}")
                    continue
                    
                scaled_value = entry.value * (10 ** scaler)
                # Formatierung: Falls es ein Fließkommawert ist, auf 2 Dezimalstellen runden
                formatted_value = f"{scaled_value:.2f}" if isinstance(scaled_value, float) else str(scaled_value)
                
                item = QTreeWidgetItem([obis_short, formatted_value, unit])
                
                # Wichtige Werte hervorheben
                if obis_short in ["1.8.0", "2.8.0", "C.1.0", "0.2.0"]:
                    # In PyQt6 müssen wir QColor verwenden statt Qt.yellow
                    from PyQt6.QtGui import QColor
                    for i in range(3):
                        item.setBackground(i, QColor(255, 255, 0))  # Gelb
                
                self.table.addTopLevelItem(item)
            except Exception as e:
                self.log_debug(f"Fehler beim Anzeigen von OBIS-Wert: {str(e)}")
        
        # Spaltenbreiten anpassen
        for i in range(3):
            self.table.resizeColumnToContents(i)
    
    def save_last_minute_data(self):
        try:
            if not self.measurement_history:
                QMessageBox.warning(self, "Keine Daten", "Keine Daten zum Speichern verfügbar!")
                return
            
            file_path, _ = QFileDialog.getSaveFileName(
                self,
                "Daten speichern",
                "",
                "Textdateien (*.txt);;Alle Dateien (*.*)"
            )
            
            if not file_path:
                return  # Benutzer hat Abbrechen gedrückt
                
            with open(file_path, "w", encoding="utf-8") as file:
                file.write("Smart Meter Daten der letzten 60 Sekunden:\n\n")
                for idx, (time_val, obis_values) in enumerate(self.measurement_history):
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
                            scaler = entry.scaler if entry.scaler is not None else 0
                            scaled_value = entry.value * (10 ** scaler)
                            file.write(f"  {entry.obis.obis_short}: {scaled_value} {unit}\n")
                        except Exception as e:
                            file.write(f"  Fehler beim Schreiben eines OBIS-Wertes: {str(e)}\n")
                    file.write("\n")
                QMessageBox.information(self, "Gespeichert", "Daten erfolgreich gespeichert!")
        except PermissionError:
            QMessageBox.critical(self, "Zugriffsfehler", "Keine Berechtigung zum Schreiben der Datei!")
        except Exception as e:
            QMessageBox.critical(self, "Fehler", f"Fehler beim Speichern: {str(e)}")
    
    def closeEvent(self, event):
        """Wird aufgerufen, wenn das Fenster geschlossen wird"""
        try:
            self.read_active = False
            self.read_timer.stop()
            self.disconnect_serial()
        finally:
            event.accept()

def main():
    app = QApplication(sys.argv)
    window = SmartMeterReader()
    window.show()
    sys.exit(app.exec())  # In PyQt6 ist es nur app.exec() ohne Klammern

if __name__ == "__main__":
    main()
