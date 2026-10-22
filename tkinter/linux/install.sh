#!/bin/sh
# Installiert den Smart Meter Ausleser für den aktuellen Benutzer (ohne root):
# - Programm nach ~/.local/bin
# - Startmenü-Eintrag (KDE, GNOME, XFCE, ...) nach ~/.local/share/applications
# Aufruf im entpackten Ordner:  sh install.sh
set -e

SRC_DIR=$(cd "$(dirname "$0")" && pwd)
BIN_DIR="$HOME/.local/bin"
APP_DIR="$HOME/.local/share/applications"

if [ ! -f "$SRC_DIR/Meter_reader" ]; then
    echo "Fehler: Meter_reader nicht gefunden in $SRC_DIR" >&2
    exit 1
fi

mkdir -p "$BIN_DIR" "$APP_DIR"
install -m 755 "$SRC_DIR/Meter_reader" "$BIN_DIR/Meter_reader"

cat > "$APP_DIR/meter_reader.desktop" <<DESKTOP
[Desktop Entry]
Type=Application
Name=Smart Meter Ausleser
Comment=OBIS-Werte eines Smart Meters über einen optischen Lesekopf auslesen
Exec=$BIN_DIR/Meter_reader
Icon=utilities-system-monitor
Terminal=false
Categories=Utility;
DESKTOP

echo "Installiert: $BIN_DIR/Meter_reader"
echo "Startmenü-Eintrag: Smart Meter Ausleser"

# Zugriff auf serielle Ports (/dev/ttyUSB*) erfordert unter Linux meist die Gruppe "dialout"
if ! id -nG | grep -qw dialout; then
    echo
    echo "Hinweis: Für den Zugriff auf den Lesekopf den Benutzer zur Gruppe dialout hinzufügen:"
    echo "  sudo usermod -aG dialout $(id -un)"
    echo "Danach ab- und wieder anmelden."
fi
