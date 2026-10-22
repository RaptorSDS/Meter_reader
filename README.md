# Meter Reader

[![Build EXE](https://github.com/RaptorSDS/Meter_reader/actions/workflows/build-exe.yml/badge.svg)](https://github.com/RaptorSDS/Meter_reader/actions/workflows/build-exe.yml)
[![License: GPL v3](https://img.shields.io/badge/license-GPLv3-blue.svg)](LICENSE)

Reads all OBIS values from a smart meter (SML protocol) over an optical USB-to-serial adapter and shows them in a simple desktop window.

## Features

- Reads **all OBIS values**, including scale and unit
- **No UAC / admin rights** needed
- Table highlights the most used values for electricity meters:

  | OBIS code | Meaning |
  |---|---|
  | `1.8.0` | Energy consumed (import) |
  | `2.8.0` | Energy delivered (export) |
  | `C.1.0` | Meter serial number |
  | `0.2.0` | Firmware version |

- Debug view of the communication
- Saves the measurements of the last 60 seconds to a text file

## Compatibility

- Windows 10/11 (tested)
- Linux (x86_64) and macOS (Apple Silicon): builds are provided, not yet tested on real hardware
- HICHI USB adapter (should work with all USB-to-serial "optical" adapters)
- Default settings for SML meters: 9600 baud, 8 data bits, no parity, 1 stop bit

## Download

The tkinter version is built automatically by GitHub Actions. Open the [Releases page](https://github.com/RaptorSDS/Meter_reader/releases) and download the file for your system from the latest release:

| System | File | How to start |
|---|---|---|
| Windows 10/11 | `Meter_reader.exe` | Start it, no installation needed |
| Linux (x86_64) | `Meter_reader-linux-x86_64.tar.gz` | Unpack and run `./Meter_reader` |
| macOS (Apple Silicon) | `Meter_reader-macos-arm64.zip` | Unpack and open `Meter_reader.app` |

### Linux

```
tar -xzf Meter_reader-linux-x86_64.tar.gz
cd Meter_reader
./Meter_reader
```

To add the program to the start menu (KDE, GNOME, XFCE, ...), run `sh install.sh` in the unpacked folder. It installs for the current user only, no root needed.

Access to the optical adapter (`/dev/ttyUSB*`) usually requires the group `dialout`:

```
sudo usermod -aG dialout $USER
```

Log out and back in afterwards.

### macOS

The app is not signed by Apple. On the first start, macOS blocks it:

1. Right-click `Meter_reader.app` and choose **Open**, then confirm with **Open**.
2. If that does not help, run once in the Terminal: `xattr -dr com.apple.quarantine Meter_reader.app`

The adapter appears as a port like `/dev/cu.usbserial-…`.

> **Note:** Versions marked *beta* are test versions. Please report problems as an [issue](https://github.com/RaptorSDS/Meter_reader/issues).

## Versions in this repository

| Folder | GUI | File |
|---|---|---|
| `tkinter/` | Tkinter (built into Python), used for the EXE | `Meter_reader.py` |
| `Qt5/` | PyQt5 | `Meter_reader.py` |
| `Qt6/` | PyQt6 | `Meter_reader.py` |
| `archiv/` | Older versions and sources | |

Each folder also contains `Meter_reader_archiv.py`, the previous version of the program.

## Build it yourself

### Option 1: auto-py-to-exe (graphical)

1. Install Python 3.12 (e.g. from the Microsoft Store).
2. In a command prompt, install the dependencies:
   ```
   pip install pyserial smllib auto-py-to-exe
   ```
   For the Qt versions also install `pyqt6` or `pyqt5`. Tkinter is already included in Python 3.12.
3. Start the tool:
   ```
   python -m auto_py_to_exe
   ```
4. In the web interface, select the Python script (e.g. `tkinter\Meter_reader.py`).
5. Compile.
6. Use the resulting EXE.

### Option 2: PyInstaller (command line)

This is the same command the GitHub build uses:

```
cd tkinter
pip install -r requirements.txt pyinstaller
pyinstaller --onefile --windowed --name Meter_reader --collect-submodules smllib Meter_reader.py
```

The EXE is created in `tkinter\dist\Meter_reader.exe`.

The same command works on Linux and creates `tkinter/dist/Meter_reader`. On macOS, leave out `--onefile` to get `tkinter/dist/Meter_reader.app`. PyInstaller always builds for the system it runs on, so each system has to be built on itself.

> **Tip:** Run the build from inside the `tkinter` folder. Started from the repository root, Python can mistake the folder named `tkinter` for the real `tkinter` module.

### Run without building

```
pip install pyserial smllib
python tkinter\Meter_reader.py
```

On Linux, tkinter may have to be installed separately (e.g. `sudo apt install python3-tk`).

## Credits and license

- Developed by Tobias Baumann ([RaptorSDS](https://github.com/RaptorSDS/))
- Uses [smllib](https://pypi.org/project/smllib/) by spaceman_spiff (GNU General Public License v3 or later, GPLv3+)
- This program is licensed under the [GNU General Public License](LICENSE).
