# Meter Reader

[![Build EXE](https://github.com/RaptorSDS/Meter_reader/actions/workflows/build-exe.yml/badge.svg)](https://github.com/RaptorSDS/Meter_reader/actions/workflows/build-exe.yml)
[![Release](https://img.shields.io/github/v/release/RaptorSDS/Meter_reader?include_prereleases&label=release)](https://github.com/RaptorSDS/Meter_reader/releases)
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
- HICHI USB adapter (should work with all USB-to-serial "optical" adapters)
- Default settings for SML meters: 9600 baud, 8 data bits, no parity, 1 stop bit

## Download (Windows EXE)

The tkinter version is built automatically by GitHub Actions as a single, small EXE (about 11 MB):

1. Open the [Releases page](https://github.com/RaptorSDS/Meter_reader/releases).
2. Download `Meter_reader.exe` from the latest release.
3. Start it, no installation needed.

> **Note:** Versions marked *beta* are pre-releases. Please report problems as an [issue](https://github.com/RaptorSDS/Meter_reader/issues).
>
> `meter_reader.exe` in the repository root is an older build. Use the release download for the current version.

## Versions in this repository

| Folder | GUI | File |
|---|---|---|
| `tkinter/` | Tkinter (built into Python), used for the EXE | `Meter_reader.py` |
| `Qt5/` | PyQt5 | `Meter_reader.py` |
| `Qt6/` | PyQt6 | `Meter_reader.py` |
| `archiv/` | Older versions and sources | |

Each folder also contains `Meter_reader_archiv.py`, the previous version of the program.

## Build the EXE yourself

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

> **Tip:** Run the build from inside the `tkinter` folder. Started from the repository root, Python can mistake the folder named `tkinter` for the real `tkinter` module.

### Run without building

```
pip install pyserial smllib
python tkinter\Meter_reader.py
```

## Releasing a new version (maintainers)

1. Change `VERSION` in `tkinter/Meter_reader.py` (e.g. `"1.2.0"` or `"1.2.0-beta"`) and push to `main`.
2. On GitHub, open **Actions → "EXE bauen (tkinter)" → "Run workflow"**, tick **"Release erstellen"** and start the run.
3. The workflow builds the EXE and publishes a release with the tag `v` + `VERSION`. Versions containing `beta` are marked as pre-releases.

## Credits and license

- Developed by Tobias Baumann ([RaptorSDS](https://github.com/RaptorSDS/))
- Uses [smllib](https://pypi.org/project/smllib/) by spaceman_spiff (GNU General Public License v3 or later, GPLv3+)
- This program is licensed under the [GNU General Public License](LICENSE).
