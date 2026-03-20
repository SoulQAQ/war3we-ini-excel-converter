# War3 Object Data Converter

[English](README.md) | [中文](README.zh-CN.md)

A Python-based tool for converting Warcraft III object data between `INI` files and `Excel` spreadsheets, making object data easier to edit and maintain.

## Features

- Convert Warcraft III object data from `INI` to `Excel`
- Convert edited `Excel` files back to `INI`
- Support for structured object data editing with better spreadsheet workflows
- Optional integration with [`w3x2lni`](https://github.com/sumneko/w3x2lni) when its installation directory is specified in the tool, enabling related `.w3x` conversion features

## Project Goal

This project is designed to help Warcraft III map creators and data editors manage object data more efficiently.

Typical workflow:

1. Export map object data to `LNI` or `INI`
2. Edit the data in `Excel`
3. Import the modified data back into the Warcraft III format
4. If [`w3x2lni`](https://github.com/sumneko/w3x2lni) is installed and configured in the tool, use it to support `.w3x`-related conversion workflows

## Current Status

This repository is still under active development.

- `INI` → `Excel` conversion is available
- `Excel` → `INI` conversion is planned / under development
- GUI support is included in [`script/gui.py`](script/gui.py)
- `w3x2lni` support depends on the installation directory configured in the tool

## Repository Structure

- [`script/`](script/)
  - Core Python conversion scripts
- [`config/`](config/)
  - Application settings
- [`webui/`](webui/)
  - Frontend files used by the desktop GUI
- [`build.bat`](build.bat)
  - Build packaged Windows executable with PyInstaller
- [`setup.bat`](setup.bat)
  - Create Python virtual environment and install dependencies
- [`start.bat`](start.bat)
  - Start the GUI application in development/local mode

## Usage

### 1. Convert `INI` to `Excel`

Run the converter script in [`script/ini_to_excel.py`](script/ini_to_excel.py).

### 2. Convert `Excel` to `INI`

Run the reverse converter in [`script/excel_to_ini.py`](script/excel_to_ini.py).

### 3. GUI mode

Launch the graphical interface in [`script/gui.py`](script/gui.py).

### 4. `w3x2lni`-assisted workflows

When you have installed [`w3x2lni`](https://github.com/sumneko/w3x2lni) and specified its installation directory in the tool, the project can use it for related `.w3x` conversion functions.

## Build From Source

The current project is intended for Windows and already includes helper batch files for environment setup, local startup, and packaging.

### Requirements

- Windows 10/11
- Python `3.12`
- Python Launcher for Windows (`py` command available in terminal)
- Internet connection for installing dependencies from [`requirements.txt`](requirements.txt)

[`setup.bat`](setup.bat) explicitly checks for Python `3.12`, creates `.venv`, and installs the required packages:

- `pyyaml`
- `openpyxl`
- `pywebview`

### 1. Clone or download the project

Clone this repository or download and extract the source code to a local folder.

### 2. Prepare the Python environment

From the project root, run:

```bat
setup.bat
```

This script will:

1. Check whether the `py` launcher exists
2. Verify that Python `3.12` is installed
3. Create or rebuild [`.venv/`](.venv/) when necessary
4. Upgrade `pip`, `setuptools`, and `wheel`
5. Install dependencies from [`requirements.txt`](requirements.txt)

### 3. Start the program locally

To launch the desktop GUI directly from source, run:

```bat
start.bat
```

[`start.bat`](start.bat) automatically re-runs [`setup.bat`](setup.bat) if the virtual environment is missing or if [`requirements.txt`](requirements.txt) has changed.

If you prefer to start it manually, use:

```bat
.venv\Scripts\python.exe script\gui.py
```

### 4. Build the executable

To package the application as a standalone Windows executable, run:

```bat
build.bat
```

[`build.bat`](build.bat) will:

1. Call [`setup.bat`](setup.bat)
2. Install `pyinstaller` into the virtual environment
3. Clean previous `build/` and `dist/` output
4. Package [`script/gui.py`](script/gui.py) into a one-file windowed executable
5. Include [`webui/`](webui/) and [`config/`](config/) in the packaged app

The generated executable will be placed at:

- [`dist/INI-Excel转换工具.exe`](dist/INI-Excel转换工具.exe)

### 5. Runtime files and configuration

- Main configuration file: [`config/setting.yaml`](config/setting.yaml)
- Default desktop UI entry: [`webui/index.html`](webui/index.html)
- The program defaults to using relative paths such as `./rundata/input` and `./rundata/output` from [`script/gui.py`](script/gui.py)

If these folders do not exist yet, create them manually before testing file conversions.

## Notes

- This project focuses on Warcraft III object data editing workflows.
- `w3x2lni`-related features are available only when the tool is configured with the correct installation path.
- Generated files are stored in [`rundata/output/`](rundata/output/).

## License

This project is licensed under the [GPL-3.0 License](https://www.gnu.org/licenses/gpl-3.0.en.html).
