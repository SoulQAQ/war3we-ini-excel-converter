# War3 Object Data Converter

[English](README.md) | [中文](README.zh-CN.md)

A Python-based tool for converting Warcraft III `LNI` object data and map object data between `INI` files and `Excel` spreadsheets, making object data easier to edit and maintain.

## Features

- Convert Warcraft III object data from `INI` to `Excel`
- Convert edited `Excel` files back to `INI`
- Support for structured object data editing with better spreadsheet workflows
- Planned integration with [`w3x2lni`](https://github.com/sumneko/w3x2lni) for direct `.w3x` ↔ `Excel` conversion when the tool is available locally

## Project Goal

This project is designed to help Warcraft III map creators and data editors manage object data more efficiently.

Typical workflow:

1. Export map object data to `LNI` or `INI`
2. Edit the data in `Excel`
3. Import the modified data back into the Warcraft III format
4. Optionally use [`w3x2lni`](https://github.com/sumneko/w3x2lni) to convert `.w3x` maps directly for a smoother workflow

## Current Status

This repository is still under active development.

- `INI` → `Excel` conversion is available
- `Excel` → `INI` conversion is planned / under development
- GUI support is included in [`script/gui.py`](script/gui.py)

## Repository Structure

- [`script/`](script/)
  - Core Python conversion scripts
- [`config/`](config/)
  - Application settings
- [`example/`](example/)
  - Example `INI` data and sample assets
- [`rundata/`](rundata/)
  - Generated output files
- [`w3x2lni/`](w3x2lni/)
  - Bundled reference copy of the upstream [`w3x2lni`](https://github.com/sumneko/w3x2lni) project

## Usage

### 1. Convert `INI` to `Excel`

Run the converter script in [`script/ini_to_excel.py`](script/ini_to_excel.py).

### 2. Convert `Excel` to `INI`

Run the reverse converter in [`script/excel_to_ini.py`](script/excel_to_ini.py).

### 3. GUI mode

Launch the graphical interface in [`script/gui.py`](script/gui.py).

## Notes

- This project focuses on Warcraft III object data editing workflows.
- If you want to use `.w3x` direct conversion, install [`w3x2lni`](https://github.com/sumneko/w3x2lni) and place it in the expected location before using the planned integration.
- Generated files are stored in [`rundata/output/`](rundata/output/).

## License

This project is licensed under the [GPL-3.0 License](https://www.gnu.org/licenses/gpl-3.0.en.html).
