# CredGen Spreadsheet Editor

A PyQt-based GUI application for editing CredGen spreadsheet files (CSV) with enhanced styling support, designed to streamline the creation and management of credits tables for media projects.

## Overview

This application provides a user-friendly interface for editing credits spreadsheets, with special features for handling styling nodes (such as page, content, and letter styles) defined in a TOML metadata file. It is tailored for workflows similar to those used in the cinecred/CredGen ecosystem.

## Key Features

- **Open/Save CSV Projects:**
  - Load and save credits spreadsheets in CSV format.
  - Project structure is based on a main CSV file (e.g., `Credits.csv`) and a styling metadata file (e.g., `Styling.toml`).

- **Styling Node Integration:**
  - Reads available style nodes (page, content, letter) from `Styling.toml`.
  - Populates dropdown menus in the spreadsheet for relevant columns, reducing user error and improving workflow.

- **Cell Reordering:**
  - Allows users to reorder selected cells via a dialog, making it easy to adjust credits order.

- **Row Management:**
  - Add or delete rows directly from the UI.

- **Menu-Driven Actions:**
  - All major actions (Open, Save, Exit, Undo, Redo, Cut, Copy, Paste, Reorder, Refresh Styling, About) are accessible from the menu bar.

- **Info Panel:**
  - Displays available page, content, and letter styles for reference.

## Application Structure

- **mainwindow.py**
  - Implements the `CredGenMainWindow` class, which manages the main window, menu bar, status bar, and info panel.
  - Handles file operations, styling data loading, and connects menu actions to spreadsheet widget methods.

- **spreadsheet_widget.py**
  - Implements the `SpreadsheetWidget` class, which provides the main spreadsheet editing interface.
  - Handles cell editing, dropdowns for styling columns, row management, and cell reordering.

- **styling_parser.py**
  - Contains the `StylingParser` class, which parses `Styling.toml` and extracts lists of available style nodes for use in the UI.

- **file_manager.py**
  - Contains the `FileManager` class, which loads and saves CSV files as lists of lists.

- **main.py**
  - Entry point for the application. Sets up the QApplication and launches the main window.

- **asset/**
  - Contains example project files, including `Credits.csv` (spreadsheet data) and `Styling.toml` (styling metadata).

## How It Works

1. **Startup:**
   - The application loads styling data from `asset/Styling.toml` (if available) and displays available styles in the info panel.

2. **Opening a Project:**
   - The user selects a project directory containing a CSV file (and optionally a TOML styling file).
   - The CSV is loaded into the spreadsheet widget; styling data is loaded and used to populate dropdowns in relevant columns.

3. **Editing:**
   - Users can edit cell values, select styles from dropdowns, add/delete rows, and reorder selected cells.
   - All changes are reflected in the UI and can be saved back to CSV.

4. **Saving:**
   - The current spreadsheet data is saved to the selected CSV file.

5. **Styling Refresh:**
   - Users can refresh styling data from the TOML file at any time, updating dropdowns and the info panel.

## Customization & Extensibility

- The codebase is modular, making it easy to add support for new file formats, additional styling features, or more advanced spreadsheet operations.
- Styling logic is isolated in `styling_parser.py` for easy adaptation to new TOML schemas.

## Requirements

- Python 3.x
- PyQt5
- toml (Python TOML parser)

## Usage

```sh
pip install pyqt5 toml
python main.py
```

## License

This project is provided as-is for educational and workflow enhancement purposes.
