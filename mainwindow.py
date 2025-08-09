#!/usr/bin/env python3
"""
CredGen Spreadsheet Editor
A PyQt application for editing CredGen spreadsheet files with enhanced styling support.

Performance note: for large CSVs, consider virtualized table views in the future.
"""

import sys
import os
from pathlib import Path
from PyQt5.QtWidgets import (
    QApplication, QMainWindow, QVBoxLayout, QHBoxLayout, QWidget,
    QMenuBar, QToolBar, QStatusBar, QFileDialog, QMessageBox,
    QSplitter, QTabWidget, QPushButton, QLabel
)
from PyQt5.QtCore import Qt, pyqtSignal
from PyQt5.QtGui import QIcon, QKeySequence, QFont

# Import our custom modules
from spreadsheet_widget import SpreadsheetWidget
from styling_parser import StylingParser
from file_manager import FileManager
from widgets.info_panel import InfoPanel
from widgets.menu_bar import MenuBar
from controller import CredGenController


class CredGenMainWindow(QMainWindow):
    """
    Main application window for CredGen Spreadsheet Editor.
    Coordinates UI, delegates data logic to controller, and manages user interactions.
    Extensible via controller plugins for new style types/features.
    """
    
    def __init__(self):
        super().__init__()
        self.controller = CredGenController()
        self.file_manager = self.controller.file_manager
        self.styling_parser = self.controller.styling_parser
        self.current_csv_file = None
        self.current_styling_file = None
        self.styling_data = None
        
        self.init_ui()
        self.setup_connections()
        
        # Load default styling data if available
        default_styling_path = str(Path('asset/Styling.toml'))
        if Path(default_styling_path).exists():
            self.styling_data = self.styling_parser.parse_styling_file(default_styling_path)
            self.current_styling_file = default_styling_path
            self.update_info_panel(self.styling_data)
            self.spreadsheet_widget.update_styling_data(self.styling_data)
        
    def init_ui(self) -> None:
        """Initialize the user interface. (Accessibility: add tooltips, ensure tab order, and add keyboard shortcuts.)"""
        self.setWindowTitle("CredGen Spreadsheet Editor")
        self.setGeometry(100, 100, 1400, 800)
        self.setWindowIcon(QIcon("icons/app_icon.png"))
        central_widget = QWidget()
        self.setCentralWidget(central_widget)
        main_layout = QVBoxLayout(central_widget)
        self.create_menu_bar()
        self.create_main_content(main_layout)
        self.create_status_bar()
        # Accessibility: set tab order (spreadsheet first, then info panel)
        self.setTabOrder(self.spreadsheet_widget, self.info_panel)
        # Add tooltips to main widgets
        self.spreadsheet_widget.setToolTip("Edit credits spreadsheet here")
        self.info_panel.setToolTip("Reference for available styles")
        
    def create_menu_bar(self) -> None:
        """Create the application menu bar. (Accessibility: keyboard shortcuts are set in MenuBar.)"""
        self.menubar = MenuBar(self)
        self.setMenuBar(self.menubar)
        self.menubar.actions['open'].setToolTip("Open a CredGen project folder")
        self.menubar.actions['save'].setToolTip("Save the current project")
        self.menubar.actions['exit'].setToolTip("Exit the application")
        self.menubar.actions['undo'].setToolTip("Undo last action")
        self.menubar.actions['redo'].setToolTip("Redo last undone action")
        self.menubar.actions['cut'].setToolTip("Cut selected cells")
        self.menubar.actions['copy'].setToolTip("Copy selected cells")
        self.menubar.actions['paste'].setToolTip("Paste cells from clipboard")
        self.menubar.actions['reorder'].setToolTip("Reorder selected cells")
        self.menubar.actions['refresh_styling'].setToolTip("Refresh styling data from TOML")
        self.menubar.actions['about'].setToolTip("About this application")
        # Connect actions to methods
        self.menubar.actions['open'].triggered.connect(self.open_project)
        self.menubar.actions['save'].triggered.connect(self.save_project)
        self.menubar.actions['exit'].triggered.connect(self.close)
        self.menubar.actions['undo'].triggered.connect(self.undo)
        self.menubar.actions['redo'].triggered.connect(self.redo)
        self.menubar.actions['cut'].triggered.connect(self.cut)
        self.menubar.actions['copy'].triggered.connect(self.copy)
        self.menubar.actions['paste'].triggered.connect(self.paste)
        self.menubar.actions['reorder'].triggered.connect(self.reorder_cells)
        self.menubar.actions['refresh_styling'].triggered.connect(self.refresh_styling)
        self.menubar.actions['about'].triggered.connect(self.show_about)
        
    # def create_toolbar(self):
    #     """Create the application toolbar."""
    #     toolbar = self.addToolBar('Main Toolbar')
    #     toolbar.setToolButtonStyle(Qt.ToolButtonTextUnderIcon)
        
    #     # New project button
    #     new_btn = toolbar.addAction('New')
    #     new_btn.setShortcut(QKeySequence.New)
    #     new_btn.triggered.connect(self.new_project)
        
    #     # Open project button
    #     open_btn = toolbar.addAction('Open')
    #     open_btn.setShortcut(QKeySequence.Open)
    #     open_btn.triggered.connect(self.open_project)
        
    #     # Save button
    #     save_btn = toolbar.addAction('Save')
    #     save_btn.setShortcut(QKeySequence.Save)
    #     save_btn.triggered.connect(self.save_project)
        
    #     toolbar.addSeparator()
        
    #     # # Reorder button
    #     # reorder_btn = toolbar.addAction('Reorder')
    #     # reorder_btn.triggered.connect(self.reorder_cells)
        
    #     # # Refresh styling button
    #     # refresh_btn = toolbar.addAction('Refresh')
    #     # refresh_btn.triggered.connect(self.refresh_styling)
        
    def create_main_content(self, layout):
        """Create the main content area."""
        splitter = QSplitter(Qt.Horizontal)
        
        # Create spreadsheet widget
        self.spreadsheet_widget = SpreadsheetWidget()
        if self.styling_data:
            self.spreadsheet_widget.update_styling_data(self.styling_data)
            
        # Use InfoPanel widget
        self.info_panel = InfoPanel()
        
        # Add widgets to splitter
        splitter.addWidget(self.spreadsheet_widget)
        splitter.addWidget(self.info_panel)
        
        # Set initial sizes
        splitter.setSizes([800, 200])
        
        layout.addWidget(splitter)
        
    def create_status_bar(self):
        """Create the status bar."""
        self.status_bar = QStatusBar()
        self.setStatusBar(self.status_bar)
        self.status_bar.showMessage("Ready")
        
    def setup_connections(self):
        """Setup signal-slot connections."""
        self.spreadsheet_widget.data_changed.connect(self.on_data_changed)
        self.spreadsheet_widget.selection_changed.connect(self.on_selection_changed)
        
    def new_project(self):
        """Create a new project."""
        try:
            self.spreadsheet_widget.clear_data()
            self.current_csv_file = None
            self.current_styling_file = None
            
            # Create new empty project structure
            self.spreadsheet_widget.create_new_project()
            
            self.status_bar.showMessage("New project created")
            self.update_window_title()
            
            if self.styling_data:
                self.spreadsheet_widget.update_styling_data(self.styling_data)
                
        except Exception as e:
            QMessageBox.critical(self, "Error", f"Failed to create new project: {str(e)}")
            
    def open_project(self) -> None:
        """Open a project folder containing Credits.csv and Styling.toml files."""
        folder_path = QFileDialog.getExistingDirectory(
            self,
            "Open Project Folder",
            "",
            QFileDialog.ShowDirsOnly | QFileDialog.DontResolveSymlinks
        )
        if folder_path:
            folder = Path(folder_path)
            credits_file = folder / 'Credits.csv'
            styling_file = folder / 'Styling.toml'
            if not credits_file.exists():
                QMessageBox.critical(
                    self,
                    "Invalid Project Folder",
                    f"Credits.csv not found in {folder_path}"
                )
                return
            if not styling_file.exists():
                QMessageBox.critical(
                    self,
                    "Invalid Project Folder",
                    f"Styling.toml not found in {folder_path}"
                )
                return
            try:
                csv_data, styling_data = self.controller.load_project(str(credits_file), str(styling_file))
                self.current_csv_file = str(credits_file)
                self.current_styling_file = str(styling_file)
                self.styling_data = styling_data
                self.spreadsheet_widget.update_styling_data(styling_data)
                self.update_info_panel(styling_data)
                self.spreadsheet_widget.load_data(csv_data, styling_data)
                self.update_window_title()
                self.statusBar().showMessage(f"Loaded project from: {folder_path}")
            except ValueError as e:
                QMessageBox.critical(self, "Error", str(e))
                return
            except Exception as e:
                QMessageBox.critical(self, "Error", f"Failed to open project: {str(e)}")
                return

    def load_project(self, csv_file_path, styling_file_path=None):
        """Load a project from files."""
        try:
            # Load CSV data
            csv_data = self.file_manager.load_csv(csv_file_path)
            self.current_csv_file = csv_file_path
            
            # Load styling data if available
            styling_data = None
            if styling_file_path and Path(styling_file_path).exists():
                styling_data = self.styling_parser.parse_styling_file(styling_file_path)
                self.current_styling_file = styling_file_path
            elif self.styling_data:
                styling_data = self.styling_data
                
            # Update spreadsheet widget
            self.spreadsheet_widget.load_data(csv_data, styling_data)
            
            # Update info panel
            self.update_info_panel(styling_data)
            
            self.status_bar.showMessage(f"Loaded project: {os.path.basename(csv_file_path)}")
            self.update_window_title()
            
        except Exception as e:
            QMessageBox.critical(self, "Error", f"Failed to load project: {str(e)}")
            
    def save_project(self) -> None:
        """Save the current project."""
        if not self.current_csv_file:
            self.save_as_project()
            return
        try:
            csv_data = self.spreadsheet_widget.get_csv_data()
            if not self.controller.validate_csv(csv_data):
                QMessageBox.critical(self, "Error", "CSV data is invalid.")
                return
            self.controller.save_project(self.current_csv_file, csv_data)
            self.status_bar.showMessage("Project saved successfully")
        except Exception as e:
            QMessageBox.critical(self, "Error", f"Failed to save project: {str(e)}")
            
    def save_as_project(self) -> None:
        """Save the project with a new name."""
        try:
            file_path, _ = QFileDialog.getSaveFileName(
                self, "Save Project As", 
                os.path.expanduser("~/Credits.csv"),
                "CSV Files (*.csv);;All Files (*)"
            )
            if not file_path:
                return
            csv_data = self.spreadsheet_widget.get_csv_data()
            if not self.controller.validate_csv(csv_data):
                QMessageBox.critical(self, "Error", "CSV data is invalid.")
                return
            self.controller.save_project(file_path, csv_data)
            self.current_csv_file = file_path
            self.status_bar.showMessage("Project saved successfully")
            self.update_window_title()
        except Exception as e:
            QMessageBox.critical(self, "Error", f"Failed to save project: {str(e)}")
            
    def update_info_panel(self, styling_data):
        """Update the information panel with styling data."""
        self.info_panel.update_info(styling_data)
        
    def update_window_title(self):
        """Update the window title with current file info."""
        if self.current_csv_file:
            filename = os.path.basename(self.current_csv_file)
            self.setWindowTitle(f"CredGen Spreadsheet Editor - {filename}")
        else:
            self.setWindowTitle("CredGen Spreadsheet Editor")
            
    def reorder_cells(self):
        """Reorder selected cells."""
        self.spreadsheet_widget.reorder_selected_cells()
        
    def refresh_styling(self):
        """Refresh styling data from the current styling file."""
        if not self.current_styling_file:
            QMessageBox.information(self, "Info", "No styling file loaded")
            return
            
        try:
            styling_data = self.styling_parser.parse_styling_file(self.current_styling_file)
            self.styling_data = styling_data
            self.spreadsheet_widget.update_styling_data(styling_data)
            self.update_info_panel(styling_data)
            self.status_bar.showMessage("Styling data refreshed")
            
        except Exception as e:
            QMessageBox.critical(self, "Error", f"Failed to refresh styling data: {str(e)}")
            
    def undo(self) -> None:
        """Undo last action."""
        prev_state = self.controller.undo()
        if prev_state:
            self.spreadsheet_widget.load_data(prev_state, self.styling_data)
            self.status_bar.showMessage("Undo performed")
        else:
            self.status_bar.showMessage("Nothing to undo")

    def redo(self) -> None:
        """Redo last undone action."""
        next_state = self.controller.redo()
        if next_state:
            self.spreadsheet_widget.load_data(next_state, self.styling_data)
            self.status_bar.showMessage("Redo performed")
        else:
            self.status_bar.showMessage("Nothing to redo")

    def on_data_changed(self) -> None:
        """Handle data change in spreadsheet."""
        # Push new state to undo stack
        state = self.spreadsheet_widget.get_csv_data()
        self.controller.push_undo(state)
        self.status_bar.showMessage("Data modified")
        
    def on_selection_changed(self):
        """Handle selection change in spreadsheet."""
        selection_info = self.spreadsheet_widget.get_selection_info()
        self.status_bar.showMessage(f"Selected: {selection_info}")
        
    def cut(self):
        """Cut selected cells."""
        self.spreadsheet_widget.cut()
        
    def copy(self):
        """Copy selected cells."""
        self.spreadsheet_widget.copy()
        
    def paste(self):
        """Paste cells from clipboard."""
        self.spreadsheet_widget.paste()
        
    def show_about(self):
        """Show about dialog."""
        QMessageBox.about(
            self, "About CredGen Spreadsheet Editor",
            "CredGen Spreadsheet Editor v1.0\n\n"
            "A specialized editor for CredGen project files with enhanced styling support.\n\n"
            "Features:\n"
            "• Edit CSV spreadsheet files\n"
            "• Styling node dropdowns\n"
            "• Cell reordering\n"
            "• Real-time validation"
        )
        
    def closeEvent(self, event):
        """Handle application close event."""
        # TODO: Check for unsaved changes
        event.accept()


def main():
    """Main application entry point."""
    app = QApplication(sys.argv)
    app.setApplicationName("CredGen Spreadsheet Editor")
    app.setApplicationVersion("1.0")
    
    # Set application style
    app.setStyle('Fusion')
    
    # Create and show main window
    window = CredGenMainWindow()
    window.show()
    
    # Start event loop
    sys.exit(app.exec_())


if __name__ == '__main__':
    main()