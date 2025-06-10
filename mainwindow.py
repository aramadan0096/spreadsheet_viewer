#!/usr/bin/env python3
"""
CredGen Spreadsheet Editor
A PyQt application for editing CredGen spreadsheet files with enhanced styling support.
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


class CredGenMainWindow(QMainWindow):
    """Main application window for CredGen Spreadsheet Editor."""
    
    def __init__(self):
        super().__init__()
        self.file_manager = FileManager()
        self.styling_parser = StylingParser()
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
        
    def init_ui(self):
        """Initialize the user interface."""
        self.setWindowTitle("CredGen Spreadsheet Editor")
        self.setGeometry(100, 100, 1400, 800)
        
        # Set application icon (if available)
        self.setWindowIcon(QIcon("icons/app_icon.png"))
        
        # Create central widget
        central_widget = QWidget()
        self.setCentralWidget(central_widget)
        
        # Create main layout
        main_layout = QVBoxLayout(central_widget)
        
        # Create toolbar
        # self.create_toolbar()
        
        # Create menu bar
        self.create_menu_bar()
        
        # Create main content area
        self.create_main_content(main_layout)
        
        # Create status bar
        self.create_status_bar()
        
    def create_menu_bar(self):
        """Create the application menu bar."""
        menubar = self.menuBar()
        
        # File menu
        file_menu = menubar.addMenu('&File')
        
        # # New project
        # new_action = file_menu.addAction('&New Project')
        # new_action.setShortcut(QKeySequence.New)
        # new_action.triggered.connect(self.new_project)
        
        # Open project
        open_action = file_menu.addAction('&Open Project')
        open_action.setShortcut(QKeySequence.Open)
        open_action.triggered.connect(self.open_project)
        
        file_menu.addSeparator()
        
        # Save
        save_action = file_menu.addAction('&Save')
        save_action.setShortcut(QKeySequence.Save)
        save_action.triggered.connect(self.save_project)
        
        # # Save As
        # save_as_action = file_menu.addAction('Save &As...')
        # save_as_action.setShortcut(QKeySequence.SaveAs)
        # save_as_action.triggered.connect(self.save_as_project)
        
        file_menu.addSeparator()
        
        # Exit
        exit_action = file_menu.addAction('E&xit')
        exit_action.setShortcut(QKeySequence.Quit)
        exit_action.triggered.connect(self.close)
        
        # Edit menu
        edit_menu = menubar.addMenu('&Edit')
        
        # Undo/Redo
        undo_action = edit_menu.addAction('&Undo')
        undo_action.setShortcut(QKeySequence.Undo)
        undo_action.triggered.connect(self.undo)
        
        redo_action = edit_menu.addAction('&Redo')
        redo_action.setShortcut(QKeySequence.Redo)
        redo_action.triggered.connect(self.redo)
        
        edit_menu.addSeparator()
        
        # Cut/Copy/Paste
        cut_action = edit_menu.addAction('Cu&t')
        cut_action.setShortcut(QKeySequence.Cut)
        cut_action.triggered.connect(self.cut)
        
        copy_action = edit_menu.addAction('&Copy')
        copy_action.setShortcut(QKeySequence.Copy)
        copy_action.triggered.connect(self.copy)
        
        paste_action = edit_menu.addAction('&Paste')
        paste_action.setShortcut(QKeySequence.Paste)
        paste_action.triggered.connect(self.paste)
        
        # Tools menu
        tools_menu = menubar.addMenu('&Tools')
        
        # Reorder cells
        reorder_action = tools_menu.addAction('&Reorder Selected Cells')
        reorder_action.setShortcut('Ctrl+R')
        reorder_action.triggered.connect(self.reorder_cells)
        
        # Refresh styling
        refresh_styling_action = tools_menu.addAction('Refresh &Styling Data')
        refresh_styling_action.setShortcut('F5')
        refresh_styling_action.triggered.connect(self.refresh_styling)
        
        # Help menu
        help_menu = menubar.addMenu('&Help')
        
        about_action = help_menu.addAction('&About')
        about_action.triggered.connect(self.show_about)
        
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
        # Create a splitter for spreadsheet and info panel
        splitter = QSplitter(Qt.Horizontal)
        
        # Create spreadsheet widget
        self.spreadsheet_widget = SpreadsheetWidget()
        if self.styling_data:
            self.spreadsheet_widget.update_styling_data(self.styling_data)
            
        # Create info panel
        self.info_panel = self.create_info_panel()
        
        # Add widgets to splitter
        splitter.addWidget(self.spreadsheet_widget)
        splitter.addWidget(self.info_panel)
        
        # Set initial sizes
        splitter.setSizes([800, 200])
        
        layout.addWidget(splitter)
        
    def create_info_panel(self):
        """Create the information panel."""
        info_widget = QWidget()
        info_layout = QVBoxLayout(info_widget)
        
        # Title
        title_label = QLabel("Styling Information")
        title_label.setFont(QFont("Arial", 12, QFont.Bold))
        info_layout.addWidget(title_label)
        
        # Tab widget for different info types
        info_tabs = QTabWidget()
        info_layout.addWidget(info_tabs)
        
        # Page Styles tab
        page_styles_widget = QWidget()
        page_styles_layout = QVBoxLayout(page_styles_widget)
        self.page_styles_label = QLabel("No styling file loaded")
        page_styles_layout.addWidget(self.page_styles_label)
        info_tabs.addTab(page_styles_widget, "Page Styles")
        
        # Content Styles tab
        content_styles_widget = QWidget()
        content_styles_layout = QVBoxLayout(content_styles_widget)
        self.content_styles_label = QLabel("No styling file loaded")
        content_styles_layout.addWidget(self.content_styles_label)
        info_tabs.addTab(content_styles_widget, "Content Styles")
        
        # Letter Styles tab
        letter_styles_widget = QWidget()
        letter_styles_layout = QVBoxLayout(letter_styles_widget)
        self.letter_styles_label = QLabel("No styling file loaded")
        letter_styles_layout.addWidget(self.letter_styles_label)
        info_tabs.addTab(letter_styles_widget, "Letter Styles")
        
        return info_widget
        
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
            
    def open_project(self):
        """Open a project folder containing Credits.csv and Styling.toml files."""
        folder_path = QFileDialog.getExistingDirectory(
            self,
            "Open Project Folder",
            "",
            QFileDialog.ShowDirsOnly | QFileDialog.DontResolveSymlinks
        )
        
        if folder_path:
            # Convert to Path object for easier path manipulation
            folder = Path(folder_path)
            
            # Check for required files
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
                # Load styling data first
                self.styling_data = self.styling_parser.parse_styling_file(str(styling_file))
                if not self.styling_data:
                    QMessageBox.critical(self, "Error", "Failed to parse Styling.toml")
                    return
                    
                self.current_styling_file = str(styling_file)
                
                # Update styling data in spreadsheet widget
                self.spreadsheet_widget.update_styling_data(self.styling_data)
                self.update_info_panel(self.styling_data)
                
                # Load the CSV data
                csv_data = self.file_manager.load_csv(str(credits_file))
                if not csv_data:
                    QMessageBox.critical(self, "Error", "Credits.csv is empty or could not be read")
                    return
                    
                self.current_csv_file = str(credits_file)
                self.spreadsheet_widget.load_data(csv_data, self.styling_data)
                
                # Update window title and status
                self.update_window_title()
                self.statusBar().showMessage(f"Loaded project from: {folder_path}")
                
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
            
    def save_project(self):
        """Save the current project."""
        if not self.current_csv_file:
            self.save_as_project()
            return
            
        try:
            # Get data from spreadsheet widget
            csv_data = self.spreadsheet_widget.get_csv_data()
            
            # Save CSV file
            self.file_manager.save_csv(self.current_csv_file, csv_data)
            
            self.status_bar.showMessage("Project saved successfully")
            
        except Exception as e:
            QMessageBox.critical(self, "Error", f"Failed to save project: {str(e)}")
            
    def save_as_project(self):
        """Save the project with a new name."""
        try:
            file_path, _ = QFileDialog.getSaveFileName(
                self, "Save Project As", 
                os.path.expanduser("~/Credits.csv"),
                "CSV Files (*.csv);;All Files (*)"
            )
            
            if not file_path:
                return
                
            # Get data from spreadsheet widget
            csv_data = self.spreadsheet_widget.get_csv_data()
            
            # Save CSV file
            self.file_manager.save_csv(file_path, csv_data)
            self.current_csv_file = file_path
            
            self.status_bar.showMessage("Project saved successfully")
            self.update_window_title()
            
        except Exception as e:
            QMessageBox.critical(self, "Error", f"Failed to save project: {str(e)}")
            
    def update_info_panel(self, styling_data):
        """Update the information panel with styling data."""
        if not styling_data:
            self.page_styles_label.setText("No styling file loaded")
            self.content_styles_label.setText("No styling file loaded")
            self.letter_styles_label.setText("No styling file loaded")
            return
            
        # Update page styles
        page_styles = styling_data.get('page_styles', [])
        page_text = "Available Page Styles:\n" + "\n".join([f"• {style}" for style in page_styles])
        self.page_styles_label.setText(page_text)
        
        # Update content styles
        content_styles = styling_data.get('content_styles', [])
        content_text = "Available Content Styles:\n" + "\n".join([f"• {style}" for style in content_styles])
        self.content_styles_label.setText(content_text)
        
        # Update letter styles
        letter_styles = styling_data.get('letter_styles', [])
        letter_text = "Available Letter Styles:\n" + "\n".join([f"• {style}" for style in letter_styles])
        self.letter_styles_label.setText(letter_text)
        
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
            
    def on_data_changed(self):
        """Handle data change in spreadsheet."""
        self.status_bar.showMessage("Data modified")
        
    def on_selection_changed(self):
        """Handle selection change in spreadsheet."""
        selection_info = self.spreadsheet_widget.get_selection_info()
        self.status_bar.showMessage(f"Selected: {selection_info}")
        
    def undo(self):
        """Undo last action."""
        self.spreadsheet_widget.undo()
        
    def redo(self):
        """Redo last undone action."""
        self.spreadsheet_widget.redo()
        
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