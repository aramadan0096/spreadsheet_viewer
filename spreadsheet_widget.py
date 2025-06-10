"""
Enhanced spreadsheet widget for CredGen with styling dropdown support and cell reordering.
"""

import csv
from PyQt5.QtWidgets import (
    QWidget, QTableWidget, QTableWidgetItem, QVBoxLayout, QHBoxLayout,
    QComboBox, QHeaderView, QAbstractItemView, QPushButton, QMessageBox,
    QDialog, QListWidget, QDialogButtonBox, QLabel, QCheckBox
)
from PyQt5.QtCore import Qt, pyqtSignal, QMimeData
from PyQt5.QtGui import QColor, QFont, QDrag, QPalette


class StyleComboBox(QComboBox):
    """Custom combo box for styling nodes with categorization."""
    
    def __init__(self, style_type="content", parent=None):
        super().__init__(parent)
        self.style_type = style_type
        self.setEditable(True)
        self.setInsertPolicy(QComboBox.NoInsert)
        
    def update_styles(self, styling_data):
        """Update the combo box with available styles."""
        self.clear()
        self.addItem("")  # Empty option
        
        if not styling_data:
            return
            
        if self.style_type == "page" and "page_styles" in styling_data:
            for style in styling_data["page_styles"]:
                self.addItem(style)
        elif self.style_type == "content" and "content_styles" in styling_data:
            for style in styling_data["content_styles"]:
                self.addItem(style)
        elif self.style_type == "letter" and "letter_styles" in styling_data:
            for style in styling_data["letter_styles"]:
                self.addItem(style)


class CellReorderDialog(QDialog):
    """Dialog for reordering selected cells."""
    
    def __init__(self, cell_data, parent=None):
        super().__init__(parent)
        self.cell_data = cell_data
        self.init_ui()
        
    def init_ui(self):
        """Initialize the dialog UI."""
        self.setWindowTitle("Reorder Cells")
        self.setModal(True)
        self.resize(400, 300)
        
        layout = QVBoxLayout(self)
        
        # Instructions
        instructions = QLabel("Drag and drop to reorder the selected cells:")
        layout.addWidget(instructions)
        
        # List widget for reordering
        self.list_widget = QListWidget()
        self.list_widget.setDragDropMode(QAbstractItemView.InternalMove)
        
        # Populate list
        for i, (row, col, value) in enumerate(self.cell_data):
            item_text = f"Row {row+1}, Col {col+1}: {value}"
            self.list_widget.addItem(item_text)
            
        layout.addWidget(self.list_widget)
        
        # Buttons
        button_box = QDialogButtonBox(QDialogButtonBox.Ok | QDialogButtonBox.Cancel)
        button_box.accepted.connect(self.accept)
        button_box.rejected.connect(self.reject)
        layout.addWidget(button_box)
        
    def get_reordered_data(self):
        """Get the reordered cell data."""
        reordered_data = []
        for i in range(self.list_widget.count()):
            item = self.list_widget.item(i)
            original_index = self.cell_data[i][0]  # This needs to be tracked properly
            reordered_data.append(self.cell_data[original_index])
        return reordered_data


class SpreadsheetWidget(QWidget):
    """Enhanced spreadsheet widget with styling support and cell reordering."""
    
    data_changed = pyqtSignal()
    selection_changed = pyqtSignal()
    
    def __init__(self, parent=None):
        super().__init__(parent)
        self.styling_data = None
        self.special_columns = {
            '@Content Style': 4,
            '@Page Style': 7,
            '@Break Harmonization': 5,
            '@Spine Position': 6,
            '@Page Runtime': 8,
            '@Page Gap': 9
        }
        
        self.init_ui()
        self.setup_connections()
        
    def init_ui(self):
        """Initialize the widget UI."""
        layout = QVBoxLayout(self)
        
        # Control panel
        control_panel = self.create_control_panel()
        layout.addWidget(control_panel)
        
        # Create table widget
        self.table = QTableWidget()
        self.table.setAlternatingRowColors(True)
        self.table.setSelectionBehavior(QAbstractItemView.SelectItems)
        self.table.setSelectionMode(QAbstractItemView.ExtendedSelection)
        
        # Enable drag and drop for reordering
        self.table.setDragDropMode(QAbstractItemView.InternalMove)
        self.table.setDefaultDropAction(Qt.MoveAction)
        
        layout.addWidget(self.table)
        
        # Set initial size
        self.table.horizontalHeader().setStretchLastSection(True)
        self.table.verticalHeader().setDefaultSectionSize(30)
        
    def create_control_panel(self):
        """Create the control panel with action buttons."""
        panel = QWidget()
        layout = QHBoxLayout(panel)

        # Add row button
        add_row_btn = QPushButton("Add Row")
        add_row_btn.clicked.connect(self.add_row)
        layout.addWidget(add_row_btn)

        # Delete row button
        delete_row_btn = QPushButton("Delete Row")
        delete_row_btn.clicked.connect(self.delete_row)
        layout.addWidget(delete_row_btn)

        layout.addStretch()
        return panel
        
    def setup_connections(self):
        """Setup signal-slot connections."""
        self.table.itemChanged.connect(self.on_item_changed)
        self.table.itemSelectionChanged.connect(self.on_selection_changed)
        
    def load_data(self, csv_data, styling_data=None):
        """Load CSV data into the table."""
        self.styling_data = styling_data

        if not csv_data:
            return

        # Set table dimensions
        self.table.setRowCount(len(csv_data))
        self.table.setColumnCount(len(csv_data[0]) if csv_data else 0)

        # Set headers if first row contains headers
        if csv_data and csv_data[0] and str(csv_data[0][0]).startswith('@'):
            headers = csv_data[0]
            self.table.setHorizontalHeaderLabels(headers)
            data_rows = csv_data[1:]
        else:
            data_rows = csv_data

        # Populate table
        for row_idx, row_data in enumerate(data_rows):
            for col_idx, cell_value in enumerate(row_data):
                self.set_cell_value(row_idx, col_idx, cell_value)

        # Setup special columns with dropdowns
        self.setup_special_columns()

        # Resize columns to content
        self.table.resizeColumnsToContents()

    def set_cell_value(self, row, col, value):
        """Set cell value with appropriate widget type."""
        header_text = self.table.horizontalHeaderItem(col)
        column_name = header_text.text() if header_text else ""

        # Check if this is a special styling column
        if column_name in self.special_columns:
            combo = self.create_style_combo(column_name)
            combo.setCurrentText(str(value))
            self.table.setCellWidget(row, col, combo)
        else:
            item = QTableWidgetItem(str(value))
            # Apply special formatting for header rows
            if str(value).startswith('@'):
                font = QFont()
                font.setBold(True)
                item.setFont(font)
            self.table.setItem(row, col, item)

    def create_style_combo(self, column_name):
        """Create a style combo box for the specified column."""
        if column_name == '@Content Style':
            combo = StyleComboBox('content')
        elif column_name == '@Page Style':
            combo = StyleComboBox('page')
        elif column_name == '@Letter Style':
            combo = StyleComboBox('letter')
        else:
            combo = StyleComboBox('content')  # Default
        combo.update_styles(self.styling_data)
        combo.currentTextChanged.connect(self.on_combo_changed)
        return combo

    def setup_special_columns(self):
        """Setup special columns with appropriate widgets."""
        if not self.styling_data:
            return
        for col in range(self.table.columnCount()):
            header_text = self.table.horizontalHeaderItem(col)
            column_name = header_text.text() if header_text else ""
            if column_name in self.special_columns:
                for row in range(self.table.rowCount()):
                    value = self.table.item(row, col).text() if self.table.item(row, col) else ""
                    combo = self.create_style_combo(column_name)
                    combo.setCurrentText(value)
                    self.table.setCellWidget(row, col, combo)

    def update_styling_data(self, styling_data):
        """Update styling data and refresh dropdowns."""
        self.styling_data = styling_data
        # Update all existing combo boxes
        for col in range(self.table.columnCount()):
            header_text = self.table.horizontalHeaderItem(col)
            column_name = header_text.text() if header_text else ""
            if column_name in self.special_columns:
                for row in range(self.table.rowCount()):
                    widget = self.table.cellWidget(row, col)
                    if isinstance(widget, StyleComboBox):
                        widget.update_styles(styling_data)

    def get_csv_data(self):
        """Get current table data as CSV-compatible list."""
        data = []
        for row in range(self.table.rowCount()):
            row_data = []
            for col in range(self.table.columnCount()):
                widget = self.table.cellWidget(row, col)
                if isinstance(widget, StyleComboBox):
                    row_data.append(widget.currentText())
                else:
                    item = self.table.item(row, col)
                    row_data.append(item.text() if item else "")
            data.append(row_data)
        return data

    def reorder_selected_cells(self):
        """Reorder selected cells using a dialog."""
        selected = self.table.selectedIndexes()
        if not selected:
            QMessageBox.information(self, "Info", "No cells selected to reorder.")
            return
        # Collect selected cell data
        cell_data = []
        for idx in selected:
            value = self.table.item(idx.row(), idx.column()).text() if self.table.item(idx.row(), idx.column()) else ""
            cell_data.append((idx.row(), idx.column(), value))
        # Show dialog
        dialog = CellReorderDialog(cell_data, self)
        if dialog.exec_() == QDialog.Accepted:
            reordered = dialog.get_reordered_data()
            for i, (row, col, value) in enumerate(reordered):
                self.set_cell_value(row, col, value)
        self.data_changed.emit()

    def on_combo_changed(self, value):
        self.data_changed.emit()

    def on_item_changed(self, item):
        """Handle item changed event from the table."""
        self.data_changed.emit()

    def on_selection_changed(self):
        """Handle selection changed event from the table."""
        self.selection_changed.emit()

    def get_selection_info(self):
        """Return a string describing the current selection."""
        selected = self.table.selectedIndexes()
        if not selected:
            return "None"
        cells = [(idx.row()+1, idx.column()+1) for idx in selected]
        return f"{len(cells)} cell(s) selected: {cells}"

    def add_row(self):
        """Add a new row to the table."""
        current_row_count = self.table.rowCount()
        self.table.insertRow(current_row_count)
        
        # Copy formatting and data from the last row if available
        if current_row_count > 0:
            for col in range(self.table.columnCount()):
                # Copy item
                item = self.table.item(current_row_count - 1, col)
                if item:
                    new_item = QTableWidgetItem(item)
                    new_item.setFlags(item.flags() & ~Qt.ItemIsEditable)  # Make non-editable
                    self.table.setItem(current_row_count, col, new_item)
                
                # Copy combo box
                widget = self.table.cellWidget(current_row_count - 1, col)
                if isinstance(widget, QComboBox):
                    combo = QComboBox()
                    combo.addItems([widget.itemText(i) for i in range(widget.count())])
                    combo.setCurrentText(widget.currentText())
                    combo.setEditable(True)
                    combo.setInsertPolicy(QComboBox.NoInsert)
                    self.table.setCellWidget(current_row_count, col, combo)
        
        self.data_changed.emit()
        
    def delete_row(self):
        """Delete the selected row(s) from the table."""
        selected_rows = set(idx.row() for idx in self.table.selectedIndexes())
        for row in sorted(selected_rows, reverse=True):
            self.table.removeRow(row)
            
        self.data_changed.emit()
        
    def add_column(self):
        """Add a new column to the table."""
        current_column_count = self.table.columnCount()
        self.table.setColumnCount(current_column_count + 1)
        
        # Set header for the new column
        self.table.setHorizontalHeaderItem(current_column_count, QTableWidgetItem(f"Column {current_column_count + 1}"))
        
        # Copy formatting from the last column if available
        if current_column_count > 0:
            for row in range(self.table.rowCount()):
                item = self.table.item(row, current_column_count - 1)
                if item:
                    new_item = QTableWidgetItem(item)
                    new_item.setFlags(item.flags() & ~Qt.ItemIsEditable)  # Make non-editable
                    self.table.setItem(row, current_column_count, new_item)
                
                widget = self.table.cellWidget(row, current_column_count - 1)
                if isinstance(widget, QComboBox):
                    combo = QComboBox()
                    combo.addItems([widget.itemText(i) for i in range(widget.count())])
                    combo.setCurrentText(widget.currentText())
                    combo.setEditable(True)
                    combo.setInsertPolicy(QComboBox.NoInsert)
                    self.table.setCellWidget(row, current_column_count, combo)
        
        self.data_changed.emit()
        
    def delete_column(self):
        """Delete the selected column(s) from the table."""
        selected_columns = set(idx.column() for idx in self.table.selectedIndexes())
        for col in sorted(selected_columns, reverse=True):
            self.table.removeColumn(col)
            
        self.data_changed.emit()
        
    def clear_data(self):
        """Clear all table data."""
        self.table.clear()
        self.table.setRowCount(0)
        self.table.setColumnCount(0)
        self.styling_data = None
        
    def create_new_project(self):
        """Create a new empty project structure."""
        # Create basic CredGen structure
        headers = [
            "@Head", "@Body", "@Tail", "@Vertical Gap", "@Content Style",
            "@Break Harmonization", "@Spine Position", "@Page Style",
            "@Page Runtime", "@Page Gap"
        ]
        
        self.clear_data()
        self.table.setColumnCount(len(headers))
        self.table.setHorizontalHeaderLabels(headers)
        
        # Add an initial empty row
        self.add_row()
        
    def load_project(self, file_path):
        """Load a project from a .csv file."""
        with open(file_path, 'r', newline='', encoding='utf-8') as csvfile:
            reader = csv.reader(csvfile)
            csv_data = list(reader)
            
            # Detect and load styling data if available
            styling_data = None
            if len(csv_data) > 0 and csv_data[0][0].startswith('@'):
                # Extract styling data from the first row
                styling_data = self.extract_styling_data(csv_data[0])
                # Remove styling row from data
                csv_data = csv_data[1:]
            
            self.load_data(csv_data, styling_data)
            
    def extract_styling_data(self, header_row):
        """Extract styling data from the header row."""
        styling_data = {
            "content_styles": [],
            "page_styles": [],
            "letter_styles": []
        }
        
        for item in header_row:
            if item.startswith('@'):
                style_type, _, style_name = item.partition(' ')
                if style_type == '@Content':
                    styling_data["content_styles"].append(style_name)
                elif style_type == '@Page':
                    styling_data["page_styles"].append(style_name)
                elif style_type == '@Letter':
                    styling_data["letter_styles"].append(style_name)
        
        return styling_data
    
    def save_project(self, file_path):
        """Save the current project to a .csv file."""
        with open(file_path, 'w', newline='', encoding='utf-8') as csvfile:
            writer = csv.writer(csvfile)
            
            # Write styling data as the first row if available
            if self.styling_data:
                header_row = self.generate_styling_header()
                writer.writerow(header_row)
            
            # Write the rest of the data
            data = self.get_csv_data()
            for row in data:
                writer.writerow(row)
                
    def generate_styling_header(self):
        """Generate the header row for styling data."""
        header = []
        
        # Content styles
        for style in self.styling_data.get("content_styles", []):
            header.append(f"@Content {style}")
            
        # Page styles
        for style in self.styling_data.get("page_styles", []):
            header.append(f"@Page {style}")
            
        # Letter styles
        for style in self.styling_data.get("letter_styles", []):
            header.append(f"@Letter {style}")
        
        return header + [""] * (self.table.columnCount() - len(header))  # Fill