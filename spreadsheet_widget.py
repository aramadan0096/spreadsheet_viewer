"""
Enhanced spreadsheet widget for CredGen with styling dropdown support and cell reordering.
"""

import csv
from PyQt5.QtWidgets import (
    QWidget, QTableWidget, QTableWidgetItem, QVBoxLayout, QHBoxLayout,
    QComboBox, QHeaderView, QAbstractItemView, QPushButton, QMessageBox,
    QDialog, QListWidget, QDialogButtonBox, QLabel, QCheckBox, QApplication
)
from PyQt5.QtCore import Qt, pyqtSignal, QMimeData
from PyQt5.QtGui import QColor, QFont, QDrag, QPalette, QClipboard
from widgets.style_combobox import StyleComboBox
from dialogs.reorder_dialog import CellReorderDialog


class SpreadsheetWidget(QWidget):
    """
    Enhanced spreadsheet widget with styling support, cell reordering, and undo/redo integration.
    Designed for extensibility and accessibility.
    """
    
    data_changed = pyqtSignal()
    selection_changed = pyqtSignal()
    
    def __init__(self, parent=None):
        super().__init__(parent)
        self.styling_data = None
        # Keep the first data row hidden in the viewer if present
        self.hidden_first_row = None
        # Column definitions - only define static properties here
        self.special_columns = {
            '@Content Style': {'type': 'content', 'col': 4},
            '@Page Style': {'type': 'page', 'col': 7},
            '@Break Harmonization': {'type': 'harmonization', 'col': 5},
            '@Spine Position': {'type': 'spine', 'col': 6},
            '@Vertical Gap': {'type': 'gap', 'col': 3},
            '@Page Runtime': {'type': 'runtime', 'col': 8},
            '@Page Gap': {'type': 'gap', 'col': 9}
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
        # Keep rows compact
        self.table.setWordWrap(False)
        # Elide long text instead of expanding columns
        self.table.setTextElideMode(Qt.ElideRight)

        # Enable drag and drop for reordering
        self.table.setDragDropMode(QAbstractItemView.InternalMove)
        self.table.setDefaultDropAction(Qt.MoveAction)

        layout.addWidget(self.table)

        # Set initial sizing behavior
        self.table.horizontalHeader().setStretchLastSection(True)
        # Make rows a fixed, shorter height
        self.table.verticalHeader().setSectionResizeMode(QHeaderView.Fixed)
        self.table.verticalHeader().setDefaultSectionSize(24)
        self.table.verticalHeader().setMinimumSectionSize(20)
        
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
        print(f"Loading data into spreadsheet. Received {len(csv_data) if csv_data else 0} rows")
        self.styling_data = styling_data or {}

        if not csv_data:
            print("No CSV data provided")
            return

        # Block signals temporarily to prevent unnecessary updates
        self.table.blockSignals(True)

        try:
            # Expect first row as headers; data follows
            headers = csv_data[0]
            data_rows = csv_data[1:] if len(csv_data) > 1 else []
            # Per request: ignore the first data row in the viewer (but preserve it)
            self.hidden_first_row = None
            if data_rows:
                print("Ignoring first data row in viewer")
                self.hidden_first_row = data_rows[0]
                data_rows = data_rows[1:]
            # Set table dimensions
            self.table.setRowCount(len(data_rows))
            self.table.setColumnCount(len(headers))
            print(f"Set table dimensions: {len(csv_data)} rows x {len(csv_data[0]) if csv_data else 0} columns")

            # Set headers if first row contains headers
            print(f"Setting headers: {headers}")
            self.table.setHorizontalHeaderLabels(headers)

            # Populate table
            print(f"Populating {len(data_rows)} data rows")
            for row_idx, row_data in enumerate(data_rows):
                for col_idx, cell_value in enumerate(row_data):
                    header_item = self.table.horizontalHeaderItem(col_idx)
                    if header_item:
                        column_name = header_item.text()
                        if column_name in self.special_columns:
                            # Create combo box for special columns
                            combo = self.create_style_combo(column_name)
                            # Set the current value
                            index = combo.findText(str(cell_value))
                            if index >= 0:
                                combo.setCurrentIndex(index)
                            combo.setToolTip(str(cell_value))
                            self.table.setCellWidget(row_idx, col_idx, combo)
                        else:
                            text = str(cell_value)
                            item = QTableWidgetItem(text)
                            # Show full text on hover when elided
                            item.setToolTip(text)
                            self.table.setItem(row_idx, col_idx, item)

        except Exception as e:
            print(f"Error loading data into spreadsheet: {str(e)}")
            raise
        finally:
            # Re-enable signals
            self.table.blockSignals(False)

        # Resize and clamp column sizes
        self.adjust_column_sizes()

    def adjust_column_sizes(self):
        """Auto-size columns, then clamp @Body to a reasonable width."""
        try:
            self.table.resizeColumnsToContents()
            # Find @Body column
            body_index = -1
            for col in range(self.table.columnCount()):
                header_item = self.table.horizontalHeaderItem(col)
                if header_item and header_item.text().strip() == '@Body':
                    body_index = col
                    break
            if body_index >= 0:
                from PyQt5.QtWidgets import QHeaderView as _QHV
                header = self.table.horizontalHeader()
                header.setSectionResizeMode(body_index, _QHV.Interactive)
                max_width = 320  # clamp width for readability
                self.table.setColumnWidth(body_index, max_width)
        except Exception as e:
            print(f"adjust_column_sizes error: {e}")

    def set_cell_value(self, row, col, value):
        """Set cell value with appropriate widget type."""
        header_text = self.table.horizontalHeaderItem(col)
        column_name = header_text.text() if header_text else ""

        # Check if this is a special styling column
        if column_name in self.special_columns:
            combo = self.create_style_combo(column_name)
            current_index = combo.findText(str(value))
            if current_index >= 0:
                combo.setCurrentIndex(current_index)
            self.table.setCellWidget(row, col, combo)
        else:
            text = str(value) if value is not None else ""
            item = QTableWidgetItem(text)
            item.setToolTip(text)
            self.table.setItem(row, col, item)

    def create_style_combo(self, column_name):
        """Create a style combo box for the specified column."""
        combo = QComboBox()
        combo.setEditable(True)
        combo.setInsertPolicy(QComboBox.NoInsert)
        
        items = ['']  # Always start with an empty option
        
        if self.styling_data:
            style_type = self.special_columns[column_name]['type']
            
            if style_type == 'content':
                items.extend(self.styling_data.get('content_styles', []))
            elif style_type == 'page':
                items.extend(self.styling_data.get('page_styles', []))
            elif style_type == 'harmonization':
                items.extend(self.styling_data.get('harmonization_values', []))
            elif style_type == 'spine':
                items.extend(self.styling_data.get('spine_positions', []))
            elif style_type == 'gap':
                items.extend(self.styling_data.get('gaps', []))
            elif style_type == 'runtime':
                items.extend(self.styling_data.get('runtimes', []))
                
        combo.addItems(items)
        combo.currentTextChanged.connect(self.on_combo_changed)
        return combo

    def setup_special_columns(self):
        """Setup special columns with appropriate widgets."""
        for col in range(self.table.columnCount()):
            header_text = self.table.horizontalHeaderItem(col)
            if header_text and header_text.text() in self.special_columns:
                # Set up dropdown for each row in this column
                for row in range(self.table.rowCount()):
                    existing_value = ""
                    existing_item = self.table.item(row, col)
                    if existing_item:
                        existing_value = existing_item.text()
                    self.set_cell_value(row, col, existing_value)
                    
    def update_styling_data(self, styling_data):
        """Update styling data and refresh dropdowns."""
        self.styling_data = styling_data
        
        # Update all existing combo boxes
        for col in range(self.table.columnCount()):
            header_item = self.table.horizontalHeaderItem(col)
            if header_item:
                column_name = header_item.text()
                if column_name in self.special_columns:
                    # Update combo boxes for each row in this column
                    for row in range(self.table.rowCount()):
                        widget = self.table.cellWidget(row, col)
                        if isinstance(widget, QComboBox):
                            current_value = widget.currentText()
                            # Create new combo with updated data
                            new_combo = self.create_style_combo(column_name)
                            # Try to restore the previous value
                            index = new_combo.findText(current_value)
                            if index >= 0:
                                new_combo.setCurrentIndex(index)
                            self.table.setCellWidget(row, col, new_combo)

    def get_csv_data(self, include_headers: bool = True):
        """Get current table data as CSV-compatible list.
        By default includes the header row as the first row so downstream loaders
        can treat headers properly and not display them as data.
        """
        data = []

        # Optionally include headers as the first row
        if include_headers:
            headers = []
            for col in range(self.table.columnCount()):
                header_item = self.table.horizontalHeaderItem(col)
                headers.append(header_item.text() if header_item else "")
            data.append(headers)

        # Re-insert hidden first row (if any) right after headers
        if include_headers and self.hidden_first_row is not None:
            data.append([str(v) for v in self.hidden_first_row])

        # Add table body rows
        for row in range(self.table.rowCount()):
            row_data = []
            for col in range(self.table.columnCount()):
                widget = self.table.cellWidget(row, col)
                if isinstance(widget, QComboBox):
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
            row, col = idx.row(), idx.column()
            # Get value based on whether it's a combo box or regular cell
            widget = self.table.cellWidget(row, col)
            if isinstance(widget, QComboBox):
                value = widget.currentText()
            else:
                item = self.table.item(row, col)
                value = item.text() if item else ""
            cell_data.append((row, col, value))
            
        # Show reorder dialog
        dialog = CellReorderDialog(cell_data, self)
        if dialog.exec_() == QDialog.Accepted:
            reordered = dialog.get_reordered_data()
            # Apply the reordered values according to new order
            for (row, col, _), (_, _, value) in zip(cell_data, reordered):
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
        
        for col in range(self.table.columnCount()):
            header_item = self.table.horizontalHeaderItem(col)
            if header_item:
                column_name = header_item.text()
                if column_name in self.special_columns:
                    # Create combo box for special columns
                    combo = self.create_style_combo(column_name)
                    self.table.setCellWidget(current_row_count, col, combo)
                else:
                    self.table.setItem(current_row_count, col, QTableWidgetItem(""))
        
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
        self.hidden_first_row = None

    # Clipboard operations
    def copy(self):
        selected = self.table.selectedIndexes()
        if not selected:
            return
        # Group by rows
        selected.sort(key=lambda x: (x.row(), x.column()))
        rows = {}
        for idx in selected:
            rows.setdefault(idx.row(), {})[idx.column()] = idx
        lines = []
        for r in sorted(rows.keys()):
            cols = rows[r]
            line_vals = []
            for c in sorted(cols.keys()):
                idx = cols[c]
                widget = self.table.cellWidget(idx.row(), idx.column())
                if isinstance(widget, QComboBox):
                    line_vals.append(widget.currentText())
                else:
                    item = self.table.item(idx.row(), idx.column())
                    line_vals.append(item.text() if item else "")
            lines.append("\t".join(line_vals))
        QApplication.clipboard().setText("\n".join(lines), mode=QClipboard.Clipboard)

    def cut(self):
        self.copy()
        # After copying, clear selected cells
        for idx in self.table.selectedIndexes():
            self.set_cell_value(idx.row(), idx.column(), "")
        self.data_changed.emit()

    def paste(self):
        text = QApplication.clipboard().text(QClipboard.Clipboard)
        if not text:
            return
        start_indexes = self.table.selectedIndexes()
        if not start_indexes:
            return
        start_row = min(idx.row() for idx in start_indexes)
        start_col = min(idx.column() for idx in start_indexes)
        rows = text.splitlines()
        for r, line in enumerate(rows):
            values = line.split("\t")
            for c, val in enumerate(values):
                row = start_row + r
                col = start_col + c
                if row < self.table.rowCount() and col < self.table.columnCount():
                    self.set_cell_value(row, col, val)
        self.data_changed.emit()
        
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