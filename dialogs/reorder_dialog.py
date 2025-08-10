from PyQt5.QtWidgets import QDialog, QVBoxLayout, QLabel, QListWidget, QAbstractItemView, QDialogButtonBox, QListWidgetItem
from PyQt5.QtCore import Qt

class CellReorderDialog(QDialog):
    """Dialog for reordering selected cells."""
    def __init__(self, cell_data, parent=None):
        super().__init__(parent)
        self.cell_data = cell_data
        self.original_order = list(range(len(cell_data)))  # Keep track of original indices
        self.init_ui()

    def init_ui(self):
        """Initialize the dialog UI."""
        self.setWindowTitle("Reorder Cells")
        self.setModal(True)
        self.resize(400, 300)
        layout = QVBoxLayout(self)
        instructions = QLabel("Drag and drop to reorder the selected cells:")
        layout.addWidget(instructions)
        self.list_widget = QListWidget()
        self.list_widget.setDragDropMode(QAbstractItemView.InternalMove)
        for i, (row, col, value) in enumerate(self.cell_data):
            item_text = f"Row {row+1}, Col {col+1}: {value}"
            item = QListWidgetItem(item_text)
            item.setData(Qt.UserRole, i)
            self.list_widget.addItem(item)
        layout.addWidget(self.list_widget)
        button_box = QDialogButtonBox(QDialogButtonBox.Ok | QDialogButtonBox.Cancel)
        button_box.accepted.connect(self.accept)
        button_box.rejected.connect(self.reject)
        layout.addWidget(button_box)

    def get_reordered_data(self):
        """Get the reordered cell data."""
        reordered_data = []
        for i in range(self.list_widget.count()):
            item = self.list_widget.item(i)
            original_index = item.data(Qt.UserRole)
            reordered_data.append(self.cell_data[original_index])
        return reordered_data
