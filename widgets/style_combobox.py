from PyQt5.QtWidgets import QComboBox

class StyleComboBox(QComboBox):
    """
    Custom combo box for styling nodes with categorization.
    Extensible for new style types via plugin/config pattern.
    """
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
        if self.style_type == "page" and "pageStyle" in styling_data:
            for style in styling_data["pageStyle"]:
                if "name" in style:
                    self.addItem(style["name"])
        elif self.style_type == "content" and "contentStyle" in styling_data:
            for style in styling_data["contentStyle"]:
                if "name" in style:
                    self.addItem(style["name"])
        elif self.style_type == "letter" and "letterStyle" in styling_data:
            for style in styling_data["letterStyle"]:
                if "name" in style:
                    self.addItem(style["name"])
