from PyQt5.QtWidgets import QWidget, QVBoxLayout, QLabel, QTabWidget

class InfoPanel(QWidget):
    """
    Panel displaying available page, content, and letter styles for reference.
    Updates dynamically based on loaded styling data.
    """

    def __init__(self, parent=None):
        super().__init__(parent)
        layout = QVBoxLayout(self)
        # Tab widget for different info types
        self.info_tabs = QTabWidget()
        layout.addWidget(self.info_tabs)
        # Page Styles tab
        self.page_styles_label = QLabel("No styling file loaded")
        page_styles_widget = QWidget()
        page_styles_layout = QVBoxLayout(page_styles_widget)
        page_styles_layout.addWidget(self.page_styles_label)
        self.info_tabs.addTab(page_styles_widget, "Page Styles")
        # Content Styles tab
        self.content_styles_label = QLabel("No styling file loaded")
        content_styles_widget = QWidget()
        content_styles_layout = QVBoxLayout(content_styles_widget)
        content_styles_layout.addWidget(self.content_styles_label)
        self.info_tabs.addTab(content_styles_widget, "Content Styles")
        # Letter Styles tab
        self.letter_styles_label = QLabel("No styling file loaded")
        letter_styles_widget = QWidget()
        letter_styles_layout = QVBoxLayout(letter_styles_widget)
        letter_styles_layout.addWidget(self.letter_styles_label)
        self.info_tabs.addTab(letter_styles_widget, "Letter Styles")

    def update_info(self, styling_data):
        if not styling_data:
            self.page_styles_label.setText("No styling file loaded")
            self.content_styles_label.setText("No styling file loaded")
            self.letter_styles_label.setText("No styling file loaded")
            return
        page_styles = styling_data.get('page_styles', [])
        page_text = "Available Page Styles:\n" + "\n".join([f"• {style}" for style in page_styles])
        self.page_styles_label.setText(page_text)
        content_styles = styling_data.get('content_styles', [])
        content_text = "Available Content Styles:\n" + "\n".join([f"• {style}" for style in content_styles])
        self.content_styles_label.setText(content_text)
        letter_styles = styling_data.get('letter_styles', [])
        letter_text = "Available Letter Styles:\n" + "\n".join([f"• {style}" for style in letter_styles])
        self.letter_styles_label.setText(letter_text)
