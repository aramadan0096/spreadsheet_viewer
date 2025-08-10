import sys
import os
from PyQt5.QtWidgets import QApplication
from mainwindow import CredGenMainWindow

def main():
    app = QApplication(sys.argv)
    app.setApplicationName("CredGen Spreadsheet Editor")
    app.setApplicationVersion("1.0")
    app.setStyle('Fusion')
    # Load application stylesheet
    try:
        base_dir = os.path.dirname(os.path.abspath(__file__))
        css_path = os.path.join(base_dir, 'UI', 'stylesheet.css')
        if os.path.exists(css_path):
            with open(css_path, 'r', encoding='utf-8') as f:
                app.setStyleSheet(f.read())
    except Exception as e:
        # Non-fatal: continue without stylesheet
        print(f"Warning: failed to load stylesheet: {e}")
    window = CredGenMainWindow()
    window.show()
    sys.exit(app.exec_())

if __name__ == '__main__':
    main()
