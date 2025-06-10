import sys
from PyQt5.QtWidgets import QApplication
from mainwindow import CredGenMainWindow

def main():
    app = QApplication(sys.argv)
    app.setApplicationName("CredGen Spreadsheet Editor")
    app.setApplicationVersion("1.0")
    app.setStyle('Fusion')
    window = CredGenMainWindow()
    window.show()
    sys.exit(app.exec_())

if __name__ == '__main__':
    main()
