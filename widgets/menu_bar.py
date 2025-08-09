from PyQt5.QtWidgets import QMenuBar
from PyQt5.QtGui import QKeySequence

class MenuBar(QMenuBar):
    """
    Application menu bar for CredGen Spreadsheet Editor.
    Provides accessible actions with keyboard shortcuts and tooltips.
    """
    def __init__(self, parent=None):
        super().__init__(parent)
        self.actions = {}
        self._create_menus()

    def _create_menus(self):
        # File menu
        file_menu = self.addMenu('&File')
        self.actions['open'] = file_menu.addAction('&Open Project')
        self.actions['open'].setShortcut(QKeySequence.Open)
        file_menu.addSeparator()
        self.actions['save'] = file_menu.addAction('&Save')
        self.actions['save'].setShortcut(QKeySequence.Save)
        file_menu.addSeparator()
        self.actions['exit'] = file_menu.addAction('E&xit')
        self.actions['exit'].setShortcut(QKeySequence.Quit)
        # Edit menu
        edit_menu = self.addMenu('&Edit')
        self.actions['undo'] = edit_menu.addAction('&Undo')
        self.actions['undo'].setShortcut(QKeySequence.Undo)
        self.actions['redo'] = edit_menu.addAction('&Redo')
        self.actions['redo'].setShortcut(QKeySequence.Redo)
        edit_menu.addSeparator()
        self.actions['cut'] = edit_menu.addAction('Cu&t')
        self.actions['cut'].setShortcut(QKeySequence.Cut)
        self.actions['copy'] = edit_menu.addAction('&Copy')
        self.actions['copy'].setShortcut(QKeySequence.Copy)
        self.actions['paste'] = edit_menu.addAction('&Paste')
        self.actions['paste'].setShortcut(QKeySequence.Paste)
        # Tools menu
        tools_menu = self.addMenu('&Tools')
        self.actions['reorder'] = tools_menu.addAction('&Reorder Selected Cells')
        self.actions['reorder'].setShortcut('Ctrl+R')
        self.actions['refresh_styling'] = tools_menu.addAction('Refresh &Styling Data')
        self.actions['refresh_styling'].setShortcut('F5')
        # Help menu
        help_menu = self.addMenu('&Help')
        self.actions['about'] = help_menu.addAction('&About')
