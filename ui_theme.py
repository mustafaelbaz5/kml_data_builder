"""Single source of truth for the desktop application's visual tokens."""

APP_STYLESHEET = """
QWidget { font-family: 'Segoe UI'; font-size: 13px; color: #172033; background: #F4F7FB; }
QFrame#header { background: #102A43; border-radius: 14px; }
QFrame#panel, QFrame#card { background: #FFFFFF; border: 1px solid #D8E2EE; border-radius: 12px; }
QFrame#toolbar { background: #EAF1F8; border: 1px solid #CBD9E7; border-radius: 12px; }
QLabel#headerTitle { color: #FFFFFF; font-size: 27px; font-weight: 700; background: transparent; }
QLabel#headerSubtitle { color: #D6E5F3; font-size: 13px; background: transparent; }
QLabel#fileChip { color: #E9F5FF; background: #1D4B70; border: 1px solid #427397; border-radius: 8px; padding: 7px 10px; }
QLabel#sectionTitle { color: #102A43; font-size: 17px; font-weight: 700; background: transparent; }
QLabel#muted { color: #52657A; background: transparent; }
QLabel#statusNote { color: #1E3A5F; background: #EAF3FF; border: 1px solid #BED8F2; border-radius: 7px; padding: 8px; }
QLabel#fieldLabel { color: #27435E; background: #EAF1F8; border: 1px solid #CBD9E7; border-radius: 6px; padding: 8px; font-weight: 700; }
QPushButton { background: #FFFFFF; color: #102A43; border: 1px solid #9EB4C8; border-radius: 7px; padding: 8px 12px; font-weight: 600; min-height: 18px; }
QPushButton:hover { background: #E7F1FA; border-color: #467FAE; }
QPushButton:disabled { background: #E9EEF3; color: #8493A3; border-color: #D5DDE6; }
QPushButton#primary { background: #0E7490; color: #FFFFFF; border-color: #0E7490; }
QPushButton#primary:hover { background: #075E72; }
QPushButton#success { background: #166534; color: #FFFFFF; border-color: #166534; }
QPushButton#warning { background: #A16207; color: #FFFFFF; border-color: #A16207; }
QLineEdit, QComboBox, QTextEdit { color: #172033; background: #FFFFFF; border: 1px solid #9EB4C8; border-radius: 7px; padding: 7px; selection-background-color: #0E7490; }
QLineEdit:focus, QComboBox:focus, QTextEdit:focus { border: 2px solid #0E7490; }
QTableWidget { color: #172033; background: #FFFFFF; alternate-background-color: #F7FAFD; border: 1px solid #CBD9E7; border-radius: 8px; gridline-color: #E0E8F0; }
QTableWidget::item { padding: 7px; }
QTableWidget::item:selected { background: #CDEBFA; color: #102A43; }
QHeaderView::section { background: #102A43; color: #FFFFFF; padding: 9px; border: 0; font-weight: 700; }
QTabWidget::pane { border: 1px solid #D8E2EE; border-radius: 8px; background: #FFFFFF; }
QTabBar::tab { background: #EAF1F8; color: #27435E; padding: 8px 13px; border: 1px solid #CBD9E7; }
QTabBar::tab:selected { background: #0E7490; color: #FFFFFF; }
QStatusBar { color: #27435E; background: #EAF1F8; }
"""

STATUS_COLORS = {
    "valid": "#166534", "unallocated": "#166534", "title_mismatch": "#A16207", "empty": "#B42318",
    "no_description": "#B42318", "incomplete": "#B42318", "unreadable": "#B42318",
    "skipped": "#A16207",
    "duplicate_different": "#A16207", "duplicate_identical": "#A16207",
}
