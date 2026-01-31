"""
UI Styling for Technical Catalog Standardization Tool
Professional, Minimal Design for Non-Technical Users
Suitable for Administrative / Medical / Management Applications
"""

MAIN_STYLE = """
/* ============================================
   MAIN WINDOW & BACKGROUND
   ============================================ */
QMainWindow {
    background-color: #f8f9fa;
}

/* ============================================
   LABELS & TEXT
   ============================================ */
QLabel {
    font-size: 14px;
    color: #495057;
    font-family: "Segoe UI", "Arial", sans-serif;
    font-weight: normal;
}

QLabel#title {
    font-size: 24px;
    font-weight: 600;
    color: #212529;
    padding: 20px;
    background-color: white;
    border-bottom: 3px solid #4a90e2;
}

QLabel#section_title {
    font-size: 16px;
    font-weight: 600;
    color: #343a40;
    padding: 10px 0px;
}

/* ============================================
   BUTTONS - LARGE & CLEAR
   ============================================ */
QPushButton {
    background-color: #4a90e2;
    color: white;
    border: none;
    padding: 14px 28px;
    font-size: 15px;
    font-weight: 600;
    border-radius: 4px;
    min-width: 140px;
    min-height: 22px;
    font-family: "Segoe UI", "Arial", sans-serif;
}

QPushButton:hover {
    background-color: #357abd;
}

QPushButton:pressed {
    background-color: #2a6aa8;
}

QPushButton:disabled {
    background-color: #dee2e6;
    color: #6c757d;
}

/* Primary action button */
QPushButton#success_button {
    background-color: #28a745;
    padding: 18px 32px;
    min-height: 28px;
    font-size: 16px;
}

QPushButton#success_button:hover {
    background-color: #218838;
}

QPushButton#success_button:disabled {
    background-color: #dee2e6;
    color: #6c757d;
}

/* Warning button */
QPushButton#warning_button {
    background-color: #ffc107;
    color: #212529;
}

QPushButton#warning_button:hover {
    background-color: #e0a800;
}

/* ============================================
   INPUT FIELDS - CLEAN & CLEAR
   ============================================ */
QLineEdit, QTextEdit {
    padding: 14px;
    border: 1px solid #ced4da;
    border-radius: 4px;
    font-size: 14px;
    background-color: white;
    color: #495057;
    font-family: "Segoe UI", "Arial", sans-serif;
    min-height: 20px;
}

QLineEdit:focus, QTextEdit:focus {
    border: 2px solid #4a90e2;
    background-color: #f8f9ff;
}

QLineEdit:disabled, QTextEdit:disabled {
    background-color: #e9ecef;
    color: #6c757d;
    border: 1px solid #dee2e6;
}

/* ============================================
   GROUP BOXES - MINIMAL BORDER
   ============================================ */
QGroupBox {
    font-size: 15px;
    font-weight: 600;
    border: 1px solid #dee2e6;
    border-radius: 6px;
    margin-top: 20px;
    padding-top: 18px;
    padding-left: 12px;
    padding-right: 12px;
    padding-bottom: 18px;
    background-color: white;
}

QGroupBox::title {
    subcontrol-origin: margin;
    subcontrol-position: top left;
    padding: 6px 12px;
    color: #343a40;
    background-color: #f8f9fa;
    border: 1px solid #dee2e6;
    border-radius: 4px;
}

/* ============================================
   RADIO BUTTONS - LARGER
   ============================================ */
QRadioButton {
    font-size: 14px;
    color: #495057;
    spacing: 8px;
    padding: 6px;
    font-family: "Segoe UI", "Arial", sans-serif;
}

QRadioButton::indicator {
    width: 18px;
    height: 18px;
}

QRadioButton::indicator:unchecked {
    border: 2px solid #adb5bd;
    border-radius: 9px;
    background-color: white;
}

QRadioButton::indicator:checked {
    border: 2px solid #4a90e2;
    border-radius: 9px;
    background-color: #4a90e2;
}

/* ============================================
   SPINBOX - LARGER
   ============================================ */
QSpinBox {
    padding: 12px;
    border: 1px solid #ced4da;
    border-radius: 4px;
    font-size: 14px;
    background-color: white;
    min-width: 100px;
    color: #495057;
    font-family: "Segoe UI", "Arial", sans-serif;
}

QSpinBox:focus {
    border: 2px solid #4a90e2;
}

QSpinBox::up-button, QSpinBox::down-button {
    width: 20px;
    border: none;
    background-color: #f8f9fa;
}

QSpinBox::up-button:hover, QSpinBox::down-button:hover {
    background-color: #e9ecef;
}

/* ============================================
   SCROLLBAR - MINIMAL
   ============================================ */
QScrollBar:vertical {
    border: none;
    background-color: #f8f9fa;
    width: 10px;
    margin: 0px;
}

QScrollBar::handle:vertical {
    background-color: #ced4da;
    border-radius: 5px;
    min-height: 30px;
}

QScrollBar::handle:vertical:hover {
    background-color: #adb5bd;
}

QScrollBar::add-line:vertical, QScrollBar::sub-line:vertical {
    height: 0px;
}

QScrollBar::add-page:vertical, QScrollBar::sub-page:vertical {
    background: none;
}
"""
