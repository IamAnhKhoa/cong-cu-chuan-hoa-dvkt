"""
UI Styling for Technical Catalog Standardization Tool
Modern, clean, and accessible design for low-tech environments
"""

MAIN_STYLE = """
QMainWindow {
    background-color: #f5f5f5;
}

QLabel {
    font-size: 13px;
    color: #333;
    font-family: "Arial", "Tahoma", "Segoe UI", sans-serif;
}

QLabel#title {
    font-size: 20px;
    font-weight: bold;
    color: #1976D2;
    padding: 10px;
}

QLabel#section_title {
    font-size: 15px;
    font-weight: bold;
    color: #555;
    padding: 8px 0px;
}

QPushButton {
    background-color: #1976D2;
    color: white;
    border: none;
    padding: 12px 24px;
    font-size: 14px;
    font-weight: bold;
    border-radius: 6px;
    min-width: 120px;
}

QPushButton:hover {
    background-color: #1565C0;
}

QPushButton:pressed {
    background-color: #0D47A1;
}

QPushButton:disabled {
    background-color: #BDBDBD;
    color: #757575;
}

QPushButton#success_button {
    background-color: #4CAF50;
}

QPushButton#success_button:hover {
    background-color: #45a049;
}

QPushButton#warning_button {
    background-color: #FF9800;
}

QPushButton#warning_button:hover {
    background-color: #F57C00;
}

QLineEdit, QTextEdit {
    padding: 8px;
    border: 2px solid #ddd;
    border-radius: 4px;
    font-size: 13px;
    background-color: white;
}

QLineEdit:focus, QTextEdit:focus {
    border: 2px solid #1976D2;
}

QLineEdit:disabled, QTextEdit:disabled {
    background-color: #f0f0f0;
    color: #999;
}

QProgressBar {
    border: 2px solid #ddd;
    border-radius: 6px;
    text-align: center;
    font-size: 12px;
    font-weight: bold;
    background-color: white;
    height: 25px;
}

QProgressBar::chunk {
    background-color: #4CAF50;
    border-radius: 4px;
}

QGroupBox {
    font-size: 14px;
    font-weight: bold;
    border: 2px solid #ddd;
    border-radius: 8px;
    margin-top: 12px;
    padding-top: 12px;
    background-color: white;
}

QGroupBox::title {
    subcontrol-origin: margin;
    subcontrol-position: top left;
    padding: 4px 10px;
    color: #1976D2;
}

QTableWidget {
    border: 2px solid #ddd;
    border-radius: 4px;
    gridline-color: #e0e0e0;
    font-size: 12px;
    background-color: white;
}

QTableWidget::item {
    padding: 5px;
}

QHeaderView::section {
    background-color: #1976D2;
    color: white;
    padding: 8px;
    border: none;
    font-weight: bold;
    font-size: 12px;
}

QScrollBar:vertical {
    border: none;
    background-color: #f0f0f0;
    width: 12px;
    margin: 0px;
}

QScrollBar::handle:vertical {
    background-color: #BDBDBD;
    border-radius: 6px;
    min-height: 20px;
}

QScrollBar::handle:vertical:hover {
    background-color: #9E9E9E;
}

QScrollBar::add-line:vertical, QScrollBar::sub-line:vertical {
    height: 0px;
}

QComboBox {
    padding: 8px;
    border: 2px solid #ddd;
    border-radius: 4px;
    font-size: 13px;
    background-color: white;
}

QComboBox:focus {
    border: 2px solid #1976D2;
}

QComboBox::drop-down {
    border: none;
    width: 30px;
}

QComboBox::down-arrow {
    image: none;
    border-left: 5px solid transparent;
    border-right: 5px solid transparent;
    border-top: 5px solid #666;
    margin-right: 5px;
}

QSpinBox {
    padding: 8px;
    border: 2px solid #ddd;
    border-radius: 4px;
    font-size: 13px;
    background-color: white;
}

QSpinBox:focus {
    border: 2px solid #1976D2;
}
"""
