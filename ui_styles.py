"""
UI Styling for Technical Catalog Standardization Tool
"Sophisticated & Simple" (Sang trọng, Tinh tế, Đơn giản)
Compact Mode (No Scroll)
"""

MAIN_STYLE = """
/* ============================================
   GLOBAL SETTINGS
   ============================================ */
QMainWindow {
    background-color: #f8fafc; /* Very light slate gray/white */
}

QWidget {
    font-family: "Segoe UI", "Inter", "system-ui", sans-serif;
    color: #334155; /* Slate 700 - Softer than black */
}

/* ============================================
   TYPOGRAPHY
   ============================================ */
QLabel {
    font-size: 13px;
    color: #475569; /* Slate 600 */
}

QLabel#title {
    font-size: 18px; /* Refined size */
    font-weight: 700;
    color: #1e293b; /* Slate 800 */
    padding: 12px;
    background-color: white;
    border-bottom: 1px solid #e2e8f0; /* Subtle divider */
    letter-spacing: 0.5px;
    text-transform: uppercase;
}

/* ============================================
   BUTTONS (Modern & Flat)
   ============================================ */
QPushButton {
    background-color: #3b82f6; /* Modern Blue */
    color: white;
    border: none;
    border-radius: 4px;
    padding: 8px 16px;
    font-size: 13px;
    font-weight: bold; /* Changed to bold */
}

QPushButton:hover {
    background-color: #2563eb;
    margin-top: -1px; /* Subtle lift */
}

QPushButton:pressed {
    background-color: #1d4ed8;
    margin-top: 1px;
}

QPushButton:disabled {
    background-color: #cbd5e1; /* Slate 300 */
    color: #94a3b8; /* Slate 400 */
}

/* Success Button (Primary Action) */
QPushButton#success_button {
    background-color: #10b981; /* Emerald 500 */
    font-size: 15px; /* Larger font */
    font-weight: 800; /* Extra Bold */
    text-transform: uppercase; /* Uppercase for impact */
    padding: 12px 24px;
    letter-spacing: 1px;
}

QPushButton#success_button:hover {
    background-color: #059669; /* Emerald 600 */
}

QPushButton#success_button:disabled {
    background-color: #cbd5e1; /* Slate 300 - Standard disabled gray */
    color: #94a3b8; /* Slate 400 - Low contrast text */
    border: none;
}

/* ============================================
   INPUT FIELDS (Clean & Minimal)
   ============================================ */
QLineEdit, QTextEdit {
    background-color: white;
    border: 1px solid #cbd5e1; /* Slate 300 */
    border-radius: 4px;
    padding: 7px 10px;
    font-size: 13px;
    selection-background-color: #bfdbfe;
    color: #334155;
}

QLineEdit:focus, QTextEdit:focus {
    border: 1px solid #3b82f6; /* Blue 500 */
    background-color: #fff;
    /* Qt doesn't support box-shadow well, simulated via border */
}

QLineEdit:disabled, QTextEdit:disabled {
    background-color: #f1f5f9; /* Slate 100 */
    color: #94a3b8;
    border: 1px solid #e2e8f0;
}

/* ============================================
   GROUP BOXES (Sophisticated Card Style)
   ============================================ */
QGroupBox {
    font-size: 13px;
    font-weight: 700;
    color: #1e293b; /* Dark Slate Header */
    border: 1px solid #e2e8f0; /* Slate 200 */
    border-radius: 6px;
    margin-top: 8px; /* Compact */
    padding-top: 20px; /* Space for title */
    padding-bottom: 10px;
    padding-left: 12px;
    padding-right: 12px;
    background-color: white;
}

QGroupBox::title {
    subcontrol-origin: margin;
    subcontrol-position: top left;
    padding: 2px 8px;
    left: 8px;
    color: #1e293b;
    background-color: transparent; /* Seamless title */
}

/* ============================================
   RADIO BUTTONS
   ============================================ */
QRadioButton {
    font-size: 13px;
    color: #475569;
    spacing: 8px;
}

QRadioButton::indicator {
    width: 16px;
    height: 16px;
    border-radius: 8px;
    border: 1px solid #cbd5e1;
    background-color: #f8fafc;
}

QRadioButton::indicator:checked {
    background-color: #3b82f6; /* Solid Blue Center */
    border: 3px solid white; /* White spacing */
    outline: 1px solid #3b82f6; /* Outer ring (if supported) or just rely on contrast */
    border-radius: 8px; /* Ensure circular */
}

QRadioButton::indicator:checked:hover {
    background-color: #2563eb;
}

QRadioButton::indicator:hover {
    border-color: #3b82f6;
}

/* ============================================
   SPINBOX
   ============================================ */
QSpinBox {
    padding: 6px 10px;
    border: 1px solid #cbd5e1;
    border-radius: 4px;
    color: #334155;
    background-color: white;
}

QSpinBox:focus {
    border: 1px solid #3b82f6;
}
"""
