# -*- coding: utf-8 -*-
"""
Technical Catalog Standardization Tool
Main Application with PyQt5 GUI

Features:
- Easy file/folder selection
- Process type selection
- Fuzzy matching with configurable threshold
- Progress tracking and results preview
"""

import sys
import os
from PyQt5.QtWidgets import (QApplication, QMainWindow, QWidget, QVBoxLayout, 
                             QHBoxLayout, QLabel, QPushButton, QLineEdit, 
                             QFileDialog, QProgressBar, QTextEdit, QGroupBox,
                             QRadioButton, QSpinBox, QTableWidget, QTableWidgetItem,
                             QMessageBox, QSplitter)
from PyQt5.QtCore import Qt, QThread, pyqtSignal, QSettings
from PyQt5.QtGui import QFont
from processor import CatalogProcessor
from ui_styles import MAIN_STYLE


class ProcessingThread(QThread):
    """Background thread for processing files"""
    progress = pyqtSignal(str)
    finished = pyqtSignal(bool, str, dict)
    
    def __init__(self, processor, process_type, input_file, output_file, threshold):
        super().__init__()
        self.processor = processor
        self.process_type = process_type
        self.input_file = input_file
        self.output_file = output_file
        self.threshold = threshold
    
    def run(self):
        """Execute processing in background"""
        self.progress.emit("Đang xử lý file...")
        
        if self.process_type == 1:
            success, message, stats = self.processor.process_quy_trinh_file(
                self.input_file, self.output_file, self.threshold
            )
        else:  # process_type == 2
            success, message, stats = self.processor.process_gia_hdnd_file(
                self.input_file, self.output_file, self.threshold
            )
        
        self.finished.emit(success, message, stats)


class MainWindow(QMainWindow):
    """Main application window"""
    
    def __init__(self):
        super().__init__()
        self.processor = CatalogProcessor()
        self.process_type = 1  # Default to process 1
        self.settings = QSettings("MyCompany", "TechnicalCatalogTool")
        self.init_ui()
        self.load_settings()
        
    def init_ui(self):
        """Initialize user interface"""
        self.setWindowTitle("Công cụ Chuẩn hóa Danh mục Kỹ thuật")
        self.setGeometry(100, 100, 1000, 700)
        
        # Central widget
        central_widget = QWidget()
        self.setCentralWidget(central_widget)
        main_layout = QVBoxLayout(central_widget)
        main_layout.setSpacing(15)
        main_layout.setContentsMargins(20, 20, 20, 20)
        
        # Title
        title = QLabel("CÔNG CỤ CHUẨN HÓA DANH MỤC KỸ THUẬT")
        title.setObjectName("title")
        title.setAlignment(Qt.AlignCenter)
        main_layout.addWidget(title)
        
        # Reference files section
        ref_group = self.create_reference_section()
        main_layout.addWidget(ref_group)
        
        # Process type selection
        process_group = self.create_process_section()
        main_layout.addWidget(process_group)
        
        # Input file section
        input_group = self.create_input_section()
        main_layout.addWidget(input_group)
        
        # Settings section
        settings_group = self.create_settings_section()
        main_layout.addWidget(settings_group)
        
        # Process button
        self.process_btn = QPushButton("⚙ BẮT ĐẦU XỬ LÝ")
        self.process_btn.setObjectName("success_button")
        self.process_btn.setMinimumHeight(50)
        self.process_btn.clicked.connect(self.start_processing)
        self.process_btn.setEnabled(False)
        main_layout.addWidget(self.process_btn)
        
        # Progress section
        progress_group = self.create_progress_section()
        main_layout.addWidget(progress_group)
        
        # Apply styles
        self.setStyleSheet(MAIN_STYLE)
        
    def create_reference_section(self):
        """Create reference files selection section"""
        group = QGroupBox("1. Chọn thư mục chứa file gốc")
        layout = QVBoxLayout()
        
        desc = QLabel("Thư mục phải chứa 3 file: QUY_TRINH_DVKT_BYT.xlsx, GIA_HDND.xlsx và DVKT_GIA_MAX.xlsx")
        desc.setWordWrap(True)
        layout.addWidget(desc)
        
        folder_layout = QHBoxLayout()
        self.ref_folder_input = QLineEdit()
        self.ref_folder_input.setPlaceholderText("Chưa chọn thư mục...")
        self.ref_folder_input.setReadOnly(True)
        folder_layout.addWidget(self.ref_folder_input, 3)
        
        browse_btn = QPushButton("📁 Chọn thư mục")
        browse_btn.clicked.connect(self.browse_reference_folder)
        folder_layout.addWidget(browse_btn, 1)
        
        layout.addLayout(folder_layout)
        
        self.ref_status = QLabel("")
        self.ref_status.setWordWrap(True)
        layout.addWidget(self.ref_status)
        
        group.setLayout(layout)
        return group
    
    def create_process_section(self):
        """Create process type selection section"""
        group = QGroupBox("2. Chọn loại xử lý")
        layout = QVBoxLayout()
        
        self.process1_radio = QRadioButton("Xử lý QUY_TRINH_DVKT_BYT → DM_DVKT_TT21_cong")
        self.process1_radio.setChecked(True)
        self.process1_radio.toggled.connect(lambda: self.set_process_type(1))
        layout.addWidget(self.process1_radio)
        
        desc1 = QLabel("   → Tạo file với các cột: STT, MA_DICH_VU, TEN_DICH_VU, DON_GIA, QUY_TRINH, CSKCB_CGKT, CSKCB_CLS")
        desc1.setStyleSheet("color: #666; font-size: 11px;")
        layout.addWidget(desc1)
        
        self.process2_radio = QRadioButton("Xử lý GIA_HDND → File chuẩn (kết hợp cả 2 file)")
        self.process2_radio.toggled.connect(lambda: self.set_process_type(2))
        layout.addWidget(self.process2_radio)
        
        desc2 = QLabel("   → Tạo file với các cột: STT, MA_TUONG_DUONG, TEN_DVKT_PHEDUYET, TEN_DVKT_GIA, PHAN_LOAI_PTTT, DON_GIA, GHI_CHU, QUYET_DINH, QUY_TRINH, TU_NGAY, DEN_NGAY, CSKCB_CGKT, CSKCB_CLS")
        desc2.setStyleSheet("color: #666; font-size: 11px;")
        layout.addWidget(desc2)
        
        group.setLayout(layout)
        return group
    
    def create_input_section(self):
        """Create input file selection section"""
        group = QGroupBox("3. Chọn file đầu vào")
        layout = QVBoxLayout()
        
        desc = QLabel('File Excel phải có cột "Tên kỹ thuật"')
        desc.setWordWrap(True)
        layout.addWidget(desc)
        
        file_layout = QHBoxLayout()
        self.input_file = QLineEdit()
        self.input_file.setPlaceholderText("Chưa chọn file...")
        self.input_file.setReadOnly(True)
        file_layout.addWidget(self.input_file, 3)
        
        browse_btn = QPushButton("📄 Chọn file Excel")
        browse_btn.clicked.connect(self.browse_input_file)
        file_layout.addWidget(browse_btn, 1)
        
        layout.addLayout(file_layout)
        
        # Output file
        output_layout = QHBoxLayout()
        output_label = QLabel("File kết quả:")
        output_layout.addWidget(output_label)
        
        self.output_file = QLineEdit()
        self.output_file.setPlaceholderText("Tự động tạo tên file...")
        output_layout.addWidget(self.output_file, 3)
        
        output_btn = QPushButton("💾 Chọn vị trí lưu")
        output_btn.clicked.connect(self.browse_output_file)
        output_layout.addWidget(output_btn, 1)
        
        layout.addLayout(output_layout)
        
        group.setLayout(layout)
        return group
    
    def create_settings_section(self):
        """Create settings section"""
        group = QGroupBox("4. Cài đặt")
        layout = QHBoxLayout()
        
        threshold_label = QLabel("Ngưỡng khớp tên (%):")
        threshold_label.setToolTip("Mức độ tương đồng tối thiểu để coi là khớp (80-100). Càng cao càng chính xác nhưng có thể bỏ sót.")
        layout.addWidget(threshold_label)
        
        self.threshold_spin = QSpinBox()
        self.threshold_spin.setRange(50, 100)
        self.threshold_spin.setValue(80)
        self.threshold_spin.setSuffix("%")
        layout.addWidget(self.threshold_spin)
        
        layout.addStretch()
        
        group.setLayout(layout)
        return group
    
    def create_progress_section(self):
        """Create progress and results section"""
        group = QGroupBox("5. Kết quả")
        layout = QVBoxLayout()
        
        # Progress bar
        self.progress_bar = QProgressBar()
        self.progress_bar.setRange(0, 100)
        self.progress_bar.setValue(0)
        layout.addWidget(self.progress_bar)
        
        # Status text
        self.status_text = QTextEdit()
        self.status_text.setReadOnly(True)
        self.status_text.setMaximumHeight(150)
        self.status_text.setPlaceholderText("Chờ xử lý...")
        layout.addWidget(self.status_text)
        
        group.setLayout(layout)
        return group
    
    def set_process_type(self, process_type):
        """Set the processing type"""
        self.process_type = process_type
        self.update_process_button_state()
    
    def load_settings(self):
        """Load saved settings"""
        saved_folder = self.settings.value("reference_folder")
        if saved_folder and os.path.exists(saved_folder):
            self.ref_folder_input.setText(saved_folder)
            success, message = self.processor.load_reference_files(saved_folder)
            if success:
                self.ref_status.setText(f"✓ {message}")
                self.ref_status.setStyleSheet("color: green; font-weight: bold;")
            else:
                self.ref_status.setText(f"✗ {message}")
                self.ref_status.setStyleSheet("color: red; font-weight: bold;")
            self.update_process_button_state()

    def browse_reference_folder(self):
        """Browse for reference files folder"""
        folder = QFileDialog.getExistingDirectory(
            self, 
            "Chọn thư mục chứa file gốc",
            self.ref_folder_input.text() or os.path.expanduser("~")
        )
        
        if folder:
            self.ref_folder_input.setText(folder)
            self.settings.setValue("reference_folder", folder)
            success, message = self.processor.load_reference_files(folder)
            
            if success:
                self.ref_status.setText(f"✓ {message}")
                self.ref_status.setStyleSheet("color: green; font-weight: bold;")
            else:
                self.ref_status.setText(f"✗ {message}")
                self.ref_status.setStyleSheet("color: red; font-weight: bold;")
            
            self.update_process_button_state()
    
    def browse_input_file(self):
        """Browse for input Excel file"""
        file, _ = QFileDialog.getOpenFileName(
            self,
            "Chọn file Excel đầu vào",
            os.path.expanduser("~"),
            "Excel Files (*.xlsx *.xls)"
        )
        
        if file:
            self.input_file.setText(file)
            # Auto-generate output filename
            if not self.output_file.text():
                base_dir = os.path.dirname(file)
                base_name = os.path.splitext(os.path.basename(file))[0]
                if self.process_type == 1:
                    output_name = f"{base_name}_DM_DVKT_TT21_cong.xlsx"
                else:
                    output_name = f"{base_name}_standardized.xlsx"
                self.output_file.setText(os.path.join(base_dir, output_name))
            
            self.update_process_button_state()
    
    def browse_output_file(self):
        """Browse for output Excel file location"""
        file, _ = QFileDialog.getSaveFileName(
            self,
            "Chọn vị trí lưu file kết quả",
            self.output_file.text() or os.path.expanduser("~"),
            "Excel Files (*.xlsx)"
        )
        
        if file:
            if not file.endswith('.xlsx'):
                file += '.xlsx'
            self.output_file.setText(file)
    
    def update_process_button_state(self):
        """Enable/disable process button based on input state"""
        has_reference = self.processor.ref_quy_trinh_df is not None
        has_input = bool(self.input_file.text())
        has_output = bool(self.output_file.text())
        
        self.process_btn.setEnabled(has_reference and has_input and has_output)
    
    def start_processing(self):
        """Start the processing in background thread"""
        # Disable controls
        self.process_btn.setEnabled(False)
        self.progress_bar.setValue(50)
        self.status_text.clear()
        self.status_text.append("🔄 Đang xử lý...")
        
        # Create and start processing thread
        self.thread = ProcessingThread(
            self.processor,
            self.process_type,
            self.input_file.text(),
            self.output_file.text(),
            self.threshold_spin.value()
        )
        self.thread.progress.connect(self.update_progress)
        self.thread.finished.connect(self.processing_finished)
        self.thread.start()
    
    def update_progress(self, message):
        """Update progress display"""
        self.status_text.append(message)
    
    def processing_finished(self, success, message, stats):
        """Handle processing completion"""
        self.progress_bar.setValue(100 if success else 0)
        
        if success:
            self.status_text.append(f"\n✓ {message}")
            self.status_text.append(f"\n📁 File đã lưu: {self.output_file.text()}")
            
            # Show match details if available
            if stats and 'match_details' in stats:
                self.status_text.append("\n--- Chi tiết khớp tên ---")
                for detail in stats['match_details'][:10]:  # Show first 10
                    if detail['score'] > 0:
                        self.status_text.append(
                            f"  • {detail['input'][:50]}... → {detail['matched'][:50]}... ({detail['score']}%)"
                        )
                    else:
                        self.status_text.append(
                            f"  • {detail['input'][:50]}... → KHÔNG TÌM THẤY"
                        )
                if len(stats['match_details']) > 10:
                    self.status_text.append(f"  ... và {len(stats['match_details']) - 10} mục khác")
            
            # Show success dialog
            QMessageBox.information(
                self,
                "Hoàn thành",
                f"Xử lý thành công!\n\n{message}\n\nFile kết quả:\n{self.output_file.text()}"
            )
        else:
            self.status_text.append(f"\n✗ Lỗi: {message}")
            QMessageBox.critical(
                self,
                "Lỗi",
                f"Xử lý thất bại!\n\n{message}"
            )
        
        # Re-enable controls
        self.process_btn.setEnabled(True)


def main():
    """Main application entry point"""
    app = QApplication(sys.argv)
    
    # Set application-wide font with Vietnamese Unicode support
    # Using Arial which has excellent Vietnamese character support
    font = QFont("Arial", 10)
    font.setStyleHint(QFont.SansSerif)
    app.setFont(font)
    
    window = MainWindow()
    window.show()
    
    sys.exit(app.exec_())


if __name__ == '__main__':
    main()
