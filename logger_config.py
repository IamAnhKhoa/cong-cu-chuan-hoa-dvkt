# -*- coding: utf-8 -*-
"""
Centralized Logging Configuration
Provides logging functionality for the entire application
"""

import logging
import os
from datetime import datetime


class AppLogger:
    """Centralized logger for the application"""
    
    def __init__(self, log_dir="logs"):
        self.log_dir = log_dir
        self.logger = None
        self.log_file = None
        
    def setup_logger(self, name="TechnicalCatalog"):
        """Setup and return a configured logger"""
        # Create logs directory if it doesn't exist
        if not os.path.exists(self.log_dir):
            os.makedirs(self.log_dir)
        
        # Create log filename with timestamp
        timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
        self.log_file = os.path.join(self.log_dir, f"processing_{timestamp}.txt")
        
        # Create logger
        self.logger = logging.getLogger(name)
        self.logger.setLevel(logging.DEBUG)
        
        # Clear existing handlers
        self.logger.handlers.clear()
        
        # Create file handler
        file_handler = logging.FileHandler(self.log_file, encoding='utf-8')
        file_handler.setLevel(logging.DEBUG)
        
        # Create console handler (optional)
        console_handler = logging.StreamHandler()
        console_handler.setLevel(logging.INFO)
        
        # Create formatter
        formatter = logging.Formatter(
            '%(asctime)s - %(levelname)s - %(message)s',
            datefmt='%Y-%m-%d %H:%M:%S'
        )
        
        file_handler.setFormatter(formatter)
        console_handler.setFormatter(formatter)
        
        # Add handlers to logger
        self.logger.addHandler(file_handler)
        self.logger.addHandler(console_handler)
        
        return self.logger
    
    def get_log_file_path(self):
        """Return the current log file path"""
        return self.log_file
    
    def log_processing_start(self, input_file, process_type):
        """Log the start of processing"""
        if self.logger:
            self.logger.info("="*60)
            self.logger.info("BẮT ĐẦU XỬ LÝ")
            self.logger.info(f"File đầu vào: {input_file}")
            self.logger.info(f"Loại xử lý: {process_type}")
            self.logger.info("="*60)
    
    def log_match(self, input_name, matched_name, score):
        """Log a successful match"""
        if self.logger:
            self.logger.debug(f"✓ Khớp: '{input_name}' → '{matched_name}' (Độ chính xác: {score}%)")
    
    def log_no_match(self, input_name, threshold):
        """Log when no match is found"""
        if self.logger:
            self.logger.warning(f"✗ Không khớp: '{input_name}' (Ngưỡng: {threshold}%)")
    
    def log_processing_end(self, stats):
        """Log processing completion with statistics"""
        if self.logger:
            self.logger.info("="*60)
            self.logger.info("KẾT THÚC XỬ LÝ")
            self.logger.info(f"Tổng số bản ghi: {stats.get('total', 0)}")
            self.logger.info(f"Khớp thành công: {stats.get('matched', 0)}")
            self.logger.info(f"Không khớp: {stats.get('unmatched', 0)}")
            if stats.get('total', 0) > 0:
                match_rate = (stats.get('matched', 0) / stats.get('total', 0)) * 100
                self.logger.info(f"Tỷ lệ khớp: {match_rate:.1f}%")
            self.logger.info("="*60)
    
    def log_error(self, error_msg):
        """Log an error"""
        if self.logger:
            self.logger.error(f"LỖI: {error_msg}")
