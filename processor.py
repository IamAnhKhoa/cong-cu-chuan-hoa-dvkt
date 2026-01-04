# -*- coding: utf-8 -*-
"""
Data Processing Module for Technical Catalog Standardization
Handles Excel file reading, matching, and transformation
"""

import pandas as pd
from fuzzywuzzy import fuzz, process
from typing import Dict, List, Tuple, Optional
import os


class CatalogProcessor:
    """Process technical catalog data with fuzzy matching"""
    
    def __init__(self):
        self.ref_quy_trinh_df = None
        self.ref_gia_hdnd_df = None
        self.ref_dvkt_gia_max_df = None
        self.reference_folder = None
        
    def load_reference_files(self, folder_path: str) -> Tuple[bool, str]:
        """Load reference Excel files from folder"""
        try:
            quy_trinh_path = os.path.join(folder_path, "QUY_TRINH_DVKT_BYT.xlsx")
            gia_hdnd_path = os.path.join(folder_path, "GIA_HDND.xlsx")
            max_gia_path = os.path.join(folder_path, "DVKT_GIA_MAX.xlsx")
            
            if not os.path.exists(quy_trinh_path):
                return False, f"Không tìm thấy file: QUY_TRINH_DVKT_BYT.xlsx"
            
            if not os.path.exists(gia_hdnd_path):
                return False, f"Không tìm thấy file: GIA_HDND.xlsx"
            
            # Helper to load and find correct sheet
            def load_with_correct_sheet(path, expected_col_keyword):
                xl = pd.ExcelFile(path, engine='openpyxl')
                target_sheet = None
                
                # Check known sheet names first
                if 'QUY TRINH' in xl.sheet_names:
                    target_sheet = 'QUY TRINH'
                
                # If not found or just to be safe, search for content
                if target_sheet is None:
                    for sheet in xl.sheet_names:
                        df = pd.read_excel(path, sheet_name=sheet, nrows=5, engine='openpyxl')
                        # Check if any column contains the keyword
                        col_match = any(expected_col_keyword.lower() in str(col).lower() for col in df.columns)
                        if col_match:
                            target_sheet = sheet
                            break
                            
                # Fallback to first sheet if nothing found (will likely fail later but better than crash)
                if target_sheet is None:
                    target_sheet = xl.sheet_names[0]
                    
                return pd.read_excel(path, sheet_name=target_sheet, engine='openpyxl')

            # Load files with smart sheet detection
            self.ref_quy_trinh_df = load_with_correct_sheet(quy_trinh_path, 'dịch vụ')
            self.ref_gia_hdnd_df = pd.read_excel(gia_hdnd_path, engine='openpyxl')
            
            # Optional third file
            if os.path.exists(max_gia_path):
                 self.ref_dvkt_gia_max_df = pd.read_excel(max_gia_path, engine='openpyxl')
                 max_msg = " và DVKT_GIA_MAX"
            else:
                 self.ref_dvkt_gia_max_df = None
                 max_msg = ""
            
            self.reference_folder = folder_path
            
            return True, f"Đã tải {len(self.ref_quy_trinh_df)} dịch vụ từ QUY_TRINH, {len(self.ref_gia_hdnd_df)} dịch vụ từ GIA_HDND{max_msg}"
            
        except Exception as e:
            return False, f"Lỗi khi tải file tham chiếu: {str(e)}"
    
    def find_best_match(self, service_name: str, reference_names: List[str], threshold: int = 80) -> Optional[Tuple[str, int]]:
        """Find best matching service name using fuzzy matching"""
        if not service_name or not reference_names:
            return None
        
        result = process.extractOne(service_name, reference_names, scorer=fuzz.token_sort_ratio)
        
        if result and result[1] >= threshold:
            return result[0], result[1]
        return None
    
    def process_quy_trinh_file(self, input_file: str, output_file: str, threshold: int = 80) -> Tuple[bool, str, Dict]:
        """
        Process file 1: QUY_TRINH_DVKT_BYT → DM_DVKT_TT21_cong
        
        Input columns: Must contain "Tên kỹ thuật"
        Output columns: STT, MA_DICH_VU, TEN_DICH_VU, DON_GIA, QUY_TRINH, CSKCB_CGKT, CSKCB_CLS
        """
        try:
            # Validate reference data loaded
            if self.ref_quy_trinh_df is None:
                return False, "Chưa tải file tham chiếu. Vui lòng chọn thư mục chứa file gốc.", {}
            
            # Read input file with proper encoding
            input_df = pd.read_excel(input_file, engine='openpyxl')
            
            # Check for required column
            name_column = None
            input_df.columns = input_df.columns.str.strip()  # Normalize headers
            
            possible_headers = ['tên kỹ thuật', 'tên dịch vụ', 'tên dvkt', 'ten dich vu', 'ten ky thuat', 'tên']
            
            # Try exact match first
            for col in input_df.columns:
                if col.lower() in possible_headers:
                    name_column = col
                    break
            
            # Fuzzy header match if exact match fails
            if name_column is None:
                for col in input_df.columns:
                    col_lower = col.lower()
                    if 'tên' in col_lower and ('thuật' in col_lower or 'dịch vụ' in col_lower or 'dvkt' in col_lower):
                        name_column = col
                        break
            
            if name_column is None:
                return False, f'Không tìm thấy cột Tên dịch vụ/kỹ thuật. Các cột có trong file: {", ".join(input_df.columns)}', {}
            
            # Prepare reference names
            ref_name_column = None
            self.ref_quy_trinh_df.columns = self.ref_quy_trinh_df.columns.str.strip()
            
            for col in self.ref_quy_trinh_df.columns:
                if 'TEN_DICH_VU' in col or ('tên' in col.lower() and 'vụ' in col.lower()):
                    ref_name_column = col
                    break
                    
            if ref_name_column is None:
                return False, "Không tìm thấy cột tên dịch vụ trong file QUY_TRINH_DVKT_BYT.xlsx", {}
            
            reference_names = self.ref_quy_trinh_df[ref_name_column].dropna().astype(str).str.strip().tolist()
            
            # Process matching
            results = []
            stats = {
                'total': len(input_df),
                'matched': 0,
                'unmatched': 0,
                'match_details': []
            }
            
            for idx, row in input_df.iterrows():
                raw_name = row[name_column]
                # Skip if empty or NaN
                if pd.isna(raw_name) or str(raw_name).strip() == "":
                    continue
                    
                service_name = str(raw_name).strip()
                
                # Find best match
                match_result = self.find_best_match(service_name, reference_names, threshold)
                
                if match_result:
                    matched_name, score = match_result
                    # Find the full row in reference
                    ref_row = self.ref_quy_trinh_df[self.ref_quy_trinh_df[ref_name_column] == matched_name].iloc[0]
                    
                    # Build output row
                    output_row = {
                        'STT': idx + 1,
                        'MA_DICH_VU': ref_row.get('MA_DICH_VU', ''),
                        'TEN_DICH_VU': ref_row.get(ref_name_column, matched_name),
                        'DON_GIA': ref_row.get('DON_GIA', ''),
                        'QUY_TRINH': ref_row.get('QUY_TRINH', ''),
                        'CSKCB_CGKT': '',  # Empty as required
                        'CSKCB_CLS': ''    # Empty as required
                    }
                    results.append(output_row)
                    stats['matched'] += 1
                    stats['match_details'].append({
                        'input': service_name,
                        'matched': matched_name,
                        'score': score
                    })
                else:
                    # No match found
                    output_row = {
                        'STT': idx + 1,
                        'MA_DICH_VU': '',
                        'TEN_DICH_VU': service_name,
                        'DON_GIA': '',
                        'QUY_TRINH': '',
                        'CSKCB_CGKT': '',
                        'CSKCB_CLS': ''
                    }
                    results.append(output_row)
                    stats['unmatched'] += 1
                    stats['match_details'].append({
                        'input': service_name,
                        'matched': 'KHÔNG TÌM THẤY',
                        'score': 0
                    })
            
            # Create output DataFrame
            output_df = pd.DataFrame(results)
            
            # Save to Excel with proper encoding
            output_df.to_excel(output_file, index=False, engine='openpyxl')
            
            message = f"Xử lý thành công!\n"
            message += f"Tổng số: {stats['total']}\n"
            message += f"Khớp: {stats['matched']}\n"
            message += f"Không khớp: {stats['unmatched']}"
            
            return True, message, stats
            
        except Exception as e:
            return False, f"Lỗi khi xử lý file: {str(e)}", {}
    
    def process_gia_hdnd_file(self, input_file: str, output_file: str, threshold: int = 80) -> Tuple[bool, str, Dict]:
        """
        Process file 2: GIA_HDND → Standardized output
        
        Output columns: STT, MA_TUONG_DUONG, TEN_DVKT_PHEDUYET, TEN_DVKT_GIA, PHAN_LOAI_PTTT,
                       DON_GIA, GHI_CHU, QUYET_DINH, QUY_TRINH, TU_NGAY, DEN_NGAY, CSKCB_CGKT, CSKCB_CLS
        """
        try:
            # Validate reference data loaded
            if self.ref_gia_hdnd_df is None or self.ref_quy_trinh_df is None:
                return False, "Chưa tải đầy đủ file tham chiếu. Vui lòng chọn thư mục chứa file gốc.", {}
            
            # Read input file with proper encoding
            input_df = pd.read_excel(input_file, engine='openpyxl')
            
            # Check for required column
            name_column = None
            input_df.columns = input_df.columns.str.strip()  # Normalize headers
            
            possible_headers = ['tên kỹ thuật', 'tên dịch vụ', 'tên dvkt', 'ten dich vu', 'ten ky thuat', 'tên']
            
            for col in input_df.columns:
                if col.lower() in possible_headers:
                    name_column = col
                    break
            
            if name_column is None:
                for col in input_df.columns:
                     col_lower = col.lower()
                     if 'tên' in col_lower and ('thuật' in col_lower or 'dịch vụ' in col_lower or 'dvkt' in col_lower):
                        name_column = col
                        break
            
            if name_column is None:
                return False, f'Không tìm thấy cột Tên dịch vụ/kỹ thuật. Các cột có trong file: {", ".join(input_df.columns)}', {}
            
            # Prepare reference names from GIA_HDND
            ref_name_column = None
            self.ref_gia_hdnd_df.columns = self.ref_gia_hdnd_df.columns.str.strip()
            
            for col in self.ref_gia_hdnd_df.columns:
                if 'Tên dịch vụ kỹ thuật' in col or ('TT23' in col and 'Tên' in col):
                    ref_name_column = col
                    break
                    
            if ref_name_column is None:
                return False, "Không tìm thấy cột tên dịch vụ trong file GIA_HDND.xlsx", {}
            
            reference_names = self.ref_gia_hdnd_df[ref_name_column].dropna().astype(str).str.strip().tolist()

            # Prepare reference names from DVKT_GIA_MAX if available
            max_reference_names = []
            max_ref_name_col = None
            if self.ref_dvkt_gia_max_df is not None:
                self.ref_dvkt_gia_max_df.columns = self.ref_dvkt_gia_max_df.columns.str.strip()
                for col in self.ref_dvkt_gia_max_df.columns:
                    if 'TEN_DVKT_GIA' in col or ('Tên' in col and 'Gia' in col):
                        max_ref_name_col = col
                        break
                if max_ref_name_col:
                    max_reference_names = self.ref_dvkt_gia_max_df[max_ref_name_col].dropna().astype(str).str.strip().tolist()

            
            # Also prepare QUY_TRINH reference for combining
            quy_trinh_name_col = None
            self.ref_quy_trinh_df.columns = self.ref_quy_trinh_df.columns.str.strip()
            
            for col in self.ref_quy_trinh_df.columns:
                if 'TEN_DICH_VU' in col or ('tên' in col.lower() and 'vụ' in col.lower()):
                    quy_trinh_name_col = col
                    break
            
            # Process matching
            results = []
            stats = {
                'total': len(input_df),
                'matched': 0,
                'unmatched': 0,
                'match_details': []
            }
            
            for idx, row in input_df.iterrows():
                raw_name = row[name_column]
                if pd.isna(raw_name) or str(raw_name).strip() == "":
                    continue
                service_name = str(raw_name).strip()
                
                # 1. Find best match in GIA_HDND
                match_result_gia = self.find_best_match(service_name, reference_names, threshold)
                
                # 2. Find best match in DVKT_GIA_MAX
                match_result_max = self.find_best_match(service_name, max_reference_names, threshold) if max_reference_names else None

                # Logic: Use MAX data if available, otherwise GIA_HDND, otherwise empty
                
                output_row = {
                    'STT': idx + 1,
                    # Default empty
                    'MA_TUONG_DUONG': '',
                    'TEN_DVKT_PHEDUYET': '',
                    'TEN_DVKT_GIA': '',
                    'PHAN_LOAI_PTTT': '',
                    'DON_GIA': '',
                    'GHI_CHU': '',
                    'QUYET_DINH': '',
                    'QUY_TRINH': '',
                    'TU_NGAY': '',
                    'DEN_NGAY': '',
                    'CSKCB_CGKT': '',
                    'CSKCB_CLS': ''
                }

                # If we have a MAX match, populate specific fields from it
                matched_name_display = 'KHÔNG TÌM THẤY'
                score_display = 0

                if match_result_max:
                    matched_name_max, score_max = match_result_max
                    
                    ref_row_max = self.ref_dvkt_gia_max_df[self.ref_dvkt_gia_max_df[max_ref_name_col] == matched_name_max].iloc[0]
                    
                    output_row['MA_TUONG_DUONG'] = ref_row_max.get('MA_TUONG_DUONG', '')
                    output_row['TEN_DVKT_PHEDUYET'] = ref_row_max.get('TEN_DVKT_PHEDUYET', '')
                    output_row['TEN_DVKT_GIA'] = matched_name_max
                    output_row['PHAN_LOAI_PTTT'] = ref_row_max.get('PHAN_LOAI_PTTT', '')
                    output_row['GHI_CHU'] = ref_row_max.get('GHI_CHU', '')
                    
                    # Try to get price from MAX as fallback
                    max_price = ref_row_max.get('DON_GIA', '')
                    if not max_price and 'GIÁ_MAX' in ref_row_max: 
                         max_price = ref_row_max['GIÁ_MAX']
                    output_row['DON_GIA'] = max_price
                    
                    matched_name_display = matched_name_max
                    score_display = score_max

                # If we have a GIA match, get Price (Priority 1) and other fields if NOT in MAX (Priority 2)
                if match_result_gia:
                    matched_name_gia, score_gia = match_result_gia
                    ref_row_gia = self.ref_gia_hdnd_df[self.ref_gia_hdnd_df[ref_name_column] == matched_name_gia].iloc[0]
                    
                    # Price Priority: GIA > MAX
                    gia_price = ref_row_gia.get('Mức giá', '')
                    # Only overwrite if GIA price is valid (not empty/NaN)
                    if not pd.isna(gia_price) and str(gia_price).strip() != '':
                        output_row['DON_GIA'] = gia_price
                    
                    output_row['QUYET_DINH'] = ref_row_gia.get('Quyết định', '') 
                    output_row['TU_NGAY'] = '' 
                    output_row['DEN_NGAY'] = ''
                    
                    # Fallbacks for other fields if not in MAX
                    if not match_result_max:
                        output_row['MA_TUONG_DUONG'] = ref_row_gia.get('Mã tương đương', '')
                        output_row['TEN_DVKT_PHEDUYET'] = ref_row_gia.get('Tên chương theo TT 23/2024', '')
                        output_row['TEN_DVKT_GIA'] = matched_name_gia
                        output_row['PHAN_LOAI_PTTT'] = '' 
                        output_row['GHI_CHU'] = ref_row_gia.get('Ghi chú', '')
                        matched_name_display = matched_name_gia
                        score_display = score_gia
                    
                    # Try to find QUY_TRINH from file 1 using GIA name (more standard)
                    if quy_trinh_name_col:
                        quy_trinh_match = self.find_best_match(matched_name_gia, 
                                                               self.ref_quy_trinh_df[quy_trinh_name_col].dropna().astype(str).tolist(),
                                                               threshold)
                        if quy_trinh_match:
                            quy_trinh_row = self.ref_quy_trinh_df[self.ref_quy_trinh_df[quy_trinh_name_col] == quy_trinh_match[0]].iloc[0]
                            output_row['QUY_TRINH'] = quy_trinh_row.get('QUY_TRINH', '')

                # Final record
                if not output_row['TEN_DVKT_GIA']:
                    output_row['TEN_DVKT_GIA'] = service_name
                
                results.append(output_row)
                
                if matched_name_display != 'KHÔNG TÌM THẤY':
                     stats['matched'] += 1
                else:
                     stats['unmatched'] += 1
                     
                stats['match_details'].append({
                    'input': service_name,
                    'matched': matched_name_display,
                    'score': score_display
                })
            
            # Create output DataFrame
            output_df = pd.DataFrame(results)
            
            # Save to Excel with proper encoding
            output_df.to_excel(output_file, index=False, engine='openpyxl')
            
            message = f"Xử lý thành công!\n"
            message += f"Tổng số: {stats['total']}\n"
            message += f"Khớp: {stats['matched']}\n"
            message += f"Không khớp: {stats['unmatched']}"
            
            if self.ref_dvkt_gia_max_df is not None:
                message += f"\n(Đã sử dụng dữ liệu ưu tiên từ DVKT_GIA_MAX)"
            
            return True, message, stats
            
        except Exception as e:
            return False, f"Lỗi khi xử lý file: {str(e)}", {}
