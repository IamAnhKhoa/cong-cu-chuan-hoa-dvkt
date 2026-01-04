import pandas as pd
import os

ref_dir = r"c:\Users\Administrator\Desktop\Danh mục kĩ thuật\File gốc để chuẩn hóa"
ref_file = os.path.join(ref_dir, "QUY_TRINH_DVKT_BYT.xlsx")

print(f"Reading {ref_file}...")
try:
    # Read without header to see raw likely structure
    df_raw = pd.read_excel(ref_file, header=None, nrows=10, engine='openpyxl')
    print("First 10 rows raw:")
    print(df_raw.to_string())
    
    # Check sheets
    xl = pd.ExcelFile(ref_file, engine='openpyxl')
    print("\nSheet names:", xl.sheet_names)
    
except Exception as e:
    print(f"Error: {e}")
