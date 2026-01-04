import pandas as pd
import os

base_dir = r"c:\Users\Administrator\Desktop\Danh mục kĩ thuật"
input_file = os.path.join(base_dir, "DMKT_BINH_MY.xlsx")
ref_dir = os.path.join(base_dir, "File gốc để chuẩn hóa")
ref_file_1 = os.path.join(ref_dir, "QUY_TRINH_DVKT_BYT.xlsx")
ref_file_2 = os.path.join(ref_dir, "GIA_HDND.xlsx")

def inspect_file(path, name):
    print(f"--- Inspecting {name} ---")
    if not os.path.exists(path):
        print(f"File not found: {path}")
        return

    try:
        df = pd.read_excel(path, nrows=5)
        print("Columns:", list(df.columns))
        print("First 2 rows:")
        print(df.head(2).to_string())
    except Exception as e:
        print(f"Error reading {name}: {e}")
    print("\n")

inspect_file(input_file, "Input File (DMKT_BINH_MY.xlsx)")
inspect_file(ref_file_1, "Reference 1 (QUY_TRINH_DVKT_BYT.xlsx)")
inspect_file(ref_file_2, "Reference 2 (GIA_HDND.xlsx)")
