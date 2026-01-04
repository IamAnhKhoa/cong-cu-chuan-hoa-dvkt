from processor import CatalogProcessor
import os

base_dir = r"c:\Users\Administrator\Desktop\Danh mục kĩ thuật"
ref_dir = os.path.join(base_dir, "File gốc để chuẩn hóa")

processor = CatalogProcessor()
success, message = processor.load_reference_files(ref_dir)

print(f"Success: {success}")
print(f"Message: {message}")
