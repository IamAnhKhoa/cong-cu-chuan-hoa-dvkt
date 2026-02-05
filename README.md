# Công cụ Chuẩn hóa Danh mục Kỹ thuật Y tế (v2.0)

![Python](https://img.shields.io/badge/Python-3.9%2B-blue)
![PyQt6](https://img.shields.io/badge/GUI-PyQt6-green)
![Design](https://img.shields.io/badge/Design-Flat%20%26%20Compact-orange)
![License](https://img.shields.io/badge/License-MIT-gray)

Một ứng dụng Desktop chuyên nghiệp được nâng cấp toàn diện lên **PyQt6**, hỗ trợ các cơ sở y tế chuẩn hóa danh mục dịch vụ kỹ thuật theo quy chuẩn của Bộ Y tế và Bảo hiểm Xã hội Việt Nam. 

Phiên bản mới tập trung vào trải nghiệm người dùng tối giản ("Sophisticated & Simple"), hiệu năng xử lý cao và tính năng hỗ trợ toàn diện.

## ✨ Tính năng mới (Cập nhật v2.0)

*   **⚡ Nâng cấp Core PyQt6**: Chuyển đổi hoàn toàn sang framework Qt6 hiện đại, mang lại giao diện sắc nét và hiệu năng ổn định hơn trên Windows 10/11.
*   **🎨 Giao diện "Sophisticated & Simple"**:
    *   **Thiết kế phẳng (Flat Design)**: Hiện đại, loại bỏ các icon rườm rà.
    *   **Super Compact**: Bố cục tối ưu không gian dọc, đảm bảo **không cần cuộn (No Scroll)** mà vẫn hiển thị đầy đủ thông tin.
    *   **Focus Visibility**: Nút bấm quan trọng sử dụng chữ **IN ĐẬM và TRẮNG** trên nền màu tương phản cao, dễ dàng nhận diện.
*   **📥 Hỗ trợ File Mẫu Tích hợp**:
    *   Tích hợp sẵn tính năng **Tải file mẫu GIA_HDND (399/NQ-HĐND)** ngay trên giao diện.
    *   File mẫu chuẩn bao gồm: **Sheet Hướng dẫn** (quy định cột bắt buộc) và **Sheet Dữ liệu mẫu** (có data thực tế).
*   **🔄 Luồng xử lý rõ ràng**:
    *   *Luồng 1 (TT21 của cổng)*: Chuẩn hóa dữ liệu đẩy cổng BHYT.
    *   *Luồng 2 (Import giá phê duyệt)*: Chuẩn hóa dữ liệu đầu vào phần mềm HIS.

## 🚀 Tính năng cốt lõi

*   **Tự động hóa cao**: Sử dụng thuật toán Fuzzy Matching (khớp mờ) để tự động tìm kiếm danh mục tương đương với độ chính xác tùy chỉnh.
*   **Xử lý đa luồng (Multi-threading)**: Đảm bảo giao diện không bao giờ bị treo (Not Responding) khi xử lý file lớn.
*   **Báo cáo chi tiết**:
    *   Xuất file Excel kết quả.
    *   Xuất file "Không khớp" riêng biệt để đối chiếu thủ công.
    *   Nhật ký (Log) chi tiết quá trình xử lý.

## 📋 Yêu cầu hệ thống

*   **Hệ điều hành**: Windows 10/11 (64-bit).
*   **Python**: 3.9 trở lên.
*   **Thư viện**: PyQt6, pandas, openpyxl, fuzzywuzzy.

## 🛠️ Cài đặt

1.  **Clone repository:**
    ```bash
    git clone https://github.com/IamAnhKhoa/cong-cu-chuan-hoa-dvkt.git
    cd cong-cu-chuan-hoa-dvkt
    ```

2.  **Cài đặt thư viện:**
    ```bash
    pip install -r requirements.txt
    ```

## 📖 Hướng dẫn sử dụng

### 1. Chuẩn bị dữ liệu
Bạn cần thư mục chứa các file Excel tham chiếu (tên file chính xác):
*   `QUY_TRINH_DVKT_BYT.xlsx`
*   `GIA_HDND.xlsx`
*   `DVKT_GIA_MAX.xlsx` (Tùy chọn)

### 2. Chạy ứng dụng
```bash
python main.py
```

### 3. Các bước thực hiện
*   **Bước 1 - Chọn thư mục tham chiếu**:
    *   Trỏ đến thư mục chứa 3 file trên.
    *   *Mới*: Bạn có thể **Tải file mẫu GIA_HDND** ngay tại đây nếu chưa có file chuẩn.
*   **Bước 2 - Chọn luồng xử lý**:
    *   Chọn **Luồng 1** hoặc **Luồng 2** tùy mục đích.
*   **Bước 3 - Chọn file đầu vào**:
    *   File Excel danh mục của bệnh viện (cần có cột `Tên dịch vụ` hoặc `Tên kỹ thuật`).
*   **Bước 4 - Xử lý**:
    *   Nhấn nút **BẮT ĐẦU XỬ LÝ** (Màu xanh, chữ đậm).
    *   Chờ thanh tiến trình chạy xong (100%).
*   **Bước 5 - Kiểm tra kết quả**:
    *   Xem file kết quả tại đường dẫn đã chọn.
    *   Dùng các nút "Xem nhật ký" hoặc "Xem dịch vụ không khớp" để kiểm tra lại.

---
*Được phát triển bởi Khoa - Phiên bản PyQt6 2026*
