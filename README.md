# Công cụ Chuẩn hóa Danh mục Kỹ thuật Y tế

![Python](https://img.shields.io/badge/Python-3.8%2B-blue)
![PyQt5](https://img.shields.io/badge/GUI-PyQt5-green)
![License](https://img.shields.io/badge/License-MIT-orange)

Một ứng dụng Desktop mạnh mẽ được xây dựng bằng Python và PyQt5, hỗ trợ các cơ sở y tế chuẩn hóa danh mục dịch vụ kỹ thuật theo quy chuẩn của Bộ Y tế và Bảo hiểm Xã hội Việt Nam.Công cụ này đặc biệt hữu ích cho việc đối chiếu, map dữ liệu và chuẩn hóa danh mục để đẩy lên cổng giám định BHYT.

## 🚀 Tính năng nổi bật

*   **Tự động hóa cao**: Sử dụng thuật toán Fuzzy Matching để tự động khớp tên dịch vụ kỹ thuật với độ chính xác cao.
*   **Giao diện trực quan**: Thiết kế thân thiện, dễ sử dụng cho nhân viên y tế và kế toán.
*   **Xử lý đa luồng**: Đảm bảo ứng dụng luôn mượt mà ngay cả khi xử lý file dữ liệu lớn.
*   **Linh hoạt**: Cho phép tùy chỉnh ngưỡng chính xác (Threshold) để phù hợp với từng loại dữ liệu.
*   **Báo cáo chi tiết**: Xuất kết quả kèm theo báo cáo chi tiết về các mục đã khớp và các mục cần kiểm tra lại.

## 📋 Yêu cầu hệ thống

*   Hệ điều hành: Windows 10/11 (Khuyến nghị), macOS, hoặc Linux.
*   Python: Phiên bản 3.7 trở lên.
*   Microsoft Excel hoặc phần mềm đọc file .xlsx tương đương.

## 🛠️ Cài đặt

1.  **Clone repository này về máy:**

    ```bash
    git clone https://github.com/username/ten-repo-cua-ban.git
    cd ten-repo-cua-ban
    ```

2.  **Tạo môi trường ảo (Khuyến nghị):**

    ```bash
    python -m venv venv
    # Windows:
    .\venv\Scripts\activate
    # Linux/Mac:
    source venv/bin/activate
    ```

3.  **Cài đặt các thư viện phụ thuộc:**

    ```bash
    pip install -r requirements.txt
    ```

## 📖 Hướng dẫn sử dụng

### 1. Chuẩn bị dữ liệu
Bạn cần chuẩn bị một thư mục chứa 2 file Excel tham chiếu bắt buộc (tên file phải chính xác):
*   `QUY_TRINH_DVKT_BYT.xlsx`: Danh mục quy trình kỹ thuật chuẩn của Bộ Y tế.
*   `GIA_HDND.xlsx`: Danh mục giá dịch vụ theo Nghị quyết HĐND.
*   *(Tùy chọn)* `DVKT_GIA_MAX.xlsx`: Nếu muốn ưu tiên lấy giá trần.

### 2. Chạy ứng dụng

```bash
python main.py
```

### 3. Quy trình xử lý

*   **Bước 1 - Chọn thư mục tham chiếu**: Trỏ đường dẫn đến thư mục chứa các file Excel đã chuẩn bị ở trên.
*   **Bước 2 - Chọn loại xử lý**:
    *   *Loại 1*: Chuẩn hóa từ Quy trình Bộ Y tế -> Danh mục TT21 (Cổng BHYT).
    *   *Loại 2*: Chuẩn hóa từ Giá HĐND -> File chuẩn tổng hợp (Dùng để import vào phần mềm HIS).
*   **Bước 3 - Chọn file đầu vào**: Chọn file Excel danh mục hiện tại của bệnh viện (yêu cầu phải có cột chứa tên dịch vụ/kỹ thuật).
*   **Bước 4 - Cấu hình**: Điều chỉnh thanh trượt "Ngưỡng khớp tên" (Mặc định 80%).
    *   *Mẹo*: Nếu danh mục của bạn viết tắt nhiều, hãy thử giảm xuống 70-75%.
*   **Bước 5 - Xử lý**: Nhấn "Bắt đầu xử lý" và đợi kết quả.

## 📂 Cấu trúc file đầu ra

Ứng dụng sẽ tự động sinh ra file Excel với các cột chuẩn hóa:

| Cột | Mô tả |
| :--- | :--- |
| `MA_DICH_VU` / `MA_TUONG_DUONG` | Mã dịch vụ chuẩn theo quy định |
| `TEN_DICH_VU` | Tên dịch vụ chuẩn hóa |
| `DON_GIA` | Đơn giá tham chiếu |
| `QUY_TRINH` | Mã quy trình kỹ thuật |
| ... | Các trường thông tin khác theo mẫu 21 |

## 🤝 Đóng góp (Contributing)

Mọi sự đóng góp đều được hoan nghênh! Nếu bạn muốn cải thiện dự án, vui lòng:

1.  Fork dự án.
2.  Tạo branch mới (`git checkout -b feature/TinhNangMoi`).
3.  Commit thay đổi (`git commit -m 'Thêm tính năng X'`).
4.  Push lên branch (`git push origin feature/TinhNangMoi`).
5.  Tạo Pull Request.

---
*Được phát triển bởi Khoa 
# cong-cu-chuan-hoa-dvkt
