# Skill: Tự động hóa Zalo Web (Zalo Group Automation Bot)

Bộ skill này hướng dẫn cách cài đặt và vận hành công cụ tự động hóa gửi tin nhắn và hình ảnh vào các **Nhóm Zalo** (Zalo Groups) bằng Selenium và Python.

## 1. Mục đích và Các Khái Niệm Quan Trọng

*   **Tính Năng:** Bot tự động quét dữ liệu từ Google Sheets (danh sách nhóm, nội dung tin nhắn, hình ảnh), đăng nhập Zalo qua Selenium, tự động tìm kiếm nhóm và gửi tin nhắn/ảnh, sau đó cập nhật trạng thái vào Google Sheets.
*   **Cơ chế vượt rào Zalo Web:** Zalo Web dùng cơ chế Lazy Loading cho các thẻ `<input type="file">`. Để giải quyết, bot sử dụng PowerShell để nạp nội dung (ảnh/văn bản) vào Clipboard của hệ điều hành, sau đó nhấn `Ctrl+V` thông qua webdriver.

## 2. Triển Khai Trên Máy Mới (Từ A đến Z)

### Bước 0: Clone repo về máy
```powershell
git clone https://github.com/quocdat16061990/SELENIUM_ZALO_GROUP.git
cd SELENIUM_ZALO_GROUP
```

### Bước 1: Chuẩn bị file JSON Service Account (Google)
⚠️ **QUAN TRỌNG:** Bot không thể chạy nếu thiếu file này!
1.  Vào Google Cloud Console, tạo Service Account với quyền **Editor** cho Google Sheets API.
2.  Tải xuống file JSON key cho Service Account đó.
3.  Copy file JSON vào thư mục gốc của dự án (cùng cấp với file `OpenZaloSendListRelative.py`).
4.  **Cập nhật cấu hình:** Mở `OpenZaloSendListRelative.py`, tìm dòng `CREDENTIALS_FILE` và đổi tên file JSON cho đúng:
    ```python
    CREDENTIALS_FILE = os.path.join(BASE_DIR, "ten-file-json-cua-ban.json")
    ```
5.  **Share Google Sheet:** Chia sẻ sheet cho email của Service Account (ví dụ: `bot-name@project-id.iam.gserviceaccount.com`) với quyền Editor.

### Bước 2: Chuẩn bị file code
Đảm bảo file code `OpenZaloSendListRelative.py` đã nằm ở thư mục gốc.

### Bước 3: Cài đặt Python và môi trường
1.  **Kiểm tra Python:** Cần Python 3.10 trở lên.
2.  **Tạo môi trường ảo:**
    ```powershell
    python -m venv venv
    .\venv\Scripts\Activate.ps1
    ```
3.  **Cài đặt thư viện:**
    ```powershell
    pip install selenium gspread pandas oauth2client
    ```

### Bước 4: Chuẩn bị Thư mục và Chrome Profile
1.  **Thư mục images/:** Tạo thư mục `images/` và thêm ảnh vào (tên ảnh trùng với keyword trong cột `Image`).
2.  **Chrome Profile:** Tạo thư mục `zalo-chrome-profile/` để lưu phiên đăng nhập Zalo.

## 3. Cấu Hình Google Sheet

Sheet cần có tên là **"Danh Sách Nhóm"** (hoặc nằm ở vị trí đầu tiên) với các cột:

| Cột | Mô tả |
| :--- | :--- |
| **Group** | Tên chính xác của nhóm Zalo cần gửi (Hoặc SĐT) |
| **Message** | Nội dung tin nhắn gửi đi |
| **Image** | Tên file ảnh trong `/images` (ví dụ: `hoc-python`) |
| **Status** | Trạng thái: `UNAPPROVED` (để bot chạy), `APPROVED` (xong), `ZALO_FAILED` (lỗi) |

*Bot sẽ tự động cập nhật trạng thái sau khi gửi.*

## 4. Cách Khởi Chạy

**Lệnh chạy nhanh:**
```powershell
.\venv\Scripts\python.exe OpenZaloSendListRelative.py
```
*Lần đầu chạy, trình duyệt mở ra -> Quét mã QR. Các lần sau sẽ tự đăng nhập.*

## 5. Các Hàm Cốt Lõi Tham Khảo
*   **search_and_click_group:** Điều hướng qua tab Danh bạ -> Nhóm để tìm kiếm chính xác.
*   **send_message:** Sử dụng PowerShell Clipboard bypass để dán ảnh và văn bản UTF-8.
*   **Stale Element Fix:** Đảm bảo không bị lỗi khi Zalo re-render DOM sau khi dán ảnh.
