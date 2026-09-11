# QuyHoanh_SCAN_IF

Ứng dụng quét và trích xuất thông tin quy hoạch xây dựng / công trình theo tuyến đường từ bản đồ quy hoạch TP.HCM ([thongtinquyhoach.hochiminhcity.gov.vn](https://thongtinquyhoach.hochiminhcity.gov.vn/)).

---

## 🌟 Tính năng chính

- **Xử lý Không gian Địa lý (Geospatial Analysis)**:
  - Đọc và phân tích file KML xuất từ Google My Maps.
  - Tự động chuyển đổi sang hệ toạ độ phẳng UTM Zone 48N (EPSG:32648) cho TP.HCM.
  - Tạo vùng đệm (buffer) dọc theo tuyến đường (mặc định 100m).
  - Tạo lưới toạ độ (grid sampling points) bao phủ toàn bộ vùng đệm để quét các thửa đất / công trình.
- **Trình thu thập dữ liệu (Web Scraper)**:
  - Sử dụng Playwright để tự động hóa tương tác với cổng thông tin quy hoạch.
  - Hỗ trợ vượt các cơ chế chống bot cơ bản (User-Agent tùy chỉnh, ẩn webdriver flag).
- **Xuất báo cáo Excel**:
  - Lưu trữ kết quả quy hoạch theo toạ độ thành file `.xlsx` rõ ràng, có cấu trúc.

---

## 📁 Cấu trúc thư mục

```text
QuyHoanh_SCAN_IF/
├── geospatial.py       # Xử lý KML, buffer 100m, tạo lưới toạ độ EPSG:32648
├── scraper.py          # Playwright automation tương tác với bản đồ quy hoạch
├── main.py             # File điều phối chính & xuất dữ liệu ra Excel
├── sample_route.kml    # Tuyến đường KML mẫu để kiểm thử
├── requirements.txt    # Danh sách các thư viện phụ thuộc
├── .gitignore          # Cấu hình bỏ qua file rác và môi trường ảo
└── README.md           # Hướng dẫn sử dụng
```

---

## 🚀 Hướng dẫn cài đặt & Chạy

### 1. Tạo môi trường ảo & Cài đặt thư viện

```bash
# Tạo môi trường ảo
python3 -m venv venv

# Kích hoạt môi trường ảo
# Trên macOS / Linux:
source venv/bin/activate
# Trên Windows:
# venv\Scripts\activate

# Cài đặt các package cần thiết
pip install -r requirements.txt

# Cài đặt trình duyệt cho Playwright
playwright install chromium
```

### 2. Chạy chương trình

```bash
python main.py sample_route.kml --buffer 100 --grid 20 --output ket_qua.xlsx
```

**Các tham số:**
- `kml_file`: Đường dẫn file KML (bắt buộc).
- `--buffer`: Bán kính vùng đệm quét theo mét (mặc định: `100`).
- `--grid`: Khoảng cách giữa các điểm quét trên lưới tính theo mét (mặc định: `20`).
- `--output`: Đường dẫn lưu file Excel kết quả (mặc định: `ket_qua_quy_hoach.xlsx`).

---

## ⚠️ Lưu ý kỹ thuật (WAF & Firewall)

Trang web `thongtinquyhoach.hochiminhcity.gov.vn` được bảo vệ bởi tường lửa F5 ASM WAF khá nghiêm ngặt. Khi triển khai cào dữ liệu quy mô lớn, nên:
- Chạy trình duyệt có giao diện (`headless=False`) hoặc sử dụng profile Chrome thật.
- Bổ sung giãn cách thời gian giữa các request (`time.sleep`) để tránh bị chặn IP.
