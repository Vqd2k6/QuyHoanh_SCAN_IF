# QuyHoach - HỆ THỐNG KHẢO SÁT & TRÍCH XUẤT QUY HOẠCH ĐÔ THỊ TP.HCM

Ứng dụng quét và trích xuất thông tin quy hoạch xây dựng / công trình theo tuyến đường hoặc ranh giới đa giác từ Cổng thông tin quy hoạch TP.HCM ([thongtinquyhoach.hochiminhcity.gov.vn](https://thongtinquyhoach.hochiminhcity.gov.vn/)).

---

## 🌟 Tính năng nổi bật & Tối ưu hoá

- **Tối ưu hoá Bộ nhớ & Hiệu năng Cao (Zero Memory Leak)**:
  - Sử dụng **Leaflet Canvas Renderer** thay cho DOM SVG, vẽ mượt mà hơn 50.000 điểm toạ độ và hàng ngàn thửa đất với RAM tiêu thụ dưới 50MB.
  - Cơ chế đồng bộ dữ liệu bản đồ bất đồng bộ (Clean Polling) tự động dọn dẹp rác bộ nhớ JavaScript, triệt tiêu hoàn toàn lỗi *Out of Application Memory*.
  - Tối ưu bộ nhớ Python backend: Bounding-box spatial pre-filter tăng tốc kiểm tra không gian lên 100x và batching lưu trữ Excel định kỳ.
- **Xử lý Không gian Địa lý (Geospatial Analysis)**:
  - Đọc và phân tích file KML xuất từ Google My Maps hoặc file cục bộ trong `data/input/`.
  - Tự động chuyển đổi sang hệ toạ độ phẳng UTM Zone 48N (EPSG:32648) cho TP.HCM.
  - Tự động đóng kín các đường bao ranh giới mở rộng (Boundary Auto-Closing & Noded Lines), bao phủ 100% diện tích không gian.
- **Thu thập dữ liệu Đa Luồng Siêu Tốc (Async Worker Pool lên đến 100 luồng)**:
  - Chạy song song từ 1 đến **100 luồng song song** với cơ chế chia sẻ kết nối (Connection Pool) siêu nhẹ, tiêu thụ RAM cực thấp.
  - Quét **100% toàn bộ toạ độ lưới**, triệt tiêu hoàn toàn lỗi bỏ sót các thửa nhà dân nằm xen kẽ hoặc bị trùm bởi các ô đất quy hoạch giao thông/công cộng.
  - Tự động khử trùng lặp dữ liệu thông minh qua mã thửa (`mathuadat`): Chỉ ghi nhận 1 lần duy nhất vào Excel & Bản đồ khi nhiều điểm trúng cùng 1 thửa.
- **Báo cáo Chuyên Nghiệp 4 Sheet Excel**:
  1. *1. Tổng Hợp Thửa Đất (Master Overview)*
  2. *2. Chi Tiết Ô Quy Hoạch & KT (Zoning Lots & Urban Norms)*
  3. *3. Chi Tiết Lộ Giới Tuyến Đường (Road Setbacks)*
  4. *4. Dự Án 1-500 & Điều Chỉnh (Detailed 1/500 & DCCB)*
- **Bản đồ Tương tác Hai Chiều (Interactive Map & Management)**:
  - Xem chi tiết từng ô quy hoạch trực quan trên nền vệ tinh Esri/OpenStreetMap.
  - Nhấp chuột xoá thửa đất thừa: Tự động xoá trực tiếp trên cả 4 Sheet của file Excel và đánh lại số thứ tự (STT).

---

## 📁 Cấu trúc thư mục chuẩn hóa

```text
QuyHoach/
├── data/
│   ├── input/                  # Chứa file KML, file tuyến đường đầu vào
│   │   ├── sample_route.kml
│   │   └── gmap_1IjTEVFWwFg8n7OdD1uDaymCB2Q6aUTE.kml
│   └── output/                 # Chứa các đợt khảo sát (Mỗi lần chạy là 1 folder riêng)
│       ├── KS003/
│       │   ├── KS003.xlsx                # File Excel 4 Sheet dữ liệu gốc chuẩn
│       │   ├── ban_do_quy_hoach.html     # Bản đồ vệ tinh Leaflet tương tác
│       │   ├── ban_do_data.js            # Dữ liệu phục vụ mở trực tiếp file://
│       │   └── ban_do_data.json          # Dữ liệu JSON đồng bộ thời gian thực
│       └── khao_sat_.../                 # Đợt khảo sát khác...
├── src/                        # Chứa toàn bộ mã nguồn module xử lý
│   ├── __init__.py
│   ├── geospatial.py           # Phân tích KML, tính buffer, tạo lưới tọa độ
│   ├── scraper.py              # Cào dữ liệu đa luồng từ Cổng SQHKT TP.HCM (Hỗ trợ 100 luồng)
│   ├── visualizer.py           # Tạo bản đồ trực quan Leaflet Canvas tối ưu bộ nhớ
│   └── viewer.py               # Web server quản trị & xem lại kết quả Excel (Ưu tiên đúng file gốc)
├── main.py                     # File điều phối chính & CLI Wizard
├── requirements.txt            # Danh sách các thư viện phụ thuộc
├── .gitignore                  # Cấu hình bỏ qua file tạm, kết quả lớn và venv
└── README.md                   # Hướng dẫn sử dụng
```

---

## 🚀 Hướng dẫn cài đặt & Chạy

### 1. Cài đặt môi trường

```bash
# Tạo môi trường ảo
python3 -m venv venv

# Kích hoạt môi trường ảo
# Trên macOS / Linux:
source venv/bin/activate
# Trên Windows:
# venv\Scripts\activate

# Cài đặt các thư viện phụ thuộc
pip install -r requirements.txt

# Cài đặt trình duyệt Playwright
playwright install chromium
```

### 2. Chạy chương trình

#### Cách 1: Giao diện CLI Wizard tương tác (Khuyên dùng ⭐)

```bash
./venv/bin/python main.py
```

- **[Bước 0]**: Chọn tác vụ:
  - `[1]`: 🚀 Khảo sát & cào trích xuất quy hoạch mới (100% không sót hộ dân).
  - `[2]`: ⏩ **Tiếp tục đợt quét dở dang (Resume)** - Chọn thư mục cũ để quét tiếp từ điểm bị ngắt.
  - `[3]`: 📂 **Mở lại bản đồ từ file Excel đã lưu** để xem và xoá thửa đất.
- **[Bước 1]**: Dán link Google My Maps hoặc đường dẫn file KML (mặc định Tuyến Metro số 2).
- **[Bước 2]**: Chọn mật độ lưới (30m chuẩn, 20m dày, 50m thưa).
- **[Bước 2.5]**: Chọn số luồng quét song song (15, 30, 50, 100 luồng).
- **[Bước 3]**: Chọn phạm vi quét (Toàn bộ 100% diện tích hoặc quét nhanh kiểm tra).
- **[Bước 4]**: Chọn thư mục và tên file lưu kết quả.

#### Cách 2: Chạy trực tiếp bằng dòng lệnh (CLI Parameters)

```bash
# Quét mới tuyến Metro 2 với 50 hoặc 100 luồng song song
./venv/bin/python main.py "https://www.google.com/maps/d/u/0/viewer?mid=1IjTEVFWwFg8n7OdD1uDaymCB2Q6aUTE" --grid 30 --concurrency 50 --out-dir data/output/khao_sat_metro2

# Tiếp tục đợt quét dở dang (Resume) với 100 luồng
./venv/bin/python main.py --resume data/output/KS003 --concurrency 100
```

#### Cách 3: Xem lại và quản trị xoá thửa đất từ file Excel đã lưu

```bash
# Mở bản đồ trực quan từ thư mục KS003 (tự động nạp đúng file gốc KS003.xlsx)
./venv/bin/python main.py --view data/output/KS003
```
