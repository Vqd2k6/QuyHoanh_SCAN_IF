import asyncio
import pandas as pd
import argparse
from geospatial import get_buffered_grid
from scraper import fetch_planning_data

async def main():
    parser = argparse.ArgumentParser(description="Trích xuất thông tin quy hoạch theo tuyến đường từ KML.")
    parser.add_argument("kml_file", help="Đường dẫn đến file KML từ Google My Maps")
    parser.add_argument("--output", default="ket_qua_quy_hoach.xlsx", help="Tên file Excel đầu ra (vd: ket_qua.xlsx)")
    parser.add_argument("--buffer", type=int, default=100, help="Bán kính vùng đệm (meters)")
    parser.add_argument("--grid", type=int, default=20, help="Khoảng cách giữa các điểm lưới (meters)")
    
    args = parser.parse_args()
    
    # 1. Phân tích KML và tạo lưới điểm
    print(f"[*] Đang đọc file KML: {args.kml_file}")
    try:
        points = get_buffered_grid(args.kml_file, buffer_meters=args.buffer, grid_spacing_meters=args.grid)
        print(f"[*] Tổng số điểm cần quét: {len(points)}")
    except Exception as e:
        print(f"[!] Lỗi khi xử lý KML: {e}")
        return

    if not points:
        print("[!] Không có toạ độ nào được tạo ra.")
        return

    # Để demo không bị quá tải, ta giới hạn số lượng điểm
    if len(points) > 100:
        print("[*] (Demo) Giới hạn số lượng điểm quét xuống 100 để tránh quá tải...")
        points = points[:100]

    # 2. Quét dữ liệu bằng Playwright
    print(f"[*] Khởi động trình duyệt tự động để thu thập dữ liệu...")
    results = await fetch_planning_data(points)
    
    # 3. Chuyển đổi dữ liệu và xuất ra Excel
    print(f"[*] Đang xuất dữ liệu ra file Excel: {args.output}")
    df = pd.DataFrame(results)
    
    # Tạo cấu trúc báo cáo giả định
    if not df.empty:
        # Nếu có dữ liệu JSON chuẩn, có thể parse ra các cột (ChucNangSDD, DienTich...)
        # Trong bản demo, cột RawData sẽ chứa JSON phản hồi.
        df.to_excel(args.output, index=False)
        print(f"[*] Hoàn tất! Dữ liệu đã được lưu tại {args.output}")
    else:
        print("[!] Không thu thập được dữ liệu nào.")

if __name__ == "__main__":
    asyncio.run(main())
