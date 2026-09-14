import asyncio
import os
import sys
import argparse
import webbrowser
import threading
import time
from datetime import datetime
import json
import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter

from src.geospatial import get_spatial_data
from src.scraper import fetch_planning_data
from src.visualizer import generate_interactive_map, update_live_data
from src.viewer import start_viewer_server, load_data_from_excel, find_excel_and_workdir
DEFAULT_URL = "https://www.google.com/maps/d/u/0/viewer?mid=1IjTEVFWwFg8n7OdD1uDaymCB2Q6aUTE&ll=10.821633009544064%2C106.62770133828562&z=16"

# ==============================================================================
# HỆ THỐNG ĐỊNH DẠNG EXCEL CHUYÊN NGHIỆP (THEME DESIGN TOKENS)
# ==============================================================================
FONT_TITLE = Font(name="Arial", size=13, bold=True, color="FFFFFF")
FONT_SUBTITLE = Font(name="Arial", size=9, italic=True, color="475569")
FONT_HEADER = Font(name="Arial", size=10, bold=True, color="FFFFFF")
FONT_DATA = Font(name="Arial", size=9.5, color="0F172A")
FONT_BOLD = Font(name="Arial", size=9.5, bold=True, color="0F172A")

FILL_NAVY = PatternFill(start_color="1E3A8A", end_color="1E3A8A", fill_type="solid")      # Sheet 1: Master
FILL_TEAL = PatternFill(start_color="0F766E", end_color="0F766E", fill_type="solid")      # Sheet 2: Ô Chức Năng
FILL_INDIGO = PatternFill(start_color="4338CA", end_color="4338CA", fill_type="solid")  # Sheet 3: Lộ Giới
FILL_AMBER = PatternFill(start_color="92400E", end_color="92400E", fill_type="solid")    # Sheet 4: 1/500 & DCCB

FILL_SUB = PatternFill(start_color="F1F5F9", end_color="F1F5F9", fill_type="solid")
FILL_ALT = PatternFill(start_color="F8FAFC", end_color="F8FAFC", fill_type="solid")

BORDER_THIN = Border(
    left=Side(style='thin', color='CBD5E1'),
    right=Side(style='thin', color='CBD5E1'),
    top=Side(style='thin', color='CBD5E1'),
    bottom=Side(style='thin', color='CBD5E1')
)

def get_or_create_workbook(filepath, source_name="KML"):
    """
    Nếu file Excel đã tồn tại (chế độ Resume), nạp workbook cũ để tiếp tục ghi nối tiếp.
    Nếu chưa tồn tại, khởi tạo mới 4 Sheet chuyên nghiệp.
    """
    if os.path.exists(filepath):
        try:
            wb = openpyxl.load_workbook(filepath)
            ws1 = wb['1. Tổng Hợp Thửa Đất'] if '1. Tổng Hợp Thửa Đất' in wb.sheetnames else wb.active
            
            ws2 = None
            for s_name in ['2. Chi Tiết Ô Quy Hoạch & KT', '2. Chi Tiết Ô Chức Năng & KT']:
                if s_name in wb.sheetnames:
                    ws2 = wb[s_name]
                    break
            if ws2 is None:
                ws2 = wb.create_sheet("2. Chi Tiết Ô Quy Hoạch & KT")
                
            ws3 = wb['3. Chi Tiết Lộ Giới Tuyến Đường'] if '3. Chi Tiết Lộ Giới Tuyến Đường' in wb.sheetnames else wb.create_sheet("3. Chi Tiết Lộ Giới Tuyến Đường")
            
            ws4 = None
            for s_name in ['4. Dự Án 1-500 & Điều Chỉnh', '4. Chi Tiết 1/500 & DCCB']:
                if s_name in wb.sheetnames:
                    ws4 = wb[s_name]
                    break
            if ws4 is None:
                ws4 = wb.create_sheet("4. Dự Án 1-500 & Điều Chỉnh")
                
            return wb, ws1, ws2, ws3, ws4, True
        except Exception as e:
            print(f"[!] Cảnh báo mở file Excel cũ ({e}), khởi tạo cấu trúc mới...")

    wb, ws1, ws2, ws3, ws4 = setup_excel_workbook(filepath, source_name=source_name)
    return wb, ws1, ws2, ws3, ws4, False

def setup_excel_workbook(filepath, source_name="KML"):
    """
    Khởi tạo file Excel 4 Sheet chuẩn báo cáo quy hoạch và khảo sát đô thị chuyên nghiệp:
    1. Tổng Hợp Thửa Đất (Master Overview)
    2. Chi Tiết Ô Chức Năng & KT (Zoning Lots & Urban Norms: Tầng cao, Mật độ, HSSDĐ)
    3. Chi Tiết Lộ Giới Tuyến Đường (Road Setbacks & Frontage)
    4. Chi Tiết Dự Án 1/500 & ĐCCB (Detailed 1/500 Projects & Local Adjustments)
    """
    wb = openpyxl.Workbook()
    current_time_str = datetime.now().strftime("%d/%m/%Y %H:%M:%S")

    # ==========================================================================
    # SHEET 1: TỔNG HỢP THỬA ĐẤT (MASTER OVERVIEW)
    # ==========================================================================
    ws1 = wb.active
    ws1.title = "1. Tổng Hợp Thửa Đất"
    ws1.views.sheetView[0].showGridLines = True
    
    # 1.1. Banner Tiêu Đề
    ws1.merge_cells("A1:P1")
    c_t1 = ws1["A1"]
    c_t1.value = "BÁO CÁO KHẢO SÁT & TRÍCH XUẤT THÔNG TIN QUY HOẠCH ĐÔ THỊ TP. HỒ CHÍ MINH"
    c_t1.font = FONT_TITLE
    c_t1.fill = FILL_NAVY
    c_t1.alignment = Alignment(horizontal="center", vertical="center")
    ws1.row_dimensions[1].height = 38
    
    # 1.2. Phụ chú nguồn dữ liệu và thời gian
    ws1.merge_cells("A2:P2")
    c_sub1 = ws1["A2"]
    c_sub1.value = f"Nguồn dữ liệu: Sở Quy hoạch - Kiến trúc TP.HCM (sqhkt-qlqh.tphcm.gov.vn) | Thời gian quét: {current_time_str} | Phạm vi: {source_name}"
    c_sub1.font = FONT_SUBTITLE
    c_sub1.fill = FILL_SUB
    c_sub1.alignment = Alignment(horizontal="center", vertical="center")
    ws1.row_dimensions[2].height = 24
    ws1.row_dimensions[3].height = 8

    # 1.3. Hàng Tiêu Đề Cột (Dòng 4 cố định)
    headers1 = [
        "STT", 
        "Số Tờ", 
        "Số Thửa", 
        "Mã Thửa Đất", 
        "Quận / Huyện", 
        "Phường / Xã", 
        "Diện Tích Thửa (m²)", 
        "Số Ô Quy Hoạch",
        "Cơ Cấu Quy Hoạch Sử Dụng Đất", 
        "Chỉ Tiêu Kiến Trúc Khái Quát (Tầng cao, Mật độ, HSSDĐ)",
        "Đồ Án Quy Hoạch 1/2000", 
        "Quy Hoạch Chi Tiết 1/500", 
        "Lộ Giới Tuyến Đường Tiếp Giáp", 
        "Toạ Độ Khảo Sát (Kinh độ, Vĩ độ)", 
        "Thời Gian Quét",
        "Toạ Độ Ranh Đa Giác (Polygon Coordinates)"
    ]
    ws1.row_dimensions[4].height = 34
    
    for col_idx, h_text in enumerate(headers1, 1):
        c = ws1.cell(row=4, column=col_idx)
        c.value = h_text
        c.font = FONT_HEADER
        c.fill = FILL_NAVY
        c.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
        c.border = BORDER_THIN
        
    col_widths1 = {
        'A': 8, 'B': 10, 'C': 10, 'D': 18, 'E': 16, 'F': 22,
        'G': 20, 'H': 16, 'I': 45, 'J': 42, 'K': 38, 'L': 32,
        'M': 35, 'N': 24, 'O': 20, 'P': 30
    }
    for col_l, w in col_widths1.items():
        ws1.column_dimensions[col_l].width = w

    ws1.freeze_panes = "A5"
    ws1.auto_filter.ref = "A4:P4"

    # ==========================================================================
    # SHEET 2: CHI TIẾT Ô CHỨC NĂNG & CHỈ TIÊU KIẾN TRÚC (ZONING LOTS)
    # ==========================================================================
    ws2 = wb.create_sheet("2. Chi Tiết Ô Quy Hoạch & KT")
    ws2.views.sheetView[0].showGridLines = True
    
    ws2.merge_cells("A1:S1")
    c_t2 = ws2["A1"]
    c_t2.value = "BẢNG PHÂN TÍCH CHI TIẾT CÁC Ô CHỨC NĂNG & CHỈ TIÊU QUY HOẠCH KIẾN TRÚC (QUY HOẠCH PHÂN KHU 1/2000)"
    c_t2.font = FONT_TITLE
    c_t2.fill = FILL_TEAL
    c_t2.alignment = Alignment(horizontal="center", vertical="center")
    ws2.row_dimensions[1].height = 38

    ws2.merge_cells("A2:S2")
    c_sub2 = ws2["A2"]
    c_sub2.value = "Trích xuất chi tiết từng ô chức năng sử dụng đất trong thửa, kèm các chỉ tiêu quy hoạch kiến trúc chính thức từ Sở QHKT"
    c_sub2.font = FONT_SUBTITLE
    c_sub2.fill = FILL_SUB
    c_sub2.alignment = Alignment(horizontal="center", vertical="center")
    ws2.row_dimensions[2].height = 24
    ws2.row_dimensions[3].height = 8

    headers2 = [
        "STT Thửa", 
        "Số Tờ", 
        "Số Thửa", 
        "Mã Thửa Đất", 
        "Quận / Huyện", 
        "Phường / Xã", 
        "STT Ô", 
        "Mã Ô Phố", 
        "Chức Năng Sử Dụng Đất", 
        "Chức Năng Chi Tiết (SQHKT)",
        "Diện Tích Ô (m²)", 
        "Tỷ Lệ Chiếm Thửa (%)", 
        "Tầng Cao Cho Phép (tầng)",
        "Chiều Cao Tối Đa (m)",
        "Mật Độ XD Tối Đa (%)",
        "Hệ Số Sử Dụng Đất (HSSDĐ)",
        "Quy Mô Dân Số Dự Kiến (người)",
        "Tổng DT Cả Ô Phố (m²)",
        "Đồ Án Quy Hoạch 1/2000"
    ]
    ws2.row_dimensions[4].height = 34
    
    for col_idx, h_text in enumerate(headers2, 1):
        c = ws2.cell(row=4, column=col_idx)
        c.value = h_text
        c.font = FONT_HEADER
        c.fill = FILL_TEAL
        c.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
        c.border = BORDER_THIN
        
    col_widths2 = {
        'A': 10, 'B': 10, 'C': 10, 'D': 18, 'E': 16, 'F': 22,
        'G': 8, 'H': 14, 'I': 28, 'J': 26, 'K': 18, 'L': 15,
        'M': 18, 'N': 16, 'O': 16, 'P': 18, 'Q': 18, 'R': 20, 'S': 36
    }
    for col_l, w in col_widths2.items():
        ws2.column_dimensions[col_l].width = w

    ws2.freeze_panes = "A5"
    ws2.auto_filter.ref = "A4:S4"

    # ==========================================================================
    # SHEET 3: CHI TIẾT LỘ GIỚI TUYẾN ĐƯỜNG (ROAD SETBACKS)
    # ==========================================================================
    ws3 = wb.create_sheet("3. Chi Tiết Lộ Giới Tuyến Đường")
    ws3.views.sheetView[0].showGridLines = True
    
    ws3.merge_cells("A1:M1")
    c_t3 = ws3["A1"]
    c_t3.value = "BẢNG CHI TIẾT LỘ GIỚI & MẶT TIỀN CÁC TUYẾN ĐƯỜNG TIẾP GIÁP CỦA THỬA ĐẤT"
    c_t3.font = FONT_TITLE
    c_t3.fill = FILL_INDIGO
    c_t3.alignment = Alignment(horizontal="center", vertical="center")
    ws3.row_dimensions[1].height = 38

    ws3.merge_cells("A2:M2")
    c_sub3 = ws3["A2"]
    c_sub3.value = "Thống kê độ rộng lộ giới quy hoạch, hướng tiếp giáp, bề rộng mặt tiền và diện tích xây dựng được phép"
    c_sub3.font = FONT_SUBTITLE
    c_sub3.fill = FILL_SUB
    c_sub3.alignment = Alignment(horizontal="center", vertical="center")
    ws3.row_dimensions[2].height = 24
    ws3.row_dimensions[3].height = 8

    headers3 = [
        "STT Thửa", 
        "Số Tờ", 
        "Số Thửa", 
        "Mã Thửa Đất", 
        "Quận / Huyện", 
        "Phường / Xã", 
        "STT Tuyến Đường",
        "Tên Tuyến Đường Tiếp Giáp", 
        "Độ Rộng Lộ Giới (m)", 
        "Hướng Tiếp Giáp", 
        "Bề Rộng Mặt Tiền Tiếp Giáp (m)", 
        "Chiều Sâu Lộ Giới Vào Thửa (m)",
        "Diện Tích Xây Dựng Ngoài Lộ Giới (m²)"
    ]
    ws3.row_dimensions[4].height = 34
    
    for col_idx, h_text in enumerate(headers3, 1):
        c = ws3.cell(row=4, column=col_idx)
        c.value = h_text
        c.font = FONT_HEADER
        c.fill = FILL_INDIGO
        c.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
        c.border = BORDER_THIN
        
    col_widths3 = {
        'A': 10, 'B': 10, 'C': 10, 'D': 18, 'E': 16, 'F': 22,
        'G': 12, 'H': 28, 'I': 18, 'J': 18, 'K': 22, 'L': 22, 'M': 26
    }
    for col_l, w in col_widths3.items():
        ws3.column_dimensions[col_l].width = w

    ws3.freeze_panes = "A5"
    ws3.auto_filter.ref = "A4:M4"

    # ==========================================================================
    # SHEET 4: CHI TIẾT DỰ ÁN 1/500 & ĐIỀU CHỈNH CỤC BỘ (1/500 & DCCB)
    # ==========================================================================
    ws4 = wb.create_sheet("4. Dự Án 1-500 & Điều Chỉnh")
    ws4.views.sheetView[0].showGridLines = True
    
    ws4.merge_cells("A1:M1")
    c_t4 = ws4["A1"]
    c_t4.value = "BẢNG CHI TIẾT DỰ ÁN QUY HOẠCH CHI TIẾT 1/500 & QUYẾT ĐỊNH ĐIỀU CHỈNH CỤC BỘ"
    c_t4.font = FONT_TITLE
    c_t4.fill = FILL_AMBER
    c_t4.alignment = Alignment(horizontal="center", vertical="center")
    ws4.row_dimensions[1].height = 38

    ws4.merge_cells("A2:M2")
    c_sub4 = ws4["A2"]
    c_sub4.value = "Theo dõi các thửa đất nằm trong ranh giới dự án quy hoạch 1/500 hoặc chịu tác động của quyết định điều chỉnh cục bộ"
    c_sub4.font = FONT_SUBTITLE
    c_sub4.fill = FILL_SUB
    c_sub4.alignment = Alignment(horizontal="center", vertical="center")
    ws4.row_dimensions[2].height = 24
    ws4.row_dimensions[3].height = 8

    headers4 = [
        "STT Thửa", 
        "Số Tờ", 
        "Số Thửa", 
        "Mã Thửa Đất", 
        "Quận / Huyện", 
        "Phường / Xã", 
        "Loại Quy Hoạch",
        "Tên Dự Án / Đồ Án Điều Chỉnh",
        "Số Quyết Định Phê Duyệt",
        "Ngày Phê Duyệt",
        "Cơ Quan Phê Duyệt",
        "Diện Tích Dự Án (m²)",
        "Tỷ Lệ Thửa Thuộc Dự Án (%)"
    ]
    ws4.row_dimensions[4].height = 34
    
    for col_idx, h_text in enumerate(headers4, 1):
        c = ws4.cell(row=4, column=col_idx)
        c.value = h_text
        c.font = FONT_HEADER
        c.fill = FILL_AMBER
        c.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
        c.border = BORDER_THIN
        
    col_widths4 = {
        'A': 10, 'B': 10, 'C': 10, 'D': 18, 'E': 16, 'F': 22,
        'G': 22, 'H': 36, 'I': 22, 'J': 16, 'K': 24, 'L': 20, 'M': 20
    }
    for col_l, w in col_widths4.items():
        ws4.column_dimensions[col_l].width = w

    ws4.freeze_panes = "A5"
    ws4.auto_filter.ref = "A4:M4"

    wb.save(filepath)
    return wb, ws1, ws2, ws3, ws4

async def run_pipeline(args):
    # 1. Xác định Thư mục đầu ra và nạp cấu hình cũ nếu đang Resume
    timestamp_str = datetime.now().strftime("%Y%m%d_%H%M%S")
    output_dir = getattr(args, 'out_dir', None)
    if not output_dir:
        if os.path.dirname(args.output):
            output_dir = os.path.dirname(args.output)
            excel_filename = os.path.basename(args.output)
        else:
            output_dir = os.path.join("data", "output", f"khao_sat_{timestamp_str}")
            excel_filename = args.output
    else:
        if not os.path.isabs(output_dir) and not output_dir.startswith("data/output") and not output_dir.startswith("data" + os.sep + "output"):
            output_dir = os.path.join("data", "output", output_dir)
        excel_filename = os.path.basename(args.output) if args.output.endswith(".xlsx") else "ket_qua.xlsx"

    os.makedirs(output_dir, exist_ok=True)
    excel_path = os.path.abspath(os.path.join(output_dir, excel_filename))
    map_html_path = os.path.abspath(os.path.join(output_dir, "ban_do_quy_hoach.html"))
    data_js_path = os.path.abspath(os.path.join(output_dir, "ban_do_data.js"))
    data_json_path = os.path.abspath(os.path.join(output_dir, "ban_do_data.json"))
    config_json_path = os.path.abspath(os.path.join(output_dir, "khao_sat_config.json"))
    args.output = excel_path
    args.map = map_html_path

    # Nếu đang chạy Resume, tự động nạp cấu hình gốc của đợt khảo sát
    is_resuming = getattr(args, 'resume', False) or os.path.exists(excel_path)
    if is_resuming and os.path.exists(config_json_path):
        try:
            with open(config_json_path, "r", encoding="utf-8") as f:
                cfg = json.load(f)
                if getattr(args, 'grid', None) is None or getattr(args, 'grid', 30) == 30:
                    args.grid = cfg.get("grid", getattr(args, 'grid', 30))
                if not getattr(args, 'kml_file', None) or args.kml_file == DEFAULT_URL:
                    args.kml_file = cfg.get("kml_file", args.kml_file)
                if getattr(args, 'buffer', 0) == 0 and "buffer" in cfg:
                    args.buffer = cfg.get("buffer", 0)
                if not getattr(args, 'focus', None) and "focus" in cfg:
                    args.focus = cfg.get("focus")
        except Exception:
            pass

    # 2. Phân tích KML hoặc Google My Maps và tính toán vùng quét
    print(f"[*] Đang xử lý nguồn dữ liệu: {args.kml_file}")
    focus_arg = None
    if getattr(args, 'focus', None):
        parts = args.focus.split(',')
        if len(parts) >= 2:
            focus_arg = (float(parts[0].strip()), float(parts[1].strip()))
            
    try:
        grid_val = getattr(args, 'grid', 30) or 30
        spatial_data = get_spatial_data(
            args.kml_file, 
            buffer_meters=getattr(args, 'buffer', 0) or 0, 
            grid_spacing_meters=grid_val,
            focus_coord=focus_arg
        )
    except Exception as e:
        print(f"[!] Lỗi khi xử lý dữ liệu toạ độ: {e}")
        return

    route_coords = spatial_data["route_coords"]
    buffer_coords = spatial_data["buffer_coords"]
    all_points = spatial_data["grid_points"]

    if not all_points:
        print("[!] Không có toạ độ nào được tạo ra trong phạm vi vùng đệm.")
        return

    # Lưu lại cấu hình chuẩn xác cho các lần Resume sau
    try:
        cfg_data = {
            "kml_file": str(args.kml_file),
            "grid": grid_val,
            "buffer": getattr(args, 'buffer', 0) or 0,
            "focus": getattr(args, 'focus', None),
            "total_points": len(all_points)
        }
        with open(config_json_path, "w", encoding="utf-8") as f:
            json.dump(cfg_data, f, ensure_ascii=False, indent=2)
    except Exception:
        pass

    # Phục hồi dữ liệu từ đợt quét trước nếu đang chạy chế độ Resume
    seen_parcels = {}
    scanned_points = {}
    parcel_count = [0]
    
    if is_resuming:
        if os.path.exists(data_json_path):
            try:
                with open(data_json_path, "r", encoding="utf-8") as f:
                    j_data = json.load(f)
                    seen_parcels = j_data.get("parcels", {})
                    scanned_points = j_data.get("scanned_points", {})
            except Exception:
                pass
        if not seen_parcels and os.path.exists(excel_path):
            try:
                seen_parcels = load_data_from_excel(excel_path)
            except Exception:
                seen_parcels = {}

    # Lọc danh sách điểm quét thực tế (Bỏ qua các điểm đã quét trong quá khứ)
    if is_resuming and scanned_points:
        unscanned_points = [pt for pt in all_points if f"{pt[0]:.6f},{pt[1]:.6f}" not in scanned_points]
        
        print("\n┌" + "─"*72 + "┐")
        print("│ ⏩ TIẾP TỤC ĐỢT QUÉT (RESUME):                                        │")
        print(f"│  • Đã phục hồi từ đợt trước: {len(seen_parcels):<5} thửa đất | {len(scanned_points):<6,} toạ độ đã quét   │")
        print(f"│  • Tổng số điểm toàn tuyến:  {len(all_points):<6,} điểm (Mật độ lưới {grid_val}m)          │")
        print(f"│  • Số điểm còn lại cần quét: {len(unscanned_points):<6,} điểm                               │")
        print("└" + "─"*72 + "┘\n")
        
        if not unscanned_points:
            print("[✨] Đợt khảo sát này đã hoàn thành 100%! Không còn điểm nào cần quét thêm.")
            return

        scan_limit = getattr(args, 'limit', 0) or 0
        if scan_limit > 0:
            points_to_scan = unscanned_points[:scan_limit]
            print(f"[*] Quét {len(points_to_scan):,} điểm tiếp theo theo giới hạn --limit={scan_limit}...")
        else:
            points_to_scan = unscanned_points
            print(f"[*] Đang tiếp tục quét toàn bộ {len(points_to_scan):,} điểm còn lại cho đến khi hoàn tất 100%...")
    else:
        total_points_available = len(all_points)
        print(f"[*] Tổng số điểm lưới bao phủ phạm vi: {total_points_available:,} (Mật độ {grid_val}m)")
        scan_limit = getattr(args, 'limit', 0) or 0
        if scan_limit > 0 and len(all_points) > scan_limit:
            print(f"[*] Giới hạn quét {scan_limit:,} điểm đầu tiên theo tham số --limit...")
            points_to_scan = all_points[:scan_limit]
        else:
            points_to_scan = all_points

    # Khởi tạo Bản đồ trực quan công trình (Lưu bản đầy đủ tĩnh lần đầu)
    generate_interactive_map(
        route_coords, 
        buffer_coords, 
        all_points, 
        output_html=map_html_path, 
        initial_parcels=seen_parcels
    )
    # Khởi tạo file ban_do_data.js / json tĩnh ban đầu
    update_live_data(data_js_path, route_coords, buffer_coords, all_points, scanned_points, seen_parcels, is_static_init=True)

    if not args.no_browser:
        print(f"[*] Tự động mở bản đồ trực quan trên trình duyệt: {map_html_path}")
        webbrowser.open(f"file://{map_html_path}")

    # 3. Khởi tạo / Nạp file Excel lưu trữ chuyên nghiệp 4 Sheet
    source_label = "Google My Maps" if "google.com/maps" in str(args.kml_file) else os.path.basename(str(args.kml_file))
    wb, ws1, ws2, ws3, ws4, is_loaded = get_or_create_workbook(args.output, source_name=source_label)
    
    if is_loaded:
        max_stt = 0
        for r in range(5, ws1.max_row + 1):
            val = ws1.cell(r, 1).value
            if isinstance(val, int):
                max_stt = max(max_stt, val)
            elif val is not None:
                max_stt += 1
        parcel_count[0] = max_stt
        print(f"[*] ⏩ TIẾP TỤC ĐỢT QUÉT (RESUME): Đã phục hồi {len(seen_parcels)} thửa đất & {len(scanned_points)} toạ độ đã quét từ đợt trước.")
    else:
        print(f"[*] Khởi tạo cấu trúc file Excel 4 Sheet chuyên nghiệp: {args.output}")
    
    unsaved_parcels = [0]
    file_io_lock = threading.Lock()

    def sync_live_map(force=False):
        try:
            update_live_data(data_js_path, route_coords, buffer_coords, points, scanned_points, seen_parcels, is_static_init=force)
        except Exception:
            pass

    # 4. Callback xử lý sau mỗi điểm quét (Siêu tốc trên RAM - 0ms Disk I/O)
    def on_point_scraped(lon, lat, parsed_info, idx, total, is_skipped=False, worker_id=1):
        time_str = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
        point_key = f"{lon:.6f},{lat:.6f}"
        
        # Xử lý trường hợp AUTO-NEXT: Điểm nằm trong ranh thửa đất đã quét
        if is_skipped and parsed_info:
            scanned_points[point_key] = True
            mathua = parsed_info.get("mathuadat", "")
            sothua = parsed_info.get("sothua", "-")
            soto = parsed_info.get("soto", "-")
            print(f"[*] [Luồng #{worker_id}] [{idx}/{total}] ⏩ AUTO-NEXT (0s): Toạ độ ({lat:.6f}, {lon:.6f}) thuộc ranh Thửa {sothua}/Tờ {soto} (Mã {mathua}) -> Bỏ qua truy vấn mạng.")
            return
            
        if parsed_info and parsed_info.get("mathuadat"):
            mathua = parsed_info["mathuadat"]
            scanned_points[point_key] = True
            
            # Kiểm tra nếu là thửa đất mới chưa từng xuất hiện
            if mathua not in seen_parcels:
                seen_parcels[mathua] = parsed_info
                with file_io_lock:
                    parcel_count[0] += 1
                    curr_stt = parcel_count[0]
                    unsaved_parcels[0] += 1
                
                # ----------------------------------------------------
                # 4.1 Ghi vào SHEET 1: Tổng Hợp Thửa Đất (Master Overview)
                # ----------------------------------------------------
                qhpk_list = parsed_info.get("qhpk_details", [])
                so_o_qh = len(qhpk_list) if qhpk_list else 1
                ranh_json = json.dumps(parsed_info.get("ranh_coords", []), ensure_ascii=False) if parsed_info.get("ranh_coords") else ""
                
                row1 = [
                    curr_stt,
                    parsed_info.get("soto", "-"),
                    parsed_info.get("sothua", "-"),
                    parsed_info.get("mathuadat", ""),
                    parsed_info.get("tenquanhuyen", ""),
                    parsed_info.get("tenphuongxa", ""),
                    parsed_info.get("dientich", 0.0),
                    so_o_qh,
                    parsed_info.get("chucnang_summary", ""),
                    parsed_info.get("chitieu_summary", ""),
                    parsed_info.get("tendoan", ""),
                    parsed_info.get("qhct_summary", ""),
                    parsed_info.get("logioi_summary", ""),
                    f"{lon:.6f}, {lat:.6f}",
                    time_str,
                    ranh_json
                ]
                ws1.append(row1)
                
                # Format hàng vừa thêm trong Sheet 1
                curr_row_idx1 = ws1.max_row
                fill_row1 = FILL_ALT if curr_stt % 2 == 0 else PatternFill(fill_type=None)
                
                for c_idx in range(1, len(row1) + 1):
                    cell = ws1.cell(row=curr_row_idx1, column=c_idx)
                    cell.font = FONT_DATA
                    cell.border = BORDER_THIN
                    if fill_row1.fill_type: cell.fill = fill_row1
                        
                    if c_idx in (1, 2, 3, 4, 8): # STT, Tờ, Thửa, Mã thửa, Số ô
                        cell.alignment = Alignment(horizontal="center", vertical="top")
                    elif c_idx == 7: # Diện tích
                        cell.alignment = Alignment(horizontal="right", vertical="top")
                        cell.number_format = "#,##0.00"
                    elif c_idx in (5, 6, 14, 15, 16):
                        cell.alignment = Alignment(horizontal="left", vertical="top")
                    else:
                        cell.alignment = Alignment(horizontal="left", vertical="top", wrap_text=True)
                        
                # ----------------------------------------------------
                # 4.2 Ghi vào SHEET 2: Chi Tiết Các Ô Chức Năng & KT
                # ----------------------------------------------------
                if qhpk_list:
                    for o_idx, o_item in enumerate(qhpk_list, 1):
                        tangcao_val = o_item.get("tangcao") or "-"
                        chieucao_val = o_item.get("chieucao") or "-"
                        matdo_val = o_item.get("matdo") or "-"
                        hesosdd_val = o_item.get("hesosdd") or "-"
                        danso_val = o_item.get("danso") or "-"
                        
                        row2 = [
                            curr_stt,
                            parsed_info.get("soto", "-"),
                            parsed_info.get("sothua", "-"),
                            parsed_info.get("mathuadat", ""),
                            parsed_info.get("tenquanhuyen", ""),
                            parsed_info.get("tenphuongxa", ""),
                            o_idx,
                            o_item.get("maopho", "") or "-",
                            o_item.get("chucnang", ""),
                            o_item.get("chucnangct", "") or o_item.get("chucnang", ""),
                            o_item.get("dientich", 0.0),
                            (o_item.get("tldientich", 0.0) / 100.0) if o_item.get("tldientich") else 0.0,
                            tangcao_val,
                            chieucao_val,
                            matdo_val,
                            hesosdd_val,
                            danso_val,
                            o_item.get("dientich_opho", None),
                            parsed_info.get("tendoan", "")
                        ]
                        ws2.append(row2)
                        
                        curr_row_idx2 = ws2.max_row
                        fill_row2 = FILL_ALT if curr_stt % 2 == 0 else PatternFill(fill_type=None)
                        for c_idx in range(1, len(row2) + 1):
                            cell = ws2.cell(row=curr_row_idx2, column=c_idx)
                            cell.font = FONT_DATA
                            cell.border = BORDER_THIN
                            if fill_row2.fill_type: cell.fill = fill_row2
                            
                            if c_idx in (1, 2, 3, 4, 7, 8, 13, 14, 15, 16, 17):
                                cell.alignment = Alignment(horizontal="center", vertical="center")
                            elif c_idx == 11:
                                cell.alignment = Alignment(horizontal="right", vertical="center")
                                cell.number_format = "#,##0.00"
                            elif c_idx == 12:
                                cell.alignment = Alignment(horizontal="right", vertical="center")
                                cell.number_format = "0.0%"
                            elif c_idx == 18 and isinstance(cell.value, (int, float)):
                                cell.alignment = Alignment(horizontal="right", vertical="center")
                                cell.number_format = "#,##0.00"
                            else:
                                cell.alignment = Alignment(horizontal="left", vertical="center")
                else:
                    # Ghi 1 dòng mặc định nếu không có ô chức năng cụ thể
                    row2_empty = [
                        curr_stt, parsed_info.get("soto", "-"), parsed_info.get("sothua", "-"),
                        parsed_info.get("mathuadat", ""), parsed_info.get("tenquanhuyen", ""), parsed_info.get("tenphuongxa", ""),
                        1, "-", "Chưa phân định", "-", parsed_info.get("dientich", 0.0), 1.0, "-", "-", "-", "-", "-", None, parsed_info.get("tendoan", "")
                    ]
                    ws2.append(row2_empty)

                # ----------------------------------------------------
                # 4.3 Ghi vào SHEET 3: Chi Tiết Lộ Giới Tuyến Đường
                # ----------------------------------------------------
                logioi_list = parsed_info.get("logioi_details", [])
                if logioi_list:
                    for lg_idx, lg_item in enumerate(logioi_list, 1):
                        row3 = [
                            curr_stt,
                            parsed_info.get("soto", "-"),
                            parsed_info.get("sothua", "-"),
                            parsed_info.get("mathuadat", ""),
                            parsed_info.get("tenquanhuyen", ""),
                            parsed_info.get("tenphuongxa", ""),
                            lg_idx,
                            lg_item.get("tenduong", ""),
                            lg_item.get("logioi", ""),
                            lg_item.get("huongtiepgiap", "") or "-",
                            lg_item.get("chieungang", None),
                            lg_item.get("chieusau", None),
                            lg_item.get("dientichxd", None)
                        ]
                        ws3.append(row3)
                        
                        curr_row_idx3 = ws3.max_row
                        fill_row3 = FILL_ALT if curr_stt % 2 == 0 else PatternFill(fill_type=None)
                        for c_idx in range(1, len(row3) + 1):
                            cell = ws3.cell(row=curr_row_idx3, column=c_idx)
                            cell.font = FONT_DATA
                            cell.border = BORDER_THIN
                            if fill_row3.fill_type: cell.fill = fill_row3
                            
                            if c_idx in (1, 2, 3, 4, 7, 9, 10):
                                cell.alignment = Alignment(horizontal="center", vertical="center")
                            elif c_idx in (11, 12, 13) and isinstance(cell.value, (int, float)):
                                cell.alignment = Alignment(horizontal="right", vertical="center")
                                cell.number_format = "#,##0.00"
                            else:
                                cell.alignment = Alignment(horizontal="left", vertical="center")
                else:
                    # Ghi 1 dòng thông báo không có lộ giới
                    row3_empty = [
                        curr_stt, parsed_info.get("soto", "-"), parsed_info.get("sothua", "-"),
                        parsed_info.get("mathuadat", ""), parsed_info.get("tenquanhuyen", ""), parsed_info.get("tenphuongxa", ""),
                        1, "Không có thông tin lộ giới", "-", "-", None, None, None
                    ]
                    ws3.append(row3_empty)

                # ----------------------------------------------------
                # 4.4 Ghi vào SHEET 4: Chi Tiết Dự Án 1/500 & ĐCCB
                # ----------------------------------------------------
                qhct_list = parsed_info.get("qhct_details", [])
                dccb_list = parsed_info.get("dccb_details", [])
                
                if qhct_list:
                    for qhct_item in qhct_list:
                        row4 = [
                            curr_stt,
                            parsed_info.get("soto", "-"),
                            parsed_info.get("sothua", "-"),
                            parsed_info.get("mathuadat", ""),
                            parsed_info.get("tenquanhuyen", ""),
                            parsed_info.get("tenphuongxa", ""),
                            "Quy hoạch chi tiết 1/500",
                            qhct_item.get("tenduan", ""),
                            qhct_item.get("soqd", ""),
                            qhct_item.get("ngayduyet", ""),
                            qhct_item.get("coquanpd", ""),
                            qhct_item.get("dientich", None),
                            (qhct_item.get("tldientich") / 100.0) if qhct_item.get("tldientich") else None
                        ]
                        ws4.append(row4)
                        curr_row_idx4 = ws4.max_row
                        for c_idx in range(1, len(row4) + 1):
                            cell = ws4.cell(row=curr_row_idx4, column=c_idx)
                            cell.font = FONT_DATA
                            cell.border = BORDER_THIN
                            if c_idx in (1, 2, 3, 4, 7, 9, 10):
                                cell.alignment = Alignment(horizontal="center", vertical="center")
                            elif c_idx == 12 and isinstance(cell.value, (int, float)):
                                cell.alignment = Alignment(horizontal="right", vertical="center")
                                cell.number_format = "#,##0.00"
                            elif c_idx == 13 and isinstance(cell.value, (int, float)):
                                cell.alignment = Alignment(horizontal="right", vertical="center")
                                cell.number_format = "0.0%"
                            else:
                                cell.alignment = Alignment(horizontal="left", vertical="center")
                                
                if dccb_list:
                    for dccb_item in dccb_list:
                        row4 = [
                            curr_stt,
                            parsed_info.get("soto", "-"),
                            parsed_info.get("sothua", "-"),
                            parsed_info.get("mathuadat", ""),
                            parsed_info.get("tenquanhuyen", ""),
                            parsed_info.get("tenphuongxa", ""),
                            "Điều chỉnh cục bộ (DCCB)",
                            dccb_item.get("tendccb", "") or dccb_item.get("tendoan", ""),
                            dccb_item.get("soqd", ""),
                            dccb_item.get("ngayduyet", ""),
                            dccb_item.get("coquanpd", ""),
                            None,
                            None
                        ]
                        ws4.append(row4)
                        curr_row_idx4 = ws4.max_row
                        for c_idx in range(1, len(row4) + 1):
                            cell = ws4.cell(row=curr_row_idx4, column=c_idx)
                            cell.font = FONT_DATA
                            cell.border = BORDER_THIN
                            if c_idx in (1, 2, 3, 4, 7, 9, 10):
                                cell.alignment = Alignment(horizontal="center", vertical="center")
                            else:
                                cell.alignment = Alignment(horizontal="left", vertical="center")

                    # Cập nhật vùng lọc tự động cho tất cả các sheet
                    ws1.auto_filter.ref = f"A4:P{ws1.max_row}"
                    ws2.auto_filter.ref = f"A4:S{ws2.max_row}"
                    ws3.auto_filter.ref = f"A4:M{ws3.max_row}"
                    if ws4.max_row > 4:
                        ws4.auto_filter.ref = f"A4:M{ws4.max_row}"
                
                print(f"[*] [Luồng #{worker_id}] [{idx}/{total}] 🏢 PHÁT HIỆN THỬA MỚI: Thửa {parsed_info['sothua']} / Tờ {parsed_info['soto']} ({parsed_info['dientich_formatted']} m²) - {parsed_info['tenphuongxa']} -> Đã ghi nhận vào bộ nhớ.")
            else:
                print(f"[*] [Luồng #{worker_id}] [{idx}/{total}] 🔁 Toạ độ ({lat:.6f}, {lon:.6f}) thuộc Thửa {parsed_info['sothua']} (Mã {mathua}) đã lưu.")
        else:
            scanned_points[point_key] = False
            print(f"[*] [Luồng #{worker_id}] [{idx}/{total}] ℹ️ Không phát hiện thửa đất tại toạ độ ({lat:.6f}, {lon:.6f})")

    # Tiến trình tự động lưu Excel và Live Map ngầm bất đồng bộ (Zero Block trên Worker Pool)
    async def auto_save_loop():
        while True:
            await asyncio.sleep(4.0)
            try:
                await asyncio.to_thread(update_live_data, data_js_path, route_coords, buffer_coords, all_points, scanned_points, seen_parcels, False)
                if unsaved_parcels[0] > 0:
                    with file_io_lock:
                        unsaved_parcels[0] = 0
                    await asyncio.to_thread(wb.save, args.output)
            except Exception:
                pass

    # 5. Khởi chạy quét dữ liệu
    print(f"[*] Khởi động trình duyệt Playwright để thu thập dữ liệu...")
    headless_mode = not args.show_browser
    auto_saver_task = asyncio.create_task(auto_save_loop())
    
    try:
        concurrency_val = getattr(args, 'concurrency', 15)
        await fetch_planning_data(
            points_to_scan, 
            on_point_scraped=on_point_scraped, 
            headless=headless_mode,
            concurrency=concurrency_val,
            initial_parcels=seen_parcels,
            initial_scanned_points=scanned_points
        )
        print(f"\n[✨] HOÀN TẤT QUÉT! Đã quét xong toàn bộ đợt khảo sát, phát hiện {parcel_count[0]} thửa đất/công trình độc lập.")
        print(f"[*] Báo cáo chuyên nghiệp 4 Sheet đã lưu đầy đủ tại: {args.output}")
        print(f"[*] Bản đồ trực quan hiển thị tại: {map_html_path}")
    except KeyboardInterrupt:
        print(f"\n[!] Người dùng đã tạm dừng tiến trình (Ctrl+C).")
        print(f"[*] Toàn bộ {parcel_count[0]} thửa đất đã phát hiện đều đã được lưu an toàn tại '{args.output}'.")
    finally:
        auto_saver_task.cancel()
        try:
            sync_live_map(force=False)
            wb.save(args.output)
        except Exception:
            pass

def interactive_wizard():
    """
    Giao diện dòng lệnh tương tác thông minh (CLI Wizard) tối ưu trải nghiệm người dùng:
    Hỏi nguồn dữ liệu, mật độ lưới, diện tích/số điểm quét, file xuất với gợi ý mặc định tiện lợi.
    """
    DEFAULT_URL = "https://www.google.com/maps/d/u/0/viewer?mid=1IjTEVFWwFg8n7OdD1uDaymCB2Q6aUTE&ll=10.821633009544064%2C106.62770133828562&z=16"
    
    print("\n" + "="*80)
    print("🏛️   HỆ THỐNG KHẢO SÁT & TRÍCH XUẤT QUY HOẠCH ĐÔ THỊ TP. HỒ CHÍ MINH")
    print("    (Đồng bộ trực tiếp Cổng thông tin SQHKT: sqhkt-qlqh.tphcm.gov.vn)")
    print("="*80)

    try:
        # [Bước 0] Chọn tác vụ
        print("\n[Bước 0] 🎯 BẠN MUỐN THỰC HIỆN TÁC VỤ GÌ?")
        print("  [1] 🚀 Khảo sát & cào trích xuất quy hoạch mới (Google My Maps / KML)")
        print("  [2] ⏩ Tiếp tục đợt quét dở dang (Resume / Quét nối tiếp từ thư mục cũ)")
        print("  [3] 📂 Mở lại bản đồ từ file Excel đã lưu để xem & xoá thửa đất (Quản trị)")
        task_mode = input("👉 Chọn tác vụ [1-3, Nhấn Enter chọn 1]: ").strip()
        
        if task_mode == "3":
            from src.viewer import start_viewer_server
            print("\n  Nhập đường dẫn file Excel (vd: data/output/khao_sat_.../ket_qua.xlsx) hoặc thư mục kết quả:")
            target_in = input("👉 Đường dẫn [Nhấn Enter để mở danh sách đợt khảo sát tại data/output/]: ").strip()
            if not target_in:
                target_in = os.path.join("data", "output")
            start_viewer_server(target_in)
            sys.exit(0)
            
        elif task_mode == "2":
            print("\n  Chọn đợt khảo sát dở dang cần tiếp tục quét:")
            target_in = input("👉 Đường dẫn thư mục hoặc file Excel [Nhấn Enter để chọn từ data/output/]: ").strip()
            if not target_in:
                target_in = os.path.join("data", "output")
            excel_path, work_dir = find_excel_and_workdir(target_in)
            
            print(f"\n[*] Đang chuẩn bị quét nối tiếp cho thư mục: {work_dir}")
            user_source = input(f"👉 Nhập link / file KML của đợt này [Nhấn Enter để dùng ranh Metro Tuyến 2]: ").strip()
            if not user_source:
                user_source = DEFAULT_URL
                
            grid_choice = input("👉 Mật độ lưới quét (mét) [Nhấn Enter chọn 30m]: ").strip()
            grid_val = int(grid_choice) if grid_choice.isdigit() and int(grid_choice) > 0 else 30
            
            concurrency_choice = input("👉 Số luồng quét song song [1-30, Nhấn Enter chọn 15]: ").strip()
            concurrency_val = int(concurrency_choice) if concurrency_choice.isdigit() and 1 <= int(concurrency_choice) <= 30 else 15
            
            limit_choice = input("👉 Giới hạn số điểm quét [0 để quét toàn bộ, Nhấn Enter chọn 0]: ").strip()
            scan_limit = int(limit_choice) if limit_choice.isdigit() and int(limit_choice) >= 0 else 0
            
            class ResumeArgs:
                pass
            args = ResumeArgs()
            args.kml_file = user_source
            args.out_dir = work_dir
            args.output = excel_path
            args.buffer = 0
            args.grid = grid_val
            args.concurrency = concurrency_val
            args.limit = scan_limit
            args.focus = None
            args.map = os.path.join(work_dir, "ban_do_quy_hoach.html")
            args.no_browser = False
            args.show_browser = False
            args.resume = True
            return args

        # [1/4] Nguồn dữ liệu
        print("\n[Bước 1/4] 🌐 NGUỒN DỮ LIỆU RANH GIỚI KHẢO SÁT")
        print("  • Dán link Google My Maps (vd: https://www.google.com/maps/d/...)")
        print("  • Hoặc đường dẫn file KML trong máy (vd: sample_route.kml hoặc data/input/sample_route.kml)")
        user_source = input("👉 Nhập link / file KML [Nhấn Enter để dùng ranh Metro Tuyến 2]: ").strip()
        if not user_source:
            user_source = DEFAULT_URL
            print("   -> Đã chọn: Tuyến Metro Số 2 Bến Thành - Tham Lương (Google My Maps)")

        # [2/4] Mật độ lưới
        print("\n[Bước 2/4] 📏 MẬT ĐỘ LƯỚI QUÉT (Khoảng cách giữa các điểm toạ độ)")
        print("  Khoảng cách càng nhỏ thì mật độ quét càng dày, bắt trọn các thửa nhà phố nhỏ.")
        print("  [1] 30 mét  (Chuẩn - Khuyên dùng cho đô thị, cân bằng tốc độ & độ bao phủ)")
        print("  [2] 20 mét  (Dày hơn - Độ chi tiết cao, khảo sát kỹ lưỡng)")
        print("  [3] 50 mét  (Thưa - Quét nhanh toàn cảnh quy mô diện rộng)")
        print("  [4] Tuỳ chỉnh số mét riêng")
        grid_choice = input("👉 Chọn mật độ lưới [1-4, Nhấn Enter chọn 1 (30m)]: ").strip()
        if grid_choice == "2":
            grid_val = 20
        elif grid_choice == "3":
            grid_val = 50
        elif grid_choice == "4":
            custom_g = input("   Nhập khoảng cách lưới (mét, vd: 25): ").strip()
            grid_val = int(custom_g) if custom_g.isdigit() and int(custom_g) > 0 else 30
        else:
            grid_val = 30
        print(f"   -> Đã chọn mật độ lưới: {grid_val} mét")

        # [2.5] Tốc độ quét (Số luồng song song)
        print("\n[Bước 2.5] ⚡ TỐC ĐỘ QUÉT (Số luồng song song)")
        print("  Chạy đa luồng song song giúp tăng tốc độ cào gấp 5x - 20x.")
        print("  [1] 15 luồng (⚡ Siêu tốc - Khuyên dùng)")
        print("  [2] 10 luồng (Tốc độ cao - Quét rất nhanh)")
        print("  [3] 5 luồng  (Tiêu chuẩn - Ổn định)")
        print("  [4] 3 luồng  (Cơ bản / Tiết kiệm)")
        print("  [5] Tuỳ chỉnh số luồng (1 đến 30)")
        concurrency_choice = input("👉 Chọn số luồng [1-5, Nhấn Enter chọn 1 (15 luồng)]: ").strip()
        if concurrency_choice == "1" or concurrency_choice == "":
            concurrency_val = 15
        elif concurrency_choice == "2":
            concurrency_val = 10
        elif concurrency_choice == "3":
            concurrency_val = 5
        elif concurrency_choice == "4":
            concurrency_val = 3
        elif concurrency_choice == "5":
            custom_c = input("   Nhập số luồng (1-30, vd: 15): ").strip()
            concurrency_val = int(custom_c) if custom_c.isdigit() and 1 <= int(custom_c) <= 30 else 15
        else:
            concurrency_val = 15
        print(f"   -> Đã chọn: {concurrency_val} luồng chạy song song")

        # [3/4] Phân tích không gian & Phạm vi quét
        print(f"\n[*] Đang phân tích ranh giới không gian địa lý...")
        try:
            spatial_data = get_spatial_data(user_source, buffer_meters=0, grid_spacing_meters=grid_val)
            total_pts = len(spatial_data["grid_points"])
            polys = spatial_data.get("buffer_coords", [])
            
            print("\n┌" + "─"*72 + "┐")
            print("│ 📍 KẾT QUẢ PHÂN TÍCH VÙNG RANH GIỚI:                                   │")
            print(f"│  • Số vùng ranh khép kín: {len(polys):<44} │")
            print(f"│  • Bước nhảy lưới toạ độ: {grid_val} mét{' '*41} │")
            print(f"│  • Tổng số điểm quét bao phủ 100% diện tích: {total_pts:,} điểm{' '*23} │")
            print("└" + "─"*72 + "┘")
        except Exception as e:
            print(f"[!] Ghi chú phân tích ranh: {e}")
            total_pts = 1000

        print("\n[Bước 3/4] 🎯 PHẠM VI & SỐ LƯỢNG ĐIỂM QUÉT")
        print(f"  [1] Quét toàn bộ 100% diện tích ({total_pts:,} điểm)")
        print(f"  [2] Quét kiểm tra nhanh (15 điểm đầu tiên)")
        print(f"  [3] Quét kiểm tra vừa (50 điểm đầu tiên)")
        print(f"  [4] Tuỳ chỉnh số lượng điểm quét")
        limit_choice = input("👉 Chọn phạm vi quét [1-4, Nhấn Enter chọn 2 (15 điểm)]: ").strip()
        if limit_choice == "1":
            scan_limit = 0
            print(f"   -> Đã chọn: Quét toàn bộ 100% ({total_pts:,} điểm)")
        elif limit_choice == "3":
            scan_limit = 50
            print("   -> Đã chọn: Quét 50 điểm kiểm tra")
        elif limit_choice == "4":
            custom_l = input("   Nhập số điểm muốn quét (vd: 100, hoặc 0 để quét tất cả): ").strip()
            scan_limit = int(custom_l) if custom_l.isdigit() and int(custom_l) >= 0 else 15
            print(f"   -> Đã chọn: Quét {scan_limit} điểm")
        else:
            scan_limit = 15
            print("   -> Đã chọn: Quét kiểm tra nhanh 15 điểm")

        # [4/4] Thiết lập thư mục & file lưu trữ kết quả
        print("\n[Bước 4/4] 💾 THIẾT LẬP THƯ MỤC & FILE LƯU TRỮ KẾT QUẢ")
        print("  Mỗi lần khảo sát sẽ tự động được lưu gọn trong 1 thư mục riêng bên trong data/output/ gồm:")
        print("    • File Excel 4 Sheet chuyên nghiệp (kèm toạ độ ranh đa giác)")
        print("    • Bản đồ vệ tinh tương tác Leaflet (HTML + JS + JSON)")
        time_tag = datetime.now().strftime("%Y%m%d_%H%M%S")
        default_dir = f"khao_sat_{time_tag}"
        user_dir_input = input(f"👉 Nhập tên thư mục đợt khảo sát [Nhấn Enter chọn '{default_dir}']: ").strip()
        if not user_dir_input:
            user_dir_input = default_dir

        if not os.path.isabs(user_dir_input) and not user_dir_input.startswith("data/output") and not user_dir_input.startswith("data" + os.sep + "output"):
            user_dir = os.path.join("data", "output", user_dir_input)
        else:
            user_dir = user_dir_input

        out_name = input("👉 Tên file Excel kết quả [Nhấn Enter chọn 'ket_qua.xlsx']: ").strip()
        if not out_name:
            out_name = "ket_qua.xlsx"
        if not out_name.endswith(".xlsx"):
            out_name += ".xlsx"
        print(f"   -> Thư mục kết quả: {user_dir}/")
        print(f"   -> File Excel: {os.path.join(user_dir, out_name)} (4 Sheet)")

        open_browser = input("👉 Tự động mở bản đồ trực quan trên trình duyệt? (Y/n) [Enter chọn Có]: ").strip().lower()
        no_browser_flag = (open_browser == 'n')

        # Tổng kết trước khi chạy
        print("\n" + "="*80)
        print("🚀 BẮT ĐẦU TIẾN TRÌNH KHẢO SÁT QUY HOẠCH...")
        print(f"  • Nguồn ranh: {user_source[:60]}...")
        print(f"  • Mật độ lưới: {grid_val}m")
        print(f"  • Số luồng song song: {concurrency_val} luồng")
        print(f"  • Số lượng quét: {'Toàn bộ 100% (' + str(total_pts) + ' điểm)' if scan_limit == 0 else str(scan_limit) + ' điểm'}")
        print(f"  • Thư mục lưu kết quả: {user_dir}/")
        print(f"  • File Excel xuất: {os.path.join(user_dir, out_name)} (4 Sheet)")
        print(f"  • Bản đồ trực quan: {os.path.join(user_dir, 'ban_do_quy_hoach.html')}")
        print("  • Lưu tức thì: BẬT (Bảo toàn dữ liệu an toàn khi bấm Ctrl+C)")
        print("="*80 + "\n")

    except (KeyboardInterrupt, EOFError):
        print("\n[!] Người dùng đã huỷ thao tác.")
        sys.exit(0)

    class Args:
        pass
    args = Args()
    args.kml_file = user_source
    args.out_dir = user_dir
    args.output = os.path.join(user_dir, out_name)
    args.buffer = 0
    args.grid = grid_val
    args.concurrency = concurrency_val
    args.limit = scan_limit
    args.focus = None
    args.map = "ban_do_quy_hoach.html"
    args.no_browser = no_browser_flag
    args.show_browser = False
    args.resume = False
    return args

def main():
    parser = argparse.ArgumentParser(description="Quét và trích xuất thông tin quy hoạch công trình từ KML hoặc Google My Maps.")
    parser.add_argument("kml_file", nargs="?", default=None, help="Đường dẫn file KML hoặc đường link Google My Maps (vd: https://www.google.com/maps/d/...)")
    parser.add_argument("--output", default="ket_qua.xlsx", help="Tên file Excel đầu ra (mặc định: ket_qua.xlsx)")
    parser.add_argument("--out-dir", "-d", help="Thư mục lưu trữ kết quả (chứa cả file Excel và Bản đồ HTML)")
    parser.add_argument("--buffer", type=int, default=0, help="Bán kính vùng đệm quét theo mét (mặc định: 0 đối với vùng ranh đa giác)")
    parser.add_argument("--grid", type=int, default=None, help="Khoảng cách giữa các điểm lưới theo mét (mặc định: 30)")
    parser.add_argument("--concurrency", "-c", type=int, default=15, help="Số luồng chạy song song (mặc định: 15)")
    parser.add_argument("--limit", type=int, default=0, help="Giới hạn số điểm quét kiểm tra (mặc định: 0 - quét toàn bộ 100%%)")
    parser.add_argument("--focus", help="Toạ độ trọng tâm khu vực khảo sát 'lat,lon'")
    parser.add_argument("--map", default="ban_do_quy_hoach.html", help="Đường dẫn file HTML bản đồ trực quan (mặc định: ban_do_quy_hoach.html)")
    parser.add_argument("--no-browser", action="store_true", help="Không tự động mở bản đồ trên trình duyệt")
    parser.add_argument("--show-browser", action="store_true", help="Hiển thị cửa sổ trình duyệt Chromium khi cào (mặc định chạy ngầm)")
    parser.add_argument("-i", "--interactive", action="store_true", help="Kích hoạt chế độ tương tác hỏi đáp từng bước (Interactive Wizard)")
    parser.add_argument("--view", "-v", help="Mở lại bản đồ trực quan từ file Excel hoặc thư mục kết quả để xem và xoá thửa")
    parser.add_argument("--resume", "-r", help="Tiếp tục đợt quét dở dang từ thư mục kết quả hoặc file Excel")
    
    # 1. Kiểm tra cờ --view trước
    for arg_idx, arg in enumerate(sys.argv):
        if arg in ("--view", "-v"):
            view_target = sys.argv[arg_idx + 1] if arg_idx + 1 < len(sys.argv) else os.path.join("data", "output")
            from src.viewer import start_viewer_server
            start_viewer_server(view_target)
            sys.exit(0)

    # 2. Tự động kích hoạt Interactive Wizard nếu không truyền tham số
    if len(sys.argv) == 1 or (len(sys.argv) == 2 and sys.argv[1] in ('-i', '--interactive')):
        args = interactive_wizard()
    else:
        args = parser.parse_args()
        if args.resume:
            excel_path, work_dir = find_excel_and_workdir(args.resume)
            args.out_dir = work_dir
            args.output = excel_path
            args.resume = True
            if '--limit' not in sys.argv:
                args.limit = 0
            if not args.kml_file:
                args.kml_file = DEFAULT_URL
        elif not args.kml_file:
            args = interactive_wizard()
    
    try:
        asyncio.run(run_pipeline(args))
    except KeyboardInterrupt:
        print(f"\n[!] Đã dừng chương trình an toàn. File kết quả '{args.output}' đã được lưu.")
        sys.exit(0)

if __name__ == "__main__":
    main()
