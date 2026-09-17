import os
import sys
import json
import webbrowser
from http.server import ThreadingHTTPServer, SimpleHTTPRequestHandler
import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side

def load_data_from_excel(excel_path):
    """
    Đọc toàn bộ dữ liệu 4 Sheet từ file Excel để phục hồi lại dữ liệu bản đồ trực quan.
    """
    if not os.path.exists(excel_path):
        raise FileNotFoundError(f"Không tìm thấy file Excel: {excel_path}")

    wb = openpyxl.load_workbook(excel_path, data_only=True)
    
    # 1. Đọc Sheet 1: Master Overview
    ws1 = wb['1. Tổng Hợp Thửa Đất'] if '1. Tổng Hợp Thửa Đất' in wb.sheetnames else wb.active
    parcels = {}
    
    for r in range(5, ws1.max_row + 1):
        stt = ws1.cell(r, 1).value
        if stt is None:
            continue
        soto = ws1.cell(r, 2).value or "-"
        sothua = ws1.cell(r, 3).value or "-"
        mathua = str(ws1.cell(r, 4).value or f"{soto}_{sothua}")
        quanhuyen = ws1.cell(r, 5).value or ""
        phuongxa = ws1.cell(r, 6).value or ""
        dientich = ws1.cell(r, 7).value or 0.0
        chucnang_summary = ws1.cell(r, 9).value or ""
        chitieu_summary = ws1.cell(r, 10).value or ""
        tendoan = ws1.cell(r, 11).value or ""
        qhct_summary = ws1.cell(r, 12).value or ""
        logioi_summary = ws1.cell(r, 13).value or ""
        toado_str = str(ws1.cell(r, 14).value or "")
        
        # Đọc toạ độ ranh đa giác từ cột 16 (P) nếu có
        ranh_coords = []
        c16_val = ws1.cell(r, 16).value
        if c16_val and isinstance(c16_val, str) and c16_val.startswith("["):
            try:
                ranh_coords = json.loads(c16_val)
            except Exception:
                ranh_coords = []
                
        # Parse toạ độ tâm
        center_lon, center_lat = 106.6277, 10.8216
        if toado_str and "," in toado_str:
            try:
                parts = toado_str.split(",")
                center_lon = float(parts[0].strip())
                center_lat = float(parts[1].strip())
            except Exception:
                pass

        parcels[mathua] = {
            "soto": str(soto),
            "sothua": str(sothua),
            "mathuadat": mathua,
            "tenquanhuyen": quanhuyen,
            "tenphuongxa": phuongxa,
            "dientich": dientich,
            "dientich_formatted": f"{float(dientich):,.2f}" if isinstance(dientich, (int, float)) else str(dientich),
            "chucnang_summary": chucnang_summary,
            "chitieu_summary": chitieu_summary,
            "tendoan": tendoan,
            "qhct_summary": qhct_summary,
            "logioi_summary": logioi_summary,
            "center": [center_lat, center_lon],
            "ranh_coords": ranh_coords,
            "qhpk_details": [],
            "logioi_details": [],
            "qhct_details": []
        }

    # 2. Đọc Sheet 2: Chi Tiết Ô Chức Năng
    sheet2_name = None
    for s_name in ['2. Chi Tiết Ô Chức Năng & KT', '2. Chi Tiết Ô Quy Hoạch & KT']:
        if s_name in wb.sheetnames:
            sheet2_name = s_name
            break
            
    if sheet2_name:
        ws2 = wb[sheet2_name]
        for r in range(5, ws2.max_row + 1):
            mathua = str(ws2.cell(r, 4).value or "")
            if mathua in parcels:
                maopho = ws2.cell(r, 8).value or ""
                chucnang = ws2.cell(r, 9).value or ""
                chucnangct = ws2.cell(r, 10).value or ""
                dt_o = ws2.cell(r, 11).value or 0.0
                tl_o = ws2.cell(r, 12).value or 0.0
                tangcao = ws2.cell(r, 13).value or ""
                chieucao = ws2.cell(r, 14).value or ""
                matdo = ws2.cell(r, 15).value or ""
                hesosdd = ws2.cell(r, 16).value or ""
                danso = ws2.cell(r, 17).value or ""
                
                parcels[mathua]["qhpk_details"].append({
                    "maopho": maopho,
                    "chucnang": chucnang,
                    "chucnangct": chucnangct,
                    "dientich": dt_o,
                    "dientich_formatted": f"{float(dt_o):,.1f}" if isinstance(dt_o, (int, float)) else str(dt_o),
                    "tldientich_formatted": f"{float(tl_o)*100:.1f}%" if isinstance(tl_o, (int, float)) and tl_o <= 1.0 else f"{tl_o}%",
                    "tangcao": tangcao,
                    "chieucao": chieucao,
                    "matdo": matdo,
                    "hesosdd": hesosdd,
                    "danso": danso
                })

    # 3. Đọc Sheet 3: Chi Tiết Lộ Giới
    if '3. Chi Tiết Lộ Giới Tuyến Đường' in wb.sheetnames:
        ws3 = wb['3. Chi Tiết Lộ Giới Tuyến Đường']
        for r in range(5, ws3.max_row + 1):
            mathua = str(ws3.cell(r, 4).value or "")
            if mathua in parcels:
                tenduong = ws3.cell(r, 8).value or ""
                logioi = ws3.cell(r, 9).value or ""
                huong = ws3.cell(r, 10).value or ""
                parcels[mathua]["logioi_details"].append({
                    "tenduong": tenduong,
                    "logioi": logioi,
                    "huongtiepgiap": huong
                })

    # 4. Đọc Sheet 4: Chi Tiết Dự Án 1/500
    if '4. Dự Án 1-500 & Điều Chỉnh' in wb.sheetnames:
        ws4 = wb['4. Dự Án 1-500 & Điều Chỉnh']
        for r in range(5, ws4.max_row + 1):
            mathua = str(ws4.cell(r, 4).value or "")
            if mathua in parcels:
                loai = ws4.cell(r, 7).value or ""
                tenduan = ws4.cell(r, 8).value or ""
                soqd = ws4.cell(r, 9).value or ""
                ngayduyet = ws4.cell(r, 10).value or ""
                coquan = ws4.cell(r, 11).value or ""
                dt_ct = ws4.cell(r, 12).value or ""
                parcels[mathua]["qhct_details"].append({
                    "loai": loai,
                    "tenduan": tenduan,
                    "soqd": soqd,
                    "ngayduyet": ngayduyet,
                    "coquanpd": coquan,
                    "dientich": dt_ct
                })

    return parcels

def delete_parcel_from_excel(excel_path, mathua_to_delete):
    """
    Xoá hoàn toàn một thửa đất khỏi cả 4 Sheet trong file Excel và đánh số lại STT từ 1 -> N.
    """
    if not os.path.exists(excel_path):
        return {"success": False, "error": "Không tìm thấy file Excel"}

    mathua_str = str(mathua_to_delete).strip()
    wb = openpyxl.load_workbook(excel_path)
    
    total_deleted_rows = 0
    
    for sheetname in wb.sheetnames:
        ws = wb[sheetname]
        rows_to_delete = []
        for r in range(5, ws.max_row + 1):
            val_col_d = str(ws.cell(r, 4).value or "").strip()
            if val_col_d == mathua_str:
                rows_to_delete.append(r)
                
        # Xoá từ dưới lên để không bị lệch chỉ số dòng
        for r in reversed(rows_to_delete):
            ws.delete_rows(r)
            total_deleted_rows += 1
            
        # Đánh số lại cột STT (Cột 1 / A)
        if sheetname == '1. Tổng Hợp Thửa Đất' or sheetname == wb.sheetnames[0]:
            new_stt = 1
            for r in range(5, ws.max_row + 1):
                if ws.cell(r, 4).value:
                    ws.cell(r, 1).value = new_stt
                    new_stt += 1
        else:
            # Các sheet chi tiết: Đánh lại STT nhóm theo mã thửa
            current_stt = 0
            last_mathua = None
            for r in range(5, ws.max_row + 1):
                curr_m = str(ws.cell(r, 4).value or "").strip()
                if curr_m:
                    if curr_m != last_mathua:
                        current_stt += 1
                        last_mathua = curr_m
                    ws.cell(r, 1).value = current_stt

        # Cập nhật lại vùng lọc AutoFilter
        max_col_letter = openpyxl.utils.get_column_letter(ws.max_column)
        if ws.max_row >= 4:
            ws.auto_filter.ref = f"A4:{max_col_letter}{ws.max_row}"

    wb.save(excel_path)
    
    # Đồng thời cập nhật file ban_do_data.js & ban_do_data.json nếu nằm cùng thư mục
    base_dir = os.path.dirname(excel_path) or "."
    data_js_path = os.path.join(base_dir, "ban_do_data.js")
    data_json_path = os.path.join(base_dir, "ban_do_data.json")
    
    for d_path in [data_js_path, data_json_path]:
        if os.path.exists(d_path):
            try:
                with open(d_path, "r", encoding="utf-8") as f:
                    content = f.read().strip()
                if content.startswith("window.LIVE_PLANNING_DATA ="):
                    raw_json = content.replace("window.LIVE_PLANNING_DATA =", "").rstrip(";")
                else:
                    raw_json = content
                live_data = json.loads(raw_json)
                if "parcels" in live_data and mathua_str in live_data["parcels"]:
                    del live_data["parcels"][mathua_str]
                    if d_path.endswith(".js"):
                        with open(d_path, "w", encoding="utf-8") as f:
                            f.write(f"window.LIVE_PLANNING_DATA = {json.dumps(live_data, ensure_ascii=False)};")
                    else:
                        with open(d_path, "w", encoding="utf-8") as f:
                            json.dump(live_data, f, ensure_ascii=False)
            except Exception as err:
                print(f"[!] Cảnh báo cập nhật {os.path.basename(d_path)}: {err}")

    return {
        "success": True, 
        "mathua": mathua_str, 
        "deleted_rows": total_deleted_rows,
        "remaining_parcels": ws1.max_row - 4 if 'ws1' in locals() and ws1.max_row >= 4 else 0
    }

def pick_primary_excel_file(dir_path):
    """
    Tìm và ưu tiên file Excel dữ liệu khảo sát gốc trong thư mục (tránh chọn nhầm file báo cáo tổng hợp).
    Thứ tự ưu tiên:
    1. File trùng chính xác tên thư mục (ví dụ: KS003/KS003.xlsx, KS002/KS002.xlsx)
    2. File có tiền tố tên thư mục (ví dụ: KS003_*.xlsx)
    3. File chuẩn ket_qua.xlsx, ketqua.xlsx, result.xlsx
    4. Bất kỳ file .xlsx nào KHÔNG bắt đầu bằng các từ khoá báo cáo ('bao_cao', 'baocao', 'report', 'thong_ke', 'tong_hop', 'readme')
    5. Các file .xlsx còn lại (sắp xếp theo dung lượng file lớn nhất)
    """
    if not os.path.exists(dir_path) or not os.path.isdir(dir_path):
        return None
        
    candidates = [f for f in os.listdir(dir_path) if f.endswith(".xlsx") and not f.startswith("~$")]
    if not candidates:
        return None
        
    dir_name = os.path.basename(os.path.abspath(dir_path))
    
    def score_file(filename):
        f_lower = filename.lower()
        d_lower = dir_name.lower()
        full_path = os.path.join(dir_path, filename)
        file_size = os.path.getsize(full_path) if os.path.exists(full_path) else 0
        
        # 1. Trùng chính xác tên folder (ví dụ: KS003.xlsx trong folder KS003)
        if f_lower == f"{d_lower}.xlsx":
            return (0, -file_size)
        # 2. Bắt đầu bằng tên folder và không phải báo cáo
        if f_lower.startswith(d_lower) and not any(p in f_lower for p in ['bao_cao', 'baocao', 'report', 'thong_ke', 'tong_hop']):
            return (1, -file_size)
        # 3. File chuẩn ket_qua.xlsx
        if f_lower in ["ket_qua.xlsx", "ketqua.xlsx", "result.xlsx", "resuilt.xlsx"]:
            return (2, -file_size)
        # 4. File dữ liệu thông thường
        if not any(f_lower.startswith(p) for p in ['bao_cao', 'baocao', 'report', 'thong_ke', 'tong_hop', 'readme']):
            return (3, -file_size)
        # 5. File báo cáo
        return (4, -file_size)
        
    sorted_candidates = sorted(candidates, key=score_file)
    return sorted_candidates[0]

def find_excel_and_workdir(target_path):
    if os.path.isfile(target_path):
        return os.path.abspath(target_path), os.path.dirname(os.path.abspath(target_path))
    
    target_dir = os.path.abspath(target_path)
    
    # 1. Quét các thư mục con đợt khảo sát (sub-runs) nếu truyền vào data/output
    sub_runs = []
    if os.path.exists(target_dir):
        for item in sorted(os.listdir(target_dir), reverse=True):
            sub_path = os.path.join(target_dir, item)
            if os.path.isdir(sub_path):
                best_xls = pick_primary_excel_file(sub_path)
                if best_xls:
                    sub_runs.append((item, os.path.join(sub_path, best_xls), sub_path))

    # 2. Tìm file Excel tốt nhất ngay trong thư mục hiện tại
    direct_best = pick_primary_excel_file(target_dir)

    # Nếu truyền vào thư mục chứa nhiều đợt khảo sát (như data/output)
    if sub_runs and len(sub_runs) > 1 and not (direct_best and os.path.basename(target_dir).lower().startswith("ks")):
        print("\n" + "="*70)
        print("📂 DANH SÁCH CÁC ĐỢT KHẢO SÁT ĐÃ LƯU TRONG THƯ MỤC:")
        print("="*70)
        for idx, (dir_name, fpath, sdir) in enumerate(sub_runs, 1):
            file_sz = os.path.getsize(fpath) / 1024
            print(f"  [{idx}] 📁 {dir_name:<30} ({os.path.basename(fpath)} - {file_sz:.1f} KB)")
        if direct_best:
            print(f"  [0] 📄 Mở file tại thư mục gốc ({direct_best})")
            
        try:
            choice = input(f"\n👉 Chọn đợt khảo sát muốn mở [1-{len(sub_runs)}, Nhấn Enter chọn 1 (mới nhất)]: ").strip()
            if choice == "0" and direct_best:
                return os.path.join(target_dir, direct_best), target_dir
            sel_idx = int(choice) - 1 if choice.isdigit() and 1 <= int(choice) <= len(sub_runs) else 0
            selected = sub_runs[sel_idx]
            print(f"   -> Đang mở đợt khảo sát: {selected[0]} ({os.path.basename(selected[1])})")
            return selected[1], selected[2]
        except Exception:
            return sub_runs[0][1], sub_runs[0][2]

    # Nếu chính thư mục này có file Excel trực tiếp
    if direct_best:
        return os.path.join(target_dir, direct_best), target_dir

    if sub_runs:
        return sub_runs[0][1], sub_runs[0][2]

    # 3. Thử tìm trong data/output
    default_out_dir = os.path.join(os.path.dirname(os.path.dirname(os.path.abspath(__file__))), "data", "output")
    if os.path.exists(default_out_dir) and target_dir != default_out_dir:
        return find_excel_and_workdir(default_out_dir)

    raise FileNotFoundError(f"Không tìm thấy file Excel hoặc đợt khảo sát nào trong: {target_dir}")

def start_viewer_server(folder_or_excel, port=8765):
    """
    Khởi động máy chủ xem bản đồ trực quan & quản trị xoá thửa tương tác hai chiều.
    """
    excel_path, work_dir = find_excel_and_workdir(folder_or_excel)

    html_path = os.path.join(work_dir, "ban_do_quy_hoach.html")
    data_js_path = os.path.join(work_dir, "ban_do_data.js")

    # Nếu chưa có ban_do_data.js hoặc html, tự động dựng từ file Excel
    try:
        from src.visualizer import generate_interactive_map, update_live_data
    except ImportError:
        from visualizer import generate_interactive_map, update_live_data
    parcels = load_data_from_excel(excel_path)
    
    # Lấy toạ độ trọng tâm
    c_lat, c_lon = 10.8216, 106.6277
    for p in parcels.values():
        if p.get("center"):
            c_lat, c_lon = p["center"]
            break

    print(f"[*] Đang nạp và đồng bộ giao diện bản đồ trực quan từ file Excel '{os.path.basename(excel_path)}' ({len(parcels)} thửa)...")
    generate_interactive_map(
        route_coords=[],
        buffer_coords=[],
        grid_points=[],
        output_html=html_path,
        initial_parcels=parcels
    )
    if not os.path.exists(data_js_path):
        update_live_data(data_js_path, [], [], [], {}, parcels)

    class MapViewerRequestHandler(SimpleHTTPRequestHandler):
        def __init__(self, *args, **kwargs):
            super().__init__(*args, directory=work_dir, **kwargs)

        def do_POST(self):
            if self.path == "/api/delete_parcel":
                content_len = int(self.headers.get('Content-Length', 0))
                post_body = self.rfile.read(content_len)
                try:
                    req_data = json.loads(post_body.decode('utf-8'))
                    mathua = req_data.get("mathua")
                    sothua = req_data.get("sothua", "")
                    soto = req_data.get("soto", "")
                    
                    res = delete_parcel_from_excel(excel_path, mathua)
                    
                    print(f"[*] 🗑️ ĐÃ XOÁ THỬA: Thửa {sothua}/Tờ {soto} (Mã {mathua}) -> Đã xóa khỏi 4 Sheet trong '{os.path.basename(excel_path)}'")
                    
                    self.send_response(200)
                    self.send_header('Content-Type', 'application/json; charset=utf-8')
                    self.end_headers()
                    self.wfile.write(json.dumps(res, ensure_ascii=False).encode('utf-8'))
                except Exception as e:
                    self.send_response(500)
                    self.send_header('Content-Type', 'application/json; charset=utf-8')
                    self.end_headers()
                    self.wfile.write(json.dumps({"success": False, "error": str(e)}).encode('utf-8'))
            else:
                self.send_response(404)
                self.end_headers()

        def log_message(self, format, *args):
            # Ẩn các log truy cập file tĩnh để terminal luôn gọn gàng
            if "POST /api/delete_parcel" in format % args:
                pass
            return

    # Tìm cổng khả dụng
    current_port = port
    server = None
    for p_offset in range(10):
        try:
            server = ThreadingHTTPServer(("127.0.0.1", current_port), MapViewerRequestHandler)
            break
        except OSError:
            current_port += 1

    map_url = f"http://127.0.0.1:{current_port}/ban_do_quy_hoach.html"
    
    print("\n" + "="*80)
    print("🏛️   BẢN ĐỒ QUẢN TRỊ & XEM LẠI QUY HOẠCH ĐÔ THỊ TP. HỒ CHÍ MINH")
    print("="*80)
    print(f"  • Thư mục kết quả: {work_dir}")
    print(f"  • File Excel gốc:  {os.path.basename(excel_path)} ({len(parcels)} thửa đất)")
    print(f"  • Bản đồ trực quan: {map_url}")
    print("┌" + "─"*78 + "┐")
    print("│ 🚀 BẢN ĐỒ ĐANG HOẠT ĐỘNG TẠI:                                                 │")
    print(f"│    {map_url:<74} │")
    print("│                                                                              │")
    print("│  ✨ HƯỚNG DẪN QUẢN TRỊ BẢN ĐỒ:                                                │")
    print("│  1. Nhấp chuột vào bất kỳ thửa đất nào để xem chi tiết 4 Sheet quy hoạch.    │")
    print("│  2. Bấm 'Xoá thửa này khỏi bản đồ': Hệ thống TỰ ĐỘNG XOÁ dòng trong file     │")
    print("│     Excel ngay lập tức và đánh số lại STT từ 1 đến N!                        │")
    print("│  3. Nhấn Ctrl + C tại đây bất cứ lúc nào để dừng máy chủ.                   │")
    print("└" + "─"*78 + "┘\n")

    webbrowser.open(map_url)
    try:
        server.serve_forever()
    except (KeyboardInterrupt, SystemExit):
        print("\n[*] Đã đóng máy chủ xem bản đồ. Chúc bạn một ngày làm việc hiệu quả!")
        server.server_close()

if __name__ == "__main__":
    target = sys.argv[1] if len(sys.argv) > 1 else "."
    start_viewer_server(target)
