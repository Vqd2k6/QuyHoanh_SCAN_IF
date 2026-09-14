import asyncio
import json
from shapely.geometry import shape, Point, Polygon
from playwright.async_api import async_playwright

def normalize_polygon_coords(ranh_str_or_obj):
    """
    Chuẩn hoá toạ độ ranh giới thửa đất thành danh sách [[lat, lon], ...] để Leaflet vẽ
    """
    if not ranh_str_or_obj:
        return []
    if isinstance(ranh_str_or_obj, str):
        try:
            coords = json.loads(ranh_str_or_obj)
        except Exception:
            return []
    else:
        coords = ranh_str_or_obj
        
    while isinstance(coords, list) and len(coords) > 0 and isinstance(coords[0], list) and len(coords[0]) > 0 and isinstance(coords[0][0], list):
        coords = coords[0]
        
    return [[p[1], p[0]] for p in coords if isinstance(p, (list, tuple)) and len(p) >= 2]

def create_shapely_polygon(ranh_str_or_obj):
    """
    Chuyển đổi ranh giới thửa đất thành đối tượng Shapely Polygon/MultiPolygon
    hỗ trợ kiểm tra toạ độ không gian điểm-trong-đa-giác (Point-in-Polygon).
    """
    if not ranh_str_or_obj:
        return None
    try:
        data = json.loads(ranh_str_or_obj) if isinstance(ranh_str_or_obj, str) else ranh_str_or_obj
        if isinstance(data, dict) and "type" in data:
            poly = shape(data)
        elif isinstance(data, list):
            coords = data
            while isinstance(coords, list) and len(coords) > 0 and isinstance(coords[0], list) and len(coords[0]) > 0 and isinstance(coords[0][0], list):
                coords = coords[0]
            pts = [(float(p[0]), float(p[1])) for p in coords if isinstance(p, (list, tuple)) and len(p) >= 2]
            if len(pts) >= 3:
                poly = Polygon(pts)
            else:
                return None
        else:
            return None
            
        if not poly.is_valid:
            poly = poly.buffer(0)
        return poly
    except Exception:
        return None

def parse_planning_data(raw_data, lat, lon, urban_block_cache=None):
    """
    Trích xuất toàn bộ thông tin chi tiết thửa đất, đồ án 1/2000, quy hoạch chi tiết 1/500,
    các ô chức năng sử dụng đất (kèm chỉ tiêu kiến trúc từ qhpksdd), lộ giới và điều chỉnh cục bộ.
    """
    if not isinstance(raw_data, dict):
        return None
        
    if raw_data.get("blocked") == 1 or "error" in raw_data:
        return {"error": raw_data.get("error", "Bị chặn bởi hệ thống")}
        
    ttc_raw = raw_data.get("ThongTinChung")
    if not ttc_raw:
        return None
        
    try:
        ttc = json.loads(ttc_raw) if isinstance(ttc_raw, str) else ttc_raw
    except Exception:
        return None
        
    if not isinstance(ttc, dict):
        return None

    # 1. Thông tin định danh thửa đất (bắt trọn vẹn cả trường hợp không có số tờ/thửa)
    mathuadat = str(ttc.get("mathuadat", "")).strip()
    sothua = str(ttc.get("sothua", "")).strip()
    soto = str(ttc.get("soto", "")).strip()
    
    if not mathuadat:
        if soto and sothua:
            mathuadat = f"TD_{soto}_{sothua}"
        else:
            mathuadat = f"TD_{lat:.5f}_{lon:.5f}"
            
    dientich_val = None
    try:
        if ttc.get("dientich") is not None:
            dientich_val = float(ttc.get("dientich"))
    except Exception:
        pass
    dientich_str = f"{dientich_val:,.2f}" if dientich_val is not None else str(ttc.get("dientich", ""))
    
    tenquanhuyen = ttc.get("tenquanhuyen", "") or "TP. Hồ Chí Minh"
    tenphuongxa = ttc.get("tenphuongxa", "") or ""
    
    # 2. Đồ án quy hoạch 1/2000
    doan_full_parts = []
    doan_details = []
    dsttdoan = ttc.get("dsttdoan", [])
    if dsttdoan and isinstance(dsttdoan, list):
        for d_item in dsttdoan:
            t_name = d_item.get("tendoan", "")
            soqd = d_item.get("soqd", "")
            ngayduyet = d_item.get("ngayduyet", "")
            coquanpd = d_item.get("coquanpd", "")
            
            p_desc = [t_name] if t_name else []
            meta_items = []
            if soqd: meta_items.append(f"QĐ: {soqd}")
            if ngayduyet: meta_items.append(f"Ngày: {ngayduyet}")
            if coquanpd: meta_items.append(f"Cơ quan: {coquanpd}")
            if meta_items:
                p_desc.append(f"({', '.join(meta_items)})")
            if p_desc:
                doan_full_parts.append(" - ".join(p_desc))
            doan_details.append({
                "tendoan": t_name,
                "soqd": soqd,
                "ngayduyet": ngayduyet,
                "coquanpd": coquanpd
            })
    elif ttc.get("dsdoan"):
        doan_full_parts = [str(x) for x in ttc.get("dsdoan")]
        
    tendoan_summary = "\n".join(doan_full_parts) if doan_full_parts else "Chưa có thông tin đồ án"
    
    # 3. Quy hoạch chi tiết 1/500 (QHChiTiet)
    qhct_raw = raw_data.get("QHChiTiet", [])
    qhct_list = json.loads(qhct_raw) if isinstance(qhct_raw, str) else qhct_raw
    qhct_details = []
    qhct_lines = []
    if isinstance(qhct_list, list):
        for item in qhct_list:
            props = item.get("properties", {}) if isinstance(item, dict) else {}
            tenduan = props.get("tenduan")
            soqd = props.get("soqd")
            ngayduyet = props.get("ngayduyet")
            coquanpd = props.get("coquanpd")
            dt_ct = props.get("dientich")
            tl_ct = props.get("tldientich")
            
            p_desc = [tenduan] if tenduan else []
            meta_sub = []
            if soqd: meta_sub.append(f"QĐ: {soqd}")
            if ngayduyet: meta_sub.append(f"Ngày: {ngayduyet}")
            if coquanpd: meta_sub.append(f"Cơ quan: {coquanpd}")
            if dt_ct:
                try:
                    meta_sub.append(f"DT: {float(dt_ct):,.1f} m²")
                except Exception:
                    pass
            if tl_ct:
                try:
                    meta_sub.append(f"Chiếm: {float(tl_ct):.1f}%")
                except Exception:
                    pass
            if meta_sub: p_desc.append(f"({', '.join(meta_sub)})")
            
            if p_desc:
                qhct_lines.append(" - ".join(p_desc))
            qhct_details.append({
                "tenduan": tenduan or "",
                "soqd": soqd or "",
                "ngayduyet": ngayduyet or "",
                "coquanpd": coquanpd or "",
                "dientich": float(dt_ct) if dt_ct is not None else None,
                "tldientich": float(tl_ct) if tl_ct is not None else None
            })
    qhct_summary = "\n".join(qhct_lines) if qhct_lines else "Không thuộc dự án QH 1/500"

    # 4. Ranh giới thửa đất (đa giác)
    ranh_str = ttc.get("ranh")
    ranh_coords = normalize_polygon_coords(ranh_str)
    polygon_geom = create_shapely_polygon(ranh_str)
    
    # 5. Quy hoạch phân khu (QHPK - các ô chức năng) + Chỉ tiêu kiến trúc (qhpksdd/{gid})
    qhpk_raw = raw_data.get("QHPK", [])
    qhpk_list = json.loads(qhpk_raw) if isinstance(qhpk_raw, str) else qhpk_raw
    qhpk_details = []
    chucnang_lines = []
    chitieu_lines = []
    
    if urban_block_cache is None:
        urban_block_cache = {}
        
    if isinstance(qhpk_list, list):
        for item in qhpk_list:
            props = item.get("properties", {}) if isinstance(item, dict) else {}
            gid = props.get("gid")
            maopho = props.get("maopho")
            chucnang = props.get("chucnang") or "Đất giao thông"
            dt = props.get("dientich")
            tl = props.get("tldientich")
            rgb = props.get("rgbcolor", "130,130,130")
            
            # Lấy thông tin chỉ tiêu kiến trúc từ cache qhpksdd
            block_info = urban_block_cache.get(gid) or {}
            chucnangct = block_info.get("chucnangct") or ""
            tangcao = block_info.get("tangcao")
            chieucao = block_info.get("chieucao")
            matdo = block_info.get("matdo")
            hesosdd = block_info.get("hesosdd")
            danso = block_info.get("danso")
            dientich_opho = block_info.get("dientich")
            
            dt_formatted = f"{float(dt):,.2f}" if dt is not None else ""
            tl_formatted = f"{float(tl):.1f}%" if tl is not None else ""
            
            # Xây dựng dòng tóm tắt chức năng
            line_str = "• "
            if maopho:
                line_str += f"[{maopho}] "
            line_str += f"{chucnang}"
            if dt_formatted:
                line_str += f": {dt_formatted} m²"
            if tl_formatted:
                line_str += f" ({tl_formatted})"
            chucnang_lines.append(line_str)
            
            # Xây dựng dòng tóm tắt chỉ tiêu kiến trúc
            ct_parts = []
            if tangcao: ct_parts.append(f"Tầng cao: {tangcao}")
            if chieucao: ct_parts.append(f"Cao: {chieucao}m")
            if matdo: ct_parts.append(f"Mật độ: {matdo}%")
            if hesosdd: ct_parts.append(f"HSSDĐ: {hesosdd}")
            if danso and danso != 0: ct_parts.append(f"Dân số: {danso}")
            
            if ct_parts:
                tag_label = f"[{maopho}]" if maopho else f"[{chucnang}]"
                chitieu_lines.append(f"• {tag_label}: " + " | ".join(ct_parts))
            
            qhpk_details.append({
                "gid": gid,
                "maopho": maopho or "",
                "chucnang": chucnang,
                "chucnangct": chucnangct,
                "dientich": float(dt) if dt is not None else 0.0,
                "dientich_formatted": dt_formatted,
                "tldientich": float(tl) if tl is not None else 0.0,
                "tldientich_formatted": tl_formatted,
                "rgbcolor": rgb,
                "tangcao": str(tangcao) if tangcao is not None else "",
                "chieucao": str(chieucao) if chieucao is not None else "",
                "matdo": str(matdo) if matdo is not None else "",
                "hesosdd": str(hesosdd) if hesosdd is not None else "",
                "danso": str(danso) if danso is not None else "",
                "dientich_opho": float(dientich_opho) if dientich_opho is not None and str(dientich_opho).replace('.', '', 1).isdigit() else None
            })
            
    chucnang_summary = "\n".join(chucnang_lines) if chucnang_lines else "Chưa xác định"
    chitieu_summary = "\n".join(chitieu_lines) if chitieu_lines else "Chưa có dữ liệu chỉ tiêu"
    
    # 6. Lộ giới (LoGioi)
    logioi_raw = raw_data.get("LoGioi", [])
    logioi_list = json.loads(logioi_raw) if isinstance(logioi_raw, str) else logioi_raw
    logioi_details = []
    logioi_lines = []
    
    if isinstance(logioi_list, list):
        for item in logioi_list:
            props = item.get("properties", {}) if isinstance(item, dict) else {}
            tenduong = props.get("tenduong")
            lg = props.get("logioi") or props.get("chieusau")
            huong = props.get("huongtiepgiap")
            chieungang = props.get("chieungang")
            chieusau = props.get("chieusau")
            dientichxd = props.get("DienTichXD")
            
            if tenduong:
                line_str = f"• {tenduong}"
                meta_sub = []
                if lg:
                    try:
                        meta_sub.append(f"Lộ giới: {float(lg):.1f}m")
                    except Exception:
                        meta_sub.append(f"Lộ giới: {lg}m")
                if huong:
                    meta_sub.append(f"Hướng: {huong}")
                if chieungang:
                    try:
                        meta_sub.append(f"Mặt tiền: {float(chieungang):.1f}m")
                    except Exception:
                        pass
                if meta_sub:
                    line_str += f" ({', '.join(meta_sub)})"
                logioi_lines.append(line_str)
                
            logioi_details.append({
                "tenduong": tenduong or "",
                "logioi": float(lg) if lg is not None and str(lg).replace('.', '', 1).isdigit() else str(lg or ""),
                "huongtiepgiap": huong or "",
                "chieungang": float(chieungang) if chieungang is not None else None,
                "chieusau": float(chieusau) if chieusau is not None else None,
                "dientichxd": float(dientichxd) if dientichxd is not None else None
            })
    logioi_summary = "\n".join(logioi_lines) if logioi_lines else "Không có thông tin lộ giới"
    
    # 7. Điều chỉnh cục bộ (DCCB)
    dccb_raw = raw_data.get("DCCB", [])
    dccb_list = json.loads(dccb_raw) if isinstance(dccb_raw, str) else dccb_raw
    dccb_details = []
    if isinstance(dccb_list, list):
        for item in dccb_list:
            props = item.get("properties", {}) if isinstance(item, dict) else (item if isinstance(item, dict) else {})
            dccb_details.append({
                "tendoan": props.get("TenDoAn", "") or "",
                "tendccb": props.get("TenDCCB", "") or "",
                "soqd": props.get("SoQD", "") or "",
                "ngayduyet": props.get("NgayDuyet", "") or "",
                "coquanpd": props.get("CoQuanPD", "") or ""
            })

    # 8. Chỉ tiêu xây dựng nhà ở riêng lẻ (CTXD)
    ctxd_raw = raw_data.get("CTXD", [])
    ctxd_list = json.loads(ctxd_raw) if isinstance(ctxd_raw, str) else ctxd_raw
    ctxd_details = []
    if isinstance(ctxd_list, list):
        for item in ctxd_list:
            if isinstance(item, dict):
                ctxd_details.append(item)

    return {
        "mathuadat": mathuadat,
        "soto": soto or "-",
        "sothua": sothua or "-",
        "dientich": dientich_val,
        "dientich_formatted": dientich_str,
        "tenquanhuyen": tenquanhuyen,
        "tenphuongxa": tenphuongxa,
        "tendoan": tendoan_summary,
        "doan_details": doan_details,
        "qhct_summary": qhct_summary,
        "qhct_details": qhct_details,
        "chucnang_summary": chucnang_summary,
        "chitieu_summary": chitieu_summary,
        "qhpk_details": qhpk_details,
        "logioi_summary": logioi_summary,
        "logioi_details": logioi_details,
        "dccb_details": dccb_details,
        "ctxd_details": ctxd_details,
        "ranh_coords": ranh_coords,
        "polygon_geom": polygon_geom,
        "latitude": lat,
        "longitude": lon
    }

async def fetch_planning_data(points, on_point_scraped=None, headless=True, concurrency=3, initial_parcels=None, initial_scanned_points=None):
    """
    Thu thập dữ liệu quy hoạch cho danh sách các điểm toạ độ thông qua API chính thức của SQHKT.
    Tự động truy vấn chỉ tiêu kiến trúc (tầng cao, mật độ, HSSDĐ) cho từng ô quy hoạch.
    Hỗ trợ chạy song song đa luồng (Worker Pool) với concurrency luồng.
    Hỗ trợ nạp sẵn initial_parcels và initial_scanned_points để tiếp tục đợt quét dở dang (Resume).
    Gọi on_point_scraped(lon, lat, parsed_info, current_idx, total_count) sau mỗi điểm.
    """
    total_points = len(points)
    urban_block_cache = {}  # gid -> details dict
    # Danh sách các bộ (polygon_geom, (minx, miny, maxx, maxy), parsed_info) để lọc không gian siêu tốc
    known_parcels = []
    
    # Nạp các thửa đất đã có từ đợt quét trước
    if initial_parcels:
        for p_key, p_info in initial_parcels.items():
            if isinstance(p_info, dict):
                p_geom = p_info.get("polygon_geom")
                if p_geom is None and p_info.get("ranh_coords"):
                    p_geom = create_shapely_polygon(p_info.get("ranh_coords"))
                if p_geom is not None:
                    known_parcels.append((p_geom, p_geom.bounds, p_info))
                    
    scanned_lookup = set(initial_scanned_points.keys()) if initial_scanned_points else set()
    lock = asyncio.Lock()
    
    concurrency = max(1, min(int(concurrency), 30))  # Hỗ trợ tối đa đến 30 luồng song song (bao gồm 15 luồng)
    
    async with async_playwright() as p:
        browser = await p.chromium.launch(
            headless=headless,
            args=['--disable-blink-features=AutomationControlled']
        )
        context = await browser.new_context(
            user_agent="Mozilla/5.0 (Macintosh; Intel Mac OS X 10_15_7) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/120.0.0.0 Safari/537.36"
        )
        await context.add_init_script("Object.defineProperty(navigator, 'webdriver', {get: () => undefined})")
        
        pages = [await context.new_page() for _ in range(concurrency)]
        
        async def init_page(page, p_idx):
            try:
                await page.goto("https://thongtinquyhoach.hochiminhcity.gov.vn/", wait_until="domcontentloaded", timeout=30000)
                await page.wait_for_timeout(1500)
            except Exception as e:
                print(f"[!] Cảnh báo kết nối luồng #{p_idx + 1}: {e}")

        print(f"[*] Khởi động {concurrency} luồng song song và kết nối tới Cổng thông tin quy hoạch TP.HCM...")
        await asyncio.gather(*[init_page(pages[i], i) for i in range(concurrency)])
        
        queue = asyncio.Queue()
        for idx, pt in enumerate(points):
            queue.put_nowait((idx, pt[0], pt[1]))
            
        async def worker(worker_id, page):
            while True:
                try:
                    idx, lon, lat = queue.get_nowait()
                except asyncio.QueueEmpty:
                    break
                    
                current_idx = idx + 1
                point_key = f"{lon:.6f},{lat:.6f}"
                pt_query = Point(lon, lat)
                
                # 0. NẾU ĐIỂM ĐÃ TỪNG QUÉT Ở ĐỢT TRƯỚC VÀ KHÔNG CÓ THỬA (CÔNG CỘNG/GIAO THÔNG)
                if point_key in scanned_lookup and initial_scanned_points.get(point_key) is False:
                    if on_point_scraped:
                        try:
                            if asyncio.iscoroutinefunction(on_point_scraped):
                                await on_point_scraped(lon, lat, None, current_idx, total_points, is_skipped=True, worker_id=worker_id + 1)
                            else:
                                on_point_scraped(lon, lat, None, current_idx, total_points, is_skipped=True, worker_id=worker_id + 1)
                        except Exception:
                            pass
                    queue.task_done()
                    continue

                # 1. KIỂM TRA AUTO-NEXT TRÊN BỘ NHỚ CHUNG (LỌC BẰNG BOUNDING BOX TRƯỚC - KHÔNG CHẶN KHÓA LUỒNG)
                cached_parcel = None
                for p_geom, (minx, miny, maxx, maxy), p_info in known_parcels:
                    if minx <= lon <= maxx and miny <= lat <= maxy:
                        try:
                            if p_geom.covers(pt_query):
                                cached_parcel = p_info
                                break
                        except Exception:
                            pass

                if cached_parcel is not None:
                    if on_point_scraped:
                        try:
                            if asyncio.iscoroutinefunction(on_point_scraped):
                                await on_point_scraped(lon, lat, cached_parcel, current_idx, total_points, is_skipped=True, worker_id=worker_id + 1)
                            else:
                                on_point_scraped(lon, lat, cached_parcel, current_idx, total_points, is_skipped=True, worker_id=worker_id + 1)
                        except Exception as cb_err:
                            print(f"[!] Lỗi callback xử lý kết quả: {cb_err}")
                    queue.task_done()
                    continue

                # 2. GỬI REQUEST TRUY VẤN MẠNG
                raw_result = None
                try:
                    raw_result = await page.evaluate(f"""async () => {{
                        try {{
                            let resp = await window.axios.post('https://sqhkt-qlqh.tphcm.gov.vn/computing/930/api/v3.1/a-z/all', 'Lat={lat}&Lon={lon}');
                            return resp.data;
                        }} catch(e) {{
                            try {{
                                let resp2 = await window.axios.post('https://thongtinquyhoach.hochiminhcity.gov.vn/computing/930/api/v3.1/a-z/all', 'Lat={lat}&Lon={lon}');
                                return resp2.data;
                            }} catch(e2) {{
                                return {{error: e2.toString()}};
                            }}
                        }}
                    }}""")
                except Exception as e:
                    raw_result = {"error": str(e)}

                # 3. TRUY VẤN CHỈ TIÊU KIẾN TRÚC QHPK NẾU CÓ
                if isinstance(raw_result, dict) and "QHPK" in raw_result:
                    qhpk_val = raw_result.get("QHPK")
                    try:
                        qhpk_items = json.loads(qhpk_val) if isinstance(qhpk_val, str) else qhpk_val
                        if isinstance(qhpk_items, list):
                            for item in qhpk_items:
                                props = item.get("properties", {}) if isinstance(item, dict) else {}
                                gid = props.get("gid")
                                if gid and gid not in urban_block_cache:
                                    try:
                                        block_data = await page.evaluate(f"""async () => {{
                                            try {{
                                                let res = await window.axios.get('https://sqhkt-qlqh.tphcm.gov.vn/api/qhpksdd/{gid}');
                                                return res.data;
                                            }} catch(e) {{
                                                return null;
                                            }}
                                        }}""")
                                        if isinstance(block_data, dict):
                                            urban_block_cache[gid] = block_data
                                        elif isinstance(block_data, list) and len(block_data) > 0 and isinstance(block_data[0], dict):
                                            urban_block_cache[gid] = block_data[0]
                                        else:
                                            urban_block_cache[gid] = {}
                                    except Exception:
                                        urban_block_cache[gid] = {}
                    except Exception:
                        pass

                # 4. PHÂN TÍCH VÀ CẬP NHẬT KẾT QUẢ
                parsed_info = parse_planning_data(raw_result, lat, lon, urban_block_cache=urban_block_cache)
                
                async with lock:
                    if parsed_info and parsed_info.get("polygon_geom"):
                        p_poly = parsed_info["polygon_geom"]
                        known_parcels.append((p_poly, p_poly.bounds, parsed_info))
                    
                if on_point_scraped:
                    try:
                        if asyncio.iscoroutinefunction(on_point_scraped):
                            await on_point_scraped(lon, lat, parsed_info, current_idx, total_points, is_skipped=False, worker_id=worker_id + 1)
                        else:
                            on_point_scraped(lon, lat, parsed_info, current_idx, total_points, is_skipped=False, worker_id=worker_id + 1)
                    except Exception as cb_err:
                        print(f"[!] Lỗi callback xử lý kết quả: {cb_err}")

                await page.wait_for_timeout(50)
                queue.task_done()

        workers = [worker(i, pages[i]) for i in range(concurrency)]
        try:
            await asyncio.gather(*workers)
        except (asyncio.CancelledError, KeyboardInterrupt):
            pass
        except Exception as e:
            if "Target page, context or browser has been closed" not in str(e):
                print(f"[!] Lỗi worker: {e}")
        finally:
            try:
                await browser.close()
            except Exception:
                pass
        
    return True

if __name__ == "__main__":
    # Test toạ độ Quận 12 (Thửa 92)
    pts = [(106.623341, 10.819917)]
    
    def test_cb(lon, lat, info, idx, total):
        print(f"[{idx}/{total}] Kết quả trích xuất:")
        print(json.dumps(info, ensure_ascii=False, indent=2))
        
    asyncio.run(fetch_planning_data(pts, on_point_scraped=test_cb, headless=True))
