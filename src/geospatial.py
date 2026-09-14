import os
import re
import urllib.request
import xml.etree.ElementTree as ET
from shapely.geometry import LineString, Polygon, MultiPolygon, Point
from shapely.ops import transform, polygonize, unary_union, linemerge
from shapely.prepared import prep
import pyproj

def resolve_kml_source(source_path_or_url):
    """
    Nhận diện đầu vào là đường dẫn file KML hoặc đường link Google My Maps.
    Nếu là link Google My Maps: tự động tải KML và trích xuất toạ độ trọng tâm (nếu có).
    Trả về: (local_kml_path, (focus_lat, focus_lon) hoặc None)
    """
    source_str = str(source_path_or_url).strip()
    
    if source_str.startswith("http://") or source_str.startswith("https://"):
        # Trích xuất mid
        mid_match = re.search(r'mid=([a-zA-Z0-9_\-]+)', source_str)
        if not mid_match:
            raise ValueError(f"Không tìm thấy tham số mid trong link Google My Maps: {source_str}")
        mid = mid_match.group(1)
        
        # Trích xuất toạ độ hiển thị (ll=lat,lon hoặc ll=lat%2Clon)
        focus_coord = None
        ll_match = re.search(r'll=([0-9\.]+)%2C([0-9\.]+)|ll=([0-9\.]+),([0-9\.]+)', source_str)
        if ll_match:
            groups = [g for g in ll_match.groups() if g]
            if len(groups) >= 2:
                focus_coord = (float(groups[0]), float(groups[1]))
                
        input_dir = os.path.join(os.path.dirname(os.path.dirname(os.path.abspath(__file__))), "data", "input")
        os.makedirs(input_dir, exist_ok=True)
        local_kml = os.path.join(input_dir, f"gmap_{mid}.kml")
        kml_url = f"https://www.google.com/maps/d/kml?mid={mid}&forcekml=1"
        
        print(f"[*] Đang tải dữ liệu từ Google My Maps (MID: {mid})...")
        req = urllib.request.Request(kml_url, headers={'User-Agent': 'Mozilla/5.0 (Macintosh; Intel Mac OS X 10_15_7)'})
        with urllib.request.urlopen(req, timeout=20) as resp:
            content = resp.read()
            with open(local_kml, 'wb') as f:
                f.write(content)
        print(f"[*] Đã tải KML thành công ({len(content) / 1024:.1f} KB) -> '{local_kml}'")
        if focus_coord:
            print(f"[*] Phát hiện toạ độ trọng tâm chỉ định: Lat {focus_coord[0]:.6f}, Lon {focus_coord[1]:.6f}")
            
        return local_kml, focus_coord
        
    if not os.path.exists(source_str):
        # Thử tìm trong data/input/
        input_dir = os.path.join(os.path.dirname(os.path.dirname(os.path.abspath(__file__))), "data", "input")
        cand = os.path.join(input_dir, source_str)
        if os.path.exists(cand):
            return cand, None

    return source_str, None

def parse_kml_geometries(kml_path, focus_coord=None, max_distance_meters=1500):
    """
    Phân tích toàn bộ hình học trong file KML:
    - Hỗ trợ cả LineString, MultiLineString, Polygon
    - Tự động ghép nối các đường bao khép kín (polygonize) tạo thành ranh giới khảo sát
    - Nếu có focus_coord (từ link Google Maps hoặc chỉ định): tập trung vào vùng ranh giới lân cận
    """
    tree = ET.parse(kml_path)
    root = tree.getroot()
    
    lines = []
    polygons = []
    
    for elem in root.iter():
        tag = elem.tag.split('}')[-1]
        
        # Đọc Polygon có sẵn
        if tag == 'Polygon':
            outer = elem.find('{*}outerBoundaryIs')
            if outer is not None:
                coords_el = outer.find('.//{*}coordinates')
            else:
                coords_el = elem.find('.//{*}coordinates')
            if coords_el is not None and coords_el.text:
                raw_pts = [p.strip().split(',') for p in coords_el.text.strip().split() if p.strip()]
                pts = [(float(p[0]), float(p[1])) for p in raw_pts if len(p) >= 2]
                if len(pts) >= 4:
                    polygons.append(Polygon(pts))
                    
        # Đọc LineString
        elif tag == 'LineString':
            coords_el = elem.find('{*}coordinates')
            if coords_el is not None and coords_el.text:
                raw_pts = [p.strip().split(',') for p in coords_el.text.strip().split() if p.strip()]
                pts = [(float(p[0]), float(p[1])) for p in raw_pts if len(p) >= 2]
                if len(pts) >= 2:
                    lines.append(LineString(pts))
                    
    print(f"[*] Tìm thấy trong KML: {len(polygons)} Polygon, {len(lines)} LineString.")
    
    # Nếu có các đoạn line, kiểm tra tự động khép góc các vòng ranh giới mở và bẻ nút giao
    if lines:
        try:
            # 1. Phát hiện các đường bao ranh giới hở (khoảng cách 2 đầu mút < 250m) và tự động nối đoạn khép kín
            closing_segments = []
            for l in lines:
                pts = list(l.coords)
                if len(pts) >= 4 and pts[0] != pts[-1]:
                    d = Point(pts[0]).distance(Point(pts[-1])) * 111000
                    if d < 250:
                        try:
                            test_p = Polygon(pts + [pts[0]])
                            if test_p.is_valid:
                                closing_segments.append(LineString([pts[-1], pts[0]]))
                        except Exception:
                            pass

            all_lines_to_node = lines + closing_segments
            if closing_segments:
                print(f"[*] Đã tự động khép kín {len(closing_segments)} đoạn ranh giới mở rộng (Boundary Envelopes).")

            # 2. Bẻ nút (node) tất cả các đoạn line giao nhau để polygonize nhận diện đầy đủ các vòng khép kín
            noded_lines = unary_union(all_lines_to_node)
            poly_from_lines = list(polygonize(noded_lines))
            if poly_from_lines:
                print(f"[*] Đã tự động tạo {len(poly_from_lines)} vùng ranh giới khép kín (Polygon) từ các đoạn line.")
                polygons.extend(poly_from_lines)
        except Exception as e:
            # Fallback nếu unary_union lỗi
            try:
                poly_from_lines = list(polygonize(lines))
                if poly_from_lines:
                    polygons.extend(poly_from_lines)
            except Exception:
                pass

    # Nếu tìm thấy các vùng ranh giới khép kín (Polygon)
    if polygons:
        # CHỈ lọc theo focus_coord nếu người dùng chủ động chỉ định qua tham số --focus
        if focus_coord:
            focus_pt = Point(focus_coord[1], focus_coord[0]) # (lon, lat)
            dist_deg = max_distance_meters / 111000.0
            matched_polys = [p for p in polygons if p.distance(focus_pt) <= dist_deg]
            if matched_polys:
                print(f"[*] Đã lọc theo tham số --focus: Lấy {len(matched_polys)}/{len(polygons)} vùng ranh giới lân cận.")
                return unary_union(matched_polys), lines, True
                
        # Mặc định lấy ĐẦY ĐỦ 100% tất cả các vùng ranh giới (toàn bộ diện tích boundary)
        print(f"[*] Chế độ quét toàn vẹn: Đang bao phủ 100% tổng diện tích ranh giới ({len(polygons)} vùng đa giác).")
        return unary_union(polygons), lines, True
        
    # Nếu chỉ là các tuyến đường (LineString đơn thuần)
    if lines:
        if len(lines) == 1:
            return lines[0], lines, False
        else:
            return unary_union(lines), lines, False
            
    raise ValueError("Không tìm thấy dữ liệu toạ độ hoặc đường bao nào trong file KML!")

def get_spatial_data(kml_source, buffer_meters=0, grid_spacing_meters=30, focus_coord=None):
    """
    Xử lý không gian địa lý:
    - kml_source: đường dẫn file KML hoặc link Google My Maps
    - buffer_meters: bán kính mở rộng (mặc định: 0 đối với vùng ranh đa giác khép kín)
    - grid_spacing_meters: khoảng cách bước nhảy lưới toạ độ quét
    - focus_coord: chỉ lọc khi người dùng chủ động truyền tham số --focus
    """
    local_kml, _ = resolve_kml_source(kml_source)
    
    # Chỉ áp dụng focus_coord nếu người dùng chủ động truyền qua CLI
    geom, all_lines, is_boundary_polygon = parse_kml_geometries(local_kml, focus_coord=focus_coord)
    
    # Hệ toạ độ UTM Zone 48N (EPSG:32648) cho TP.HCM
    project_to_utm = pyproj.Transformer.from_crs("EPSG:4326", "EPSG:32648", always_xy=True).transform
    project_to_wgs84 = pyproj.Transformer.from_crs("EPSG:32648", "EPSG:4326", always_xy=True).transform
    
    geom_utm = transform(project_to_utm, geom)
    
    # Xử lý vùng bao quét
    if is_boundary_polygon:
        print(f"[*] Chế độ quét: 100% Vùng ranh giới đa giác khép kín (Boundary Polygons). Diện tích: {geom_utm.area / 10000:.1f} ha.")
        survey_zone_utm = geom_utm if buffer_meters <= 0 else geom_utm.buffer(buffer_meters)
    else:
        actual_buffer = buffer_meters if buffer_meters > 0 else 100
        print(f"[*] Chế độ quét: Vùng đệm {actual_buffer}m dọc theo tuyến đường.")
        survey_zone_utm = geom_utm.buffer(actual_buffer)
        
    survey_zone_wgs84 = transform(project_to_wgs84, survey_zone_utm)
    
    # Tạo lưới điểm bên trong toàn bộ vùng quét (Tối ưu hoá Prepared Geometry siêu tốc)
    minx, miny, maxx, maxy = survey_zone_utm.bounds
    prep_zone = prep(survey_zone_utm)
    
    xs = []
    ys = []
    
    x = minx
    while x <= maxx:
        y = miny
        while y <= maxy:
            p = Point(x, y)
            if prep_zone.contains(p):
                xs.append(x)
                ys.append(y)
            y += grid_spacing_meters
        x += grid_spacing_meters
        
    # Vectorized coordinate transform sang WGS84
    if xs:
        transformer = pyproj.Transformer.from_crs("EPSG:32648", "EPSG:4326", always_xy=True)
        lons, lats = transformer.transform(xs, ys)
        grid_points = list(zip(lons, lats))
    else:
        grid_points = []
    
    # Trích xuất toạ độ ranh giới vùng quét cho Leaflet (hỗ trợ cả MultiPolygon tách biệt từng vòng)
    buffer_coords = []
    if hasattr(survey_zone_wgs84, 'exterior'):
        buffer_coords = [[[lat, lon] for lon, lat in survey_zone_wgs84.exterior.coords]]
    elif hasattr(survey_zone_wgs84, 'geoms'):
        for g in survey_zone_wgs84.geoms:
            if hasattr(g, 'exterior'):
                buffer_coords.append([[lat, lon] for lon, lat in g.exterior.coords])
                
    # Trích xuất toạ độ tất cả các đường line
    route_coords = []
    if all_lines:
        for l in all_lines:
            route_coords.append([[lat, lon] for lon, lat in l.coords])
            
    print(f"[*] Đã sinh lưới gồm {len(grid_points)} điểm quét (bước nhảy {grid_spacing_meters}m) bao phủ toàn bộ 100% diện tích ranh giới.")

    return {
        "route_coords": route_coords,
        "buffer_coords": buffer_coords,
        "grid_points": grid_points,
        "is_polygon_boundary": is_boundary_polygon
    }

if __name__ == "__main__":
    import sys
    test_src = sys.argv[1] if len(sys.argv) > 1 else "sample_route.kml"
    data = get_spatial_data(test_src, buffer_meters=0, grid_spacing_meters=30)
    print(f"Sample grid points count: {len(data['grid_points'])}")
