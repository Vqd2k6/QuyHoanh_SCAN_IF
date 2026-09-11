import xml.etree.ElementTree as ET
from shapely.geometry import LineString, Polygon, Point
from shapely.ops import transform
import pyproj

def parse_kml_coordinates(kml_path):
    tree = ET.parse(kml_path)
    root = tree.getroot()
    
    # Tìm thẻ coordinates
    coords_text = ""
    for elem in root.iter():
        if 'coordinates' in elem.tag:
            coords_text = elem.text.strip()
            # Lấy geometry đầu tiên tìm được
            break
            
    if not coords_text:
        raise ValueError("Không tìm thấy toạ độ trong file KML")
        
    coords = []
    for pair in coords_text.split():
        if not pair.strip():
            continue
        parts = pair.split(',')
        lon = float(parts[0])
        lat = float(parts[1])
        coords.append((lon, lat))
        
    if len(coords) == 1:
        return Point(coords[0])
    elif coords[0] == coords[-1] and len(coords) >= 4:
        return Polygon(coords)
    else:
        return LineString(coords)

def get_buffered_grid(kml_path, buffer_meters=100, grid_spacing_meters=20):
    """
    Đọc KML, tạo vùng đệm (buffer) và trả về danh sách toạ độ (lon, lat) nằm trong vùng đệm.
    """
    geometry = parse_kml_coordinates(kml_path)
    
    # Chuyển đổi từ WGS84 (EPSG:4326) sang hệ toạ độ phẳng UTM Zone 48N (EPSG:32648) dùng cho TP.HCM
    project_to_utm = pyproj.Transformer.from_crs("EPSG:4326", "EPSG:32648", always_xy=True).transform
    project_to_wgs84 = pyproj.Transformer.from_crs("EPSG:32648", "EPSG:4326", always_xy=True).transform
    
    geom_utm = transform(project_to_utm, geometry)
    
    # Tạo vùng đệm
    buffered_utm = geom_utm.buffer(buffer_meters)
    
    # Tạo lưới điểm (grid) bên trong vùng đệm
    minx, miny, maxx, maxy = buffered_utm.bounds
    points = []
    
    x = minx
    while x <= maxx:
        y = miny
        while y <= maxy:
            p = Point(x, y)
            if buffered_utm.contains(p):
                points.append(p)
            y += grid_spacing_meters
        x += grid_spacing_meters
        
    # Chuyển các điểm về lại WGS84
    wgs84_points = [transform(project_to_wgs84, p) for p in points]
    
    print(f"[Info] Đã tạo lưới gồm {len(wgs84_points)} điểm trong vùng đệm {buffer_meters}m.")
    return [(p.x, p.y) for p in wgs84_points]

if __name__ == "__main__":
    # Test script
    import sys
    if len(sys.argv) > 1:
        pts = get_buffered_grid(sys.argv[1], buffer_meters=100, grid_spacing_meters=20)
        print(f"Sample points: {pts[:5]}")
