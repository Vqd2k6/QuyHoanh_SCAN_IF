"""
Bus Route Data Extractor & GIS Converter for Ho Chi Minh City Bus Network
Extracts 30 bus routes, converts to GeoJSON, KML (individual & merged), Excel/CSV summary,
and prepares datasets for Interactive Web GIS Map.
"""

import os
import json
import html
import concurrent.futures
import requests
import pandas as pd

TARGET_ROUTES_INFO = [
    {"stt": 1, "code": "04", "display_code": "B 04", "open": "5:00", "close": "20:15", "start_point": "Bến Thành", "end_point": "Bến xe An Sương", "length_km": 16},
    {"stt": 2, "code": "13", "display_code": "B 13", "open": "4:00", "close": "19:45", "start_point": "Bến xe buýt Sài Gòn", "end_point": "Bến xe Củ Chi", "length_km": 36},
    {"stt": 3, "code": "23", "display_code": "B 23", "open": "4:00", "close": "20:30", "start_point": "Bến xe buýt Chợ Lớn", "end_point": "Cầu Lớn", "length_km": 23},
    {"stt": 4, "code": "24", "display_code": "B 24", "open": "4:00", "close": "20:30", "start_point": "Bến xe Miền Đông", "end_point": "Hóc Môn", "length_km": 27},
    {"stt": 5, "code": "27", "display_code": "B 27", "open": "5:00", "close": "20:00", "start_point": "Bến xe buýt Sài Gòn", "end_point": "Bến xe An Sương", "length_km": 15},
    {"stt": 6, "code": "30", "display_code": "B 30", "open": "5:00", "close": "22:00", "start_point": "Chợ Tân Hương", "end_point": "Bến xe buýt Văn Thánh", "length_km": 18},
    {"stt": 7, "code": "33", "display_code": "B 33", "open": "4:30", "close": "22:00", "start_point": "Bến xe An Sương", "end_point": "Đại học Quốc gia", "length_km": 25},
    {"stt": 8, "code": "41", "display_code": "B 41", "open": "5:00", "close": "19:20", "start_point": "Bến xe Miền Tây", "end_point": "Bến xe An Sương", "length_km": 22},
    {"stt": 9, "code": "48", "display_code": "B 48", "open": "4:30", "close": "20:00", "start_point": "Bến xe buýt Tân Phú", "end_point": "Chợ Hiệp Thành", "length_km": 21},
    {"stt": 10, "code": "60-1", "display_code": "B 60-1", "open": "4:45", "close": "18:30", "start_point": "Bến xe Miền Tây", "end_point": "Bến xe Biên Hòa", "length_km": 62},
    {"stt": 11, "code": "61-3", "display_code": "B 61-3", "open": "6:00", "close": "16:30", "start_point": "Bến xe An Sương", "end_point": "Thủ Dầu Một", "length_km": 34},
    {"stt": 12, "code": "62", "display_code": "B 62", "open": "5:00", "close": "19:00", "start_point": "Bến xe buýt Quận 8", "end_point": "Thới An", "length_km": 24},
    {"stt": 13, "code": "62-5", "display_code": "B 62-5", "open": "4:00", "close": "19:30", "start_point": "Bến xe An Sương", "end_point": "Bến xe Hậu Nghĩa", "length_km": 31},
    {"stt": 14, "code": "65", "display_code": "B 65", "open": "4:45", "close": "20:45", "start_point": "Bến Thành", "end_point": "Bến xe An Sương", "length_km": 16},
    {"stt": 15, "code": "70-1", "display_code": "B 70-1", "open": "3:00", "close": "19:00", "start_point": "Bến xe Củ Chi", "end_point": "Bến xe Tây Ninh", "length_km": 66},
    {"stt": 16, "code": "70-2", "display_code": "B 70-2", "open": "2:55", "close": "19:00", "start_point": "Bến xe Củ Chi", "end_point": "Khu du lịch Núi Bà Đen", "length_km": 73},
    {"stt": 17, "code": "71", "display_code": "B 71", "open": "5:00", "close": "20:30", "start_point": "Bến xe An Sương", "end_point": "Phật Cô Đơn", "length_km": 24},
    {"stt": 18, "code": "74", "display_code": "B 74", "open": "3:30", "close": "21:00", "start_point": "Bến xe An Sương", "end_point": "Bến xe Củ Chi", "length_km": 21},
    {"stt": 19, "code": "78", "display_code": "B 78", "open": "5:00", "close": "20:00", "start_point": "Bến xe buýt Thới An", "end_point": "Cầu Lớn", "length_km": 18},
    {"stt": 20, "code": "79", "display_code": "B 79", "open": "5:00", "close": "20:00", "start_point": "Bến xe Củ Chi", "end_point": "Đền Bến Dược", "length_km": 25},
    {"stt": 21, "code": "85", "display_code": "B 85", "open": "4:35", "close": "20:30", "start_point": "Bến xe An Sương", "end_point": "Khu công nghiệp Nhị Xuân", "length_km": 14},
    {"stt": 22, "code": "87", "display_code": "B 87", "open": "5:00", "close": "20:30", "start_point": "Bến xe Củ Chi", "end_point": "An Nhơn Tây", "length_km": 20},
    {"stt": 23, "code": "94", "display_code": "B 94", "open": "4:00", "close": "19:00", "start_point": "Bến xe buýt Chợ Lớn", "end_point": "Bến xe Củ Chi", "length_km": 36},
    {"stt": 24, "code": "100", "display_code": "B 100", "open": "4:45", "close": "21:00", "start_point": "Bến xe Củ Chi", "end_point": "Cầu Tân Thái", "length_km": 16},
    {"stt": 25, "code": "104", "display_code": "B 104", "open": "4:30", "close": "22:00", "start_point": "Bến xe An Sương", "end_point": "Bến xe buýt Văn Thánh", "length_km": 16},
    {"stt": 26, "code": "107", "display_code": "B 107", "open": "4:45", "close": "20:00", "start_point": "Bến xe Củ Chi", "end_point": "Bố Heo", "length_km": 15},
    {"stt": 27, "code": "122", "display_code": "B 122", "open": "4:45", "close": "19:30", "start_point": "Bến xe An Sương", "end_point": "Tân Quy", "length_km": 19},
    {"stt": 28, "code": "126", "display_code": "B 126", "open": "4:45", "close": "20:00", "start_point": "Bến xe Củ Chi", "end_point": "Bình Mỹ", "length_km": 19},
    {"stt": 29, "code": "145", "display_code": "B 145", "open": "4:30", "close": "20:30", "start_point": "Bến xe buýt Chợ Lớn", "end_point": "Chợ Hiệp Thành", "length_km": 23},
    {"stt": 30, "code": "151", "display_code": "B 151", "open": "4:00", "close": "20:00", "start_point": "Bến xe Miền Tây", "end_point": "Bến xe An Sương", "length_km": 17}
]

# Color palette for 30 distinct routes
PALETTE = [
    "#E53935", "#D81B60", "#8E24AA", "#5E35B1", "#3949AB",
    "#1E88E5", "#039BE5", "#00ACC1", "#00897B", "#43A047",
    "#7CB342", "#C0CA33", "#FDD835", "#FFB300", "#FB8C00",
    "#F4511E", "#6D4C41", "#546E7A", "#78909C", "#E91E63",
    "#9C27B0", "#673AB7", "#3F51B5", "#2196F3", "#00BCD4",
    "#009688", "#4CAF50", "#8BC34A", "#FF9800", "#FF5722"
]

BASE_DIR = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
OUTPUT_DIR = os.path.join(BASE_DIR, "data", "output", "bus_routes")
INDIVIDUAL_GEOJSON_DIR = os.path.join(OUTPUT_DIR, "individual", "geojson")
INDIVIDUAL_KML_DIR = os.path.join(OUTPUT_DIR, "individual", "kml")
MERGED_DIR = os.path.join(OUTPUT_DIR, "merged")
SUMMARY_DIR = os.path.join(OUTPUT_DIR, "summary")
VISUALIZER_DIR = os.path.join(OUTPUT_DIR, "visualizer")

os.makedirs(INDIVIDUAL_GEOJSON_DIR, exist_ok=True)
os.makedirs(INDIVIDUAL_KML_DIR, exist_ok=True)
os.makedirs(MERGED_DIR, exist_ok=True)
os.makedirs(SUMMARY_DIR, exist_ok=True)
os.makedirs(VISUALIZER_DIR, exist_ok=True)

HEADERS = {
    "User-Agent": "Mozilla/5.0 (Macintosh; Intel Mac OS X 10_15_7) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/120.0.0.0 Safari/537.36",
    "Accept": "application/json, text/plain, */*"
}

def get_with_retry(url, max_retries=5, delay=1.0):
    for attempt in range(max_retries):
        try:
            resp = requests.get(url, headers=HEADERS, timeout=12)
            if resp.status_code == 200:
                return resp
        except Exception as e:
            if attempt == max_retries - 1:
                raise e
        import time
        time.sleep(delay * (attempt + 1))
    return None

def fetch_all_routes():
    """Fetch all routes list from official EBMS server"""
    url = "http://apicms.ebms.vn/businfo/getallroute"
    print(f"Fetching route list from {url}...")
    resp = get_with_retry(url)
    routes = resp.json()
    return {rt["RouteNo"].strip(): rt for rt in routes}

def fetch_route_details(route_meta, ebms_info, color):
    """Fetch variants, paths, and stops for a single route"""
    route_no = route_meta["code"]
    route_id = ebms_info["RouteId"]
    official_name = ebms_info["RouteName"]
    
    # 1. Get variants (Lượt đi / Lượt về)
    var_url = f"http://apicms.ebms.vn/businfo/getvarsbyroute/{route_id}"
    v_resp = get_with_retry(var_url)
    variants = v_resp.json() if v_resp and v_resp.status_code == 200 else []
    
    variant_details = []
    total_stops = 0
    total_points = 0
    
    for v in variants:
        v_id = v["RouteVarId"]
        v_name = v.get("RouteVarName", "")
        v_short = v.get("RouteVarShortName", "")
        is_outbound = v.get("Outbound", True)
        dist_m = v.get("Distance", 0)
        run_min = v.get("RunningTime", 0)
        
        # Path coordinates
        p_url = f"http://apicms.ebms.vn/businfo/getpathsbyvar/{route_id}/{v_id}"
        p_resp = get_with_retry(p_url)
        coordinates = []
        if p_resp and p_resp.status_code == 200:
            p_data = p_resp.json()
            lats = p_data.get("lat", [])
            lngs = p_data.get("lng", [])
            for lat, lng in zip(lats, lngs):
                if lat and lng:
                    coordinates.append([float(lng), float(lat)])
        
        # Stops
        s_url = f"http://apicms.ebms.vn/businfo/getstopsbyvar/{route_id}/{v_id}"
        s_resp = get_with_retry(s_url)
        stops = []
        if s_resp and s_resp.status_code == 200:
            stops_raw = s_resp.json()
            for s in stops_raw:
                if s.get("Lat") and s.get("Lng"):
                    stops.append({
                        "stop_id": s.get("StopId"),
                        "code": s.get("Code"),
                        "name": s.get("Name"),
                        "stop_type": s.get("StopType"),
                        "zone": s.get("Zone"),
                        "ward": s.get("Ward"),
                        "address": s.get("AddressNo", "") + " " + s.get("Street", ""),
                        "street": s.get("Street"),
                        "lat": float(s.get("Lat")),
                        "lng": float(s.get("Lng")),
                        "routes": s.get("Routes")
                    })
        
        total_stops += len(stops)
        total_points += len(coordinates)
        
        variant_details.append({
            "var_id": v_id,
            "var_name": v_name,
            "var_short_name": v_short,
            "outbound": is_outbound,
            "distance_m": dist_m,
            "distance_km": round(dist_m / 1000.0, 2) if dist_m else 0,
            "running_time_min": run_min,
            "start_stop": v.get("StartStop"),
            "end_stop": v.get("EndStop"),
            "coordinates": coordinates,
            "stops": stops
        })
    
    return {
        "stt": route_meta["stt"],
        "route_id": route_id,
        "route_no": route_no,
        "display_code": route_meta["display_code"],
        "official_name": official_name,
        "start_point": route_meta["start_point"],
        "end_point": route_meta["end_point"],
        "listed_length_km": route_meta["length_km"],
        "operating_open": route_meta["open"],
        "operating_close": route_meta["close"],
        "operating_hours": f"{route_meta['open']} - {route_meta['close']}",
        "color": color,
        "total_stops": total_stops,
        "total_points": total_points,
        "variants": variant_details
    }

def route_to_geojson_feature_collection(route):
    """Convert a single route object to GeoJSON FeatureCollection"""
    features = []
    
    # 1. LineString Features for each variant
    for v in route["variants"]:
        if v["coordinates"]:
            dir_label = "Lượt đi" if v["outbound"] else "Lượt về"
            feature = {
                "type": "Feature",
                "geometry": {
                    "type": "LineString",
                    "coordinates": v["coordinates"]
                },
                "properties": {
                    "feature_type": "route_path",
                    "stt": route["stt"],
                    "route_id": route["route_id"],
                    "route_no": route["route_no"],
                    "display_code": route["display_code"],
                    "route_name": route["official_name"],
                    "direction": dir_label,
                    "outbound": v["outbound"],
                    "start_point": route["start_point"],
                    "end_point": route["end_point"],
                    "start_stop": v["start_stop"],
                    "end_stop": v["end_stop"],
                    "distance_km": v["distance_km"],
                    "listed_length_km": route["listed_length_km"],
                    "running_time_min": v["running_time_min"],
                    "operating_hours": route["operating_hours"],
                    "stroke": route["color"],
                    "stroke_width": 4,
                    "stroke_opacity": 0.85
                }
            }
            features.append(feature)
            
    # 2. Point Features for Stops (unique by StopId to avoid exact duplicates)
    seen_stop_ids = set()
    for v in route["variants"]:
        dir_label = "Lượt đi" if v["outbound"] else "Lượt về"
        for s in v["stops"]:
            stop_id = s["stop_id"]
            if stop_id in seen_stop_ids:
                continue
            seen_stop_ids.add(stop_id)
            
            p_feature = {
                "type": "Feature",
                "geometry": {
                    "type": "Point",
                    "coordinates": [s["lng"], s["lat"]]
                },
                "properties": {
                    "feature_type": "bus_stop",
                    "stop_id": s["stop_id"],
                    "stop_code": s["code"],
                    "stop_name": s["name"],
                    "stop_type": s["stop_type"],
                    "route_no": route["route_no"],
                    "display_code": route["display_code"],
                    "direction": dir_label,
                    "address": s["address"],
                    "street": s["street"],
                    "zone": s["zone"],
                    "ward": s["ward"],
                    "all_routes": s["routes"],
                    "marker_color": route["color"]
                }
            }
            features.append(p_feature)
            
    return {
        "type": "FeatureCollection",
        "properties": {
            "stt": route["stt"],
            "route_no": route["route_no"],
            "display_code": route["display_code"],
            "route_name": route["official_name"],
            "start_point": route["start_point"],
            "end_point": route["end_point"],
            "listed_length_km": route["listed_length_km"],
            "operating_hours": route["operating_hours"],
            "color": route["color"],
            "num_variants": len(route["variants"]),
            "total_stops": route["total_stops"],
            "total_points": route["total_points"]
        },
        "features": features
    }

def hex_to_kml_color(hex_color, alpha_hex="ff"):
    """Convert #RRGGBB hex to KML aabbggrr format"""
    hex_color = hex_color.lstrip("#")
    if len(hex_color) == 6:
        r, g, b = hex_color[0:2], hex_color[2:4], hex_color[4:6]
        return f"{alpha_hex}{b}{g}{r}"
    return "ffff0000"

def route_to_kml(route):
    """Generate KML content for a single route"""
    kml_color = hex_to_kml_color(route["color"], "e6")
    route_name_esc = html.escape(f"{route['display_code']} - {route['official_name']}")
    
    kml_lines = [
        '<?xml version="1.0" encoding="UTF-8"?>',
        '<kml xmlns="http://www.opengis.net/kml/2.2">',
        '  <Document>',
        f'    <name>{route_name_esc}</name>',
        f'    <description>Lộ trình tuyến xe buýt {html.escape(route["display_code"])} ({html.escape(route["operating_hours"])})</description>',
        '    <Style id="routeLineStyle">',
        '      <LineStyle>',
        f'        <color>{kml_color}</color>',
        '        <width>4</width>',
        '      </LineStyle>',
        '    </Style>',
        '    <Style id="busStopStyle">',
        '      <IconStyle>',
        '        <scale>0.8</scale>',
        '        <Icon>',
        '          <href>http://maps.google.com/mapfiles/kml/shapes/bus.png</href>',
        '        </Icon>',
        '      </IconStyle>',
        '    </Style>'
    ]
    
    for v in route["variants"]:
        dir_label = "Lượt đi" if v["outbound"] else "Lượt về"
        folder_name = html.escape(f"{dir_label}: {v['start_stop']} → {v['end_stop']} ({v['distance_km']} km)")
        kml_lines.append('    <Folder>')
        kml_lines.append(f'      <name>{folder_name}</name>')
        
        # Route path LineString
        if v["coordinates"]:
            coord_str = " ".join([f"{c[0]},{c[1]},0" for c in v["coordinates"]])
            desc = html.escape(
                f"<b>Tuyến:</b> {route['display_code']} - {route['official_name']}<br/>"
                f"<b>Lượt:</b> {dir_label}<br/>"
                f"<b>Cự ly:</b> {v['distance_km']} km<br/>"
                f"<b>Thời gian hành trình:</b> {v['running_time_min']} phút<br/>"
                f"<b>Thời gian hoạt động:</b> {route['operating_hours']}"
            )
            kml_lines.extend([
                '      <Placemark>',
                f'        <name>{html.escape(route["display_code"])} - Đường đi {dir_label}</name>',
                f'        <description><![CDATA[{desc}]]></description>',
                '        <styleUrl>#routeLineStyle</styleUrl>',
                '        <LineString>',
                '          <tessellate>1</tessellate>',
                f'          <coordinates>{coord_str}</coordinates>',
                '        </LineString>',
                '      </Placemark>'
            ])
            
        # Stops Folder
        if v["stops"]:
            kml_lines.append('      <Folder>')
            kml_lines.append(f'        <name>Trạm dừng ({len(v["stops"])} trạm)</name>')
            for s in v["stops"]:
                s_name = html.escape(f"[{s['code']}] {s['name']}")
                s_desc = html.escape(
                    f"<b>Trạm:</b> {s['name']}<br/>"
                    f"<b>Mã trạm:</b> {s['code']}<br/>"
                    f"<b>Địa chỉ:</b> {s['address']}<br/>"
                    f"<b>Khu vực:</b> {s['ward']}, {s['zone']}<br/>"
                    f"<b>Các tuyến qua trạm:</b> {s['routes']}"
                )
                kml_lines.extend([
                    '        <Placemark>',
                    f'          <name>{s_name}</name>',
                    f'          <description><![CDATA[{s_desc}]]></description>',
                    '          <styleUrl>#busStopStyle</styleUrl>',
                    '          <Point>',
                    f'            <coordinates>{s["lng"]},{s["lat"]},0</coordinates>',
                    '          </Point>',
                    '        </Placemark>'
                ])
            kml_lines.append('      </Folder>')
            
        kml_lines.append('    </Folder>')
        
    kml_lines.append('  </Document>')
    kml_lines.append('</kml>')
    return "\n".join(kml_lines)

def generate_merged_kml(all_routes):
    """Generate a single unified KML containing all 30 bus routes with structured folders"""
    kml_lines = [
        '<?xml version="1.0" encoding="UTF-8"?>',
        '<kml xmlns="http://www.opengis.net/kml/2.2">',
        '  <Document>',
        '    <name>HỆ THỐNG 30 TUYẾN XE BUÝT TP.HCM (TỔNG HỢP)</name>',
        '    <description>Dữ liệu hợp nhất lộ trình và trạm dừng 30 tuyến xe buýt TP.HCM trích xuất từ EBMS</description>',
        '    <Style id="busStopMergedStyle">',
        '      <IconStyle>',
        '        <scale>0.7</scale>',
        '        <Icon>',
        '          <href>http://maps.google.com/mapfiles/kml/shapes/bus.png</href>',
        '        </Icon>',
        '      </IconStyle>',
        '    </Style>'
    ]
    
    # Styles for each route
    for r in all_routes:
        kml_color = hex_to_kml_color(r["color"], "e6")
        kml_lines.extend([
            f'    <Style id="style_route_{r["stt"]}">',
            '      <LineStyle>',
            f'        <color>{kml_color}</color>',
            '        <width>4</width>',
            '      </LineStyle>',
            '    </Style>'
        ])
        
    for r in all_routes:
        r_title = html.escape(f"TUYẾN {r['display_code']}: {r['start_point']} - {r['end_point']}")
        kml_lines.append('    <Folder>')
        kml_lines.append(f'      <name>{r_title}</name>')
        
        for v in r["variants"]:
            dir_label = "Lượt đi" if v["outbound"] else "Lượt về"
            if v["coordinates"]:
                coord_str = " ".join([f"{c[0]},{c[1]},0" for c in v["coordinates"]])
                desc = html.escape(
                    f"<b>Tuyến:</b> {r['display_code']} - {r['official_name']}<br/>"
                    f"<b>Hành trình:</b> {v['start_stop']} → {v['end_stop']}<br/>"
                    f"<b>Lượt:</b> {dir_label}<br/>"
                    f"<b>Cự ly:</b> {v['distance_km']} km (Định mức: {r['listed_length_km']} km)<br/>"
                    f"<b>Giờ hoạt động:</b> {r['operating_hours']}<br/>"
                    f"<b>Thời gian chạy:</b> {v['running_time_min']} phút"
                )
                kml_lines.extend([
                    '      <Placemark>',
                    f'        <name>{html.escape(r["display_code"])} ({dir_label})</name>',
                    f'        <description><![CDATA[{desc}]]></description>',
                    f'        <styleUrl>#style_route_{r["stt"]}</styleUrl>',
                    '        <LineString>',
                    '          <tessellate>1</tessellate>',
                    f'          <coordinates>{coord_str}</coordinates>',
                    '        </LineString>',
                    '      </Placemark>'
                ])
        kml_lines.append('    </Folder>')
        
    kml_lines.append('  </Document>')
    kml_lines.append('</kml>')
    return "\n".join(kml_lines)

def run_extraction_and_export():
    """Main extraction and export workflow"""
    print("=" * 60)
    print("STARTING HCMC 30 BUS ROUTES EXTRACTION & EXPORT PIPELINE")
    print("=" * 60)
    
    ebms_routes_map = fetch_all_routes()
    print(f"Loaded {len(ebms_routes_map)} routes from EBMS API.")
    
    # Process all 30 routes concurrently
    all_processed_routes = []
    
    with concurrent.futures.ThreadPoolExecutor(max_workers=8) as executor:
        future_to_meta = {}
        for idx, meta in enumerate(TARGET_ROUTES_INFO):
            color = PALETTE[idx % len(PALETTE)]
            code = meta["code"]
            ebms_info = ebms_routes_map.get(code)
            if not ebms_info:
                print(f"Warning: Route {code} not found in EBMS!")
                continue
            future = executor.submit(fetch_route_details, meta, ebms_info, color)
            future_to_meta[future] = meta
            
        for future in concurrent.futures.as_completed(future_to_meta):
            meta = future_to_meta[future]
            try:
                res = future.result()
                all_processed_routes.append(res)
                print(f"✓ Processed [{res['display_code']}] {res['official_name']} - {res['total_stops']} stops, {res['total_points']} pts")
            except Exception as e:
                print(f"✗ Failed to process {meta['display_code']}: {e}")
                
    # Sort back by STT 1 -> 30
    all_processed_routes.sort(key=lambda x: x["stt"])
    
    print("\n" + "=" * 60)
    print("GENERATING INDIVIDUAL & MERGED GEOJSON & KML FILES...")
    print("=" * 60)
    
    merged_features = []
    summary_rows = []
    
    for r in all_processed_routes:
        clean_code = r["display_code"].replace(" ", "_").replace("-", "_")
        
        # 1. Individual GeoJSON
        geojson_data = route_to_geojson_feature_collection(r)
        geojson_path = os.path.join(INDIVIDUAL_GEOJSON_DIR, f"tuyen_{clean_code}.geojson")
        with open(geojson_path, "w", encoding="utf-8") as f:
            json.dump(geojson_data, f, ensure_ascii=False, indent=2)
            
        # 2. Individual KML
        kml_str = route_to_kml(r)
        kml_path = os.path.join(INDIVIDUAL_KML_DIR, f"tuyen_{clean_code}.kml")
        with open(kml_path, "w", encoding="utf-8") as f:
            f.write(kml_str)
            
        # Collect for merged
        merged_features.extend(geojson_data["features"])
        
        # Summary row
        outbound_var = next((v for v in r["variants"] if v["outbound"]), None)
        inbound_var = next((v for v in r["variants"] if not v["outbound"]), None)
        
        summary_rows.append({
            "STT": r["stt"],
            "Mã số tuyến": r["display_code"],
            "Tên đầy đủ theo EBMS": r["official_name"],
            "Điểm đầu tuyến": r["start_point"],
            "Điểm cuối tuyến": r["end_point"],
            "Giờ mở tuyến": r["operating_open"],
            "Giờ đóng tuyến": r["operating_close"],
            "Chiều dài theo danh mục (km)": r["listed_length_km"],
            "Chiều dài Lượt đi thực tế (km)": outbound_var["distance_km"] if outbound_var else None,
            "Chiều dài Lượt về thực tế (km)": inbound_var["distance_km"] if inbound_var else None,
            "Số trạm Lượt đi": len(outbound_var["stops"]) if outbound_var else 0,
            "Số trạm Lượt về": len(inbound_var["stops"]) if inbound_var else 0,
            "Tổng số điểm tọa độ lộ trình": r["total_points"],
            "Thời gian chạy lượt đi (phút)": outbound_var["running_time_min"] if outbound_var else None,
            "Màu sắc định dạng": r["color"]
        })
        
    # 3. Merged GeoJSON
    merged_geojson = {
        "type": "FeatureCollection",
        "properties": {
            "title": "Hệ Thống 30 Tuyến Xe Buýt TP.HCM (Hợp Nhất)",
            "total_routes": len(all_processed_routes),
            "generated_at": pd.Timestamp.now().isoformat(),
            "source": "EBMS - Trung tâm Quản lý Giao thông công cộng TP.HCM"
        },
        "features": merged_features
    }
    merged_geojson_path = os.path.join(MERGED_DIR, "hcmc_30_bus_routes_merged.geojson")
    with open(merged_geojson_path, "w", encoding="utf-8") as f:
        json.dump(merged_geojson, f, ensure_ascii=False, indent=2)
        
    # 4. Merged KML
    merged_kml_str = generate_merged_kml(all_processed_routes)
    merged_kml_path = os.path.join(MERGED_DIR, "hcmc_30_bus_routes_merged.kml")
    with open(merged_kml_path, "w", encoding="utf-8") as f:
        f.write(merged_kml_str)
        
    # 5. Excel & CSV Summaries
    df_summary = pd.DataFrame(summary_rows)
    excel_path = os.path.join(SUMMARY_DIR, "danh_sach_30_tuyen_bus.xlsx")
    csv_path = os.path.join(SUMMARY_DIR, "danh_sach_30_tuyen_bus.csv")
    
    with pd.ExcelWriter(excel_path, engine="openpyxl") as writer:
        df_summary.to_excel(writer, index=False, sheet_name="30_Tuyen_Bus_TPHCM")
    df_summary.to_csv(csv_path, index=False, encoding="utf-8-sig")
    
    # 6. Save JS data payload for Visualizer Map
    js_data_path = os.path.join(VISUALIZER_DIR, "routes_data.js")
    with open(js_data_path, "w", encoding="utf-8") as f:
        f.write(f"const BUS_ROUTES_DATA = {json.dumps(all_processed_routes, ensure_ascii=False)};\n")
        
    print("\n" + "=" * 60)
    print("ALL FILES GENERATED SUCCESSFULLY!")
    print(f"1. Individual GeoJSON files (30): {INDIVIDUAL_GEOJSON_DIR}")
    print(f"2. Individual KML files (30):     {INDIVIDUAL_KML_DIR}")
    print(f"3. Merged GeoJSON:                {merged_geojson_path}")
    print(f"4. Merged KML:                    {merged_kml_path}")
    print(f"5. Excel Summary:                 {excel_path}")
    print(f"6. CSV Summary:                   {csv_path}")
    print(f"7. Visualizer Data JS:            {js_data_path}")
    print("=" * 60)

if __name__ == "__main__":
    run_extraction_and_export()
