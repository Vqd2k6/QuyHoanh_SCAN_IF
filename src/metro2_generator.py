#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Generator for Metro Line 2 Stations (S1 - S26) Data & GIS Formats
Corridor: Phan Văn Hớn (Q.12) -> An Sương -> Hóc Môn -> Củ Chi -> Thái Mỹ
"""

import os
import json
import math
import pandas as pd
from shapely.geometry import LineString, Point, Polygon, mapping
from pyproj import Transformer
from shapely.ops import transform

# 26 stations provided by user
STATIONS_DATA = [
    {"stt": 1, "code": "S1", "name": "Phan Văn Hớn", "lat": 10.82840562, "lng": 106.62438281, "district": "Quận 12", "ward": "P. Tân Thới Nhất", "role": "Ga đầu mối kết nối Tuyến Metro 2 Giai đoạn 1 & Depot Tham Lương", "is_terminal": True, "is_hub": True},
    {"stt": 2, "code": "S2", "name": "Đông Hưng Thuận", "lat": 10.83683484, "lng": 106.61897037, "district": "Quận 12", "ward": "P. Đông Hưng Thuận", "role": "Ga trên cao dọc trục Trường Chinh", "is_terminal": False, "is_hub": False},
    {"stt": 3, "code": "S3", "name": "An Sương", "lat": 10.84510004, "lng": 106.61361613, "district": "Quận 12", "ward": "Nút giao An Sương", "role": "Ga trung chuyển đa phương thức / Bến xe khách An Sương", "is_terminal": False, "is_hub": True},
    {"stt": 4, "code": "S4", "name": "Nguyễn Ảnh Thủ", "lat": 10.85394911, "lng": 106.60786161, "district": "Hóc Môn", "ward": "X. Bà Điểm - Trung Chánh", "role": "Ga trên cao cửa ngõ phía Nam Huyện Hóc Môn", "is_terminal": False, "is_hub": False},
    {"stt": 5, "code": "S5", "name": "Nguyễn Thị Sóc", "lat": 10.86237159, "lng": 106.60213891, "district": "Hóc Môn", "ward": "X. Xuân Thới Đông", "role": "Ga tiếp cận trung tâm Chợ đầu mối Nông sản Hóc Môn", "is_terminal": False, "is_hub": False},
    {"stt": 6, "code": "S6", "name": "Bà Triệu", "lat": 10.87618576, "lng": 106.59233429, "district": "Hóc Môn", "ward": "Thị trấn Hóc Môn", "role": "Ga trung tâm hành chính Đô thị Hóc Môn", "is_terminal": False, "is_hub": True},
    {"stt": 7, "code": "S7", "name": "Lý Thường Kiệt", "lat": 10.88455120, "lng": 106.58645867, "district": "Hóc Môn", "ward": "X. Tân Hiệp", "role": "Ga trên cao trục Quốc Lộ 22", "is_terminal": False, "is_hub": False},
    {"stt": 8, "code": "S8", "name": "Nguyễn Thị Nuôi", "lat": 10.89412744, "lng": 106.57975756, "district": "Hóc Môn", "ward": "X. Tân Hiệp", "role": "Ga cửa ngõ phía Bắc Hóc Môn giáp Củ Chi", "is_terminal": False, "is_hub": False},
    {"stt": 9, "code": "S9", "name": "Vòng cung Tây Bắc 2", "lat": 10.90256984, "lng": 106.57382749, "district": "Củ Chi", "ward": "X. Tân Phú Trung", "role": "Ga kết nối khu đô thị mới Vòng cung Tây Bắc", "is_terminal": False, "is_hub": False},
    {"stt": 10, "code": "S10", "name": "Vành Đai 3", "lat": 10.91237123, "lng": 106.56706798, "district": "Củ Chi", "ward": "X. Tân Phú Trung", "role": "Ga nút giao liên vùng Vành Đai 3 TP.HCM & Cao tốc Mộc Bài", "is_terminal": False, "is_hub": True},
    {"stt": 11, "code": "S11", "name": "An Hạ", "lat": 10.92166257, "lng": 106.56103416, "district": "Củ Chi", "ward": "X. Tân Phú Trung", "role": "Ga tiếp cận Khu công nghiệp & Cụm tiểu thủ công nghiệp An Hạ", "is_terminal": False, "is_hub": False},
    {"stt": 12, "code": "S12", "name": "Tân Phú Trung", "lat": 10.92976802, "lng": 106.55484287, "district": "Củ Chi", "ward": "X. Tân Phú Trung", "role": "Ga trung tâm dân cư Xã Tân Phú Trung", "is_terminal": False, "is_hub": False},
    {"stt": 13, "code": "S13", "name": "Hương Lộ 2", "lat": 10.93512731, "lng": 106.54872461, "district": "Củ Chi", "ward": "X. Tân Phú Trung", "role": "Ga giao cắt tuyến huyết mạch Hương Lộ 2", "is_terminal": False, "is_hub": False},
    {"stt": 14, "code": "S14", "name": "Đào Thị Kiểm", "lat": 10.94215656, "lng": 106.53424850, "district": "Củ Chi", "ward": "X. Tân Thông Hội", "role": "Ga tiếp cận khu dân cư hiện hữu Tân Thông Hội", "is_terminal": False, "is_hub": False},
    {"stt": 15, "code": "S15", "name": "Lê Minh Nhựt", "lat": 10.94779196, "lng": 106.52294597, "district": "Củ Chi", "ward": "X. Tân Thông Hội", "role": "Ga giao cắt tuyến đường Lê Minh Nhựt", "is_terminal": False, "is_hub": False},
    {"stt": 16, "code": "S16", "name": "Tân Thông Hội", "lat": 10.95333211, "lng": 106.51500496, "district": "Củ Chi", "ward": "X. Tân Thông Hội", "role": "Ga trung tâm Xã văn hóa nông thôn mới Tân Thông Hội", "is_terminal": False, "is_hub": False},
    {"stt": 17, "code": "S17", "name": "Trần Tử Bình", "lat": 10.95807963, "lng": 106.50711265, "district": "Củ Chi", "ward": "Thị trấn Củ Chi", "role": "Ga cửa ngõ phía Nam Đô thị Thị trấn Củ Chi", "is_terminal": False, "is_hub": False},
    {"stt": 18, "code": "S18", "name": "Huỳnh Văn Cọ", "lat": 10.96255437, "lng": 106.49882580, "district": "Củ Chi", "ward": "Thị trấn Củ Chi", "role": "Ga tiếp cận khu dân cư đông đúc trung tâm Củ Chi", "is_terminal": False, "is_hub": False},
    {"stt": 19, "code": "S19", "name": "Củ Chi", "lat": 10.96746525, "lng": 106.48825815, "district": "Củ Chi", "ward": "Thị trấn Củ Chi", "role": "Ga trung tâm hành chính - kinh tế Đô thị Củ Chi / Bến xe Củ Chi - Tỉnh Lộ 8", "is_terminal": False, "is_hub": True},
    {"stt": 20, "code": "S20", "name": "BV Củ Chi", "lat": 10.97319014, "lng": 106.47848023, "district": "Củ Chi", "ward": "X. Tân An Hội", "role": "Ga Bệnh viện Đa khoa Khu vực Củ Chi", "is_terminal": False, "is_hub": False},
    {"stt": 21, "code": "S21", "name": "Nguyễn Thị Rành", "lat": 10.97827776, "lng": 106.46900914, "district": "Củ Chi", "ward": "X. Tân An Hội", "role": "Ga giao cắt đường Nguyễn Thị Rành (hướng Địa đạo Bến Dược)", "is_terminal": False, "is_hub": False},
    {"stt": 22, "code": "S22", "name": "Tân An Hội", "lat": 10.98544243, "lng": 106.45723526, "district": "Củ Chi", "ward": "X. Tân An Hội", "role": "Ga trung tâm Xã Tân An Hội & Cụm tiểu thủ công nghiệp", "is_terminal": False, "is_hub": False},
    {"stt": 23, "code": "S23", "name": "Ba Sa", "lat": 10.99142068, "lng": 106.44962812, "district": "Củ Chi", "ward": "X. Phước Hiệp", "role": "Ga giao cắt đường Ba Sa / Quốc Lộ 22", "is_terminal": False, "is_hub": False},
    {"stt": 24, "code": "S24", "name": "Phước Hiệp", "lat": 11.00180658, "lng": 106.43812079, "district": "Củ Chi", "ward": "X. Phước Hiệp", "role": "Ga trung tâm Xã Phước Hiệp", "is_terminal": False, "is_hub": False},
    {"stt": 25, "code": "S25", "name": "Tỉnh Lộ 7", "lat": 11.00628863, "lng": 106.42606194, "district": "Củ Chi", "ward": "X. Phước Hiệp", "role": "Ga nút giao Tỉnh Lộ 7 huyết mạch kết nối Bến Súc - Dầu Tiếng", "is_terminal": False, "is_hub": True},
    {"stt": 26, "code": "S26", "name": "Thái Mỹ", "lat": 11.01344302, "lng": 106.41244102, "district": "Củ Chi", "ward": "X. Thái Mỹ", "role": "Ga cuối ranh giới TP.HCM - Tây Ninh / Depot Thái Mỹ / Trạm thu phí An Sương - Mộc Bài", "is_terminal": True, "is_hub": True}
]

def format_chainage(meters):
    km = int(meters // 1000)
    m = int(round(meters % 1000))
    return f"Km {km}+{m:03d}"

def main():
    out_dir = 'data/output/metro2_stations'
    os.makedirs(out_dir, exist_ok=True)

    # 1. Load high-precision corridor line from mit2s_data.json
    mit2s_path = 'data/output/bus_routes/visualizer/mit2s_data.json'
    with open(mit2s_path, 'r', encoding='utf-8') as f:
        mit_data = json.load(f)
    
    # Reverse coordinates so they run from S1 (Phan Văn Hớn, SE) to S26 (Thái Mỹ, NW)
    line_coords_se_nw = list(reversed(mit_data['coordinates']))

    t_to_3857 = Transformer.from_crs('EPSG:4326', 'EPSG:3857', always_xy=True)
    t_to_4326 = Transformer.from_crs('EPSG:3857', 'EPSG:4326', always_xy=True)

    line_geom = LineString(line_coords_se_nw)
    line_3857 = transform(t_to_3857.transform, line_geom)

    # 2. Enrich stations with projection chainage, step distances, and 500m catchment buffer
    processed_stations = []
    excel_rows = []

    for idx, s in enumerate(STATIONS_DATA):
        pt_geo = Point(s['lng'], s['lat'])
        pt_3857 = transform(t_to_3857.transform, pt_geo)
        
        # Project onto centerline
        proj_dist_m = line_3857.project(pt_3857)
        offset_m = line_3857.distance(pt_3857)

        # 500m catchment buffer circle
        buffer_3857 = pt_3857.buffer(500, resolution=24)
        buffer_geo = transform(t_to_4326.transform, buffer_3857)
        buffer_coords = [[round(c[0], 6), round(c[1], 6)] for c in list(buffer_geo.exterior.coords)]

        processed_stations.append({
            **s,
            "chainage_m": round(proj_dist_m, 1),
            "chainage_km": round(proj_dist_m / 1000.0, 2),
            "chainage_text": format_chainage(proj_dist_m),
            "offset_from_centerline_m": round(offset_m, 1),
            "catchment_radius_m": 500,
            "catchment_coords": buffer_coords
        })

    # Calculate step distances between adjacent stations
    for idx, s in enumerate(processed_stations):
        prev_m = 0.0
        next_m = 0.0
        if idx > 0:
            prev_s = processed_stations[idx - 1]
            prev_m = s['chainage_m'] - prev_s['chainage_m']
        if idx < len(processed_stations) - 1:
            next_s = processed_stations[idx + 1]
            next_m = next_s['chainage_m'] - s['chainage_m']
        
        s['dist_from_prev_m'] = round(prev_m, 1)
        s['dist_from_prev_km'] = round(prev_m / 1000.0, 2)
        s['dist_to_next_m'] = round(next_m, 1)
        s['dist_to_next_km'] = round(next_m / 1000.0, 2)

        excel_rows.append({
            "STT": s['stt'],
            "Mã Ga": s['code'],
            "Tên Ga": s['name'],
            "Vĩ độ (Latitude)": s['lat'],
            "Kinh độ (Longitude)": s['lng'],
            "Lý trình": s['chainage_text'],
            "Cự ly tích lũy (km)": s['chainage_km'],
            "Khoảng cách ga trước (km)": s['dist_from_prev_km'],
            "Khoảng cách ga sau (km)": s['dist_to_next_km'],
            "Quận / Huyện": s['district'],
            "Phường / Xã": s['ward'],
            "Loại ga": "Ga đầu/cuối" if s['is_terminal'] else ("Ga trung chuyển" if s['is_hub'] else "Ga trên cao"),
            "Chức năng / Quy hoạch": s['role']
        })

    # Overall line stats
    total_stations = len(processed_stations)
    total_length_km = round(processed_stations[-1]['chainage_km'], 2)
    avg_spacing_km = round(total_length_km / (total_stations - 1), 2)

    meta_summary = {
        "project_name": "Tuyến Đường Sắt Đô Thị Số 2 (Metro 2) TP.HCM - Giai đoạn 2 & 3",
        "corridor_name": "Hành lang Quốc Lộ 22 (Bến xe An Sương - Củ Chi - Tây Ninh)",
        "start_station": "S1 - Phan Văn Hớn",
        "end_station": "S26 - Thái Mỹ",
        "total_stations": total_stations,
        "total_length_km": total_length_km,
        "average_station_spacing_km": avg_spacing_km,
        "districts": ["Quận 12 (3 ga)", "Hóc Môn (5 ga)", "Củ Chi (18 ga)"],
        "line_color": "#dc2626",
        "track_casing_color": "#1e293b",
        "catchment_color": "#0284c7"
    }

    full_payload = {
        "meta": meta_summary,
        "stations": processed_stations,
        "line_geometry": {
            "type": "LineString",
            "coordinates": line_coords_se_nw
        }
    }

    # 3. Write stations_data.json & stations_data.js
    json_path = os.path.join(out_dir, 'stations_data.json')
    with open(json_path, 'w', encoding='utf-8') as f:
        json.dump(full_payload, f, ensure_ascii=False, indent=2)

    js_path = os.path.join(out_dir, 'stations_data.js')
    with open(js_path, 'w', encoding='utf-8') as f:
        f.write(f"const METRO2_DATA = {json.dumps(full_payload, ensure_ascii=False)};\n")

    # 4. Write Excel & CSV
    df = pd.DataFrame(excel_rows)
    xlsx_path = os.path.join(out_dir, 'metro2_stations.xlsx')
    with pd.ExcelWriter(xlsx_path, engine='openpyxl') as writer:
        df.to_excel(writer, sheet_name='Danh_Sach_26_Ga', index=False)
        worksheet = writer.sheets['Danh_Sach_26_Ga']
        # Autofit column width
        for col in worksheet.columns:
            max_len = max(len(str(cell.value or '')) for cell in col)
            col_letter = col[0].column_letter
            worksheet.column_dimensions[col_letter].width = max(max_len + 3, 12)

    csv_path = os.path.join(out_dir, 'metro2_stations.csv')
    df.to_csv(csv_path, index=False, encoding='utf-8-sig')

    # 5. Write GeoJSON
    geojson_features = []
    # 5.1 Route LineString
    geojson_features.append({
        "type": "Feature",
        "geometry": {
            "type": "LineString",
            "coordinates": line_coords_se_nw
        },
        "properties": {
            "name": "Hướng tuyến Tuyến Metro Số 2 (An Sương - Củ Chi - Tây Ninh)",
            "type": "Metro_Corridor",
            "total_stations": total_stations,
            "length_km": total_length_km
        }
    })

    # 5.2 Station Catchment Buffers (Polygons)
    for s in processed_stations:
        geojson_features.append({
            "type": "Feature",
            "geometry": {
                "type": "Polygon",
                "coordinates": [s['catchment_coords']]
            },
            "properties": {
                "station_code": s['code'],
                "station_name": s['name'],
                "type": "Catchment_Buffer_500m",
                "radius_m": 500
            }
        })

    # 5.3 Station Points
    for s in processed_stations:
        geojson_features.append({
            "type": "Feature",
            "geometry": {
                "type": "Point",
                "coordinates": [s['lng'], s['lat']]
            },
            "properties": {
                "stt": s['stt'],
                "code": s['code'],
                "name": s['name'],
                "district": s['district'],
                "ward": s['ward'],
                "chainage": s['chainage_text'],
                "chainage_km": s['chainage_km'],
                "dist_prev_km": s['dist_from_prev_km'],
                "dist_next_km": s['dist_to_next_km'],
                "role": s['role'],
                "is_terminal": s['is_terminal'],
                "is_hub": s['is_hub']
            }
        })

    geojson_payload = {
        "type": "FeatureCollection",
        "name": "Hệ thống 26 Ga Tuyến Metro Số 2 TP.HCM (QL22)",
        "features": geojson_features
    }

    geojson_path = os.path.join(out_dir, 'metro2_stations.geojson')
    with open(geojson_path, 'w', encoding='utf-8') as f:
        json.dump(geojson_payload, f, ensure_ascii=False, indent=2)

    # 6. Write KML
    kml_placemarks = []
    # Line
    line_kml_coords = ' '.join([f"{c[0]},{c[1]},0" for c in line_coords_se_nw])
    kml_placemarks.append(f"""    <Placemark>
      <name>Trục tuyến Tuyến Metro Số 2 (An Sương - Củ Chi)</name>
      <styleUrl>#metroLineStyle</styleUrl>
      <LineString>
        <tessellate>1</tessellate>
        <coordinates>{line_kml_coords}</coordinates>
      </LineString>
    </Placemark>""")

    # Stations
    for s in processed_stations:
        kml_placemarks.append(f"""    <Placemark>
      <name>{s['code']} - {s['name']}</name>
      <description><![CDATA[
        <b>Mã Ga:</b> {s['code']}<br/>
        <b>Tên Ga:</b> {s['name']}<br/>
        <b>Lý trình:</b> {s['chainage_text']} ({s['chainage_km']} km)<br/>
        <b>Quận / Huyện:</b> {s['district']} - {s['ward']}<br/>
        <b>Tọa độ:</b> {s['lat']:.8f}, {s['lng']:.8f}<br/>
        <b>Quy hoạch:</b> {s['role']}
      ]]></description>
      <styleUrl>#{'terminalStyle' if s['is_terminal'] else ('hubStyle' if s['is_hub'] else 'stationStyle')}</styleUrl>
      <Point>
        <coordinates>{s['lng']:.8f},{s['lat']:.8f},0</coordinates>
      </Point>
    </Placemark>""")

    kml_content = f"""<?xml version="1.0" encoding="UTF-8"?>
<kml xmlns="http://www.opengis.net/kml/2.2">
  <Document>
    <name>Hệ Thống 26 Ga Tuyến Metro Số 2 (An Sương - Củ Chi - Tây Ninh)</name>
    <description>Lộ trình khảo sát 26 ga Metro 2 dọc hành lang QL22 từ Phan Văn Hớn đến Thái Mỹ</description>
    <Style id="metroLineStyle">
      <LineStyle>
        <color>ff0000ff</color>
        <width>4</width>
      </LineStyle>
    </Style>
    <Style id="terminalStyle">
      <IconStyle>
        <color>ff00d7ff</color>
        <scale>1.3</scale>
        <Icon><href>http://maps.google.com/mapfiles/kml/paddle/ylw-stars.png</href></Icon>
      </IconStyle>
    </Style>
    <Style id="hubStyle">
      <IconStyle>
        <color>ffff00aa</color>
        <scale>1.2</scale>
        <Icon><href>http://maps.google.com/mapfiles/kml/paddle/purple-circle.png</href></Icon>
      </IconStyle>
    </Style>
    <Style id="stationStyle">
      <IconStyle>
        <color>ff0000ff</color>
        <scale>1.0</scale>
        <Icon><href>http://maps.google.com/mapfiles/kml/paddle/red-circle.png</href></Icon>
      </IconStyle>
    </Style>
{chr(10).join(kml_placemarks)}
  </Document>
</kml>"""

    kml_path = os.path.join(out_dir, 'metro2_stations.kml')
    with open(kml_path, 'w', encoding='utf-8') as f:
        f.write(kml_content)

    print(f"Generated successfully: {total_stations} stations, {total_length_km} km total length.")

if __name__ == '__main__':
    main()
