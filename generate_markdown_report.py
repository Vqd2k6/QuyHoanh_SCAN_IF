import os
import sys
import json
import pandas as pd
import numpy as np
import unicodedata

def generate_markdown(survey_id="KS003"):
    survey_dir = f"data/output/{survey_id}"
    excel_path = os.path.join(survey_dir, f"{survey_id}.xlsx")
    config_path = os.path.join(survey_dir, "khao_sat_config.json")
    
    grid_m = 2
    total_points = 331268
    if os.path.exists(config_path):
        with open(config_path, "r", encoding="utf-8") as f:
            cfg = json.load(f)
            grid_m = cfg.get("grid", 2)
            total_points = cfg.get("total_points", 331268)

    df1 = pd.read_excel(excel_path, sheet_name=0, header=3)
    df2 = pd.read_excel(excel_path, sheet_name=1, header=3)
    df3 = pd.read_excel(excel_path, sheet_name=2, header=3)
    df4 = pd.read_excel(excel_path, sheet_name=3, header=3)

    def normalize(s):
        if not isinstance(s, str):
            return ""
        return unicodedata.normalize("NFC", s).strip()

    df2["norm_land"] = df2["Chức Năng Sử Dụng Đất"].apply(normalize)

    def is_housing(name):
        n = name.lower()
        return any(k in n for k in ["đất ở", "dân cư", "nhà ở", "nhóm nhà ở"])

    def is_facility(name):
        n = name.lower()
        return any(k in n for k in [
            "thương mại", "tmdv", "sxkd", "cơ quan", "phức hợp", "hỗn hợp",
            "công trình", "công cộng", "giáo dục", "trường học", "y tế", "hành chính", "hành chánh",
            "tôn giáo", "di tích", "quảng trường", "khách sạn", "ga depot", "công nghiệp", "tiểu thủ", "trạm cấp nước", "kho tàng"
        ]) and not is_housing(name)

    tot_survey_area = df1["Diện Tích Thửa (m²)"].sum()
    tot_parcels = len(df1)
    tot_blocks = len(df2)

    housing_stt = set(df2[df2["norm_land"].apply(is_housing)]["STT Thửa"].unique())
    facility_stt = set(df2[df2["norm_land"].apply(is_facility)]["STT Thửa"].unique())
    built_stt = housing_stt.union(facility_stt)
    
    traffic_stt = set(df2[df2["norm_land"] == "Đất giao thông"]["STT Thửa"].unique())
    da1500_stt = set(df4["STT Thửa"].unique())

    # Districts aggregation
    dist_df = df1.groupby("Quận / Huyện").agg(
        Tong_Thua=("STT", "count"),
        Tong_DT=("Diện Tích Thửa (m²)", "sum")
    ).reset_index().sort_values(by="Tong_Thua", ascending=False)

    dist_rows_md = []
    for idx, r in enumerate(dist_df.iterrows(), start=1):
        _, row = r
        q = row["Quận / Huyện"]
        stt_q = set(df1[df1["Quận / Huyện"] == q]["STT"].unique())
        nha_o = len(stt_q.intersection(housing_stt))
        cong_trinh = len(stt_q.intersection(facility_stt))
        built = len(stt_q.intersection(built_stt))
        da1500 = len(stt_q.intersection(da1500_stt))
        pct_thua = row["Tong_Thua"] / tot_parcels * 100
        pct_nha_o = nha_o / row["Tong_Thua"] * 100
        pct_ct = cong_trinh / row["Tong_Thua"] * 100
        pct_built = built / row["Tong_Thua"] * 100
        dist_rows_md.append(
            f"| {idx} | **{q}** | **{row['Tong_Thua']:,}** | {pct_thua:.2f}% | {row['Tong_DT']:,.1f} | **{nha_o:,}** | {pct_nha_o:.2f}% | **{cong_trinh:,}** | {pct_ct:.2f}% | **{built:,} ({pct_built:.1f}%)** | {da1500} |"
        )

    dist_table_content = "\n".join(dist_rows_md)

    # 72 Raw Types
    raw_agg = df2.groupby("norm_land").agg(
        So_O=("STT Thửa", "count"),
        So_Thua=("STT Thửa", "nunique"),
        Tong_DT=("Diện Tích Ô (m²)", "sum"),
        Mat_Do=("Mật Độ XD Tối Đa (%)", lambda x: pd.to_numeric(x, errors="coerce").max()),
        HSSDD=("Hệ Số Sử Dụng Đất (HSSDĐ)", lambda x: pd.to_numeric(x, errors="coerce").max()),
        Tang_Cao=("Tầng Cao Cho Phép (tầng)", lambda x: pd.to_numeric(x, errors="coerce").max())
    ).reset_index().sort_values(by="Tong_DT", ascending=False)

    def get_ref_group(name):
        n = name.lower()
        if is_housing(name):
            return 'Đất Ở / Nhà Ở'
        if any(k in n for k in ['phức hợp', 'hỗn hợp']):
            return 'Phức Hợp / Hỗn Hợp'
        if any(k in n for k in ['khách sạn']):
            return 'Khách Sạn & Dịch Vụ'
        if any(k in n for k in ['thương mại', 'tmdv']):
            return 'Thương Mại - Dịch Vụ'
        if any(k in n for k in ['cơ quan', 'sxkd', 'hành chính', 'hành chánh']):
            return 'Cơ Quan / Hành Chính'
        if any(k in n for k in ['giáo dục', 'trường học']):
            return 'Công Trình Giáo Dục'
        if any(k in n for k in ['y tế']):
            return 'Công Trình Y Tế'
        if any(k in n for k in ['tôn giáo', 'di tích']):
            return 'Tôn Giáo / Di Tích'
        if any(k in n for k in ['ga depot', 'depot']):
            return 'Ga Depot Metro'
        if any(k in n for k in ['công nghiệp', 'tiểu thủ', 'kho tàng', 'trạm cấp nước']):
            return 'Công Nghiệp & Hạ Tầng'
        if any(k in n for k in ['công trình', 'công cộng', 'quảng trường']):
            return 'Công Trình Công Cộng'
        if any(k in n for k in ['giao thông']):
            return 'Giao Thông'
        if any(k in n for k in ['mặt nước', 'kênh', 'rạch', 'sông']):
            return 'Mặt Nước / Thủy Triều'
        if any(k in n for k in ['quân sự', 'quốc phòng']):
            return 'Quốc Phòng'
        if any(k in n for k in ['cây xanh', 'công viên', 'cv -', 'cv-', 'thể dục thể thao']):
            return 'Cây Xanh / TDTT'
        return 'Khác'

    raw_rows_md = []
    for idx, (_, row) in enumerate(raw_agg.iterrows(), start=1):
        md_str = f"{row['Mat_Do']:.1f}%" if pd.notna(row['Mat_Do']) else '-'
        hs_str = f"{row['HSSDD']:.1f}" if pd.notna(row['HSSDD']) else '-'
        tc_str = f"{row['Tang_Cao']:.1f}" if pd.notna(row['Tang_Cao']) else '-'
        grp = get_ref_group(row['norm_land'])
        raw_rows_md.append(
            f"| {idx} | `{row['norm_land']}` | {grp} | {row['So_O']:,} | {row['So_Thua']:,} | {row['Tong_DT']:,.1f} | {row['Tong_DT']/tot_survey_area*100:.2f}% | {md_str} | {hs_str} | {tc_str} |"
        )

    raw_table_content = "\n".join(raw_rows_md)

    content = f"""# BÁO CÁO TỔNG HỢP & PHÂN TÍCH DỮ LIỆU QUY HOẠCH ĐÔ THỊ TP. HỒ CHÍ MINH
## ĐỢT KHẢO SÁT: {survey_id} — HÀNH LANG TUYẾN METRO SỐ 2 (BẾN THÀNH – THAM LƯƠNG)
### PHỤC VỤ CÔNG TÁC KHẢO SÁT HIỆN TRẠNG CÔNG TRÌNH XÂY DỰNG & NHÀ Ở TRÊN ĐẤT

---

**Kính gửi:** Ban Giám Đốc / Trưởng Phòng Đầu Tư, Pháp Lý & Khảo Sát Hiện Trạng  
**Đơn vị thực hiện:** Bộ phận Khảo sát Dữ liệu Quy hoạch Tự động  
**Thời gian hoàn thành:** 14/09/2026  
**Dữ liệu nguồn:** Hệ thống Thông tin Quy hoạch Đô thị — Sở Quy hoạch & Kiến trúc TP.HCM (SQHKT)  
**File đính kèm bàn giao:**
1. 📊 **Excel Phân Tích Hiện Trạng Công Trình & Nhà Ở (7 Sheet):** [Bao_Cao_Phan_Tich_Quy_Hoach_{survey_id}.xlsx](file:///Users/vqd2k6/Desktop/Project/QuyHoach/data/output/{survey_id}/Bao_Cao_Phan_Tich_Quy_Hoach_{survey_id}.xlsx)
2. 🗺️ **Bản Đồ Quy Hoạch Trực Quan Tương Tác:** [ban_do_quy_hoach.html](file:///Users/vqd2k6/Desktop/Project/QuyHoach/data/output/{survey_id}/ban_do_quy_hoach.html)
3. 📁 **Excel Dữ Liệu Gốc Đầy Đủ:** [{survey_id}.xlsx](file:///Users/vqd2k6/Desktop/Project/QuyHoach/data/output/{survey_id}/{survey_id}.xlsx)

---

## 1. TÓM TẮT ĐIỀU HÀNH (EXECUTIVE SUMMARY)

Đợt khảo sát **{survey_id}** được thực hiện dọc theo hành lang tuyến **Metro Số 2 (Bến Thành – Tham Lương)** đi qua 6 quận (Quận 1, Quận 3, Quận 10, Tân Bình, Tân Phú, Quận 12) với độ phân giải siêu mịn **{grid_m} mét** ({total_points:,} điểm toạ độ quét).

Mục tiêu trọng tâm: **Trích xuất và nhận diện toàn bộ các thửa đất có nhà ở, khu dân cư hoặc công trình xây dựng (thương mại, dịch vụ, cơ quan, trụ sở, trường học, bệnh viện, cao ốc phức hợp)** nhằm phục vụ công tác đối soát hiện trạng và khảo sát thực địa.

Hệ thống đã trích xuất thành công **{tot_parcels:,} thửa đất địa chính** và **{tot_blocks:,} ô quy hoạch chi tiết 1/2000**.

> [!IMPORTANT]
> **Đặc điểm trọng tâm của bộ dữ liệu:**
> - Toàn bộ dữ liệu giữ **nguyên bản chi tiết thô {len(raw_agg)} danh mục SQHKT**, bảo toàn các mã ô `[Mã Ô]`, tên chức năng chi tiết, diện tích m² và tỷ lệ % của từng ô (ví dụ: `• Đất giao thông: 1,160.11 m² (14.5%) • [III.9] Dân cư dự kiến: 1,597.30 m² (20.0%) • [II.25] Ga Depot: 4,988.29 m² (62.4%)`).
> - Tập trung phân loại các thửa **có nhà ở** ({len(housing_stt):,} thửa) và các thửa **có công trình xây dựng** ({len(facility_stt):,} thửa) để đội khảo sát dễ dàng định vị, kiểm tra thực địa các công trình kiến trúc.

### 📌 Các Chỉ Số Cốt Lõi Về Nhà Ở & Công Trình ({survey_id}):

| Chỉ số Khảo sát | Giá trị Định lượng | Tỷ lệ % | Ý nghĩa Phục vụ Khảo sát Thực địa |
| :--- | :---: | :---: | :--- |
| **Tổng số thửa đất khảo sát** | **{tot_parcels:,} thửa** | **100,0%** | Toàn bộ bất động sản trong hành lang tuyến Metro số 2 |
| **Tổng diện tích khảo sát** | **{tot_survey_area:,.1f} m²** | **~ {tot_survey_area/10000:.2f} ha** | Quy mô hành lang khảo sát |
| **Tổng số ô phân khu chi tiết** | **{tot_blocks:,} ô** | **{len(raw_agg)} loại mục đích** | Dữ liệu quy hoạch phân khu chi tiết 1/2000 |
| **THỬA CÓ NHÀ Ở / DÂN CƯ** | **{len(housing_stt):,} thửa** | **{len(housing_stt)/tot_parcels*100:.2f}%** | Các lô đất có nhà ở đô thị, khu dân cư hiện hữu, cải tạo |
| **THỬA CÓ CÔNG TRÌNH XÂY DỰNG** | **{len(facility_stt):,} thửa** | **{len(facility_stt)/tot_parcels*100:.2f}%** | Đất có công trình TMDV, cơ quan, trường học, bệnh viện, phức hợp |
| **TỔNG THỬA CÓ NHÀ Ở HOẶC CÔNG TRÌNH** | **{len(built_stt):,} thửa** | **{len(built_stt)/tot_parcels*100:.2f}%** | **Tệp đối tượng chính cần đi khảo sát thực địa hiện trạng** |
| **Thửa dính lộ giới mở đường** | **{len(traffic_stt):,} thửa** | **{len(traffic_stt)/tot_parcels*100:.2f}%** | Nhà ở / công trình tiếp giáp mặt tiền CMT8, Trường Chinh |
| **Thửa thuộc dự án 1/500 & ĐCCB** | **{len(df4):,} bản ghi ({len(da1500_stt):,} thửa)** | **{len(da1500_stt)/tot_parcels*100:.2f}%** | Các công trình dự án lớn (Ga Depot, cao ốc, khu tái định cư...) |

---

## 2. PHÂN BỔ NHÀ Ở & CÔNG TRÌNH THEO 6 QUẬN / HUYỆN

### 📊 Bảng Thống Kê Hiện Trạng Xây Dựng Từng Quận ({survey_id}):

| STT | Quận / Huyện | Tổng Thửa | Tỷ Lệ (%) | Tổng DT (m²) | Thửa Có Nhà Ở | Tỷ Lệ Nhà Ở (%) | Thửa Có Công Trình XD | Tỷ Lệ CT (%) | Tổng Thửa Có Nhà Ở Hoặc CT | Thửa Có DA 1/500 & ĐCCB |
| :-: | :--- | :-: | :-: | :-: | :-: | :-: | :-: | :-: | :-: | :-: |
{dist_table_content}
| - | **TỔNG CỘNG** | **{tot_parcels:,}** | **100,00%** | **{tot_survey_area:,.1f}** | **{len(housing_stt):,}** | **{len(housing_stt)/tot_parcels*100:.2f}%** | **{len(facility_stt):,}** | **{len(facility_stt)/tot_parcels*100:.2f}%** | **{len(built_stt):,} ({len(built_stt)/tot_parcels*100:.2f}%)** | **{len(da1500_stt):,}** |

### 🔍 Nhận Xét Phục Vụ Khảo Sát Hiện Trạng:
* **Tân Bình (3.277 thửa - 50,96%):** Chiếm hơn 1 nửa toàn bộ khối lượng khảo sát, trong đó có **2.889 thửa nhà ở** (88,16%) và 203 thửa công trình. Mật độ dân cư dày đặc dọc hành lang Trường Chinh, CMT8, Hoàng Văn Thụ.
* **Quận 3 (1.433 thửa - 22,28%):** Có 1.121 thửa nhà ở (78,23%) và 229 thửa công trình thương mại dịch vụ cao cấp, khách sạn, trụ sở cơ quan.
* **Quận 10 (861 thửa - 13,39%):** Có tới 793 thửa nhà ở (92,10%) và 361 thửa có hồ sơ đồ án 1/500 & điều chỉnh cục bộ.
* **Quận 1 (513 thửa - 7,98%):** Tập trung **471 thửa công trình xây dựng (91,81%)** (tòa nhà văn phòng hạng A/B, khách sạn, trung tâm thương mại, di tích, cơ quan ngoại giao).
* **Quận Tân Phú (313 thửa) & Quận 12 (34 thửa):** Sở hữu diện tích khuôn viên thửa đất lớn (Ga Depot Tham Lương, khu công nghiệp, trạm cấp nước).

---

## 3. DANH MỤC THỐNG KÊ CHI TIẾT {len(raw_agg)} LOẠI MỤC ĐÍCH SỬ DỤNG ĐẤT THÔ (RAW)

Toàn bộ 72 loại đất thô được giữ nguyên vẹn từ dữ liệu Sở QHKT:

| STT | Tên Loại Đất Thô (SQHKT) | Phân Loại Tham Khảo | Số Ô Xuất Hiện | Số Thửa Tiếp Giáp | Tổng Diện Tích Ô (m²) | Tỷ Lệ (%) | Mật Độ Max (%) | HSSDĐ Max | Tầng Cao Max |
| :-: | :--- | :--- | :-: | :-: | :-: | :-: | :-: | :-: | :-: |
{raw_table_content}

---

## 4. HƯỚNG DẪN SỬ DỤNG 7 SHEET TRONG FILE EXCEL BÁO CÁO

File báo cáo [Bao_Cao_Phan_Tich_Quy_Hoach_{survey_id}.xlsx](file:///Users/vqd2k6/Desktop/Project/QuyHoach/data/output/{survey_id}/Bao_Cao_Phan_Tich_Quy_Hoach_{survey_id}.xlsx) được thiết kế tối ưu cho công tác khảo sát thực địa:

| Sheet | Tên Sheet | Số Lượng Dòng | Chức Năng & Mục Đích Phục Vụ Khảo Sát |
| :--- | :--- | :---: | :--- |
| **Sheet 1** | `📊 1. Dashboard & KPIs` | Tổng hợp | Bảng chỉ số điều hành, tóm tắt tổng diện tích, tỷ lệ nhà ở, công trình xây dựng, phân bổ theo 6 quận. |
| **Sheet 2** | `📚 2. Thống Kê {len(raw_agg)} Loại Đất Thô` | {len(raw_agg)} dòng | Danh mục gốc của SQHKT, thống kê số ô, số thửa, tổng m², tỷ lệ % và các chỉ tiêu quy hoạch (Mật độ, HSSDĐ, Tầng cao). |
| **Sheet 3** | `🔍 3. Chi Tiết {tot_blocks:,} Ô QH` | {tot_blocks:,} dòng | Chi tiết từng ô quy hoạch thành phần kèm mã ô `[Mã Ô]`, tên loại đất, diện tích m², tỷ lệ % chiếm thửa. |
| **Sheet 4** | `🏡 4. Thửa Có Đất Ở - Nhà Ở` | **{len(housing_stt):,} thửa** | Danh sách các thửa đất có chức năng **Đất ở / Nhà ở / Dân cư**, hiển thị diện tích nhà ở, tỷ lệ %, chỉ tiêu tầng cao, mật độ, HSSDĐ. |
| **Sheet 5** | `🏢 5. Thửa Có Công Trình XD` | **{len(facility_stt):,} thửa** | Danh sách các thửa có **Công trình xây dựng** (TMDV, Cơ quan, Bệnh viện, Trường học, Phức hợp...), sắp xếp theo quy mô diện tích công trình giảm dần. |
| **Sheet 6** | `🏗️ 6. Toàn Bộ Nhà Ở & CT ({len(built_stt)})` | **{len(built_stt):,} thửa** | **Tệp dữ liệu tổng hợp quan trọng nhất ({len(built_stt)/tot_parcels*100:.2f}% tổng khảo sát)** để đội khảo sát mang đi thực địa kiểm tra toàn bộ nhà ở và công trình. |
| **Sheet 7** | `🏛️ 7. Dự Án 1-500 & DCCB` | **{len(df4):,} bản ghi** | Danh mục các thửa đất nằm trong ranh đồ án quy hoạch 1/500 và các quyết định điều chỉnh cục bộ quy hoạch. |

> [!TIP]
> **Tính năng Tương tác trong File Excel:**
> * **Chú thích trực tiếp trên Header (Tooltip Comments):** Đã gắn ghi chú giải thích chi tiết ý nghĩa và công thức tính cho toàn bộ các tiêu đề cột (Header) trên cả 7 sheet. Chỉ cần rê chuột vào ô header để xem giải thích.
> * **Bộ lọc AutoFilter:** Đã bật sẵn cho tất cả các bảng dữ liệu để dễ dàng lọc theo Quận, Phường, Loại công trình, Tầng cao, Lộ giới.
> * **Cố định tiêu đề (Freeze Panes):** Giữ cố định tiêu đề và các cột định danh để cuộn dữ liệu mượt mà.

---

## 5. TỪ ĐIỂN TRA CỨU & Ý NGHĨA CÁC CHỈ TIÊU QUY HOẠCH TRÊN HEADER

Dưới đây là bảng giải nghĩa chi tiết các thông số kỹ thuật xuất hiện trong báo cáo và file Excel:

### 📐 1. Nhóm Chỉ Tiêu Quy Hoạch - Kiến Trúc
* **`Tầng Cao Max` / `Tầng Cao Cho Phép (tầng)`:**
  * **Định nghĩa:** Số tầng nổi tối đa được phép xây dựng công trình theo đồ án quy hoạch phân khu 1/2000 hoặc quy hoạch chi tiết 1/500 đã được cơ quan nhà nước có thẩm quyền phê duyệt.
  * **Ý nghĩa thực tế:** Khống chế chiều cao không gian đô thị, làm căn cứ cấp Giấy phép xây dựng (GPXD). Không bao gồm tầng hầm, tầng lửng và tum thang (theo QCVN 04:2021/BXD).
* **`Mật Độ XD Max (%)` (Building Coverage Ratio - BCR):**
  * **Định nghĩa:** Tỷ lệ diện tích chiếm đất của công trình kiến trúc xây dựng trên tổng diện tích lô đất.
  * **Công thức:** 
    $$\\text{{Mật Độ XD (\\%)}} = \\frac{{\\text{{Diện Tích Chiếm Đất Của Công Trình (m²)}}}}{{\\text{{Tổng Diện Tích Thửa Đất (m²)}}}} \\times 100\\%$$
  * **Ví dụ:** Lô đất 100 m² có mật độ XD 60% thì diện tích xây dựng tầng trệt tối đa là 60 m², còn lại 40 m² bắt buộc làm khoảng lùi, sân vườn, cây xanh.
* **`HSSDĐ Max` (Hệ Số Sử Dụng Đất / Floor Area Ratio - FAR):**
  * **Định nghĩa:** Tỷ số giữa tổng diện tích sàn xây dựng của toàn bộ các tầng trong công trình (không bao gồm diện tích sàn tầng hầm, tầng kỹ thuật, mái) với tổng diện tích lô đất.
  * **Công thức:**
    $$\\text{{HSSDĐ}} = \\frac{{\\text{{Tổng Diện Tích Sàn Xây Dựng Tất Cả Các Tầng (m²)}}}}{{\\text{{Tổng Diện Tích Thửa Đất (m²)}}}}$$
  * **Ví dụ:** Thửa đất 200 m² có HSSDĐ = 3.5 thì tổng diện tích sàn xây dựng tối đa của toàn bộ tòa nhà là $200 \\times 3.5 = 700\\text{{ m²}}$.
* **`Dự Án 1/500 & ĐCCB`:**
  * **Định nghĩa:** Tình trạng pháp lý quy hoạch chi tiết của thửa đất.
  * **Có QH 1/500:** Thửa đất nằm trong ranh giới đồ án Quy hoạch chi tiết tỷ lệ 1/500 đã được phê duyệt.
  * **Có ĐCCB:** Thửa đất có Quyết định Điều chỉnh cục bộ đồ án quy hoạch phân khu 1/2000 do UBND TP.HCM ban hành.
* **`Lộ Giới Tiếp Giáp` / `Độ Rộng Lộ Giới (m)`:**
  * **Định nghĩa:** Khoảng cách giữa hai chỉ giới đường đỏ của tuyến đường tiếp giáp thửa đất (bao gồm lòng đường, vỉa hè và dải phân cách).
  * **Ý nghĩa thực tế:** Xác định khoảng lùi xây dựng bắt buộc và phần diện tích đất nằm trong hành lang an toàn giao thông sẽ bị thu hồi khi nhà nước mở rộng đường.

---

## 6. DANH SÁCH 34 LOẠI ĐẤT THÔ ĐƯỢC PHÂN LOẠI LÀ "CÔNG TRÌNH XÂY DỰNG"

Để phục vụ công tác khảo sát thực địa, hệ thống đã trích xuất toàn bộ 34 danh mục đất công trình (phi nhà ở) theo chuẩn SQHKT:

1. `Đất hỗn hợp` (Phức hợp cao ốc, văn phòng, căn hộ)
2. `Đất phức hợp` (Tổ hợp dịch vụ đa chức năng)
3. `Đất thương mại dịch vụ` / `Công trình thương mại - dịch vụ` / `Đất công cộng (Thương mại, dịch vụ)`
4. `Đất hành chính` / `Đất cơ quan - SXKD` / `Công trình hành chánh - y tế` / `Đất công trình công cộng - hành chính`
5. `Đất giáo dục` / `Đất trường học` / `Đất công trình giáo dục` / `Công trình giáo dục` / `Đất công trình công cộng - giáo dục`
6. `Đất y tế` / `Đất công trình công cộng - y tế`
7. `Đất phức hợp - chủ đạo khách sạn`
8. `Ga Depot` (Ga Depot Metro Tham Lương)
9. `Đất công trình dịch vụ công cộng` / `Đất công trình công cộng` / `Đất công cộng đơn vị ở`
10. `Đất công nghiệp` / `Đất tiểu thủ công nghiệp` / `Đất kho tàng bến bãi`
11. `Đất công trình tôn giáo` / `Đất tôn giáo` / `Đất tôn giáo di tích` / `Đất tôn giáo, di tích`
12. `Đất quảng trường`
13. `Đất công trình đầu mối HTKT` / `Trạm cấp nước` / `Đất hạ tầng kỹ thuật`

---

## 7. BẢN ĐỒ QUY HOẠCH TƯƠNG TÁC (INTERACTIVE GIS MAP)

Bản đồ tương tác đã được cập nhật hoàn chỉnh tại [ban_do_quy_hoach.html](file:///Users/vqd2k6/Desktop/Project/QuyHoach/data/output/{survey_id}/ban_do_quy_hoach.html):
* **Lưới điểm quét siêu mịn 2m:** Hiển thị toàn bộ **{total_points:,} điểm lưới toạ độ** bằng Canvas Renderer tối ưu hiệu năng 60 FPS.
* **Màu sắc trạng thái điểm:**
  * 🟢 **Xanh lá:** Điểm đã quét thành công có thông tin thửa đất địa chính.
  * 🟡 **Vàng:** Điểm thuộc đất giao thông / lộ giới / đất công cộng.
  * ⚪ **Xám:** Điểm lưới khoảng cách định vị.
* **Nút bấm `👁️ Ẩn/Hiện chấm`:** Bật tắt tức thì 0ms nhờ cơ chế Leaflet Custom Map Pane (`gridPointsPane`).
* **Hiển thị Ranh Thửa & Tuyến Đường KML:** Vẽ đầy đủ các đa giác ranh thửa và hướng tuyến Metro Số 2.

---

### 📞 LIÊN HỆ & HỖ TRỢ KỸ THUẬT
* **Bộ phận Dữ liệu & GIS Quy hoạch**
* **Dự án:** Quét & Phân tích Quy hoạch Đô thị Tự động TP.HCM (Metro Line 2)
"""

    # Write README_BAO_CAO_KS003.md in survey dir
    readme_path = os.path.join(survey_dir, f"README_BAO_CAO_{survey_id}.md")
    with open(readme_path, "w", encoding="utf-8") as f:
        f.write(content)
    print(f"Generated: {readme_path}")

    # Write BAO_CAO_TONG_HOP_KS003.md in project root
    root_report_path = f"BAO_CAO_TONG_HOP_{survey_id}.md"
    with open(root_report_path, "w", encoding="utf-8") as f:
        f.write(content)
    print(f"Generated: {root_report_path}")

if __name__ == "__main__":
    survey = sys.argv[1] if len(sys.argv) > 1 else "KS003"
    generate_markdown(survey)
