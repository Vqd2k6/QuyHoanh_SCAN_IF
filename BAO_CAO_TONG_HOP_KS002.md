# BÁO CÁO TỔNG HỢP & PHÂN TÍCH DỮ LIỆU QUY HOẠCH ĐÔ THỊ TP. HỒ CHÍ MINH
## ĐỢT KHẢO SÁT: KS002 — HÀNH LANG TUYẾN METRO SỐ 2 (BẾN THÀNH – THAM LƯƠNG)
### PHỤC VỤ CÔNG TÁC KHẢO SÁT HIỆN TRẠNG CÔNG TRÌNH XÂY DỰNG & NHÀ Ở TRÊN ĐẤT

---

**Kính gửi:** Ban Giám Đốc / Trưởng Phòng Đầu Tư, Pháp Lý & Khảo Sát Hiện Trạng  
**Đơn vị thực hiện:** Bộ phận Khảo sát Dữ liệu Quy hoạch Tự động  
**Thời gian hoàn thành:** 13/09/2026  
**Dữ liệu nguồn:** Hệ thống Thông tin Quy hoạch Đô thị — Sở Quy hoạch & Kiến trúc TP.HCM (SQHKT)  
**File đính kèm bàn giao:**
1. 📊 **Excel Phân Tích Hiện Trạng Công Trình & Nhà Ở (7 Sheet):** [Bao_Cao_Phan_Tich_Quy_Hoach_KS002.xlsx](file:///Users/vqd2k6/Desktop/Project/QuyHoach/data/output/KS002/Bao_Cao_Phan_Tich_Quy_Hoach_KS002.xlsx)
2. 🗺️ **Bản Đồ Quy Hoạch Trực Quan Tương Tác:** [ban_do_quy_hoach.html](file:///Users/vqd2k6/Desktop/Project/QuyHoach/data/output/KS002/ban_do_quy_hoach.html)
3. 📁 **Excel Dữ Liệu Gốc Đầy Đủ:** [KS002.xlsx](file:///Users/vqd2k6/Desktop/Project/QuyHoach/data/output/KS002/KS002.xlsx)

---

## 1. TÓM TẮT ĐIỀU HÀNH (EXECUTIVE SUMMARY)

Đợt khảo sát **KS002** được thực hiện dọc theo hành lang tuyến **Metro Số 2 (Bến Thành – Tham Lương)** đi qua 6 quận (Quận 1, Quận 3, Quận 10, Tân Bình, Tân Phú, Quận 12) với mục tiêu **trích xuất và nhận diện toàn bộ các thửa đất có nhà ở, khu dân cư hoặc công trình xây dựng (thương mại, dịch vụ, cơ quan, trụ sở, trường học, bệnh viện, cao ốc phức hợp)** phục vụ công tác đối soát hiện trạng và khảo sát thực địa.

Hệ thống đã quét lưới vi mô đa luồng độ phân giải **4 mét** (82.824 điểm toạ độ), trích xuất thành công **6.321 thửa đất địa chính** và **10.932 ô quy hoạch chi tiết 1/2000**.

> [!IMPORTANT]
> **Đặc điểm trọng tâm của bộ dữ liệu:**
> - Toàn bộ dữ liệu giữ **nguyên bản chi tiết thô**, bảo toàn các mã ô `[Mã Ô]`, tên chức năng chi tiết, diện tích m² và tỷ lệ % của từng ô (ví dụ: `• Đất giao thông: 1,160.11 m² (14.5%) • [III.9] Dân cư dự kiến: 1,597.30 m² (20.0%) • [II.25] Ga Depot: 4,988.29 m² (62.4%)`).
> - Tập trung phân loại các thửa **có nhà ở** và các thửa **có công trình xây dựng** để đội khảo sát dễ dàng định vị, kiểm tra thực địa các công trình kiến trúc.

### 📌 Các Chỉ Số Cốt Lõi Về Nhà Ở & Công Trình:

| Chỉ số Khảo sát | Giá trị Định lượng | Tỷ lệ % | Ý nghĩa Phục vụ Khảo sát Thực địa |
| :--- | :---: | :---: | :--- |
| **Tổng số thửa đất khảo sát** | **6.321 thửa** | **100,0%** | Toàn bộ bất động sản trong hành lang tuyến Metro số 2 |
| **Tổng diện tích khảo sát** | **2.856.129,4 m²** | **~ 285,61 ha** | Quy mô hành lang khảo sát |
| **Tổng số ô phân khu chi tiết** | **10.932 ô** | **72 loại mục đích** | Dữ liệu quy hoạch phân khu chi tiết 1/2000 |
| **THỬA CÓ NHÀ Ở / DÂN CƯ** | **4.798 thửa** | **75,91%** | Các lô đất có nhà ở đô thị, khu dân cư hiện hữu, cải tạo |
| **THỬA CÓ CÔNG TRÌNH XÂY DỰNG** | **1.131 thửa** | **17,89%** | Đất có công trình TMDV, cơ quan, trường học, bệnh viện, phức hợp |
| **TỔNG THỬA CÓ NHÀ Ở HOẶC CÔNG TRÌNH** | **5.715 thửa** | **90,41%** | **Tệp đối tượng chính cần đi khảo sát thực địa hiện trạng** |
| **Thửa dính lộ giới mở đường** | **4.282 thửa** | **67,74%** | Nhà ở / công trình tiếp giáp mặt tiền CMT8, Trường Chinh |
| **Thửa thuộc dự án 1/500 & ĐCCB** | **632 bản ghi** | **10,00%** | Các công trình dự án lớn (Ga Depot, cao ốc, khu tái định cư...) |

---

## 2. PHÂN BỔ NHÀ Ở & CÔNG TRÌNH THEO 6 QUẬN / HUYỆN

```
[Quận 1: 500 thửa] ➔ [Quận 3: 1,419 thửa] ➔ [Quận 10: 841 thửa] ➔ [Quận Tân Bình: 3,208 thửa] ➔ [Quận Tân Phú: 318 thửa] ➔ [Quận 12: 35 thửa]
```

### 📊 Bảng Thống Kê Hiện Trạng Xây Dựng Từng Quận:

| STT | Quận / Huyện | Tổng Thửa | Tỷ Lệ (%) | Tổng DT (m²) | Thửa Có Nhà Ở | Tỷ Lệ Nhà Ở (%) | Thửa Có Công Trình XD | Tỷ Lệ CT (%) | Tổng Thửa Có Nhà Ở Hoặc CT | Thửa Có DA 1/500 |
| :-: | :--- | :-: | :-: | :-: | :-: | :-: | :-: | :-: | :-: | :-: |
| 1 | **Quận Tân Bình** | **3.208** | 50,75% | 754.664,6 | **2.585** | 80,58% | **334** | 10,41% | **2.880 (89,8%)** | 34 |
| 2 | **Quận 3** | **1.419** | 22,45% | 325.109,5 | **1.134** | 79,92% | **188** | 13,25% | **1.309 (92,2%)** | 316 |
| 3 | **Quận 10** | **841** | 13,30% | 393.770,0 | **682** | 81,09% | **150** | 17,84% | **819 (97,4%)** | 108 |
| 4 | **Quận 1** | **500** | 7,91% | 393.270,7 | **185** | 37,00% | **354** | 70,80% | **456 (91,2%)** | 165 |
| 5 | **Quận Tân Phú** | **318** | 5,03% | 737.241,9 | **206** | 64,78% | **76** | 23,90% | **222 (69,8%)** | 8 |
| 6 | **Quận 12** | **35** | 0,55% | 252.072,7 | **6** | 17,14% | **29** | 82,86% | **29 (82,9%)** | 1 |
| - | **TỔNG CỘNG** | **6.321** | **100,00%** | **2.856.129,4** | **4.798** | **75,91%** | **1.131** | **17,89%** | **5.715 (90,41%)** | **632** |

### 🔍 Nhận Xét Phục Vụ Khảo Sát Hiện Trạng:
* **Tân Bình, Quận 3, Quận 10:** Tỷ lệ nhà ở dân cư chiếm áp đảo (80 - 81%), mật độ nhà phố liền kề san sát, chủ yếu là nhà ở hiện hữu cải tạo và chỉnh trang.
* **Quận 1:** 70,8% số thửa là **Công trình xây dựng** (Tòa nhà văn phòng, khách sạn, trung tâm thương mại, cao ốc phức hợp).
* **Quận 12 & Tân Phú:** Chứa các cụm công trình quy mô diện tích lớn (Ga Depot Metro Tham Lương, khu kho bãi, cơ sở sản xuất).

---

## 3. DANH MỤC THỐNG KÊ CHI TIẾT 72 LOẠI MỤC ĐÍCH SỬ DỤNG ĐẤT THÔ (RAW)

| STT | Tên Loại Đất Thô (SQHKT) | Phân Loại Tham Khảo | Số Ô Xuất Hiện | Số Thửa Tiếp Giáp | Tổng Diện Tích Ô (m²) | Tỷ Lệ (%) | Mật Độ Max (%) | HSSDĐ Max | Tầng Cao Max |
| :-: | :--- | :--- | :-: | :-: | :-: | :-: | :-: | :-: | :-: |
| 1 | `Đất giao thông` | Giao Thông | 4.587 | 4.279 | 976.451,1 | 34,19% | - | - | - |
| 2 | `Đất hỗn hợp` | Phức Hợp / Hỗn Hợp | 59 | 55 | 444.971,1 | 15,58% | 45,0% | 13,5 | - |
| 3 | `Đất ở` | Đất Ở / Nhà Ở | 1.843 | 1.840 | 156.115,2 | 5,47% | 66,8% | 5,3 | - |
| 4 | `Đất công viên cây xanh` | Cây Xanh | 27 | 26 | 151.953,6 | 5,32% | 10,0% | 0,1 | - |
| 5 | `Đất quân sự` | Quốc Phòng | 2 | 2 | 140.245,2 | 4,91% | - | - | - |
| 6 | `Đất ở - hiện hữu` | Đất Ở / Nhà Ở | 905 | 892 | 82.629,9 | 2,89% | - | 1,7 | - |
| 7 | `Dân cư dự kiến` | Đất Ở / Dân Cư | 27 | 20 | 71.475,9 | 2,50% | 30,0% | 4,0 | - |
| 8 | `Đất công trình dịch vụ công cộng` | Công Trình Công Cộng | 170 | 169 | 67.211,4 | 2,35% | 49,0% | 5,8 | - |
| 9 | `Đất phức hợp` | Công Trình Phức Hợp | 483 | 463 | 64.903,9 | 2,27% | 100,0% | 16,0 | - |
| 10 | `Mặt nước` | Mặt Nước | 12 | 8 | 53.887,8 | 1,89% | - | - | - |
| 11 | `Đất ở hiện hữu` | Đất Ở / Nhà Ở | 452 | 452 | 50.897,0 | 1,78% | - | - | - |
| 12 | `Đất giáo dục` | Công Trình Giáo Dục | 53 | 53 | 48.461,2 | 1,70% | 60,0% | 3,5 | 6.0 |
| 13 | `Đất công trình công cộng - y tế` | Công Trình Y Tế | 5 | 4 | 40.845,8 | 1,43% | 60,0% | 3,2 | - |
| 14 | `Đất công nghiệp` | Công Trình Công Nghiệp | 5 | 5 | 38.235,8 | 1,34% | - | - | - |
| 15 | `Đất cây xanh - thể dục thể thao` | Thể Thao | 2 | 2 | 34.578,8 | 1,21% | 10,0% | 0,1 | - |
| 16 | `Cây xanh cách ly` | Cách Ly | 19 | 19 | 33.616,7 | 1,18% | - | - | - |
| 17 | `Đất cây xanh cách ly` | Cách Ly | 31 | 31 | 29.526,3 | 1,03% | - | - | - |
| 18 | `Đất công trình tôn giáo` | Công Trình Tôn Giáo | 31 | 31 | 24.065,3 | 0,84% | 60,0% | 2,5 | - |
| 19 | `Đất sử dụng hỗn hợp` | Công Trình Hỗn Hợp | 80 | 80 | 23.989,6 | 0,84% | 65,0% | 13,0 | - |
| 20 | `Dân cư hiện hữu cải tạo` | Đất Ở / Nhà Ở | 304 | 304 | 23.508,6 | 0,82% | 60,0% | 3,0 | - |
| 21 | `Đất ở chỉnh trang` | Đất Ở / Nhà Ở | 458 | 458 | 23.406,3 | 0,82% | - | 1,8 | - |
| 22 | `Đất phức hợp - chủ đạo khách sạn` | Khách Sạn & Dịch Vụ | 44 | 36 | 21.280,0 | 0,75% | 80,0% | 5,0 | - |
| 23 | `Đất quảng trường` | Không Gian Công Cộng | 32 | 32 | 20.035,6 | 0,70% | 5,0% | 0,1 | - |
| 24 | `Đất nhóm nhà ở hiện trạng` | Đất Ở / Nhà Ở | 335 | 335 | 19.935,9 | 0,70% | - | - | - |
| 25 | `Đất ở hiện hữu - cải tạo` | Đất Ở / Nhà Ở | 125 | 121 | 17.782,4 | 0,62% | 60,0% | 3,5 | 5.0 |
| 26 | `Đất thương mại dịch vụ` | Thương Mại - Dịch Vụ | 49 | 48 | 16.499,6 | 0,58% | 40,0% | 3,5 | 1.0 |
| 27 | `Đất công trình công cộng` | Công Trình Công Cộng | 55 | 54 | 16.442,0 | 0,58% | 66,8% | 5,3 | - |
| 28 | `Đất hành chính` | Trụ Sở Cơ Quan / Hành Chính | 40 | 40 | 14.653,7 | 0,51% | 80,0% | 5,0 | - |
| 29 | `Đất dân cư cải tạo` | Đất Ở / Nhà Ở | 284 | 284 | 14.121,3 | 0,49% | 75,0% | 1,8 | - |
| 30 | `CV - thể dục thể thao - VH` | Cây Xanh TDTT | 99 | 97 | 13.872,9 | 0,49% | - | - | - |
| 31 | `Đất cơ quan - SXKD` | Trụ Sở Cơ Quan / SXKD | 8 | 8 | 11.233,4 | 0,39% | 60,0% | 4,2 | - |
| 32 | `Đất trường học` | Công Trình Giáo Dục | 25 | 25 | 9.771,9 | 0,34% | 45,0% | 1,8 | - |
| 33 | `Đất tôn giáo` | Công Trình Tôn Giáo | 9 | 9 | 6.840,9 | 0,24% | 40,0% | 1,2 | - |
| 34 | `Đất nhóm nhà ở hiện hữu` | Đất Ở / Nhà Ở | 50 | 50 | 5.180,9 | 0,18% | 60,0% | 2,4 | - |
| 35 | `Ga Depot` | Ga Depot Metro | 5 | 5 | 5.018,7 | 0,18% | - | - | - |
| 36 | `Đất y tế` | Công Trình Y Tế | 21 | 21 | 4.871,8 | 0,17% | 49,0% | 5,9 | - |
| 37 | `Đất tôn giáo di tích` | Tôn Giáo Di Tích | 3 | 3 | 4.545,3 | 0,16% | - | - | - |
| 38 | `Công trình thương mại - dịch vụ` | Thương Mại - Dịch Vụ | 9 | 9 | 3.659,9 | 0,13% | 80,0% | 4,0 | - |
| 39 | `Khu nhà ở hiện hữu không cho XD mới` | Đất Ở Hiện Trạng | 24 | 24 | 3.327,1 | 0,12% | - | - | - |
| 40 | `Đất phức hợp - chủ đạo văn hóa/GT` | Phức Hợp Văn Hóa | 9 | 9 | 3.090,3 | 0,11% | 60,0% | 5,0 | - |
| 41 | `Đất công trình công cộng - giáo dục` | Công Trình Giáo Dục | 6 | 6 | 2.684,2 | 0,09% | - | - | - |
| 42 | `Công trình hành chánh - y tế` | Hành Chính & Y Tế | 6 | 6 | 2.508,4 | 0,09% | 60,0% | 3,5 | - |
| 43 | `Công trình giáo dục` | Công Trình Giáo Dục | 4 | 4 | 2.378,8 | 0,08% | - | - | - |
| 44 | `Phức hợp` | Phức Hợp | 9 | 9 | 2.339,2 | 0,08% | 70,0% | 4,9 | - |
| 45 | `Đất ở dự kiến chỉnh trang - XD mới` | Đất Ở / Nhà Ở | 40 | 40 | 1.955,0 | 0,07% | 40,0% | 5,2 | - |
| 46 | `Đất tôn giáo, di tích` | Tôn Giáo | 4 | 4 | 1.944,9 | 0,07% | - | - | - |
| 47 | `Đất phức hợp - chủ đạo chức năng ở` | Phức Hợp Ở | 8 | 8 | 1.777,2 | 0,06% | 60,0% | 6,0 | - |
| 48 | `Đất công cộng (Thương mại, dịch vụ)` | Thương Mại Dịch Vụ | 3 | 3 | 1.745,5 | 0,06% | 60,0% | 4,0 | - |
| 49 | `Kênh - rạch` | Kênh Rạch | 3 | 3 | 1.259,5 | 0,04% | - | - | - |
| 50 | `Đất nhóm nhà ở hiện hữu chỉnh trang` | Đất Ở / Nhà Ở | 12 | 12 | 1.157,4 | 0,04% | - | - | - |
| 51 | `Đất công cộng đơn vị ở` | Công Cộng Đơn Vị Ở | 3 | 3 | 871,9 | 0,03% | 40,0% | 2,4 | - |
| 52 | `Khu dân cư hiện hữu cải tạo chỉnh trang`| Đất Ở / Nhà Ở | 11 | 11 | 829,3 | 0,03% | - | - | - |
| 53 | `Công viên cây xanh` | Cây Xanh | 4 | 4 | 741,4 | 0,03% | - | - | - |
| 54 | `Đất công trình giáo dục` | Công Trình Giáo Dục | 3 | 3 | 647,0 | 0,02% | - | - | - |
| 55 | `Đất công trình công cộng - hành chính`| Hành Chính | 3 | 3 | 557,2 | 0,02% | - | - | - |
| 56 | `Đất ở - xây dựng mới` | Đất Ở / Nhà Ở | 3 | 3 | 521,8 | 0,02% | - | - | - |
| 57 | `Dân cư hiện hữu` | Đất Ở / Nhà Ở | 5 | 5 | 451,5 | 0,02% | - | - | - |
| 58 | `Đất khác` | Khác | 5 | 5 | 390,7 | 0,01% | - | - | - |
| 59 | `Sông rạch, mặt nước` | Mặt Nước | 2 | 2 | 344,4 | 0,01% | - | - | - |
| 60 | `Đất công trình đầu mối HTKT` | Hạ Tầng Kỹ Thuật | 2 | 2 | 321,1 | 0,01% | - | - | - |
| 61 | `Đất nhóm nhà ở xây dựng mới` | Đất Ở / Nhà Ở | 2 | 2 | 283,2 | 0,01% | - | - | - |
| 62 | `Đất cây xanh` | Cây Xanh | 2 | 2 | 165,3 | 0,01% | - | - | - |
| 63 | `Đất cây xanh CV - TDTT` | Cây Xanh TDTT | 2 | 2 | 160,8 | 0,01% | - | - | - |
| 64 | `Cây xanh CV - TDTT` | Cây Xanh TDTT | 1 | 1 | 134,8 | 0,00% | - | - | - |
| 65 | `Trạm cấp nước` | Trạm Cấp Nước | 1 | 1 | 91,4 | 0,00% | - | - | - |
| 66 | `Đất tiểu thủ công nghiệp` | Tiểu Thủ CN | 1 | 1 | 87,4 | 0,00% | - | - | - |
| 67 | `Đất quốc phòng` | Quốc Phòng | 1 | 1 | 79,6 | 0,00% | - | - | - |
| 68 | `Đất kho tàng bến bãi` | Kho Bãi | 1 | 1 | 74,4 | 0,00% | - | - | - |
| 69 | `Đất cây xanh phục vụ công cộng` | Cây Xanh | 1 | 1 | 63,4 | 0,00% | - | - | - |
| 70 | `Đất CV CX - TDTT` | Cây Xanh TDTT | 1 | 1 | 48,2 | 0,00% | - | - | - |
| 71 | `Đất ở xây dựng mới` | Đất Ở / Nhà Ở | 1 | 1 | 29,2 | 0,00% | - | - | - |
| 72 | `Đất hạ tầng kỹ thuật` | Hạ Tầng Kỹ Thuật | 1 | 1 | 6,3 | 0,00% | - | - | - |

---

## 4. HƯỚNG DẪN SỬ DỤNG 7 SHEET TRONG FILE EXCEL BÁO CÁO

File báo cáo [Bao_Cao_Phan_Tich_Quy_Hoach_KS002.xlsx](file:///Users/vqd2k6/Desktop/Project/QuyHoach/data/output/KS002/Bao_Cao_Phan_Tich_Quy_Hoach_KS002.xlsx) được thiết kế tối ưu cho công tác khảo sát thực địa:

1. **Sheet `📊 1. Dashboard & KPIs`:** Bảng điều khiển tóm tắt số lượng thửa nhà ở, công trình xây dựng, lộ giới theo từng quận.
2. **Sheet `📚 2. Thống Kê 72 Loại Đất Thô`:** Bảng tra cứu catalog 72 loại đất nguyên bản SQHKT (diện tích, tỷ lệ %, số ô, mật độ, tầng cao).
3. **Sheet `🔍 3. Chi Tiết 10.932 Ô QH` (CÔNG CỤ LỌC TỪNG Ô THÔ):** Toàn bộ 10.932 ô quy hoạch độc lập kèm cột `Nhãn Ô Chi Tiết Đầy Đủ` (`[Mã Ô] Tên loại đất: m² (%)`) có sẵn AutoFilter để lọc bất kỳ mục đích đất nào.
4. **Sheet `🏡 4. Thửa Có Đất Ở - Nhà Ở` (4.798 THỬA):** Danh sách toàn bộ các thửa có nhà ở / đất ở dân cư, hiển thị diện tích đất ở (m²), tỷ lệ %, tầng cao max, mật độ max, HSSDĐ max và lộ giới.
5. **Sheet `🏢 5. Thửa Có Công Trình XD` (1.130 THỬA):** Danh sách toàn bộ các thửa có công trình xây dựng (TMDV, Cao ốc phức hợp, Văn phòng, Trường học, Bệnh viện, Cơ quan, Khách sạn, Ga Depot...) kèm diện tích công trình (m²), tỷ lệ % và chỉ tiêu kiến trúc.
6. **Sheet `🏗️ 6. Toàn Bộ Nhà Ở & CT (5715)` (5.715 THỬA - 90,41%):** **BẢNG TỔNG HỢP TOÀN DIỆN PHỤC VỤ ĐI KHẢO SÁT THỰC ĐỊA**, bao gồm toàn bộ các thửa có nhà ở hoặc công trình, phân loại rõ hiện trạng trên đất, diện tích nhà ở (m²), diện tích công trình (m²), lộ giới và toạ độ GPS.
7. **Sheet `🏛️ 7. Dự Án 1-500 & DCCB` (632 BẢN GHI):** Chi tiết 632 bản ghi dự án quy hoạch 1/500 và Quyết định điều chỉnh cục bộ.

---

## 5. BẢNG TRA CỨU & GIẢI THÍCH CHI TIẾT Ý NGHĨA CÁC TIÊU ĐỀ CỘT (HEADERS)

> [!TIP]
> **Đã tích hợp Tooltip chú thích tương tác trực tiếp trong Excel:** Khi mở file [Bao_Cao_Phan_Tich_Quy_Hoach_KS002.xlsx](file:///Users/vqd2k6/Desktop/Project/QuyHoach/data/output/KS002/Bao_Cao_Phan_Tich_Quy_Hoach_KS002.xlsx), chỉ cần **rê chuột (hover) vào bất kỳ tiêu đề cột nào** có dấu tam giác đỏ ở góc trên bên phải, hộp thoại ghi chú giải thích ý nghĩa cột sẽ tự động bật lên.

### 5.1. Nhóm Chỉ Tiêu Quy Hoạch - Kiến Trúc & Xây Dựng

| Tiêu đề cột (Header) | Tên đầy đủ & Khái niệm | Ý nghĩa thực tiễn & Cách hiểu | Công thức / Ví dụ minh họa |
| :--- | :--- | :--- | :--- |
| **`Tầng Cao Max` / `Tầng Cao Cho Phép`** | **Tầng cao tối đa** | Số tầng cao tối đa được phép xây dựng công trình theo đồ án quy hoạch phân khu 1/2000 được phê duyệt. | Ví dụ: `6 tầng` nghĩa là công trình xây dựng tại vị trí này không được vượt quá 6 tầng nổi (chưa kể tầng hầm/lửng/tum theo quy chuẩn). |
| **`Mật Độ XD Max (%)`** | **Mật độ xây dựng thuần tối đa** | Tỷ lệ phần trăm giữa diện tích chiếm đất của công trình xây dựng trên tổng diện tích lô đất. | $\text{Mật độ XD (\%)} = \frac{\text{Diện tích xây dựng tầng trệt}}{\text{Tổng diện tích thửa đất}} \times 100\%$<br>Ví dụ: Thửa đất $100\text{ m}^2$ có mật độ max $60\%$ thì chỉ được xây nhà tối đa $60\text{ m}^2$, $40\text{ m}^2$ còn lại phải làm sân lùi, giếng trời hoặc cây xanh. |
| **`HSSDĐ Max`** | **Hệ số sử dụng đất (FAR - Floor Area Ratio)** | Tỷ số giữa tổng diện tích sàn xây dựng của tất cả các tầng (không tính tầng hầm, tầng kỹ thuật, tum thang) trên tổng diện tích thửa đất. | $\text{HSSDĐ} = \frac{\text{Tổng diện tích sàn XD tất cả các tầng}}{\text{Tổng diện tích thửa đất}}$<br>Ví dụ: Thửa đất $100\text{ m}^2$ có HSSDĐ = $3.5$ thì tổng diện tích sàn xây dựng tối đa được phép là $350\text{ m}^2$. |
| **`Dự Án 1/500 & ĐCCB`** | **Dự án 1/500 & Điều chỉnh cục bộ** | Tình trạng pháp lý quy hoạch đặc biệt của thửa đất. Cho biết thửa đất có nằm trong ranh Đồ án Quy hoạch chi tiết 1/500 (dự án phát triển cụ thể như Ga Depot, Khu tái định cư...) hoặc Quyết định Điều chỉnh cục bộ (ĐCCB) của UBND TP hay không. | • `Có QH 1/500`: Đã có quy hoạch chi tiết cấp dự án.<br>• `Có ĐCCB`: Có quyết định điều chỉnh cục bộ quy hoạch.<br>• `Không`: Tuân theo quy hoạch phân khu 1/2000 chung. |
| **`Lộ Giới Tiếp Giáp`** | **Chỉ giới đường đỏ & Hành lang an toàn giao thông** | Bề rộng lộ giới của tuyến đường/hẻm tiếp giáp thửa đất tính từ tim đường sang hai bên. Xác định khoảng lùi xây dựng và phần diện tích đất sẽ bị thu hồi/không được cấp phép xây dựng kiên cố khi Nhà nước mở đường. | Ví dụ: `Đường Cách Mạng Tháng Tám (Lộ giới 35m)` nghĩa là tim đường mở rộng mỗi bên $17.5\text{ m}$. Phần đất nằm trong phạm vi này là `DT Giao Thông Lộ Giới`. |

---

### 5.2. Nhóm Định Danh, Vị Trí & Khảo Sát Thực Địa

| Tiêu đề cột (Header) | Ý nghĩa nghiệp vụ |
| :--- | :--- |
| **`Mã Thửa Đất`** | Mã định danh chuẩn hóa duy nhất của thửa đất theo hệ thống địa chính (Ghép: `Quận-Phường-SốTờ-SốThửa`). |
| **`Số Tờ` / `Số Thửa`** | Số hiệu tờ bản đồ địa chính và số thửa đất theo Giấy chứng nhận quyền sử dụng đất (Sổ hồng/Sổ đỏ). |
| **`Mã Ô Phố`** | Mã định danh ô quy hoạch phân khu trong bản đồ 1/2000 (Ví dụ: `[II.26]`, `[III.9]`, `[I/89]`). Dùng để đối chiếu với bản vẽ quy hoạch phân khu của Quận. |
| **`Phân Loại Thực Địa`** | Phân nhóm thửa đất phục vụ trực tiếp công tác khảo sát thực tế hiện trường:<br>• `🏡 Đất Có Nhà Ở / Dân Cư`: Thửa đất phục vụ ở.<br>• `🏢 Đất Có Công Trình Xây Dựng`: Thửa đất thương mại dịch vụ, cơ quan, hạ tầng.<br>• `🏡🏢 Hỗn Hợp Nhà Ở & Công Trình`: Thửa vừa có nhà ở vừa có chức năng công trình.<br>• `🛣️ Đất Lộ Giới Giao Thông Thuần`: Thửa đất thuần dính hành lang giao thông. |
| **`Toạ Độ Khảo Sát`** | Tọa độ GPS tâm thửa đất (`Kinh độ, Vĩ độ` hệ WGS84 chuẩn), có thể dán trực tiếp lên Google Maps để dẫn đường đến vị trí thực địa. |

---

### 5.3. Nhóm Diện Tích, Tỷ Lệ & Phân Bổ Chức Năng

| Tiêu đề cột (Header) | Ý nghĩa nghiệp vụ |
| :--- | :--- |
| **`Tổng DT Thửa (m²)`** | Tổng diện tích khuôn viên thửa đất theo ranh địa chính đo đạc ($m^2$). |
| **`DT Nhà Ở (m²)` & `Tỷ Lệ Nhà Ở (%)`** | Diện tích ($m^2$) và tỷ lệ phần trăm thửa đất thuộc quy hoạch đất ở / nhóm nhà ở / khu dân cư. |
| **`DT Công Trình (m²)` & `Tỷ Lệ Công Trình (%)`** | Diện tích ($m^2$) và tỷ lệ phần trăm thửa đất thuộc quy hoạch công trình kiến trúc (TMDV, cơ quan, y tế, giáo dục, tôn giáo, ga metro...). |
| **`DT Giao Thông Lộ Giới (m²)`** | Diện tích phần thửa đất bị vướng vào chỉ giới đường đỏ giao thông (phần đất có nguy cơ bị giải tỏa mở đường). |
| **`Cơ Cấu Quy Hoạch Thô (SQHKT)`** | Chuỗi dữ liệu tổng hợp toàn bộ các ô quy hoạch thành phần cấu thành nên thửa đất kèm diện tích và tỷ lệ %. |
| **`Nhãn Ô Chi Tiết Đầy Đủ`** | Cấu trúc nhãn hiển thị trực quan: `[Mã Ô] Tên loại đất: Diện tích m² (Tỷ lệ %)`. |

---

## 6. HƯỚNG DẪN KHAI THÁC BẢN ĐỒ QUY HOẠCH TRỰC QUAN

1. Mở file [ban_do_quy_hoach.html](file:///Users/vqd2k6/Desktop/Project/QuyHoach/data/output/KS002/ban_do_quy_hoach.html) trên trình duyệt.
2. Bấm nút `👁️ Ẩn/Hiện chấm` để bật/tắt tức thì 82.824 chấm quét.
3. Gõ Số tờ - Số thửa (ví dụ: `Tờ 47 Thửa 1` hoặc `Tờ 15 Thửa 24`) trên thanh tìm kiếm để bản đồ tự động bay tới ranh thửa đất.
4. Click trực tiếp lên thửa đất để xem Popup chi tiết các ô quy hoạch nhà ở và công trình xây dựng.

---
*Báo cáo được trích xuất tự động và lập trình phân tích bởi Hệ thống Quét Quy Hoạch Tự Động.*

