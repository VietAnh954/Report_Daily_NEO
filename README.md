# 📊 BÁO CÁO DAILY KÊNH NEO (AFFINA DAILY NEO REPORT)

Hệ thống tự động hóa trích xuất báo cáo Daily Kênh Neo (Affina) từ dữ liệu Google Sheets & Google Drive, chạy tự động định kỳ qua **GitHub Actions** hoặc chạy thủ công tại máy cục bộ / Jupyter Notebook.

- **Thư mục lưu báo cáo Google Drive**: **[Report_daily_NEO](https://drive.google.com/drive/folders/1uGHy8E3FLPgc-TPDum9u_ELNPU4nUf-u?usp=sharing)** (ID: `1uGHy8E3FLPgc-TPDum9u_ELNPU4nUf-u`).
- **Nguồn Danh sách Nhân sự (DSNS)**: Tự động tải trực tiếp từ **OneDrive** qua Direct Download link (hoạt động 100% kể cả khi tắt máy tính!).

---

## ⏰ 1. Lịch Chạy Báo Cáo Tự Động (GitHub Actions)

Workflow được cấu hình tại `.github/workflows/daily_neo.yml`:

| Ca chạy | Giờ Việt Nam | Cấu hình Cron UTC | Ghi chú |
| :--- | :--- | :--- | :--- |
| **Sáng** | **07:22 AM** | `22 0 * * *` | Báo cáo đầu ngày |
| **Chiều 1** | **16:11 PM** | `11 9 * * *` | Cập nhật giữa ca chiều |
| **Chiều 2** | **17:28 PM** | `28 10 * * *` | Chốt doanh số cuối ngày |

*Ngoài ra, bạn có thể bấm **Run workflow** trong tab Actions trên GitHub bất cứ lúc nào.*

---

## 📑 2. Cấu Trúc Báo Cáo Excel (8 Sheet Tiêu Chuẩn)

Báo cáo Excel xuất ra bao gồm chính xác 8 sheet theo yêu cầu:

1. **`detail`**: 
   - Danh sách chi tiết toàn bộ các hợp đồng / case phát sinh của Kênh Neo trong tháng báo cáo.
   - Bao gồm quy tắc tự động chuẩn hóa: Nếu nhân sự là `PHẠM TRƯỜNG KHÁNH` (`LD4641`) có kênh là `HO` thì tự động chuyển thành `CD` (Kênh Neo).
2. **`tracking AG`**: 
   - Theo dõi tiến độ doanh số và hoạt động của từng Đại lý (AG) Kênh Neo.
   - Thống kê số lượng hợp đồng, tổng doanh số thực thu và phí quy đổi tương ứng theo từng AG.
3. **`tracking SM`**: 
   - Theo dõi hiệu quả quản lý của cấp Trưởng phòng / BDM (SM) và đội ngũ tư vấn viên thuộc quyền.
   - Doanh số nhóm, số lượng hoạt động, và tổng hợp toàn bộ cấp dưới.
4. **`tracking SD`**: 
   - Theo dõi hiệu quả quản lý cấp Giám đốc Vùng / BDD (SD) toàn bộ Kênh Neo.
   - Đánh giá chỉ tiêu doanh số, số lượng SM trực thuộc và tỷ lệ đóng góp của từng vùng.
5. **`tracking CD`**: 
   - Báo cáo tổng hợp dành cho cấp Giám đốc Kênh (CD - PHẠM TRƯỜNG KHÁNH).
   - Chi tiết theo từng nhánh SD thuộc quyền và tổng toàn kênh.
6. **`theo dõi tái tục detail`**: 
   - Bảng chi tiết toàn bộ các hợp đồng đến hạn tái tục trong kỳ của Kênh Neo.
   - Tình trạng tái tục (Đã tái tục / Chưa tái tục), thông tin khách hàng và tư vấn viên phụ trách.
7. **`theo dõi tái tục SM`**: 
   - Thống kê tỷ lệ tái tục và số hợp đồng / doanh số tái tục theo từng Trưởng phòng (SM).
8. **`theo dõi tái tục SD`**: 
   - Thống kê tỷ lệ tái tục và số hợp đồng / doanh số tái tục theo từng Giám đốc Vùng (SD).

---

## 🎨 3. Tính Năng Định Dạng Excel Nổi Bật

- **Màu sắc chuyên nghiệp**: Mỗi bảng được phân bổ mã màu nhận diện thương hiệu rõ ràng (Xanh lam, Xanh ngọc, Tím thạch anh, Cam nhạt,...).
- **Bộ lọc động (AutoFilter)**: Tự động kích hoạt bộ lọc cho tất cả các cột trên toàn bộ 8 sheet.
- **Cố định dòng tiêu đề (Freeze Panes)**: Cố định dòng Header đầu tiên giúp cuộn dữ liệu mượt mà, dễ đối chiếu.
- **Dynamic SUBTOTAL**: Dòng **TỔNG CỘNG** ở chân mỗi bảng sử dụng công thức Excel động:
  ```excel
  =SUBTOTAL(9, [Cột_Bắt_Đầu]:[Cột_Kết_Thúc])
  ```
  Khi người dùng bấm chọn lọc theo bất kỳ SM, SD hay sản phẩm nào, dòng tổng cộng sẽ **tự động tính toán lại chỉ trên các dòng đang hiển thị**.
- **Tự động căn chỉnh**: Tự tính toán độ rộng cột vừa vặn, viền kẻ (border) thanh mảnh, số tiền được định dạng phân cách hàng nghìn rõ ràng (`#,##0`).

---

## ☁️ 4. Nguồn Dữ Liệu Input (Tự Động 100%)

1. **Cấp đơn**: Tải trực tiếp qua Google Sheets API từ Google Sheet private online:  
   `https://docs.google.com/spreadsheets/d/1qc_QhrvpoLLp6w9RkGBEkm8qBO49GJE8oMlwkCdJOsk/edit?usp=sharing`
2. **Danh sách Nhân sự (DSNS)**: Tự động tải từ link OneDrive:  
   `https://1drv.ms/x/c/506a9d11fc30ada1/IQCsopTcUW2nSZJ_dhCCC9nwAb-1Wkmo0xYa5HzEyaIQIVU?e=TFjv1Y`  
   *(Khi người quản lý nhân sự chỉnh sửa file trên OneDrive, GitHub Actions luôn tải bản mới nhất trực tiếp từ Microsoft Cloud kể cả khi bạn tắt máy tính).*
3. **Quy đổi**: Tải từ Google Drive ID `1SDVXT33gHfIKR17x2xdWiO5xgabVXOWH`.
