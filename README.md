# 📊 BÁO CÁO DAILY KÊNH NEO (AFFINA DAILY NEO REPORT)

Hệ thống tự động hóa trích xuất báo cáo Daily Kênh Neo (Affina) từ dữ liệu Google Sheets & Google Drive, chạy tự động định kỳ qua **GitHub Actions** hoặc chạy thủ công tại máy cục bộ / Jupyter Notebook.

Tự động lưu báo cáo vào thư mục Google Drive: **[Report_daily_NEO](https://drive.google.com/drive/folders/1uGHy8E3FLPgc-TPDum9u_ELNPU4nUf-u?usp=sharing)** (ID: `1uGHy8E3FLPgc-TPDum9u_ELNPU4nUf-u`).

---

## 📑 1. Cấu Trúc Báo Cáo Excel (8 Sheet Tiêu Chuẩn)

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

## 🎨 2. Tính Năng Định Dạng Excel Nổi Bật

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

## ☁️ 3. Thư Mục Google Drive Đích & Quyền Upload

- **Thư mục lưu trữ**: `Report_daily_NEO`
- **Link thư mục**: [Google Drive Folder](https://drive.google.com/drive/folders/1uGHy8E3FLPgc-TPDum9u_ELNPU4nUf-u?usp=sharing)
- **Folder ID**: `1uGHy8E3FLPgc-TPDum9u_ELNPU4nUf-u`

> ⚠️ **LƯU Ý QUAN TRỌNG VỀ QUYỀN TẢI LÊN GOOGLE DRIVE:**
> - Nếu thư mục nằm trong **Google Drive cá nhân (@gmail.com)**: Chính sách của Google không cho phép Service Account (`daily-report-bot@...`) tạo file mới trong Drive cá nhân vì Service Account có 0 storage quota (`storageQuotaExceeded`).
> - Để GitHub Actions tự động upload file thành công vào folder cá nhân, bạn chỉ cần dùng **OAuth Refresh Token** (giống hệt dự án `AnLoan` bạn đã làm).
> - Script `get_refresh_token.py` đã được tích hợp sẵn: Chạy `python get_refresh_token.py` trên máy, đăng nhập Gmail 1 lần và dán 3 secret vào GitHub.

---

## ⚙️ 4. Tự Động Hóa Qua GitHub Actions

Hệ thống được cấu hình workflow tại `.github/workflows/daily_neo.yml`:

### Lịch chạy tự động (Schedule):
- **08:00 AM (Giờ VN)**: Cập nhật doanh số đầu ngày.
- **17:30 PM (Giờ VN)**: Tổng hợp báo cáo chốt ca chiều.

### Kích hoạt thủ công (Manual Trigger):
- Vào tab **Actions** trên GitHub Repo -> Chọn workflow **Daily Report NEO** -> Bấm **Run workflow**.
- Có thể nhập tham số tùy chọn:
  - `report_month`: Tháng muốn xuất báo cáo (ví dụ: `2` hoặc để trống lấy tháng hiện tại).
  - `report_year`: Năm muốn xuất báo cáo (ví dụ: `2026` hoặc để trống lấy năm hiện tại).

### Cấu hình GitHub Secrets:
Vào GitHub Repository (**Settings > Secrets and variables > Actions > New repository secret**) và thêm các Secrets:

| Tên Secret | Ý Nghĩa / Cách Lấy |
| :--- | :--- |
| `GOOGLE_SERVICE_ACCOUNT_JSON` | Nội dung JSON file Service Account (Dùng để đọc dữ liệu Cấp đơn, Nhân sự, Quy đổi). |
| `GOOGLE_CLIENT_ID` | Client ID OAuth (đã có sẵn trong file `oauth_credentials.json`). |
| `GOOGLE_CLIENT_SECRET` | Client Secret OAuth (đã có sẵn trong file `oauth_credentials.json`). |
| `GOOGLE_REFRESH_TOKEN` | Refresh Token của tài khoản Gmail sở hữu folder Drive (chạy `python get_refresh_token.py` để lấy). |

---

## 💻 5. Hướng Dẫn Chạy Báo Cáo Tại Máy Cục Bộ (Local)

### Cài đặt môi trường:
```bash
pip install -r requirements.txt
```

### Chạy bằng file Python script:
```bash
python Daily_NEO.py
```

### Chạy bằng Jupyter Notebook:
Mở file `Daily_NEO.ipynb` bằng VS Code, JupyterLab hoặc Google Colab, sau đó bấm **Run All**.
