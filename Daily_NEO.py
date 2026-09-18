# -*- coding: utf-8 -*-
"""
==================================================================================
DAILY REPORT NEO - HỆ THỐNG XUẤT BÁO CÁO DAILY TỰ ĐỘNG KÊNH NEO (AFFINA)
==================================================================================
Tự động chạy hàng ngày qua GitHub Actions hoặc chạy trực tiếp trên máy cục bộ / Colab.

Bao gồm 8 Sheet chuẩn theo yêu cầu:
  1. detail                : Chi tiết toàn bộ các case Kênh Neo phát sinh trong tháng
  2. tracking AG           : Theo dõi tiến độ doanh số & tuyển dụng từng Đại lý (AG)
  3. tracking SM           : Theo dõi hiệu quả quản lý cấp SM (BDM) & đội ngũ
  4. tracking SD           : Theo dõi hiệu quả quản lý cấp SD (BDD) toàn kênh
  5. tracking CD           : Tổng hợp cấp Giám đốc Kênh (CD - PHẠM TRƯỜNG KHÁNH)
  6. theo dõi tái tục detail: Chi tiết toàn bộ các hợp đồng tái tục Kênh Neo
  7. theo dõi tái tục SM   : Tỉ lệ & doanh số tái tục theo từng SM (BDM)
  8. theo dõi tái tục SD   : Tỉ lệ & doanh số tái tục theo từng SD (BDD)

Tính năng nổi bật:
  - 100% tự động tải dữ liệu từ Google Drive & Google Sheets API
  - Động cơ truy vấn DuckDB trong bộ nhớ cực nhanh, không phụ thuộc database ngoài
  - Định dạng Excel chuyên nghiệp (Header tô màu, căn chỉnh, border, auto-fit độ rộng)
  - Cố định dòng tiêu đề (Freeze Panes) & Bật bộ lọc (AutoFilter) tự động cho mọi sheet
  - Công thức TỔNG CỘNG linh hoạt: Dùng =SUBTOTAL(9, ...) tự động cập nhật khi filter
  - Tự động upload báo cáo hoàn chỉnh lên Google Drive (Thư mục Report_NEO)
==================================================================================
"""

import os
import sys
import io
import re
import unicodedata
import warnings
from datetime import datetime
from pathlib import Path

import pandas as pd
import duckdb
import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter

from google.oauth2 import service_account
from google.oauth2.credentials import Credentials
from google.auth.transport.requests import Request
from googleapiclient.discovery import build
from googleapiclient.http import MediaIoBaseDownload, MediaFileUpload

warnings.filterwarnings('ignore')
if hasattr(sys.stdout, 'reconfigure'):
    sys.stdout.reconfigure(encoding='utf-8', line_buffering=True)

# ============================================================================
# CẤU HÌNH THỜI GIAN BÁO CÁO (Mặc định: Tháng & Năm hiện tại)
# ============================================================================
_month_env = str(os.environ.get('REPORT_MONTH', '')).strip()
_year_env  = str(os.environ.get('REPORT_YEAR', '')).strip()
REPORT_MONTH = int(_month_env) if _month_env else datetime.now().month
REPORT_YEAR  = int(_year_env) if _year_env else datetime.now().year


# Thư mục làm việc tạm thời
WORK_DIR = os.environ.get('AFFINA_WORK_DIR', os.path.join(os.getcwd(), 'temp_affina_neo'))
OUTPUT_DIR = os.path.join(os.getcwd(), 'output')
os.makedirs(WORK_DIR, exist_ok=True)
os.makedirs(OUTPUT_DIR, exist_ok=True)

# Google Drive File IDs & Sheet IDs
SHEET_CAPDON_ID = '1qc_QhrvpoLLp6w9RkGBEkm8qBO49GJE8oMlwkCdJOsk'
DSNS_FILE_ID    = '1_Mr_wnoJ2zBQJ0Pb9xFPwAH8IsIQiyuk'  # DSNS CTV sale Affina FINAL V2.xlsx
QUYDOI_FILE_ID  = '1SDVXT33gHfIKR17x2xdWiO5xgabVXOWH'  # 26_02_04_sửa ngày_quy_doi_all.xlsx
DRIVE_FOLDER_ID = os.environ.get('DRIVE_FOLDER_ID', '1uGHy8E3FLPgc-TPDum9u_ELNPU4nUf-u')  # Target Drive Folder: Report_daily_NEO


# Đường dẫn file nội bộ (fallback khi chạy offline)
LOCAL_SA_KEY_PATHS = [
    'google_service_account.json',
    r'C:\Users\ADMIN\Desktop\AFFINA\CODE\google_service_account.json',
    os.path.join(os.path.dirname(__file__), 'google_service_account.json')
]


# ============================================================================
# PHẦN 1: XÁC THỰC GOOGLE DRIVE & SHEETS API
# ============================================================================
def init_google_services():
    """
    Khởi tạo kết nối Google Drive & Sheets API.
    Hỗ trợ linh hoạt:
    1. Service Account JSON từ biến môi trường GOOGLE_SERVICE_ACCOUNT_JSON (GitHub Actions)
    2. File Service Account JSON cục bộ (google_service_account.json)
    3. OAuth Refresh Token (GOOGLE_CLIENT_ID, GOOGLE_CLIENT_SECRET, GOOGLE_REFRESH_TOKEN)
    """
    drive_service = None
    sheets_service = None
    creds = None
    scopes = [
        'https://www.googleapis.com/auth/drive',
        'https://www.googleapis.com/auth/spreadsheets.readonly'
    ]

    # Cách 1: OAuth Credentials (GOOGLE_CLIENT_ID + GOOGLE_CLIENT_SECRET + GOOGLE_REFRESH_TOKEN)
    # Ưu tiên OAuth vì tài khoản cá nhân có dung lượng Drive để tạo & upload file mới
    client_id = os.environ.get('GOOGLE_CLIENT_ID')
    client_secret = os.environ.get('GOOGLE_CLIENT_SECRET')
    refresh_token = os.environ.get('GOOGLE_REFRESH_TOKEN')
    if client_id and client_secret and refresh_token:
        try:
            creds = Credentials(
                token=None,
                refresh_token=refresh_token,
                client_id=client_id,
                client_secret=client_secret,
                token_uri='https://oauth2.googleapis.com/token',
                scopes=scopes
            )
            creds.refresh(Request())
            print("  🔑 Xác thực thành công qua OAuth Refresh Token (tài khoản cá nhân có quota upload)!")
        except Exception as e:
            print(f"  ⚠️ Lỗi xác thực OAuth: {e}")

    # Cách 2: Chuỗi JSON Service Account từ Secret GitHub Actions
    if not creds:
        sa_json_str = os.environ.get('GOOGLE_SERVICE_ACCOUNT_JSON')
        if sa_json_str and sa_json_str.strip():
            try:
                import json
                info = json.loads(sa_json_str)
                creds = service_account.Credentials.from_service_account_info(info, scopes=scopes)
                print("  🔑 Xác thực thành công qua GOOGLE_SERVICE_ACCOUNT_JSON (GitHub Secrets)!")
            except Exception as e:
                print(f"  ⚠️ Lỗi parse GOOGLE_SERVICE_ACCOUNT_JSON: {e}")

    # Cách 3: File Service Account cục bộ
    if not creds:
        sa_path = os.environ.get('GOOGLE_SA_KEY_PATH')
        candidates = [sa_path] if sa_path else LOCAL_SA_KEY_PATHS
        for p in candidates:
            if p and os.path.exists(p):
                try:
                    creds = service_account.Credentials.from_service_account_file(p, scopes=scopes)
                    print(f"  🔑 Xác thực thành công qua file Service Account: {p}")
                    break
                except Exception as e:
                    print(f"  ⚠️ Không thể nạp key từ {p}: {e}")


    if creds:
        drive_service = build('drive', 'v3', credentials=creds)
        sheets_service = build('sheets', 'v4', credentials=creds)
    else:
        print("  ⚠️ Không tìm thấy thông tin xác thực Google API. Sẽ ưu tiên dùng file local nếu có.")

    return drive_service, sheets_service


# ============================================================================
# PHẦN 2: TẢI VÀ NẠP DỮ LIỆU TỪ GOOGLE DRIVE / SHEETS
# ============================================================================
def download_drive_file(drive_service, file_id, local_path):
    """Tải file từ Google Drive về máy cục bộ / runner."""
    req = drive_service.files().get_media(fileId=file_id)
    fh = io.BytesIO()
    downloader = MediaIoBaseDownload(fh, req)
    done = False
    while not done:
        _, done = downloader.next_chunk()
    fh.seek(0)
    with open(local_path, 'wb') as f:
        f.write(fh.read())
    print(f"  ✅ Đã tải file ID [{file_id}] -> {os.path.basename(local_path)}")


def export_google_sheet_data(sheets_service, spreadsheet_id, target_sheets, local_excel_path):
    """
    Xuất các sheet cần thiết từ Google Sheet về file Excel nội bộ bằng Sheets API.
    Bỏ qua giới hạn 10MB của lệnh Drive export thông thường.
    """
    print(f"  ⏳ Đang tải {len(target_sheets)} sheet Cấp đơn từ Google Sheet...")
    with pd.ExcelWriter(local_excel_path, engine='openpyxl') as writer:
        sheets_written = 0
        for s_name in target_sheets:
            try:
                result = sheets_service.spreadsheets().values().get(
                    spreadsheetId=spreadsheet_id,
                    range=s_name,
                    valueRenderOption='FORMATTED_VALUE',
                    dateTimeRenderOption='FORMATTED_STRING'
                ).execute()
                values = result.get('values', [])
                if values:
                    max_cols = max(len(row) for row in values)
                    padded = [row + [None] * (max_cols - len(row)) for row in values]
                    header = [str(c) if c is not None else '' for c in padded[0]]
                    seen = {}
                    unique_header = []
                    for col in header:
                        if col in seen:
                            seen[col] += 1
                            unique_header.append(f"{col}_{seen[col]}")
                        else:
                            seen[col] = 0
                            unique_header.append(col)
                    data_rows = padded[1:] if len(padded) > 1 else []
                    df = pd.DataFrame(data_rows, columns=unique_header)
                else:
                    df = pd.DataFrame()
                
                safe_name = re.sub(r'[\\/?*\[\]:]', '', s_name)[:31]
                df.to_excel(writer, sheet_name=safe_name, index=False)
                sheets_written += 1
                print(f"     ✓ Đã tải sheet '{safe_name}': {len(df)} dòng")
            except Exception as e:
                print(f"     ⚠️ Không thể tải sheet '{s_name}': {e}")
        if sheets_written == 0:
            pd.DataFrame().to_excel(writer, sheet_name='Sheet1', index=False)
    print(f"  ✅ Đã xuất {sheets_written} sheet Cấp đơn thành công.")


def upload_to_drive(drive_service, local_path, folder_id=DRIVE_FOLDER_ID, target_filename=None):
    """
    Upload file kết quả lên Google Drive.
    Mặc định lưu vào thư mục Report_daily_NEO (ID: 1uGHy8E3FLPgc-TPDum9u_ELNPU4nUf-u).
    """
    if not drive_service:
        print("  ⚠️ Bỏ qua upload Google Drive (chưa cấu hình credentials).")
        return
    
    if not target_filename:
        target_filename = os.path.basename(local_path)
        
    print(f"\n☁️ Đang upload báo cáo lên Google Drive (Folder ID: {folder_id}): {target_filename}...")
    try:
        # 1. Kiểm tra file cũ trong thư mục để ghi đè (update) hoặc tạo mới
        q_file = f"name = '{target_filename}' and '{folder_id}' in parents and trashed = false"
        res_file = drive_service.files().list(
            q=q_file,
            fields="files(id, name)",
            supportsAllDrives=True,
            includeItemsFromAllDrives=True
        ).execute()
        existing = res_file.get('files', [])
        
        media = MediaFileUpload(
            local_path,
            mimetype='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
            resumable=True
        )
        if existing:
            file_id = existing[0]['id']
            drive_service.files().update(
                fileId=file_id,
                media_body=media,
                supportsAllDrives=True
            ).execute()
            print(f"  ✅ Đã cập nhật (ghi đè) file trên Drive: {target_filename} (ID: {file_id})")
        else:
            meta = {'name': target_filename, 'parents': [folder_id]}
            new_file = drive_service.files().create(
                body=meta,
                media_body=media,
                fields='id',
                supportsAllDrives=True
            ).execute()
            print(f"  ✅ Đã tải file mới lên Drive: {target_filename} (ID: {new_file['id']})")
    except Exception as e:
        err_msg = str(e)
        print(f"  ❌ Lỗi khi upload Google Drive: {e}")
        if 'storageQuotaExceeded' in err_msg or 'quota' in err_msg.lower():
            print("\n" + "=" * 75)
            print("⚠️ LƯU Ý VỀ DUNG LƯỢNG GOOGLE DRIVE (PERSONAL DRIVE QUOTA):")
            print("Google Service Account không có dung lượng lưu trữ trên Google Drive cá nhân (@gmail.com).")
            print("👉 Để tự động upload thành công vào folder cá nhân trên GitHub Actions:")
            print("   Vui lòng thêm 3 Secret OAuth giống dự án AnLoan vào GitHub:")
            print("   1. GOOGLE_CLIENT_ID")
            print("   2. GOOGLE_CLIENT_SECRET")
            print("   3. GOOGLE_REFRESH_TOKEN (lấy từ script get_refresh_token.py)")
            print("   (Hoặc nếu dùng Google Workspace, chuyển folder vào Shared Drive - Bộ nhớ dùng chung).")
            print("=" * 75 + "\n")



# ============================================================================
# PHẦN 3: HÀM LÀM SẠCH VÀ CHUẨN HÓA DỮ LIỆU
# ============================================================================
def standardize_date_format(date_value):
    """Chuẩn hóa mọi định dạng ngày tháng về YYYY-MM-DD."""
    if pd.isna(date_value): return None
    date_str = str(date_value).strip()
    if date_str in ['None', 'nan', '', 'NaT']: return None
    try:
        match = re.match(r'^(\d{4})-(\d{1,2})-(\d{1,2})', date_str)
        if match:
            year, p1, p2 = match.groups()
            if int(p1) > 12 and int(p2) <= 12:
                try: return pd.to_datetime(f"{year}-{p2.zfill(2)}-{p1.zfill(2)}", format='%Y-%m-%d').strftime('%Y-%m-%d')
                except: return date_str
            else: return f"{year}-{p1.zfill(2)}-{p2.zfill(2)}"
        if '/' in date_str:
            parts = date_str.split('/')
            if len(parts) == 3:
                p1, p2, p3 = [p.strip() for p in parts]
                if len(p3) >= 3 or (p3.isdigit() and int(p3) > 31): d, m, y = p1, p2, p3
                elif len(p1) >= 3 or (p1.isdigit() and int(p1) > 31): y, m, d = p1, p2, p3
                else: d, m, y = p1, p2, p3
                y = f"20{y}" if len(y) == 2 and int(y) <= 30 else f"19{y}" if len(y) == 2 else f"20{y[1:]}" if len(y) == 3 else f"202{y}" if len(y) == 1 else y
                try: return pd.to_datetime(f"{y.zfill(4)}-{m.zfill(2)}-{d.zfill(2)}", format='%Y-%m-%d').strftime('%Y-%m-%d')
                except:
                    try: return pd.to_datetime(f"{y.zfill(4)}-{d.zfill(2)}-{m.zfill(2)}", format='%Y-%m-%d').strftime('%Y-%m-%d')
                    except: return date_str
        parsed = pd.to_datetime(date_str, dayfirst=True, errors='coerce')
        if pd.notna(parsed) and parsed.year > 1900: return parsed.strftime('%Y-%m-%d')
    except: return date_str
    return date_str


def clean_currency_column(series):
    """Làm sạch định dạng tiền tệ VN (chấm/phẩy) trả về Int64."""
    series_str = series.astype(str).str.replace(r'[^\d,\.]', '', regex=True).str.strip()
    mask_vn_format = series_str.str.contains(r'\.') & series_str.str.contains(',')
    series_str[mask_vn_format] = series_str[mask_vn_format].str.replace('.', '', regex=False).str.replace(',', '.', regex=False)
    mask_comma_only = (~series_str.str.contains(r'\.')) & series_str.str.contains(',')
    series_str[mask_comma_only] = series_str[mask_comma_only].str.replace(',', '.', regex=False)
    mask_multi_dot = series_str.str.count(r'\.') > 1
    if mask_multi_dot.any():
        def fix_multi_dot(x):
            if not isinstance(x, str): return x
            parts = x.split('.')
            return ''.join(parts[:-1]) + '.' + parts[-1] if len(parts[-1]) <= 2 and all(p.isdigit() for p in parts[:-1]) else x.replace('.', '')
        series_str[mask_multi_dot] = series_str[mask_multi_dot].apply(fix_multi_dot)
    mask_single_dot = series_str.str.count(r'\.') == 1
    series_str[mask_single_dot] = series_str[mask_single_dot].apply(lambda x: x if isinstance(x, str) and len(x.split('.')[1]) <= 2 else x.replace('.', '') if isinstance(x, str) else x)
    num = pd.to_numeric(series_str, errors='coerce')
    return num.apply(lambda x: x * 1000 if pd.notna(x) and x < 1000 else x).round(0).astype('Int64')


def clean_whitespace(df):
    """Loại bỏ ký tự xuống dòng, khoảng trắng thừa."""
    for col in df.columns:
        if df[col].dtype == 'object':
            df[col] = df[col].str.replace(r'[\n\r\t]+', ' ', regex=True).str.replace(r'\s{2,}', ' ', regex=True).str.strip()
    return df


def clean_contract(value):
    """Chuẩn hóa ký hiệu số hợp đồng."""
    if pd.isna(value): return value
    return re.sub(r'\s+', ' ', re.sub(r'[⁄/]', '/', re.sub(r'[–—−-]', '-', unicodedata.normalize('NFC', str(value))))).strip()


def convert_channel(val):
    """Quy chuẩn tên Kênh."""
    val_str = str(val).strip()
    if val_str in ["Core Agency", "core Agency", "Standard"]: return "Core Agency"
    if val_str in ["CTV_TSA (TSA 2)", "CTV_TSA (TSA2)"]: return "CTV_TSA (TSA 2)"
    if val_str in ["Elite", "H.O", "Neo", "TSA"]: return val_str
    if "NEO" in val_str.upper(): return "Neo"
    return "Core Agency"


# ============================================================================
# PHẦN 4: NẠP VÀ GỘP DỮ LIỆU CẤP ĐƠN
# ============================================================================
def load_and_clean_all_data(capdon_excel_path, nhansu_excel_path, quydoi_excel_path):
    """
    Nạp toàn bộ dữ liệu Cấp đơn (7 sheet), Nhân sự và Tỷ lệ quy đổi vào RAM.
    Tự động áp dụng quy tắc đặc biệt cho anh PHẠM TRƯỜNG KHÁNH (CD Kênh Neo).
    """
    print("\n⏳ Đang xử lý file Quy Đổi...")
    df_quydoi = pd.read_excel(quydoi_excel_path, sheet_name='Updating')
    print(f"   ✓ Xong Quy đổi: {len(df_quydoi)} dòng")

    print("⏳ Đang xử lý file Nhân Sự...")
    df_nhansu = pd.read_excel(nhansu_excel_path, dtype=str, sheet_name='DSNS AGENCY 2025').dropna(subset=['Họ tên'])
    df_nhansu = clean_whitespace(df_nhansu)
    df_nhansu['Channel'] = df_nhansu['Channal'].apply(convert_channel)
    for c in ['Thời gian bắt đầu', 'Ngày hiệu lực chức danh', 'Ngày Sinh']:
        if c in df_nhansu.columns: df_nhansu[c] = df_nhansu[c].apply(standardize_date_format)
    df_nhansu['Điện thoại'] = df_nhansu['Điện thoại'].astype(str).str.lstrip('0')
    df_nhansu['Người giới thiệu'] = df_nhansu['Người giới thiệu'].astype(str).str.lstrip('0')
    df_nhansu['Code UM'] = df_nhansu['Code UM'].astype(str).str.lstrip('0')

    # Đảm bảo PHẠM TRƯỜNG KHÁNH là CD Kênh Neo
    khanh_ns_mask = df_nhansu['Họ tên'].astype(str).str.upper().str.contains('PHẠM TRƯỜNG KHÁNH', na=False)
    if khanh_ns_mask.any():
        df_nhansu.loc[khanh_ns_mask, 'Channel'] = 'Neo'
        df_nhansu.loc[khanh_ns_mask, 'Chức danh'] = 'CD'
        df_nhansu.loc[khanh_ns_mask, 'QUẢN LÝ CẤP 2 (BDD)'] = 'PHẠM TRƯỜNG KHÁNH'
        df_nhansu.loc[khanh_ns_mask, 'QUẢN LÝ CẤP 1 (BDM)'] = 'PHẠM TRƯỜNG KHÁNH'
    print(f"   ✓ Xong Nhân sự: {len(df_nhansu)} dòng")

    print("⏳ Đang xử lý file Cấp Đơn (7 sheet)...")
    # 1. BHSK
    df_BHSK_raw = pd.read_excel(capdon_excel_path, sheet_name='Sức khỏe', header=None)
    hdr_sk = 2
    for r in range(min(5, len(df_BHSK_raw))):
        row_text = " ".join(str(c) for c in df_BHSK_raw.iloc[r] if pd.notna(c)).lower()
        if 'stt' in row_text and ('ngày' in row_text or 'người được' in row_text):
            hdr_sk = r
            break

    bhsk_cols = ['Ngày update', 'STT', 'Tên Người được BH', 'Ngày Sinh', 'Giới tính', 'Email', 'Số hộ chiếu', 'CMND', 'ĐỊA CHỈ', 'Tên', 'Quan hệ', 'Ngày sinh người mua Bảo hiểm', 'Số CMND_CCCD NMBH', 'Số điện thoại NMBH', 'Địa chỉ NMBH', 'Email NMBH', 'Chương trình bảo hiểm', 'Ngoại trú', 'Nha khoa', 'Thai sản', 'Topup', 'Phí bảo hiểm', 'Giảm phí_refund', 'Giảm phí_deduct', 'Tổng giảm phí', 'Số tiền thanh toán', 'Ngày thanh toán', 'Ngày bắt đầu', 'Ngày kết thúc', 'Số Giấy Chứng Nhận', 'Số hợp đồng', 'Thông tin xuất hoá đơn', 'Phí điều chỉnh ( Nếu có)', 'Giảm phí ( Nếu có)', 'Nguyên nhân (Nếu có)', 'Lead ID (nếu có)', 'Phone trên lead', 'Code sale', 'Phone Khách hàng', 'Tên liên hệ', 'Người giới thiệu', 'Hình thức thanh toán', 'Note', 'Đối tác nhà bảo hiểm', 'Sản phẩm', 'Channel', 'Mã hợp đồng Affina', 'Đã gửi mail cho khách', 'Nộp claim', 'Ủy quyền bồi thường (Nếu có)']
    df_BHSK = df_BHSK_raw.iloc[hdr_sk+1:].copy().reset_index(drop=True)
    df_BHSK.columns = bhsk_cols[:len(df_BHSK.columns)] + [f'C_{i}' for i in range(len(df_BHSK.columns) - len(bhsk_cols))] if len(df_BHSK.columns) > len(bhsk_cols) else bhsk_cols[:len(df_BHSK.columns)]
    df_BHSK['Loại bảo hiểm'] = 'BHSK'
    df_BHSK = df_BHSK.rename(columns={'Tên': 'Tên NMBH', 'Giới tính': 'Giới tính NNBH', 'Ngày sinh người mua Bảo hiểm': 'Ngày sinh NMBH', 'Ngày Sinh': 'Ngày Sinh NNBH', 'Chương trình bảo hiểm': 'Gói bảo hiểm', 'CMND': 'CCCD'})
    df_BHSK = df_BHSK[[c for c in ['STT', 'Ngày thanh toán', 'Code sale', 'Sản phẩm', 'Đối tác nhà bảo hiểm', 'Số tiền thanh toán', 'Channel', 'Tên Người được BH', 'Ngày Sinh NNBH', 'Phone Khách hàng', 'Ngày bắt đầu', 'Ngày kết thúc', 'Loại bảo hiểm', 'Số hợp đồng', 'Tên NMBH', 'Ngày sinh NMBH', 'Gói bảo hiểm'] if c in df_BHSK.columns]]

    # 2. BHXM
    df_BHXM_raw = pd.read_excel(capdon_excel_path, sheet_name='Thông tin cấp Bảo hiểm xe máy', header=None)
    hdr_xm = 1
    for r in range(min(5, len(df_BHXM_raw))):
        row_text = " ".join(str(c) for c in df_BHXM_raw.iloc[r] if pd.notna(c)).lower()
        if 'biển số xe' in row_text or 'khách hàng' in row_text:
            hdr_xm = r
            break
    df_BHXM = df_BHXM_raw.iloc[hdr_xm+1:].copy().dropna(subset=[0]).reset_index(drop=True)
    bhxm_cols = ['Ngày update', 'STT', 'BIỂN SỐ XE', 'SỐ KHUNG', 'SỐ MÁY', 'TÊN KHÁCH HÀNG', 'PHÍ BẢO HIỂM TNDS BẮT BUỘC', 'PHÍ BẢO HIỂM TAI NẠN NNTX', 'SỐ NĂM', 'TỔNG PHÍ BẢO HIỂM', 'NGÀY CẤP ĐƠN', 'NGÀY BẮT ĐẦU', 'NGÀY KẾT THÚC', 'LOẠI XE', 'NHÃN HIỆU XE', 'SỐ ĐIẸN THOẠI', 'Email', 'Chương trình', 'Code sale', 'Hình thức thanh toán', 'Đối tác nhà bảo hiểm', 'Sản phẩm', 'Channel', 'Số hợp đồng', 'Note']
    df_BHXM.columns = bhxm_cols[:len(df_BHXM.columns)]
    df_BHXM_done = df_BHXM[[c for c in ['Ngày update', 'Code sale', 'Chương trình', 'Đối tác nhà bảo hiểm', 'TỔNG PHÍ BẢO HIỂM', 'Channel', 'TÊN KHÁCH HÀNG', 'NGÀY BẮT ĐẦU', 'NGÀY KẾT THÚC', 'Số hợp đồng', 'BIỂN SỐ XE'] if c in df_BHXM.columns]].copy()
    if not df_BHXM_done.empty:
        df_BHXM_done.columns = ['Ngày thanh toán', 'Code sale', 'Sản phẩm', 'Đối tác nhà bảo hiểm', 'Số tiền thanh toán', 'Channel', 'Tên Người được BH', 'Ngày bắt đầu', 'Ngày kết thúc', 'Số hợp đồng', 'Biển số xe']
        df_BHXM_done['Loại bảo hiểm'] = 'BHXM'

    # 3. BHYT
    df_BHYT_raw = pd.read_excel(capdon_excel_path, sheet_name='BHYTBHXH')
    df_BHYT = df_BHYT_raw.copy()
    df_BHYT['Loại bảo hiểm'] = 'BHYT/BHXH'
    df_BHYT = df_BHYT.rename(columns={'Họ tên NĐBH': 'Tên Người được BH', 'Họ tên BMBH': 'Tên NMBH', 'Ngày sinh': 'Ngày Sinh NNBH', 'Code sales': 'Code sale', 'Đối tác NBH': 'Đối tác nhà bảo hiểm', 'Mã tờ khai': 'Số hợp đồng', 'Phí Bảo hiểm': 'Số tiền thanh toán', 'Ngày duyệt': 'Ngày thanh toán', 'Ngày thanh toán': 'Ngày thanh toán (real)'})
    df_BHYT_done = df_BHYT[[c for c in ['STT', 'Ngày thanh toán', 'Code sale', 'Sản phẩm', 'Đối tác nhà bảo hiểm', 'Số tiền thanh toán', 'Channel', 'Tên Người được BH', 'Ngày Sinh NNBH', 'Ngày bắt đầu', 'Ngày kết thúc', 'Loại bảo hiểm', 'Số hợp đồng', 'Tên NMBH', 'Ngày sinh NMBH'] if c in df_BHYT.columns]].copy()

    # 4. BHOTO
    df_BHOTO = pd.read_excel(capdon_excel_path, sheet_name='Bao hiem oto', dtype={"Số tiền": str})
    df_BHOTO_done = df_BHOTO[['Ngày thanh toán', 'Code sale', 'Chương trình', 'Đối tác nhà bảo hiểm', 'Số tiền', 'Channel', 'Tên khách hàng', 'Số GCN', 'Biển số', 'Ngày bắt đầu hiệu lực', 'Ngày kết thúc hiệu lực']].copy()
    df_BHOTO_done.columns = ['Ngày thanh toán', 'Code sale', 'Sản phẩm', 'Đối tác nhà bảo hiểm', 'Số tiền thanh toán', 'Channel', 'Tên Người được BH', 'Số hợp đồng', 'Biển số xe', 'Ngày bắt đầu', 'Ngày kết thúc']
    df_BHOTO_done['Loại bảo hiểm'] = 'BHOTO'

    # 5. Du lịch
    df_BHDL = pd.read_excel(capdon_excel_path, sheet_name='Du lịch', dtype={"Phí bảo hiểm": str}).rename(columns={'Họ Và Tên': 'Tên Người được BH'})
    df_BHDL_done = df_BHDL[['Ngày thanh toán', 'Tên sale', 'Sản phẩm', 'Đối tác nhà BH', 'Phí bảo hiểm', 'Channel', 'Tên Người được BH', 'Số hợp đồng']].copy()
    df_BHDL_done.columns = ['Ngày thanh toán', 'Code sale', 'Sản phẩm', 'Đối tác nhà bảo hiểm', 'Số tiền thanh toán', 'Channel', 'Tên Người được BH', 'Số hợp đồng']
    df_BHDL_done['Loại bảo hiểm'] = 'BHDL'

    # 6. Trách nhiệm sản phẩm
    df_TNSP = pd.read_excel(capdon_excel_path, sheet_name='Trách nhiệm sản phẩm', dtype={"Phí Bảo hiểm": str}).rename(columns={'Nhà Bảo hiểm': 'Đối tác nhà bảo hiểm', 'Phí Bảo hiểm': 'Số tiền thanh toán'})
    df_TNSP_done = df_TNSP[['Ngày thanh toán', 'Code sale', 'Sản phẩm', 'Đối tác nhà bảo hiểm', 'Số tiền thanh toán', 'Channel']].copy()
    df_TNSP_done['Loại bảo hiểm'] = 'TNDS'

    # 7. Bảo hiểm rủi ro
    df_BHRR = pd.read_excel(capdon_excel_path, sheet_name='Bảo hiểm rủi ro', dtype={"Số tiền thanh toán": str}).rename(columns={'Tên khách hàng': 'Tên Người được BH', 'Mã hợp đồng': 'Số hợp đồng'})
    df_BHRR_done = df_BHRR[['Ngày thanh toán', 'Code sale', 'Sản phẩm', 'Đối tác nhà bảo hiểm', 'Số tiền thanh toán', 'Channel', 'Tên Người được BH', 'Số hợp đồng', 'Ngày bắt đầu', 'Ngày kết thúc']].copy()
    df_BHRR_done['Loại bảo hiểm'] = 'BHRR'

    # Gộp tất cả 7 sheet
    df_union = pd.concat([df_BHSK, df_BHOTO_done, df_BHXM_done, df_BHDL_done, df_TNSP_done, df_BHRR_done, df_BHYT_done], axis=0, ignore_index=True)
    df_union["Số hợp đồng"] = df_union["Số hợp đồng"].apply(clean_contract)
    for col in ['Ngày thanh toán', 'Ngày bắt đầu', 'Ngày kết thúc', 'Ngày Sinh NNBH', 'Ngày sinh NMBH']:
        if col in df_union.columns: df_union[col] = df_union[col].apply(standardize_date_format)
    df_union['Số tiền thanh toán'] = clean_currency_column(df_union['Số tiền thanh toán'])
    df_union = clean_whitespace(df_union)

    # Quy tắc đặc biệt: Nếu Code sale chứa 'PHẠM TRƯỜNG KHÁNH' hoặc 'LD4641' và Channel là 'HO' -> tự động chuyển Channel thành 'CD'
    khanh_mask = (
        df_union['Code sale'].astype(str).str.upper().str.contains('PHẠM TRƯỜNG KHÁNH', na=False) |
        df_union['Code sale'].astype(str).str.upper().str.contains('LD4641', na=False) |
        (df_union['Code sale'].astype(str).str.lstrip('0') == '903926873')
    ) & (df_union['Channel'].astype(str).str.upper().isin(['HO', 'H.O']))
    khanh_switched = khanh_mask.sum()
    if khanh_switched > 0:
        df_union.loc[khanh_mask, 'Channel'] = 'CD'
        print(f"   ✓ Đã tự động chuyển {khanh_switched} case của PHẠM TRƯỜNG KHÁNH từ 'HO' sang 'CD' (Kênh NEO)!")

    print(f"   ✓ Xong Cấp đơn tổng hợp: {len(df_union)} dòng")
    return df_union, df_nhansu, df_quydoi


# ============================================================================
# PHẦN 5: ĐỘNG CƠ SQL DUCKDB - TÍNH TOÁN 8 SHEET BÁO CÁO
# ============================================================================
def execute_all_reports(df_union, df_nhansu, df_quydoi, report_year, report_month):
    """
    Truy vấn và tính toán toàn bộ 8 DataFrame cho báo cáo Daily Kênh Neo.
    """
    print(f"\n🚀 Đang khởi động DuckDB truy vấn dữ liệu Tháng {report_month}/{report_year}...")
    con = duckdb.connect()

    con.register('df_capdon', df_union)
    con.register('df_union', df_union)
    con.register('df_quydoi', df_quydoi)
    con.register('df_ns', df_nhansu)
    con.register('df_nhansu', df_nhansu)

    # ------------------------------------------------------------------------
    # Sheet 1: DETAIL
    # ------------------------------------------------------------------------
    q_detail = f"""
    WITH t1 AS (
        SELECT dnsa.*, TRY_CAST(uadcd."Ngày thanh toán" AS DATE) as "Ngày thanh toán", uadcd."Code sale", uadcd."Sản phẩm", uadcd."Đối tác nhà bảo hiểm", uadcd."Số tiền thanh toán", upper(uadcd."Channel") as "Channel Sales", uadcd."Loại bảo hiểm", uadcd."Tên Người được BH", uadcd."Tên NMBH", uadcd."Số hợp đồng"
        FROM df_nhansu dnsa JOIN df_capdon uadcd ON LTRIM(dnsa."Điện thoại", '0') = LTRIM(uadcd."Code sale", '0')
    ),
    detail AS (
        SELECT "Code", "Họ tên", "Code UM", "Chức danh", "Channel", "Ngày thanh toán", "Sản phẩm", "Đối tác nhà bảo hiểm", CAST("Số tiền thanh toán" AS DOUBLE) as "Số tiền thanh toán", "QUẢN LÝ CẤP 1 (BDM)" as "QUẢN LÝ CẤP 1 (SM)", "QUẢN LÝ CẤP 2 (BDD)" as "QUẢN LÝ CẤP 2 (SD)", "Channel Sales", "Thời gian bắt đầu", "Loại bảo hiểm", "Tên Người được BH", "Tên NMBH", "Số hợp đồng"
        FROM t1 WHERE "Channel" = 'Neo' AND "Ngày thanh toán" IS NOT NULL AND extract(year from "Ngày thanh toán") = {report_year} AND extract(month from "Ngày thanh toán") = {report_month}
    ),
    AG1 AS (
        SELECT d."Code", d."Họ tên", d."Chức danh", d."Code UM", d."Channel", d."Ngày thanh toán", d."Sản phẩm", d."Đối tác nhà bảo hiểm", d."QUẢN LÝ CẤP 1 (SM)",
            CASE WHEN d."QUẢN LÝ CẤP 1 (SM)" IS NULL THEN d."QUẢN LÝ CẤP 2 (SD)" ELSE d."QUẢN LÝ CẤP 1 (SM)" END as transfer_when_missing_SM,
            d."QUẢN LÝ CẤP 2 (SD)", d."Channel Sales", d."Thời gian bắt đầu", d."Loại bảo hiểm", d."Số tiền thanh toán", d."Tên Người được BH", d."Tên NMBH", d."Số hợp đồng",
            CASE
                WHEN d."Sản phẩm" LIKE '%Trách nhiệm Dân sự%' OR d."Sản phẩm" LIKE '%VCOTO%' OR d."Sản phẩm" LIKE '%vật chất ô tô%' OR d."Sản phẩm" LIKE '%Lái phụ xe và người ngồi trên xe%' OR d."Loại bảo hiểm" = 'BHRR' THEN ROUND(d."Số tiền thanh toán" / 1.1, 0)
                WHEN d."Sản phẩm" = 'Trách nhiệm sản phẩm' THEN ROUND(d."Số tiền thanh toán" / 1.05, 0)
                ELSE ROUND(d."Số tiền thanh toán", 0)
            END as "Doanh thu trước thuế",
            qd_main."rate_bonus", qd_main."RMM_OR" AS "RMM_OR_rate", qd_main."RMD_OR" AS "RMD_OR_rate",
            CASE WHEN extract(month from d."Ngày thanh toán") BETWEEN 2 AND 6 AND qd_main."sp_contest_neo" = 1 THEN qd_main."RMM_IO" ELSE NULL END as "RMM_IO_rate",
            CASE WHEN extract(month from d."Ngày thanh toán") BETWEEN 2 AND 6 AND qd_main."sp_contest_neo" = 1 THEN qd_main."RMD_IO" ELSE NULL END as "RMD_IO_rate"
        FROM detail d LEFT JOIN df_quydoi qd_main ON UPPER(TRIM(qd_main."provider")) = UPPER(TRIM(d."Đối tác nhà bảo hiểm")) AND UPPER(qd_main."product") = UPPER(d."Sản phẩm") AND d."Ngày thanh toán" >= TRY_CAST(qd_main."Effective_Date" AS DATE) AND d."Ngày thanh toán" <= TRY_CAST(qd_main."Valid_to" AS DATE)
    ),
    AG2 AS (
        SELECT AG1.*, (CAST(AG1."RMM_OR_rate" AS DOUBLE) * AG1."Doanh thu trước thuế") as "SM_OR", (CAST(AG1."RMD_OR_rate" AS DOUBLE) * AG1."Doanh thu trước thuế") as "SD_OR", (CAST(AG1."RMM_IO_rate" AS DOUBLE) * AG1."Doanh thu trước thuế") as "SM_IO", (CAST(AG1."RMD_IO_rate" AS DOUBLE) * AG1."Doanh thu trước thuế") as "SD_IO", ROUND(CAST(AG1."rate_bonus" AS DOUBLE) * AG1."Doanh thu trước thuế", 0) as "EST_bonus"
        FROM AG1
    )
    SELECT DISTINCT "Code", "Họ tên", "Chức danh", "Channel", "Ngày thanh toán", "Sản phẩm", "Đối tác nhà bảo hiểm", "Loại bảo hiểm", upper("Tên Người được BH") as "Tên Người được BH", upper("Tên NMBH") as "Tên Bên mua BH", "Số hợp đồng", "QUẢN LÝ CẤP 1 (SM)", "QUẢN LÝ CẤP 2 (SD)", "Thời gian bắt đầu", "Số tiền thanh toán", "Doanh thu trước thuế", "SM_OR", "SD_OR", "EST_bonus" as "AG_bonus", "SM_IO", "SD_IO"
    FROM AG2 ORDER BY "Ngày thanh toán"
    """
    df_detail = con.execute(q_detail).df()
    print(f"  ✓ Sheet 1: detail ({len(df_detail)} dòng)")

    # ------------------------------------------------------------------------
    # Sheet 2: TRACKING AG
    # ------------------------------------------------------------------------
    q_ag = f"""
    WITH t1 AS (
        SELECT dnsa.*, TRY_CAST(uadcd."Ngày thanh toán" AS DATE) as "Ngày thanh toán", uadcd."Code sale", uadcd."Sản phẩm", uadcd."Đối tác nhà bảo hiểm", uadcd."Số tiền thanh toán", upper(uadcd."Channel") as "Channel Sales", uadcd."Loại bảo hiểm"
        FROM df_nhansu dnsa JOIN df_capdon uadcd ON LTRIM(dnsa."Điện thoại", '0') = LTRIM(uadcd."Code sale", '0')
    ),
    detail AS (
        SELECT "Code", "Họ tên", "Chức danh", "Channel", "Ngày thanh toán", "Sản phẩm", "Đối tác nhà bảo hiểm", CAST("Số tiền thanh toán" AS DOUBLE) as "Số tiền thanh toán", "QUẢN LÝ CẤP 1 (BDM)", "QUẢN LÝ CẤP 2 (BDD)", "Channel Sales", "Thời gian bắt đầu", "Loại bảo hiểm"
        FROM t1 WHERE "Channel" = 'Neo' AND "Ngày thanh toán" IS NOT NULL AND extract(year from "Ngày thanh toán") = {report_year} AND extract(month from "Ngày thanh toán") = {report_month}
    ),
    t2 AS (
        SELECT "Code", "Họ tên", "Chức danh", "Channel", "Ngày thanh toán", d."Sản phẩm", "Đối tác nhà bảo hiểm", "QUẢN LÝ CẤP 1 (BDM)", "QUẢN LÝ CẤP 2 (BDD)", "Channel Sales", "Thời gian bắt đầu", "Loại bảo hiểm", "Số tiền thanh toán",
            CASE
                WHEN d."Sản phẩm" LIKE '%Trách nhiệm Dân sự%' OR d."Sản phẩm" LIKE '%VCOTO%' OR d."Sản phẩm" LIKE '%vật chất ô tô%' OR d."Sản phẩm" LIKE '%Lái phụ xe và người ngồi trên xe%' OR "Loại bảo hiểm" = 'BHRR' THEN ROUND(d."Số tiền thanh toán" / 1.1, 0)
                WHEN d."Sản phẩm" = 'Trách nhiệm sản phẩm' THEN ROUND(d."Số tiền thanh toán" / 1.05, 0) ELSE ROUND(d."Số tiền thanh toán", 0)
            END as "Doanh thu trước thuế",
            qd."rate_bonus",
            CASE
                WHEN UPPER(d."Sản phẩm") IN ('B-ONE_NEW', 'B-ONE_RENEW', 'SKTA_NEW', 'SKTA_RENEW', 'AFFINA_CARE', 'TAI NẠN 24/7', 'AFFINA_CARE_NEW', 'AFFINA_CARE_RENEW') THEN 1
                WHEN ("Ngày thanh toán" >= DATE '2025-09-01' AND UPPER(d."Sản phẩm") LIKE '%XE MÁY%' AND UPPER(d."Đối tác nhà bảo hiểm") IN ('BSH', 'PVI DIGITAL', 'AAA')) THEN 1
                WHEN ("Ngày thanh toán" >= DATE '2025-09-01' AND UPPER(d."Sản phẩm") LIKE '%LÁI PHỤ XE VÀ NGƯỜI NGỒI TRÊN XE%' AND UPPER(d."Đối tác nhà bảo hiểm") IN ('BSH', 'PTI', 'TASCO', 'MIC', 'PVI DIGITAL', 'AAA')) THEN 1
                WHEN ("Ngày thanh toán" >= DATE '2025-09-01' AND UPPER(d."Sản phẩm") LIKE '%TRÁCH NHIỆM DÂN SỰ Ô TÔ%' AND UPPER(d."Đối tác nhà bảo hiểm") IN ('BSH', 'PTI', 'TASCO', 'MIC', 'PVI DIGITAL', 'AAA')) THEN 1
                ELSE 0
            END AS exchange_core
        FROM detail d LEFT JOIN df_quydoi qd ON UPPER(TRIM(qd."provider")) = UPPER(TRIM(d."Đối tác nhà bảo hiểm")) AND UPPER(qd."product") = UPPER(d."Sản phẩm") AND d."Ngày thanh toán" >= TRY_CAST(qd."Effective_Date" AS DATE) AND d."Ngày thanh toán" <= TRY_CAST(qd."Valid_to" AS DATE)
    ),
    t3 AS (
        SELECT t2.*, CASE WHEN "Ngày thanh toán" = current_date THEN "Số tiền thanh toán" END as "Doanh số hôm nay", (CAST("rate_bonus" AS DOUBLE) * "Doanh thu trước thuế") as "EST_Bonus", (exchange_core * "Doanh thu trước thuế") as "Doanh số qui đổi"
        FROM t2
    ),
    DSA AS (
        SELECT "Code", "Họ tên", "Chức danh", "Channel", "Thời gian bắt đầu", "QUẢN LÝ CẤP 1 (BDM)", "QUẢN LÝ CẤP 2 (BDD)", SUM("Số tiền thanh toán") as "Số tiền thanh toán", SUM("Doanh số qui đổi") as "Doanh số qui đổi", SUM("Doanh thu trước thuế") as "Doanh thu trước thuế", SUM("Doanh số hôm nay") as "Doanh số hôm nay", SUM("EST_Bonus") as "EST_Bonus"
        FROM t3 GROUP BY 1,2,3,4,5,6,7
    ),
    t6 AS (
        SELECT dnsa."Code", dnsa."Họ tên", dnsa."Chức danh", dnsa."Channel", dnsa."QUẢN LÝ CẤP 1 (BDM)", dnsa."QUẢN LÝ CẤP 2 (BDD)", DSA."Số tiền thanh toán", DSA."Doanh số qui đổi", DSA."Doanh thu trước thuế", DSA."Doanh số hôm nay", DSA."EST_Bonus"
        FROM df_nhansu dnsa LEFT JOIN DSA ON TRIM(DSA."Họ tên") = TRIM(dnsa."Họ tên") AND TRIM(DSA."Code") = TRIM(dnsa."Code") WHERE dnsa."Chức danh" = 'AG' AND dnsa."Channel" = 'Neo'
    ),
    tx AS (
        SELECT "Code", "Họ tên", "Chức danh", "Điện thoại" FROM df_nhansu WHERE "Channel" = 'Neo' AND "Chức danh" = 'AG'
    ),
    ngt AS (
        SELECT c."Code", c."Họ tên" as "ten_DSA", COUNT(d."Người giới thiệu") as "Số người giới thiệu"
        FROM tx c LEFT JOIN df_nhansu d ON LTRIM(c."Điện thoại", '0') = LTRIM(d."Người giới thiệu", '0') GROUP BY 1, 2
    )
    SELECT t6."Code", t6."Họ tên", t6."Chức danh", t6."Channel", t6."QUẢN LÝ CẤP 1 (BDM)" as "QUẢN LÝ CẤP 1 (SM)", t6."QUẢN LÝ CẤP 2 (BDD)" as "QUẢN LÝ CẤP 2 (SD)", COALESCE(t6."Số tiền thanh toán", 0) as "Số tiền thanh toán", COALESCE(t6."Doanh số qui đổi", 0) as "Doanh số qui đổi", COALESCE(t6."Doanh thu trước thuế", 0) as "Doanh thu trước thuế", COALESCE(t6."Doanh số hôm nay", 0) as "Doanh số hôm nay", COALESCE(t6."EST_Bonus", 0) as "AG_bonus", COALESCE(ngt."Số người giới thiệu", 0) as "Số người giới thiệu"
    FROM t6 LEFT JOIN ngt ON ngt."Code" = t6."Code" ORDER BY t6."Họ tên"
    """
    df_ag = con.execute(q_ag).df()
    print(f"  ✓ Sheet 2: tracking AG ({len(df_ag)} dòng)")

    # ------------------------------------------------------------------------
    # Sheet 3: TRACKING SM
    # ------------------------------------------------------------------------
    q_sm = f"""
    WITH t1 AS (
        SELECT dnsa.*, TRY_CAST(uadcd."Ngày thanh toán" AS DATE) as "Ngày thanh toán", uadcd."Code sale", uadcd."Sản phẩm", uadcd."Đối tác nhà bảo hiểm", uadcd."Số tiền thanh toán", upper(uadcd."Channel") as "Channel Sales", uadcd."Loại bảo hiểm"
        FROM df_nhansu dnsa JOIN df_capdon uadcd ON LTRIM(dnsa."Điện thoại", '0') = LTRIM(uadcd."Code sale", '0')
    ),
    detail AS (
        SELECT "Code", "Họ tên", "Code UM", "Chức danh", "Channel", "Ngày thanh toán", "Sản phẩm", "Đối tác nhà bảo hiểm", CAST("Số tiền thanh toán" AS DOUBLE) as "Số tiền thanh toán", "QUẢN LÝ CẤP 1 (BDM)" as "QUẢN LÝ CẤP 1 (SM)", "QUẢN LÝ CẤP 2 (BDD)" as "QUẢN LÝ CẤP 2 (SD)", "Channel Sales", "Thời gian bắt đầu", "Loại bảo hiểm"
        FROM t1 WHERE "Channel" = 'Neo' AND "Ngày thanh toán" IS NOT NULL AND extract(year from "Ngày thanh toán") = {report_year} AND extract(month from "Ngày thanh toán") = {report_month}
    ),
    AG1 AS (
        SELECT "Code", "Họ tên", "Chức danh", "Code UM", "Channel", "Ngày thanh toán", d."Sản phẩm", "Đối tác nhà bảo hiểm", "QUẢN LÝ CẤP 1 (SM)", CASE WHEN "QUẢN LÝ CẤP 1 (SM)" IS NULL THEN "QUẢN LÝ CẤP 2 (SD)" ELSE "QUẢN LÝ CẤP 1 (SM)" END as transfer_when_missing_SM, "QUẢN LÝ CẤP 2 (SD)", "Channel Sales", "Thời gian bắt đầu", "Loại bảo hiểm", "Số tiền thanh toán",
            CASE WHEN d."Sản phẩm" LIKE '%Trách nhiệm Dân sự%' OR d."Sản phẩm" LIKE '%VCOTO%' OR d."Sản phẩm" LIKE '%vật chất ô tô%' OR d."Sản phẩm" LIKE '%Lái phụ xe và người ngồi trên xe%' OR "Loại bảo hiểm" = 'BHRR' THEN ROUND(d."Số tiền thanh toán" / 1.1, 0) WHEN d."Sản phẩm" = 'Trách nhiệm sản phẩm' THEN ROUND(d."Số tiền thanh toán" / 1.05, 0) ELSE ROUND(d."Số tiền thanh toán", 0) END as "Doanh thu trước thuế",
            qd."rate_bonus", CASE WHEN extract(month from d."Ngày thanh toán") BETWEEN 2 AND 6 AND qd."sp_contest_neo" = 1 THEN qd."RMM_IO" ELSE NULL END as "RMM_IO"
        FROM detail d LEFT JOIN df_quydoi qd ON UPPER(TRIM(qd."provider")) = UPPER(TRIM(d."Đối tác nhà bảo hiểm")) AND UPPER(qd."product") = UPPER(d."Sản phẩm") AND d."Ngày thanh toán" >= TRY_CAST(qd."Effective_Date" AS DATE) AND d."Ngày thanh toán" <= TRY_CAST(qd."Valid_to" AS DATE)
    ),
    AG2 AS (
        SELECT AG1.*, ROUND(CAST(qd."RMM_OR" AS DOUBLE), 4) as "RMM_OR_rate", (CAST(qd."RMM_OR" AS DOUBLE) * AG1."Doanh thu trước thuế") as "RMM_OR", ROUND(CAST(qd."RMD_OR" AS DOUBLE), 4) as "RMD_OR_rate", (CAST(qd."RMD_OR" AS DOUBLE) * AG1."Doanh thu trước thuế") as "RMD_OR", ROUND(CAST(qd."rate_bonus" AS DOUBLE) * AG1."Doanh thu trước thuế", 0) as "EST_bonus", (CAST(qd."RMD_IO" AS DOUBLE) * AG1."Doanh thu trước thuế") as "RMD_IO", (CAST(qd."RMM_IO" AS DOUBLE) * AG1."Doanh thu trước thuế") as "RMM_IO", CASE WHEN "Ngày thanh toán" = current_date THEN "Số tiền thanh toán" END as "Doanh số hôm nay"
        FROM AG1 LEFT JOIN df_quydoi qd ON UPPER(TRIM(qd."provider")) = UPPER(TRIM(AG1."Đối tác nhà bảo hiểm")) AND UPPER(qd."product") = UPPER(AG1."Sản phẩm") AND AG1."Ngày thanh toán" >= TRY_CAST(qd."Effective_Date" AS DATE) AND AG1."Ngày thanh toán" <= TRY_CAST(qd."Valid_to" AS DATE)
    ),
    AG3 AS (
        SELECT "Code", "Họ tên", "Code UM", "Chức danh", "Channel", "Ngày thanh toán", "Sản phẩm", "Thời gian bắt đầu", "QUẢN LÝ CẤP 1 (SM)", "QUẢN LÝ CẤP 2 (SD)", COALESCE(SUM("Doanh số hôm nay"), 0) as "Doanh số hôm nay", COALESCE(SUM("Số tiền thanh toán"), 0) as "Số tiền thanh toán", COALESCE(SUM("Doanh thu trước thuế"), 0) as "Doanh thu trước thuế", COALESCE(SUM("RMM_OR"), 0) as "RMM_OR", COALESCE(SUM("EST_bonus"), 0) as "EST_bonus", COALESCE(SUM("RMM_IO"), 0) as "RMM_IO"
        FROM AG2 GROUP BY 1,2,3,4,5,6,7,8,9,10
    ),
    SM_track AS (
        SELECT dnsa."Code", dnsa."Họ tên", dnsa."Code UM", dnsa."Chức danh", TRY_CAST(dnsa."Thời gian bắt đầu" AS DATE) as "Thời gian bắt đầu", dnsa."QUẢN LÝ CẤP 2 (BDD)" as "QUẢN LÝ CẤP 2 (SD)", SUM(f."Số tiền thanh toán") as "Số tiền thanh toán", COALESCE(SUM(f."Doanh thu trước thuế"), 0) as "Doanh thu trước thuế", SUM(f."EST_bonus") as "AG_bonus", SUM(f."Doanh số hôm nay") as "Doanh số hôm nay", COALESCE(ROUND(SUM(f."RMM_OR"), 0), 0) as "RMM_OR", COALESCE(SUM(f."RMM_IO"), 0) as "RMM_IO"
        FROM df_nhansu dnsa LEFT JOIN AG3 f ON TRIM(dnsa."Code UM") = TRIM(f."Code UM") AND TRIM(f."QUẢN LÝ CẤP 1 (SM)") = TRIM(dnsa."Họ tên") WHERE UPPER(dnsa."Chức danh") LIKE '%SM%' AND dnsa."Channel" = 'Neo' GROUP BY 1,2,3,4,5,6
    ),
    SM_track2 AS (
        SELECT t7.*, COUNT(DISTINCT CASE WHEN dnsa."Status" IN ('A','P') AND dnsa."Chức danh" = 'AG' THEN dnsa."Code" END) as "Số AG thực tế",
            COUNT(DISTINCT CASE WHEN dnsa."Status" IN ('A','P') AND dnsa."Chức danh" = 'AG' AND extract(year from TRY_CAST(dnsa."Thời gian bắt đầu" AS DATE)) = {report_year} AND extract(month from TRY_CAST(dnsa."Thời gian bắt đầu" AS DATE)) = {report_month} THEN dnsa."Code" END) as "Tuyển dụng mới AG",
            COUNT(DISTINCT CASE WHEN d."Chức danh" NOT LIKE '%SM%' AND d."Chức danh" NOT LIKE '%SD%' THEN d."Code" END) as "AG Active"
        FROM SM_track t7 JOIN df_nhansu dnsa ON TRIM(dnsa."Code UM") = TRIM(t7."Code UM") AND TRIM(dnsa."QUẢN LÝ CẤP 1 (BDM)") = TRIM(t7."Họ tên") LEFT JOIN AG3 d ON TRIM(dnsa."Code UM") = TRIM(d."Code UM") AND d."QUẢN LÝ CẤP 1 (SM)" = t7."Họ tên" GROUP BY 1,2,3,4,5,6,7,8,9,10,11,12
    )
    SELECT "Code", "Họ tên", "Code UM", "Chức danh", "Thời gian bắt đầu", "QUẢN LÝ CẤP 2 (SD)", COALESCE("Số tiền thanh toán", 0) as "Số tiền thanh toán", COALESCE("Doanh thu trước thuế", 0) as "Doanh thu trước thuế", COALESCE("AG_bonus", 0) as "AG_bonus", COALESCE("Doanh số hôm nay", 0) as "Doanh số hôm nay", COALESCE("RMM_OR", 0) as "SM_OR", COALESCE("Số AG thực tế", 0) as "Số AG thực tế", COALESCE("Tuyển dụng mới AG", 0) as "Tuyển dụng mới AG", COALESCE("AG Active", 0) as "AG Active", COALESCE("RMM_IO", 0) as "SM_IO"
    FROM SM_track2 ORDER BY "Họ tên"
    """
    df_sm = con.execute(q_sm).df()
    print(f"  ✓ Sheet 3: tracking SM ({len(df_sm)} dòng)")

    # ------------------------------------------------------------------------
    # Sheet 4: TRACKING SD
    # ------------------------------------------------------------------------
    q_sd = f"""
    WITH t1 AS (
        SELECT dnsa.*, TRY_CAST(uadcd."Ngày thanh toán" AS DATE) as "Ngày thanh toán", uadcd."Số hợp đồng", uadcd."Code sale", uadcd."Sản phẩm", uadcd."Đối tác nhà bảo hiểm", uadcd."Số tiền thanh toán", upper(uadcd."Channel") as "Channel Sales", uadcd."Loại bảo hiểm"
        FROM df_nhansu dnsa JOIN df_capdon uadcd ON LTRIM(dnsa."Điện thoại", '0') = LTRIM(uadcd."Code sale", '0')
        WHERE uadcd."Ngày thanh toán" IS NOT NULL
    ),
    detail AS (
        SELECT "Code", "Họ tên", "Số hợp đồng", "Code UM", "Chức danh", "Channel", "Ngày thanh toán", "Sản phẩm", "Đối tác nhà bảo hiểm", CAST("Số tiền thanh toán" AS DOUBLE) as "Số tiền thanh toán", "QUẢN LÝ CẤP 1 (BDM)", "QUẢN LÝ CẤP 2 (BDD)", "Channel Sales", "Thời gian bắt đầu", "Loại bảo hiểm"
        FROM t1 WHERE "Channel" = 'Neo' AND "Ngày thanh toán" IS NOT NULL AND extract(year from "Ngày thanh toán") = {report_year} AND extract(month from "Ngày thanh toán") = {report_month}
    ),
    agg_detail AS (
        SELECT "Code", "Họ tên", "Số hợp đồng", "Code UM", "Chức danh", "Channel", "Ngày thanh toán", "Sản phẩm", "Đối tác nhà bảo hiểm", "QUẢN LÝ CẤP 1 (BDM)", CASE WHEN "QUẢN LÝ CẤP 1 (BDM)" IS NULL THEN "QUẢN LÝ CẤP 2 (BDD)" ELSE "QUẢN LÝ CẤP 1 (BDM)" END as transfer_when_missing_SM, "QUẢN LÝ CẤP 2 (BDD)", "Channel Sales", "Thời gian bắt đầu", "Loại bảo hiểm", SUM("Số tiền thanh toán") as "Số tiền thanh toán"
        FROM detail GROUP BY "Code", "Họ tên", "Số hợp đồng", "Code UM", "Chức danh", "Channel", "Ngày thanh toán", "Sản phẩm", "Đối tác nhà bảo hiểm", "QUẢN LÝ CẤP 1 (BDM)", "QUẢN LÝ CẤP 2 (BDD)", "Channel Sales", "Thời gian bắt đầu", "Loại bảo hiểm"
    ),
    t2 AS (
        SELECT d."Code", d."Họ tên", d."Code UM", d."Chức danh", d."Channel", d."Số hợp đồng", d."Ngày thanh toán", d."Sản phẩm", d."Đối tác nhà bảo hiểm", d."QUẢN LÝ CẤP 1 (BDM)", d.transfer_when_missing_SM, d."QUẢN LÝ CẤP 2 (BDD)", d."Channel Sales", d."Thời gian bắt đầu", d."Loại bảo hiểm", d."Số tiền thanh toán",
            CASE WHEN d."Sản phẩm" LIKE '%Trách nhiệm Dân sự%' OR d."Sản phẩm" LIKE '%VCOTO%' OR d."Sản phẩm" LIKE '%vật chất ô tô%' OR d."Sản phẩm" LIKE '%Lái phụ xe và người ngồi trên xe%' OR d."Loại bảo hiểm" = 'BHRR' THEN ROUND(d."Số tiền thanh toán" / 1.1, 0) WHEN d."Sản phẩm" = 'Trách nhiệm sản phẩm' THEN ROUND(d."Số tiền thanh toán" / 1.05, 0) ELSE ROUND(d."Số tiền thanh toán", 0) END as "Doanh thu trước thuế",
            dsqd."rate_bonus",
            CASE WHEN UPPER(d."Sản phẩm") IN ('B-ONE_NEW', 'B-ONE_RENEW', 'SKTA_NEW', 'SKTA_RENEW', 'AFFINA_CARE', 'TAI NẠN 24/7', 'AFFINA_CARE_NEW', 'AFFINA_CARE_RENEW') THEN 1 WHEN ("Ngày thanh toán" >= DATE '2025-09-01' AND UPPER(d."Sản phẩm") LIKE '%XE MÁY%' AND UPPER(d."Đối tác nhà bảo hiểm") IN ('BSH', 'PVI DIGITAL', 'AAA')) THEN 1 WHEN ("Ngày thanh toán" >= DATE '2025-09-01' AND UPPER(d."Sản phẩm") LIKE '%LÁI PHỤ XE VÀ NGƯỜI NGỒI TRÊN XE%' AND UPPER(d."Đối tác nhà bảo hiểm") IN ('BSH', 'PTI', 'TASCO', 'MIC', 'PVI DIGITAL', 'AAA')) THEN 1 WHEN ("Ngày thanh toán" >= DATE '2025-09-01' AND UPPER(d."Sản phẩm") LIKE '%TRÁCH NHIỆM DÂN SỰ Ô TÔ%' AND UPPER(d."Đối tác nhà bảo hiểm") IN ('BSH', 'PTI', 'TASCO', 'MIC', 'PVI DIGITAL', 'AAA')) THEN 1 ELSE 0 END AS exchange_core,
            ROUND(CAST(dsqd."RMM_OR" AS DOUBLE), 4) as "RMM_OR_rate", ROUND(CAST(dsqd."RMD_OR" AS DOUBLE), 4) as "RMD_OR_rate",
            CASE WHEN extract(month from d."Ngày thanh toán") BETWEEN 2 AND 6 AND dsqd."sp_contest_neo" = 1 THEN ROUND(CAST(dsqd."RMM_IO" AS DOUBLE), 4) ELSE NULL END as "RMM_IO_rate",
            CASE WHEN extract(month from d."Ngày thanh toán") BETWEEN 2 AND 6 AND dsqd."sp_contest_neo" = 1 THEN ROUND(CAST(dsqd."RMD_IO" AS DOUBLE), 4) ELSE NULL END as "RMD_IO_rate"
        FROM agg_detail d LEFT JOIN df_quydoi dsqd ON UPPER(TRIM(dsqd."provider")) = UPPER(TRIM(d."Đối tác nhà bảo hiểm")) AND UPPER(dsqd."product") = UPPER(d."Sản phẩm") AND d."Ngày thanh toán" >= TRY_CAST(dsqd."Effective_Date" AS DATE) AND d."Ngày thanh toán" <= TRY_CAST(dsqd."Valid_to" AS DATE)
    ),
    t3 AS (
        SELECT t2.*, CASE WHEN "Ngày thanh toán" = current_date THEN "Số tiền thanh toán" END as "Doanh số hôm nay", (CAST("rate_bonus" AS DOUBLE) * "Doanh thu trước thuế")  as "EST_Bonus", (exchange_core * "Doanh thu trước thuế") as "Doanh số qui đổi", ("RMM_OR_rate" * "Doanh thu trước thuế") as "RMM_OR", ("RMD_OR_rate" * "Doanh thu trước thuế") as "RMD_OR", (CAST("RMM_IO_rate" AS DOUBLE) * "Doanh thu trước thuế") as "RMM_IO", (CAST("RMD_IO_rate" AS DOUBLE) * "Doanh thu trước thuế") as "RMD_IO"
        FROM t2
    ),
    DSA AS (
        SELECT "Code", "Code UM", "Họ tên", "Chức danh", "Channel", "Thời gian bắt đầu", "QUẢN LÝ CẤP 1 (BDM)", "QUẢN LÝ CẤP 2 (BDD)", SUM("Số tiền thanh toán") as "Số tiền thanh toán", SUM("Doanh số qui đổi") as "Doanh số qui đổi", SUM("Doanh thu trước thuế") as "Doanh thu trước thuế", SUM("Doanh số hôm nay") as "Doanh số hôm nay", SUM("EST_Bonus") as "EST_Bonus", SUM("RMM_OR") as "RMM_OR", SUM("RMD_OR") as "RMD_OR", SUM("RMM_IO") as "RMM_IO", SUM("RMD_IO") as "RMD_IO"
        FROM t3 GROUP BY 1,2,3,4,5,6,7,8
    ),
    t7 AS (
        SELECT dnsa."Code", dnsa."Họ tên", dnsa."Code UM", dnsa."Chức danh", TRY_CAST(dnsa."Thời gian bắt đầu" AS DATE) as "Thời gian bắt đầu", dnsa."QUẢN LÝ CẤP 2 (BDD)",
            COALESCE(ag_sum."Số tiền thanh toán", 0) + COALESCE(sm_self."Số tiền thanh toán", 0) as "Số tiền thanh toán",
            COALESCE(ag_sum."Doanh thu trước thuế", 0) + COALESCE(sm_self."Doanh thu trước thuế", 0) as "Doanh thu trước thuế",
            COALESCE(ag_sum."Doanh số qui đổi", 0) + COALESCE(sm_self."Doanh số qui đổi", 0) as "Doanh số qui đổi",
            COALESCE(ag_sum."EST_Bonus", 0) + COALESCE(sm_self."EST_Bonus", 0) as "EST_Bonus",
            COALESCE(ag_sum."Doanh số hôm nay", 0) + COALESCE(sm_self."Doanh số hôm nay", 0) as "Doanh số hôm nay",
            COALESCE(ag_sum."RMM_OR", 0) + COALESCE(sm_self."RMM_OR", 0) as "RMM_OR",
            COALESCE(ag_sum."RMD_OR", 0) + COALESCE(sm_self."RMD_OR", 0) as "RMD_OR",
            COALESCE(ag_sum."RMM_IO", 0) + COALESCE(sm_self."RMM_IO", 0) as "RMM_IO",
            COALESCE(ag_sum."RMD_IO", 0) + COALESCE(sm_self."RMD_IO", 0) as "RMD_IO"
        FROM df_nhansu dnsa
        LEFT JOIN (SELECT "Code UM", SUM("Số tiền thanh toán") as "Số tiền thanh toán", SUM("Doanh thu trước thuế") as "Doanh thu trước thuế", SUM("Doanh số qui đổi") as "Doanh số qui đổi", SUM("EST_Bonus") as "EST_Bonus", SUM("Doanh số hôm nay") as "Doanh số hôm nay", SUM("RMM_OR") as "RMM_OR", SUM("RMD_OR") as "RMD_OR", SUM("RMM_IO") as "RMM_IO", SUM("RMD_IO") as "RMD_IO" FROM DSA WHERE UPPER("Chức danh") NOT LIKE '%SM%' AND UPPER("Chức danh") NOT LIKE '%SD%' GROUP BY "Code UM") ag_sum ON TRIM(dnsa."Code UM") = TRIM(ag_sum."Code UM")
        LEFT JOIN (SELECT "Code", SUM("Số tiền thanh toán") as "Số tiền thanh toán", SUM("Doanh thu trước thuế") as "Doanh thu trước thuế", SUM("Doanh số qui đổi") as "Doanh số qui đổi", SUM("EST_Bonus") as "EST_Bonus", SUM("Doanh số hôm nay") as "Doanh số hôm nay", SUM("RMM_OR") as "RMM_OR", SUM("RMD_OR") as "RMD_OR", SUM("RMM_IO") as "RMM_IO", SUM("RMD_IO") as "RMD_IO" FROM DSA WHERE UPPER("Chức danh") LIKE '%SM%' GROUP BY "Code") sm_self ON TRIM(dnsa."Code") = TRIM(sm_self."Code")
        WHERE UPPER(dnsa."Chức danh") LIKE '%SM%' AND dnsa."Channel" = 'Neo'
    ),
    t8 AS (
        SELECT t7."QUẢN LÝ CẤP 2 (BDD)",
            COUNT(DISTINCT CASE WHEN dnsa."Status" IN ('A', 'P') AND dnsa."Chức danh" = 'AG' THEN dnsa."Code" END) as "Số DSA thực tế",
            COUNT(DISTINCT CASE WHEN dnsa."Status" IN ('A', 'P') AND dnsa."Chức danh" = 'AG' AND extract(year from TRY_CAST(dnsa."Thời gian bắt đầu" AS DATE)) = {report_year} AND extract(month from TRY_CAST(dnsa."Thời gian bắt đầu" AS DATE)) = {report_month} THEN dnsa."Code" END) as "Tuyển dụng mới DSA",
            COUNT(DISTINCT CASE WHEN d."Chức danh" NOT LIKE '%SM%' AND d."Chức danh" NOT LIKE '%SD%' THEN d."Code" END) as "DSA Active"
        FROM t7
        JOIN df_nhansu dnsa ON TRIM(dnsa."Code UM") = TRIM(t7."Code UM") AND TRIM(dnsa."QUẢN LÝ CẤP 1 (BDM)") = TRIM(t7."Họ tên")
        LEFT JOIN DSA d ON TRIM(d."Code UM") = TRIM(t7."Code UM") AND TRIM(d."QUẢN LÝ CẤP 1 (BDM)") = TRIM(t7."Họ tên")
        GROUP BY t7."QUẢN LÝ CẤP 2 (BDD)"
    ),
    nearly_BDD1 AS (
        SELECT "QUẢN LÝ CẤP 2 (BDD)", SUM("Số tiền thanh toán") as "Số tiền thanh toán", SUM("Doanh số qui đổi") as "Doanh số qui đổi", SUM("Doanh thu trước thuế") as "Doanh thu trước thuế", SUM("EST_Bonus") as "EST_Bonus", SUM("RMM_OR") as "RMM_OR", SUM("RMD_OR") as "RMD_OR", SUM("RMM_IO") as "RMM_IO", SUM("RMD_IO") as "RMD_IO"
        FROM DSA WHERE "QUẢN LÝ CẤP 2 (BDD)" IS NOT NULL GROUP BY 1
    ),
    nearly_BDD2 AS (
        SELECT a.*, d."Tuyển dụng mới DSA", d."Số DSA thực tế", d."DSA Active"
        FROM nearly_BDD1 AS a
        LEFT JOIN t8 AS d ON UPPER(d."QUẢN LÝ CẤP 2 (BDD)") = UPPER(a."QUẢN LÝ CẤP 2 (BDD)")
    ),
    SM_OR_by_SD AS (
        SELECT "QUẢN LÝ CẤP 2 (BDD)", ROUND(SUM("RMM_OR"), 0) as "SM_OR" FROM t7 GROUP BY 1
    ),
    SM_IO_by_SD AS (
        SELECT "QUẢN LÝ CẤP 2 (BDD)", ROUND(SUM("RMM_IO"), 0) as "SM_IO" FROM t7 GROUP BY 1
    )
    SELECT nbdd."QUẢN LÝ CẤP 2 (BDD)" as "QUẢN LÝ CẤP 2 (SD)",
        COALESCE(nbdd."Số tiền thanh toán", 0) as "Số tiền thanh toán",
        COALESCE(nbdd."Doanh thu trước thuế", 0) as "Doanh thu trước thuế",
        COALESCE(nbdd."EST_Bonus", 0) as "AG_bonus",
        COALESCE(sm_or."SM_OR", 0) as "SM_OR",
        COALESCE(nbdd."RMD_OR", 0) as "SD_OR",
        COALESCE(nbdd."Tuyển dụng mới DSA", 0) as "Tuyển dụng mới AG",
        COALESCE(nbdd."Số DSA thực tế", 0) as "Số AG thực tế",
        COALESCE(nbdd."DSA Active", 0) as "AG Active",
        COALESCE(sm_io."SM_IO", 0) as "SM_IO",
        COALESCE(nbdd."RMD_IO", 0) as "SD_IO"
    FROM nearly_BDD2 nbdd
    LEFT JOIN SM_OR_by_SD sm_or ON UPPER(sm_or."QUẢN LÝ CẤP 2 (BDD)") = UPPER(nbdd."QUẢN LÝ CẤP 2 (BDD)")
    LEFT JOIN SM_IO_by_SD sm_io ON UPPER(sm_io."QUẢN LÝ CẤP 2 (BDD)") = UPPER(nbdd."QUẢN LÝ CẤP 2 (BDD)")
    WHERE nbdd."QUẢN LÝ CẤP 2 (BDD)" IS NOT NULL
    ORDER BY "Số tiền thanh toán" DESC
    """
    df_sd = con.execute(q_sd).df()
    print(f"  ✓ Sheet 4: tracking SD ({len(df_sd)} dòng)")

    # ------------------------------------------------------------------------
    # Sheet 5: TRACKING CD (Giám đốc Kênh Neo - PHẠM TRƯỜNG KHÁNH)
    # ------------------------------------------------------------------------
    con.register('df_sd_calc', df_sd)
    q_cd = """
    SELECT 
        'LD4641' as "Code CD",
        'PHẠM TRƯỜNG KHÁNH' as "Giám đốc Kênh (CD)",
        "QUẢN LÝ CẤP 2 (SD)" as "SD phụ trách",
        "Số tiền thanh toán",
        "Doanh thu trước thuế",
        "AG_bonus",
        "SM_OR",
        "SD_OR",
        "Tuyển dụng mới AG",
        "Số AG thực tế",
        "AG Active",
        "SM_IO",
        "SD_IO"
    FROM df_sd_calc
    """
    df_cd = con.execute(q_cd).df()
    print(f"  ✓ Sheet 5: tracking CD ({len(df_cd)} dòng)")

    # ------------------------------------------------------------------------
    # Sheet 6: THEO DÕI TÁI TỤC DETAIL
    # ------------------------------------------------------------------------
    q_theodoi = """
    WITH LocKhacRenew AS (
        SELECT * FROM df_union WHERE "Loại bảo hiểm" IN ('BHSK', 'BHOTO', 'BHXM')
    ),
    KhachNew AS (
        SELECT 
            "Tên Người được BH" as "Tên Người được BH new",
            "Tên NMBH" as "Tên NMBH new",
            TRY_CAST(lkr."Ngày Sinh NNBH" AS DATE) as "Ngày Sinh NNBH new",
            concat(replace(CAST("Ngày Sinh NNBH" AS VARCHAR),'-',''), strip_accents(LOWER(REPLACE("Tên Người được BH", ' ', '')))) as "Code riêng",
            "Channel" as "Channel new",
            "Code sale" as "Code sales new",
            "Sản phẩm" as "Sản phẩm new",
            coalesce("Đối tác nhà bảo hiểm", '') as "Đối tác nhà bảo hiểm new",
            "Số tiền thanh toán" as "Số tiền thanh toán new",
            "Số hợp đồng" as "Số hợp đồng new",
            TRY_CAST("Ngày bắt đầu" AS DATE) as "Ngày hiệu lực khách new",
            TRY_CAST("Ngày kết thúc" AS DATE) as "Ngày kết thúc khách new",
            extract(year from TRY_CAST("Ngày bắt đầu" AS DATE)) as "Năm hiệu lực khách new",
            lkr.*
        FROM LocKhacRenew lkr
        WHERE TRY_CAST("Ngày kết thúc" AS DATE) >= date_trunc('month', current_date)
          AND TRY_CAST("Ngày kết thúc" AS DATE) < date_trunc('month', current_date + interval '2 month')
          AND "Loại bảo hiểm" IN ('BHSK', 'BHOTO', 'BHXM')
    ),
    LocRenew AS (
        SELECT * FROM df_union WHERE "Channel" like '%enew%' AND "Loại bảo hiểm" IN ('BHSK', 'BHOTO', 'BHXM')
    ),
    AllRenew AS (
        SELECT 
            "Tên Người được BH" as "Tên Người được BH Renew",
            "Tên NMBH" as "Tên NMBH Renew",
            TRY_CAST(lr."Ngày Sinh NNBH" AS DATE) as "Ngày Sinh NNBH Renew",
            concat(replace(CAST("Ngày Sinh NNBH" AS VARCHAR),'-',''), strip_accents(LOWER(REPLACE("Tên Người được BH", ' ', '')))) as "Code riêng",
            "Channel" as "Channel Renew",
            "Code sale" as "Code sales Renew",
            "Sản phẩm" as "Sản phẩm Renew",
            "Số tiền thanh toán" as "Số tiền thanh toán Renew",
            "Số hợp đồng" as "Số hợp đồng Renew",
            TRY_CAST("Ngày bắt đầu" AS DATE) as "Ngày hiệu lực renew",
            TRY_CAST("Ngày kết thúc" AS DATE) as "Ngày kết thúc renew",
            extract(year from TRY_CAST("Ngày bắt đầu" AS DATE)) as "Năm hiệu lực renew",
            lr.*
        FROM LocRenew lr
    ),
    Join2Bang AS (
        SELECT 
            kn."Code riêng", kn."Tên Người được BH new", kn."Tên NMBH new", kn."Ngày Sinh NNBH new",
            "Channel new", "Channel Renew", "Số hợp đồng new", "Sản phẩm new",
            "Đối tác nhà bảo hiểm new", "Code sales new", "Ngày hiệu lực khách new",
            "Ngày kết thúc khách new", "Số tiền thanh toán new", "Số hợp đồng Renew",
            "Tên NMBH Renew", "Sản phẩm Renew", "Code sales Renew", "Ngày hiệu lực renew",
            "Ngày kết thúc renew", "Số tiền thanh toán Renew", "Năm hiệu lực khách new", "Năm hiệu lực renew"
        FROM KhachNew kn
        LEFT JOIN AllRenew ar ON kn."Code riêng" = ar."Code riêng" AND "Ngày hiệu lực renew" > "Ngày hiệu lực khách new"
    ),
    RankedRenewals AS (
        SELECT *, ROW_NUMBER() OVER (PARTITION BY "Code riêng", "Channel new" ORDER BY "Ngày hiệu lực renew" ASC) as renewal_rank
        FROM Join2Bang
    ),
    detail_fix AS (
        SELECT 
            "Code riêng", "Channel new", "Tên Người được BH new", "Tên NMBH new", "Ngày Sinh NNBH new",
            "Đối tác nhà bảo hiểm new", "Năm hiệu lực khách new", "Ngày hiệu lực khách new",
            "Ngày kết thúc khách new", "Số hợp đồng new", "Sản phẩm new", "Code sales new", "Số tiền thanh toán new",
            MAX(CASE WHEN renewal_rank = 1 and "Năm hiệu lực renew" - "Năm hiệu lực khách new" = 1 THEN "Năm hiệu lực renew" END) AS "Năm gia hạn 1",
            MAX(CASE WHEN renewal_rank = 1 and "Năm hiệu lực renew" - "Năm hiệu lực khách new" = 1 THEN "Ngày hiệu lực renew" END) AS "Ngày hiệu lực renew 1",
            MAX(CASE WHEN renewal_rank = 1 and "Năm hiệu lực renew" - "Năm hiệu lực khách new" = 1 THEN "Ngày kết thúc renew" END) AS "Ngày kết thúc renew 1",
            MAX(CASE WHEN renewal_rank = 1 and "Năm hiệu lực renew" - "Năm hiệu lực khách new" = 1 THEN "Channel Renew" END) AS "Kênh gia hạn 1",
            MAX(CASE WHEN renewal_rank = 1 and "Năm hiệu lực renew" - "Năm hiệu lực khách new" = 1 THEN "Số tiền thanh toán Renew" END) AS "Số tiền thanh toán 1",
            MAX(CASE WHEN renewal_rank = 1 and "Năm hiệu lực renew" - "Năm hiệu lực khách new" = 1 THEN "Code sales Renew" END) AS "Code sales Renew 1",
            MAX(CASE WHEN renewal_rank = 1 and "Năm hiệu lực renew" - "Năm hiệu lực khách new" = 1 THEN "Sản phẩm Renew" END) AS "Sản phẩm Renew 1",
            MAX(CASE WHEN renewal_rank = 1 and "Năm hiệu lực renew" - "Năm hiệu lực khách new" = 1 THEN "Tên NMBH Renew" END) AS "Tên NMBH Renew 1",
            MAX(CASE WHEN renewal_rank = 1 and "Năm hiệu lực renew" - "Năm hiệu lực khách new" = 1 THEN "Số hợp đồng Renew" END) AS "Số hợp đồng Renew 1"
        FROM RankedRenewals GROUP BY 1,2,3,4,5,6,7,8,9,10,11,12,13
    ),
    dsrenew AS (
        SELECT ROW_NUMBER() OVER (PARTITION BY "Code riêng" ORDER BY "Ngày hiệu lực khách new") AS stt, *
        FROM detail_fix
    ),
    Renewal_Base AS (
        SELECT 
            upper(trim(leading '0' from "Code sales new")) AS "Code Sale",
            upper("Tên Người được BH new") as "NĐBH",
            "Ngày Sinh NNBH new" as "Ngày sinh NĐBH",
            upper("Tên NMBH new") as "BMBH",
            "Đối tác nhà bảo hiểm new" as "Đối tác nhà bảo hiểm",
            "Sản phẩm new" as "Sản phẩm",
            "Channel new" as "Channel",
            "Số hợp đồng new" as "Số hợp đồng",
            "Ngày hiệu lực khách new" as "Ngày bắt đầu hiệu lực",
            "Ngày kết thúc khách new" as "Ngày kết thúc hiệu lực",
            "Số tiền thanh toán new" AS "Số tiền thanh toán",
            "Code sales Renew 1" as "Code Sale tái tục",
            "Tên NMBH Renew 1" as "BMBH tái tục",
            "Số hợp đồng Renew 1" as "Số hợp đồng tái tục",
            "Sản phẩm Renew 1" as "Sản phẩm tái tục",
            "Kênh gia hạn 1" as "Channel tái tục",
            "Ngày hiệu lực renew 1" as "Ngày bắt đầu hiệu lực đơn tái tục",
            "Số tiền thanh toán 1" AS "Số tiền tái tục"
        FROM dsrenew
    ),
    theodoi AS (
        SELECT Renewal_Base.*, ns."Code", ns."Họ tên", ns."Chức danh", ns."Channal", ns."Channel" as "Channel Nhân sự", ns."Code UM", 
               ns."QUẢN LÝ CẤP 1 (BDM)", ns."Column1" as "Code BDD", ns."QUẢN LÝ CẤP 2 (BDD)", "Ngày kết thúc hiệu lực" - 3 as "Hạn thanh toán", 
               case when "Số tiền tái tục" > 0 then 'Đã tái tục' when current_date > "Ngày kết thúc hiệu lực" then 'Quá hạn tái tục' 
                    when current_date >= ("Ngày kết thúc hiệu lực" - 3) - 7 then 'Gần đến hạn' else 'Chưa đến hạn tái tục' end as "Trạng thái"
        FROM Renewal_Base LEFT JOIN df_ns ns ON Renewal_Base."Code Sale" = upper(trim(leading '0' from ns."Họ tên"))
        UNION ALL
        SELECT Renewal_Base.*, ns."Code", ns."Họ tên", ns."Chức danh", ns."Channal", ns."Channel" as "Channel Nhân sự", ns."Code UM", 
               ns."QUẢN LÝ CẤP 1 (BDM)", ns."Column1" as "Code BDD", ns."QUẢN LÝ CẤP 2 (BDD)", "Ngày kết thúc hiệu lực" - 3 as "Hạn thanh toán", 
               case when "Số tiền tái tục" > 0 then 'Đã tái tục' when current_date > "Ngày kết thúc hiệu lực" then 'Quá hạn tái tục' 
                    when current_date >= ("Ngày kết thúc hiệu lực" - 3) - 7 then 'Gần đến hạn' else 'Chưa đến hạn tái tục' end as "Trạng thái"
        FROM Renewal_Base LEFT JOIN df_ns ns ON Renewal_Base."Code Sale" = upper(trim(leading '0' from ns."Điện thoại"))
    ),
    detail_tt AS (
        SELECT 
            "Code Sale", "NĐBH", "Ngày sinh NĐBH", "BMBH", coalesce("Đối tác nhà bảo hiểm", '') as "Đối tác nhà bảo hiểm", "Sản phẩm", 
            "Channel" as "Channel Sale", "Số hợp đồng", "Ngày bắt đầu hiệu lực", "Ngày kết thúc hiệu lực", "Số tiền thanh toán", 
            "Code" as "Code Nhân sự", "Họ tên" as "Tên Sale", "Chức danh", "Channel Nhân sự", "Code UM", 
            "QUẢN LÝ CẤP 1 (BDM)" as "QUẢN LÝ CẤP 1 (SM)", "Code BDD" as "Code SD", "QUẢN LÝ CẤP 2 (BDD)" as "QUẢN LÝ CẤP 2 (SD)", 
            "Code Sale tái tục", "BMBH tái tục", "Số hợp đồng tái tục", "Sản phẩm tái tục", "Channel tái tục", 
            "Ngày bắt đầu hiệu lực đơn tái tục", "Số tiền tái tục", extract(month from "Ngày kết thúc hiệu lực") as "Tháng", 
            "Ngày kết thúc hiệu lực" - 3 as "Hạn thanh toán", "Trạng thái" 
        FROM theodoi 
        WHERE "Channel Nhân sự" = 'Neo' OR "Channel" LIKE '%NEO%' OR "Channel" = 'CD' OR "QUẢN LÝ CẤP 2 (BDD)" = 'PHẠM TRƯỜNG KHÁNH'
    )
    SELECT * FROM detail_tt ORDER BY list_position(['Gần đến hạn', 'Quá hạn tái tục', 'Chưa đến hạn tái tục','Đã tái tục'], "Trạng thái"), "Ngày kết thúc hiệu lực" ASC
    """
    df_tt_detail = con.execute(q_theodoi).df()
    print(f"  ✓ Sheet 6: theo dõi tái tục detail ({len(df_tt_detail)} dòng)")
    con.register('detail_tt', df_tt_detail)

    # ------------------------------------------------------------------------
    # Sheet 7: THEO DÕI TÁI TỤC SM
    # ------------------------------------------------------------------------
    q_tt_sm = """
    SELECT 
        d."Tháng", bdm."Code", bdm."Họ tên", bdm."Chức danh", bdm."QUẢN LÝ CẤP 2 (BDD)" as "QUẢN LÝ CẤP 2 (SD)", 
        count(d."Số hợp đồng") as "Số lượng", sum(d."Số tiền thanh toán") as "Tổng phí BH cần tái tục", 
        sum(d."Số tiền tái tục") as "Tổng phí BH đã tái tục", 
        case when sum(d."Số tiền thanh toán") > 0 then round(sum(d."Số tiền tái tục") / sum(d."Số tiền thanh toán"), 4) else 0 end as "Tỉ lệ tái tục" 
    FROM detail_tt d 
    LEFT JOIN (SELECT DISTINCT "Code", "Họ tên", "Chức danh", "QUẢN LÝ CẤP 2 (BDD)" FROM df_ns WHERE "Channel" = 'Neo') as bdm 
      ON upper(bdm."Họ tên") = upper(d."QUẢN LÝ CẤP 1 (SM)") 
    GROUP BY 1, 2, 3, 4, 5
    ORDER BY d."Tháng", "Tổng phí BH cần tái tục" desc
    """
    df_tt_sm = con.execute(q_tt_sm).df()
    print(f"  ✓ Sheet 7: theo dõi tái tục SM ({len(df_tt_sm)} dòng)")

    # ------------------------------------------------------------------------
    # Sheet 8: THEO DÕI TÁI TỤC SD
    # ------------------------------------------------------------------------
    q_tt_sd = """
    SELECT 
        d."Tháng", sd."Code" as "Code SD", d."QUẢN LÝ CẤP 2 (SD)" as "Họ tên SD", sd."Chức danh", 
        count(d."Số hợp đồng") as "Số lượng", sum(d."Số tiền thanh toán") as "Tổng phí BH cần tái tục", 
        sum(d."Số tiền tái tục") as "Tổng phí BH đã tái tục", 
        case when sum(d."Số tiền thanh toán") > 0 then round(sum(d."Số tiền tái tục") / sum(d."Số tiền thanh toán"), 4) else 0 end as "Tỉ lệ tái tục" 
    FROM detail_tt d 
    LEFT JOIN (SELECT DISTINCT "Code", "Họ tên", "Chức danh" FROM df_ns WHERE "Channel" = 'Neo') as sd 
      ON upper(sd."Họ tên") = upper(d."QUẢN LÝ CẤP 2 (SD)")
    WHERE d."QUẢN LÝ CẤP 2 (SD)" IS NOT NULL
    GROUP BY 1, 2, 3, 4
    ORDER BY d."Tháng", "Tổng phí BH cần tái tục" desc
    """
    df_tt_sd = con.execute(q_tt_sd).df()
    print(f"  ✓ Sheet 8: theo dõi tái tục SD ({len(df_tt_sd)} dòng)")

    return {
        'detail': df_detail,
        'tracking AG': df_ag,
        'tracking SM': df_sm,
        'tracking SD': df_sd,
        'tracking CD': df_cd,
        'theo dõi tái tục detail': df_tt_detail,
        'theo dõi tái tục SM': df_tt_sm,
        'theo dõi tái tục SD': df_tt_sd
    }


# ============================================================================
# PHẦN 6: XUẤT EXCEL CHUYÊN NGHIỆP VỚI CÔNG THỨC DYNAMIC SUBTOTAL
# ============================================================================
def build_and_format_excel_report(sheets_dict, output_excel_path):
    """
    Tạo file Excel 8 sheet với:
    - Header bảng màu nhận diện chuẩn
    - AutoFilter mọi cột
    - Freeze panes hàng 1
    - Căn chỉnh & định dạng số/tiền/ngày/tỷ lệ
    - Dòng TỔNG CỘNG dùng công thức =SUBTOTAL(9, ...) tự động co giãn theo bộ lọc!
    """
    print(f"\n🎨 Đang định dạng Excel chuyên nghiệp & thêm công thức SUBTOTAL động...")
    wb = openpyxl.Workbook()
    wb.remove(wb.active)

    font_family = 'Calibri'
    font_title = Font(name=font_family, size=11, bold=True, color='FFFFFF')
    font_regular = Font(name=font_family, size=11)
    font_bold = Font(name=font_family, size=11, bold=True)

    thin_side = Side(style='thin', color='D9D9D9')
    thin_border = Border(left=thin_side, right=thin_side, top=thin_side, bottom=thin_side)

    tot_side_top = Side(style='thin', color='000000')
    tot_side_bot = Side(style='double', color='000000')
    tot_border = Border(left=thin_side, right=thin_side, top=tot_side_top, bottom=tot_side_bot)
    fill_tot = PatternFill(start_color='EAEDED', end_color='EAEDED', fill_type='solid')

    align_center = Alignment(horizontal='center', vertical='center')
    align_right = Alignment(horizontal='right', vertical='center')
    align_left = Alignment(horizontal='left', vertical='center')
    align_hdr = Alignment(horizontal='center', vertical='center', wrap_text=True)

    # Cấu hình màu sắc Header & các cột số cần tính SUBTOTAL cho từng sheet
    sheets_config = {
        'detail': {
            'hdr_color': '1F4E78',  # Navy Blue
            'label_col_name': 'Họ tên',
            'numeric_cols': ['Số tiền thanh toán', 'Doanh thu trước thuế', 'SM_OR', 'SD_OR', 'AG_bonus', 'SM_IO', 'SD_IO']
        },
        'tracking AG': {
            'hdr_color': '2E4053',  # Slate Steel
            'label_col_name': 'Họ tên',
            'numeric_cols': ['Số tiền thanh toán', 'Doanh số qui đổi', 'Doanh thu trước thuế', 'Doanh số hôm nay', 'AG_bonus', 'Số người giới thiệu']
        },
        'tracking SM': {
            'hdr_color': '117864',  # Emerald Teal
            'label_col_name': 'Họ tên',
            'numeric_cols': ['Số tiền thanh toán', 'Doanh thu trước thuế', 'AG_bonus', 'Doanh số hôm nay', 'SM_OR', 'Số AG thực tế', 'Tuyển dụng mới AG', 'AG Active', 'SM_IO']
        },
        'tracking SD': {
            'hdr_color': '7D6608',  # Antique Bronze
            'label_col_name': 'QUẢN LÝ CẤP 2 (SD)',
            'numeric_cols': ['Số tiền thanh toán', 'Doanh thu trước thuế', 'AG_bonus', 'SM_OR', 'SD_OR', 'Tuyển dụng mới AG', 'Số AG thực tế', 'AG Active', 'SM_IO', 'SD_IO']
        },
        'tracking CD': {
            'hdr_color': '78281F',  # Imperial Burgundy
            'label_col_name': 'SD phụ trách',
            'numeric_cols': ['Số tiền thanh toán', 'Doanh thu trước thuế', 'AG_bonus', 'SM_OR', 'SD_OR', 'Tuyển dụng mới AG', 'Số AG thực tế', 'AG Active', 'SM_IO', 'SD_IO']
        },
        'theo dõi tái tục detail': {
            'hdr_color': '4A235A',  # Royal Purple
            'label_col_name': 'NĐBH',
            'numeric_cols': ['Số tiền thanh toán', 'Số tiền tái tục']
        },
        'theo dõi tái tục SM': {
            'hdr_color': '1B4F72',  # Deep Sapphire
            'label_col_name': 'Họ tên',
            'numeric_cols': ['Số lượng', 'Tổng phí BH cần tái tục', 'Tổng phí BH đã tái tục']
        },
        'theo dõi tái tục SD': {
            'hdr_color': '0E6655',  # Forest Pine
            'label_col_name': 'Họ tên SD',
            'numeric_cols': ['Số lượng', 'Tổng phí BH cần tái tục', 'Tổng phí BH đã tái tục']
        }
    }

    for sheet_name, cfg in sheets_config.items():
        df = sheets_dict.get(sheet_name, pd.DataFrame())
        ws = wb.create_sheet(title=sheet_name)
        ws.views.sheetView[0].showGridLines = True
        ws.freeze_panes = 'A2'

        fill_hdr = PatternFill(start_color=cfg['hdr_color'], end_color=cfg['hdr_color'], fill_type='solid')
        num_cols = cfg['numeric_cols']

        # 1. Vẽ Header
        for c_idx, col_name in enumerate(df.columns, 1):
            cell = ws.cell(1, c_idx, col_name)
            cell.font = font_title
            cell.fill = fill_hdr
            cell.alignment = align_hdr
            cell.border = thin_border

        num_rows = len(df)
        end_data_r = max(num_rows + 1, 2)
        tot_r = end_data_r + 1

        # 2. Đổ dữ liệu
        for r_idx, row in df.iterrows():
            row_num = r_idx + 2
            for c_idx, col_name in enumerate(df.columns, 1):
                val = row[col_name]
                cell = ws.cell(row_num, c_idx)
                cell.font = font_regular
                cell.border = thin_border

                if col_name in num_cols:
                    if pd.notna(val) and str(val).strip() not in ['', 'nan', 'None']:
                        try:
                            cell.value = float(val)
                            cell.number_format = '#,##0'
                        except:
                            cell.value = val
                    else:
                        cell.value = 0
                        cell.number_format = '#,##0'
                    cell.alignment = align_right
                elif 'Tỉ lệ' in col_name or 'rate' in col_name.lower():
                    if pd.notna(val) and str(val).strip() not in ['', 'nan', 'None']:
                        try:
                            cell.value = float(val)
                            cell.number_format = '0.00%'
                        except:
                            cell.value = val
                    cell.alignment = align_right
                elif any(k in col_name.lower() for k in ['ngày', 'hạn', 'date']):
                    cell.value = str(val)[:10] if pd.notna(val) and str(val) != 'None' else ''
                    cell.alignment = align_center
                elif col_name in ['STT', 'Code', 'Code Sale', 'Code Nhân sự', 'Code UM', 'Code SD', 'Code CD', 'Tháng', 'Trạng thái', 'Chức danh']:
                    cell.value = str(val) if pd.notna(val) and str(val) != 'None' else ''
                    cell.alignment = align_center
                else:
                    cell.value = str(val) if pd.notna(val) and str(val) != 'None' else ''
                    cell.alignment = align_left

        # 3. Dòng TỔNG CỘNG với SUBTOTAL(9, ...)
        if num_rows > 0:
            # Xác định cột đặt nhãn "TỔNG CỘNG"
            label_col = 1
            if cfg['label_col_name'] in df.columns:
                label_col = list(df.columns).index(cfg['label_col_name']) + 1

            for c_idx in range(1, len(df.columns) + 1):
                cell = ws.cell(tot_r, c_idx)
                cell.fill = fill_tot
                cell.border = tot_border

            c_lbl = ws.cell(tot_r, label_col, "TỔNG CỘNG")
            c_lbl.font = font_bold
            c_lbl.alignment = align_center if label_col in [1, 2] else align_right

            for c_idx, col_name in enumerate(df.columns, 1):
                col_let = get_column_letter(c_idx)
                cell = ws.cell(tot_r, c_idx)

                if col_name in num_cols:
                    cell.value = f"=SUBTOTAL(9, {col_let}2:{col_let}{end_data_r})"
                    cell.number_format = '#,##0'
                    cell.font = font_bold
                    cell.alignment = align_right
                elif 'Tỉ lệ tái tục' in col_name:
                    if 'Tổng phí BH cần tái tục' in df.columns and 'Tổng phí BH đã tái tục' in df.columns:
                        can_idx = list(df.columns).index('Tổng phí BH cần tái tục') + 1
                        da_idx = list(df.columns).index('Tổng phí BH đã tái tục') + 1
                        can_let = get_column_letter(can_idx)
                        da_let = get_column_letter(da_idx)
                        cell.value = f'=IF({can_let}{tot_r}>0, {da_let}{tot_r}/{can_let}{tot_r}, 0)'
                        cell.number_format = '0.00%'
                        cell.font = font_bold
                        cell.alignment = align_right

        # 4. Kích hoạt AutoFilter
        last_col_let = get_column_letter(len(df.columns))
        ws.auto_filter.ref = f"A1:{last_col_let}{end_data_r}"

        # 5. Tự động căn chỉnh độ rộng cột
        for c_idx, col_name in enumerate(df.columns, 1):
            col_let = get_column_letter(c_idx)
            max_len = len(str(col_name))
            sample_vals = df[col_name].dropna().astype(str).tolist()[:50]
            if sample_vals:
                max_val_len = max(len(s) for s in sample_vals)
                max_len = max(max_len, max_val_len)
            ws.column_dimensions[col_let].width = min(max(max_len + 4, 12), 40)

    wb.save(output_excel_path)
    print(f"  ✅ Đã lưu file Excel hoàn chỉnh vào: {output_excel_path}")


# ============================================================================
# MAIN ENTRY POINT
# ============================================================================
def main():
    print("=" * 80)
    print("CHƯƠNG TRÌNH DAILY REPORT KÊNH NEO (AFFINA)")
    print(f"Thời gian: Ngày {datetime.now():%d/%m/%Y %H:%M:%S} | Tháng báo cáo: {REPORT_MONTH}/{REPORT_YEAR}")
    print("=" * 80)

    # 1. Khởi tạo Google Services
    drive_service, sheets_service = init_google_services()

    # 2. Thiết lập đường dẫn các file dữ liệu input
    local_capdon_path = os.path.join(WORK_DIR, 'CapDon_Auto.xlsx')
    local_nhansu_path = os.path.join(WORK_DIR, 'DSNS_CTV_sale_Affina.xlsx')
    local_quydoi_path = os.path.join(WORK_DIR, 'quy_doi_all.xlsx')

    # 2.1 Tải Cấp đơn từ Google Sheet nếu có kết nối
    target_capdon_sheets = [
        'Sức khỏe', 'Thông tin cấp Bảo hiểm xe máy', 'BHYT/BHXH',
        'Bao hiem oto', 'Du lịch', 'Trách nhiệm sản phẩm', 'Bảo hiểm rủi ro'
    ]
    if sheets_service:
        export_google_sheet_data(sheets_service, SHEET_CAPDON_ID, target_capdon_sheets, local_capdon_path)
    elif not os.path.exists(local_capdon_path):
        # Fallback tìm file local
        fallback_cd = r"C:\Users\ADMIN\Desktop\AFFINA\CODE\build_test\CapDon_Auto.xlsx"
        if os.path.exists(fallback_cd):
            import shutil
            shutil.copyfile(fallback_cd, local_capdon_path)
            print(f"  ℹ Dùng file Cấp đơn fallback: {fallback_cd}")
        else:
            raise FileNotFoundError(f"Không tìm thấy dữ liệu Cấp đơn!")

    # 2.2 Tải Nhân sự
    if drive_service:
        download_drive_file(drive_service, DSNS_FILE_ID, local_nhansu_path)
    elif not os.path.exists(local_nhansu_path):
        fallback_ns = r"C:\Users\ADMIN\OneDrive\Nhân sự sales\DSNS CTV sale Affina FINAL V2.xlsx"
        if os.path.exists(fallback_ns):
            import shutil
            shutil.copyfile(fallback_ns, local_nhansu_path)
            print(f"  ℹ Dùng file Nhân sự fallback: {fallback_ns}")
        else:
            raise FileNotFoundError("Không tìm thấy dữ liệu Nhân sự!")

    # 2.3 Tải Quy đổi
    if drive_service:
        download_drive_file(drive_service, QUYDOI_FILE_ID, local_quydoi_path)
    elif not os.path.exists(local_quydoi_path):
        fallback_qd = r"C:\Users\ADMIN\Desktop\AFFINA\DA\DATA_3_input\Copy of 26_02_04_sửa ngày_quy_doi_all.xlsx"
        if os.path.exists(fallback_qd):
            import shutil
            shutil.copyfile(fallback_qd, local_quydoi_path)
            print(f"  ℹ Dùng file Quy đổi fallback: {fallback_qd}")
        else:
            raise FileNotFoundError("Không tìm thấy dữ liệu Quy đổi!")

    # 3. Làm sạch dữ liệu
    df_union, df_nhansu, df_quydoi = load_and_clean_all_data(
        local_capdon_path, local_nhansu_path, local_quydoi_path
    )

    # 4. Chạy truy vấn DuckDB tạo 8 DataFrame
    sheets_dict = execute_all_reports(
        df_union, df_nhansu, df_quydoi, REPORT_YEAR, REPORT_MONTH
    )

    # 5. Xuất báo cáo Excel với đầy đủ 8 sheet và định dạng chuyên nghiệp
    today_tag = datetime.now().strftime("%d%m")
    out_filename = f"{today_tag}_Daily_Report_NEO_Thang_{REPORT_MONTH}_{REPORT_YEAR}.xlsx"
    out_excel_path = os.path.join(OUTPUT_DIR, out_filename)

    build_and_format_excel_report(sheets_dict, out_excel_path)

    # 6. Upload file kết quả lên Google Drive
    if drive_service:
        upload_to_drive(drive_service, out_excel_path, DRIVE_FOLDER_ID, out_filename)


    print("\n" + "=" * 80)
    print(f"🎉 BÁO CÁO DAILY NEO ĐÃ HOÀN TẤT XUẤT THÀNH CÔNG!")
    print(f"📁 Đường dẫn file Excel: {out_excel_path}")
    print(f"📊 Các sheet đã tạo: {list(sheets_dict.keys())}")
    print("=" * 80)


if __name__ == '__main__':
    main()
