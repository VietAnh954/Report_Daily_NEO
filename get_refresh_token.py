"""
get_refresh_token.py — Chạy 1 LẦN trên máy bạn để lấy Refresh Token
=====================================================================
Bước 1: Vào Google Cloud Console → APIs & Services → Credentials
Bước 2: Click "+ CREATE CREDENTIALS" → "OAuth client ID"
Bước 3: Application type: "Desktop app" → Name: "Daily Report" → Create
Bước 4: Download file JSON → đổi tên thành "oauth_credentials.json"
Bước 5: Đặt file JSON cạnh file này rồi chạy: python get_refresh_token.py
Bước 6: Browser mở ra → đăng nhập Gmail → cho phép quyền
Bước 7: Copy refresh_token hiện ra → dán vào GitHub Secrets
"""

import json
from google_auth_oauthlib.flow import InstalledAppFlow

SCOPES = [
    "https://www.googleapis.com/auth/drive",
    "https://www.googleapis.com/auth/spreadsheets.readonly",
]

flow = InstalledAppFlow.from_client_secrets_file("oauth_credentials.json", SCOPES)
creds = flow.run_local_server(port=0)

print("\n" + "=" * 60)
print("✅ THÀNH CÔNG! Copy 3 giá trị sau vào GitHub Secrets:")
print("=" * 60)

# Đọc client_id và client_secret từ file credentials
with open("oauth_credentials.json") as f:
    oauth_info = json.load(f)
    installed = oauth_info.get("installed", oauth_info.get("web", {}))

print(f"\nGOOGLE_CLIENT_ID:\n{installed['client_id']}")
print(f"\nGOOGLE_CLIENT_SECRET:\n{installed['client_secret']}")
print(f"\nGOOGLE_REFRESH_TOKEN:\n{creds.refresh_token}")
print("\n" + "=" * 60)
