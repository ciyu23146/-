import requests
from bs4 import BeautifulSoup
import re
import csv
from datetime import datetime
from zoneinfo import ZoneInfo
import os

from google.oauth2 import service_account
from googleapiclient.discovery import build
from googleapiclient.http import MediaFileUpload

# ================================
# ① Google Drive API 設定
# ================================
SERVICE_ACCOUNT_FILE = r"C:\path\to\your_service_account.json"  # ←秘密鍵JSONファイルの絶対パス
SCOPES = ['https://www.googleapis.com/auth/drive.file']  # Driveファイルへのアクセス権

# Google Drive API認証
creds = service_account.Credentials.from_service_account_file(
    SERVICE_ACCOUNT_FILE, scopes=SCOPES)
drive_service = build('drive', 'v3', credentials=creds)

# アップロード先フォルダID（Google Drive上のフォルダを使いたい場合）
# 例）URLが https://drive.google.com/drive/folders/XXXXXXXXXXXX の場合
FOLDER_ID = 'XXXXXXXXXXXX'  # ←任意のフォルダIDに置き換える

# ================================
# ② 保存先フォルダ（ローカル）
# ================================
save_dir = r"C:\Users\あなたのユーザー名\Documents\スクレイピング"
os.makedirs(save_dir, exist_ok=True)

# ================================
# ③ ファイル名生成
# ================================
user_input = input("ファイル名に入れるキーワードを入力してください: ").strip()
now = datetime.now(ZoneInfo("Asia/Tokyo"))
date_str = now.strftime("%y%m%d")
time_str = now.strftime("%H%M")
filename = f"{date_str}_{time_str}_{user_input}.csv"
save_path = os.path.join(save_dir, filename)

# ================================
# ④ スクレイピング
# ================================
url = "https://mimorin2014.com/"
res = requests.get(url)
res.encoding = res.apparent_encoding
soup = BeautifulSoup(res.text, "html.parser")

entry = soup.select_one(".entry-body") or soup.select_one("#main") or soup
text = entry.get_text("\n", strip=True)

# --- 正規表現パターン ---
pattern = re.compile(
    r"\*?(\d+)[\s　]+"
    r"([\d\*]+)[\s　]+"
    r"([\d\*]+)[\s　]+"
    r"([\d\.]+)[\s　]+"
    r"([\d\.]+)[\s　]+"
    r"([\d\*]+)[\s　]+"
    r"([\d\*]+)[\s　]+"
    r"([\d\.%*]+)[\s　]+"
    r"(.+)"
)

columns = ["順位","座席数","回数","箱平均","番箱","取得館","上映館","取得率","タイトル"]
data = []

last_rank = 0
for match in pattern.finditer(text):
    row = list(match.groups())
    rank = int(row[0])
    if rank == 1 and last_rank > 1:
        break
    data.append(row)
    last_rank = rank

# --- 整形して表示 ---
print("\t".join(columns))
for row in data[:25]:
    print("\t".join(row))

# --- CSV保存 ---
with open(save_path, "w", newline="", encoding="utf-8-sig") as f:
    writer = csv.writer(f)
    writer.writerow(columns)
    writer.writerows(data)

print(f"\n✅ ローカルに保存完了: {save_path}")
print(f"データ件数: {len(data)} 件")

# ================================
# ⑤ Google Drive へアップロード
# ================================
print("Google Driveにアップロード中...")

file_metadata = {'name': filename}
if FOLDER_ID:
    file_metadata['parents'] = [FOLDER_ID]

media = MediaFileUpload(save_path, mimetype='text/csv')
uploaded_file = drive_service.files().create(
    body=file_metadata,
    media_body=media,
    fields='id'
).execute()

file_id = uploaded_file.get('id')
drive_url = f"https://drive.google.com/file/d/{file_id}/view"

print(f"✅ Google Driveにアップロード完了: {drive_url}")
