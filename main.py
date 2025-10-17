import os
import io
import csv
import requests
import re
from bs4 import BeautifulSoup
from datetime import datetime, timedelta, timezone
from google.oauth2 import service_account
from googleapiclient.discovery import build
from googleapiclient.http import MediaIoBaseUpload, MediaIoBaseDownload

# ================================
# ① Google Drive サービスアカウント設定
# ================================
SCOPES = ['https://www.googleapis.com/auth/drive']
CREDENTIALS_FILE = os.path.join(os.path.dirname(__file__), "credentials.json")

creds = service_account.Credentials.from_service_account_file(
    CREDENTIALS_FILE, scopes=SCOPES
)
drive_service = build("drive", "v3", credentials=creds)

# ================================
# ② Driveのファイル設定vhttps://drive.google.com/file/d/1EJeJ5215gngU_7YxlFnj_3kFA4q--Koe/view?usp=drive_link
# ================================
FILE_ID = "1EJeJ5215gngU_7YxlFnj_3kFA4q--Koe" \
""  # ←追記したいCSVファイルのID
print("Google DriveファイルID:", FILE_ID)

# ================================
# ③ 既存CSVをダウンロード
# ================================
print("既存ファイルを取得中...")
request = drive_service.files().get_media(fileId=FILE_ID)
file_data = io.BytesIO()
downloader = MediaIoBaseDownload(file_data, request)
done = False
while not done:
    status, done = downloader.next_chunk()
file_data.seek(0)

# CSVをリストとして読み込み
existing_rows = []
try:
    text_wrapper = io.TextIOWrapper(file_data, encoding="utf-8-sig")
    reader = csv.reader(text_wrapper)
    existing_rows = list(reader)
except Exception as e:
    print("⚠️ CSVの読み込みに失敗:", e)
    existing_rows = []

# ================================
# ④ スクレイピング
# ================================
url = "https://mimorin2014.com/"
res = requests.get(url)
res.encoding = res.apparent_encoding
soup = BeautifulSoup(res.text, "html.parser")

entry = soup.select_one(".entry-body") or soup.select_one("#main") or soup
text = entry.get_text("\n", strip=True)

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

# ================================
# ⑤ 追記処理
# ================================
JST = timezone(timedelta(hours=9))
now = datetime.now(JST).strftime("%Y/%m/%d %H:%M")

# 空ファイルならヘッダーを追加
if not existing_rows:
    existing_rows.append(["取得日時"] + columns)

# データに取得日時を追加
for row in data:
    existing_rows.append([now] + row)

# ================================
# ⑥ CSVを再アップロード（上書き・ファイル名保持）
# ================================
# 元のファイル名を取得
file_metadata = drive_service.files().get(fileId=FILE_ID, fields="name").execute()
original_name = file_metadata["name"]

# CSVを書き出し
output = io.BytesIO()
text_wrapper = io.TextIOWrapper(output, encoding="utf-8-sig", newline="")
writer = csv.writer(text_wrapper)
writer.writerows(existing_rows)
text_wrapper.flush()  # ←これでBytesIOに反映
output.seek(0)

# Drive上書き（ファイル名保持）
media = MediaIoBaseUpload(output, mimetype="text/csv", resumable=True)
updated_file = drive_service.files().update(
    fileId=FILE_ID,
    media_body=media,
    body={"name": original_name}
).execute()

print(f"\n✅ Drive上のファイルに追記完了: https://drive.google.com/file/d/{FILE_ID}/view")
print(f"追記データ件数: {len(data)} 件")
