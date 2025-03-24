import gspread
from google.oauth2 import service_account
import os
from google.cloud import storage

# 認証情報の設定
SCOPES = [
    'https://www.googleapis.com/auth/spreadsheets',
    'https://www.googleapis.com/auth/drive'
]
SERVICE_ACCOUNT_FILE = 'config/credentials.json'  # JSONファイルのパス

# 認証情報の読み込み
credentials = service_account.Credentials.from_service_account_file(
    SERVICE_ACCOUNT_FILE, scopes=SCOPES)

# gspreadクライアントの作成
gc = gspread.authorize(credentials)

# スプレッドシートを開く（実際のスプレッドシート名を指定）
try:
    # IDでスプレッドシートを開く
    spreadsheet = gc.open_by_key('1dTnuPyLHjYANic63d24fWiugxQi_qA0QKFlE0Jd-nXE')
    worksheet = spreadsheet.sheet1
    print("スプレッドシートへの接続成功！")
    print(f"シート名: {worksheet.title}")
    print(f"行数: {worksheet.row_count}")
except Exception as e:
    print("スプレッドシートへの接続エラー:", e)

# Google Cloud Storageの設定
os.environ['GOOGLE_APPLICATION_CREDENTIALS'] = SERVICE_ACCOUNT_FILE

# ストレージクライアントの作成
try:
    storage_client = storage.Client()
    print("Google Cloud Storage接続成功！")
    # 以下の行はバケットが実際に存在する場合のみコメントを外してください
    # bucket = storage_client.bucket('your-bucket-name')
    # print(f"バケット {bucket.name} にアクセスできます")
except Exception as e:
    print("Google Cloud Storage接続エラー:", e)