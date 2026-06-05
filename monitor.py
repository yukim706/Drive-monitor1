# ============================================================
# Googleドライブ ファイル監視スクリプト（GitHub Actions用）
# 【認証方式】サービスアカウント JSON（環境変数 GOOGLE_SA_JSON）
# 【メール送信】GAS WebアプリへのHTTPリクエスト（環境変数 GAS_MAIL_URL）
# 【セキュリティ】GAS_SECRET_TOKEN によるトークン検証付き
# ============================================================

import subprocess
subprocess.run(['pip', 'install', '--quiet', 'gspread', 'google-auth',
                'google-api-python-client'], check=True)

import gspread
from google.oauth2.service_account import Credentials
from googleapiclient.discovery import build
from datetime import datetime, timezone, timedelta
import json
import os
import urllib.request

# ── タイムゾーン（JST）──────────────────────────────────────
JST = timezone(timedelta(hours=9))

# ── 認証（サービスアカウント）────────────────────────────────
SCOPES = [
    'https://www.googleapis.com/auth/drive.readonly',
    'https://www.googleapis.com/auth/spreadsheets',
]
sa_info = json.loads(os.environ['GOOGLE_SA_JSON'])
creds = Credentials.from_service_account_info(sa_info, scopes=SCOPES)

gc            = gspread.Client(auth=creds)
drive_service = build('drive', 'v3', credentials=creds)

# ── 設定値（環境変数 or デフォルト値）────────────────────────
FOLDER_ID        = os.environ.get('FOLDER_ID',      '1m7VuFYKcticG68MkGG8yT6u2NA8K-TES')
SPREADSHEET_ID   = os.environ.get('SPREADSHEET_ID', '146fJr4d1TL1PWx_jGwNpznNzqp5Q_BwBH2_jutdjuhs')
GAS_MAIL_URL     = os.environ['GAS_MAIL_URL']
GAS_SECRET_TOKEN = os.environ['GAS_SECRET_TOKEN']
EMAIL_TO         = os.environ.get('EMAIL_TO', 'yukimgidai2020@gmail.com')
HISTORY_SHEET    = '削除or名前変更履歴'
FILELIST_SHEET   = 'ファイルリスト'

# ── MIMEタイプ → 拡張子 変換辞書 ─────────────────────────────
MIME_EXTENSIONS = {
    'application/vnd.google-apps.document':     '.gdoc',
    'application/vnd.google-apps.spreadsheet':  '.gsheet',
    'application/vnd.google-apps.presentation': '.gslides',
    'application/vnd.google-apps.drawing':      '.gdraw',
    'application/pdf':                          '',
    'application/vnd.openxmlformats-officedocument.wordprocessingml.document':   '.docx',
    'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet':         '.xlsx',
    'application/vnd.openxmlformats-officedocument.presentationml.presentation': '.pptx',
    'image/jpeg': '.jpg',
    'image/png':  '.png',
}

# ── Drive API でフォルダ内ファイルを再帰取得 ──────────────────
def get_all_files(folder_id, folder_path, result=None):
    if result is None:
        result = []

    page_token = None
    while True:
        response = drive_service.files().list(
            q=f"'{folder_id}' in parents and mimeType != 'application/vnd.google-apps.folder' and trashed = false",
            fields='nextPageToken, files(name, mimeType)',
            pageSize=1000,
            pageToken=page_token
        ).execute()
        for f in response.get('files', []):
            ext = MIME_EXTENSIONS.get(f['mimeType'], '')
            result.append({'fileName': f['name'] + ext, 'folderPath': folder_path})
        page_token = response.get('nextPageToken')
        if not page_token:
            break

    sub_page_token = None
    while True:
        sub_response = drive_service.files().list(
            q=f"'{folder_id}' in parents and mimeType = 'application/vnd.google-apps.folder' and trashed = false",
            fields='nextPageToken, files(id, name)',
            pageSize=1000,
            pageToken=sub_page_token
        ).execute()
        for sub in sub_response.get('files', []):
            get_all_files(sub['id'], folder_path + '/' + sub['name'], result)
        sub_page_token = sub_response.get('nextPageToken')
        if not sub_page_token:
            break

    return result

# ── スプレッドシート操作ヘルパー ──────────────────────────────
def get_or_create_sheet(spreadsheet, sheet_name, headers=None):
    try:
        return spreadsheet.worksheet(sheet_name)
    except gspread.exceptions.WorksheetNotFound:
        sheet = spreadsheet.add_worksheet(title=sheet_name, rows=1, cols=10)
        if headers:
            sheet.append_row(headers)
        return sheet

def load_previous_files(spreadsheet):
    sheet = get_or_create_sheet(spreadsheet, FILELIST_SHEET)
    records = sheet.get_all_values()
    if records and records[0] and records[0][0]:
        try:
            return json.loads(records[0][0])
        except json.JSONDecodeError:
            return None
    return None

def save_current_files(spreadsheet, file_list):
    sheet = get_or_create_sheet(spreadsheet, FILELIST_SHEET)
    sheet.clear()
    sheet.update(range_name='A1', values=[[json.dumps(file_list, ensure_ascii=False)]])
    print(f"  ファイルリスト保存完了：{len(file_list)} 件")

def record_deletion_history(spreadsheet, deleted_files):
    sheet = get_or_create_sheet(
        spreadsheet, HISTORY_SHEET,
        headers=['削除or名前変更日時', 'ファイル名', 'フォルダパス']
    )
    timestamp = datetime.now(JST).strftime('%Y/%m/%d %H:%M:%S JST')
    sheet.append_rows([[timestamp, f['fileName'], f['folderPath']] for f in deleted_files])
    print(f"  履歴シートに {len(deleted_files)} 件記録しました")

# ── メール送信（GAS Webアプリ経由・トークン検証付き）──────────
def send_email(deleted_files):
    body = '以下のファイルが削除or名前変更されました:\n' + '\n'.join(
        f"  {f['fileName']}（フォルダ: {f['folderPath']}）" for f in deleted_files
    )
    payload = json.dumps({
        'token':   GAS_SECRET_TOKEN,
        'to':      EMAIL_TO,
        'subject': 'ファイルが削除or名前変更されました',
        'body':    body,
    }).encode('utf-8')

    req = urllib.request.Request(
        GAS_MAIL_URL,
        data=payload,
        headers={'Content-Type': 'application/json'},
        method='POST'
    )
    try:
        with urllib.request.urlopen(req, timeout=30) as res:
            result = res.read().decode('utf-8')
            print(f"  メール送信完了 → {EMAIL_TO}（GAS応答: {result}）")
    except Exception as e:
        print(f"  ⚠️ メール送信失敗: {e}")

# ── メイン処理 ────────────────────────────────────────────────
def monitor_folder():
    print("=" * 50)
    print(f"監視開始：{datetime.now(JST).strftime('%Y/%m/%d %H:%M:%S JST')}")
    print("=" * 50)

    spreadsheet = gc.open_by_key(SPREADSHEET_ID)

    print("\n[1] ファイル一覧取得中...")
    root_name = drive_service.files().get(
        fileId=FOLDER_ID, fields='name'
    ).execute().get('name', 'Root')
    current_files = get_all_files(FOLDER_ID, root_name)
    print(f"  現在のファイル数：{len(current_files)} 件")

    print("\n[2] 前回リストと比較中...")
    previous_files = load_previous_files(spreadsheet)

    if previous_files is None:
        print("  前回リストなし → 今回のリストを初回保存します")
    else:
        print(f"  前回のファイル数：{len(previous_files)} 件")
        current_set   = {(f['fileName'], f['folderPath']) for f in current_files}
        deleted_files = [f for f in previous_files
                         if (f['fileName'], f['folderPath']) not in current_set]

        if deleted_files:
            print(f"\n  ⚠️ 削除or名前変更：{len(deleted_files)} 件検出")
            for f in deleted_files:
                print(f"    - {f['fileName']}（{f['folderPath']}）")
            print("\n[3] 履歴シートに記録中...")
            record_deletion_history(spreadsheet, deleted_files)
            print("\n[4] メール通知送信中...")
            send_email(deleted_files)
        else:
            print("  変更なし（削除・名前変更は検出されませんでした）")

    print("\n[5] 現在のファイルリストを保存中...")
    save_current_files(spreadsheet, current_files)
    print("\n✅ 監視処理完了")
    print("=" * 50)

# ── 実行 ──────────────────────────────────────────────────────
monitor_folder()
