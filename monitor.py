# ============================================================
# Googleドライブ ファイル監視スクリプト（GitHub Actions用）
# 【認証方式】サービスアカウント JSON（環境変数 GOOGLE_SA_JSON）
# 【メール送信】GAS WebアプリへのHTTPリクエスト（環境変数 GAS_MAIL_URL）
# 【セキュリティ】GAS_SECRET_TOKEN によるトークン検証付き
# 【検知内容】削除 / 名前変更 / 新規追加
# 【シート構成】1シート（ファイルリスト兼変更履歴）
#   A:更新日時 B:ファイル名 C:変更前ファイル名 D:URL E:ファイルID F:種別 G:フォルダパス
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

# ── タイムゾーン（JST）
JST = timezone(timedelta(hours=9))

# ── 認証（サービスアカウント）
SCOPES = [
    'https://www.googleapis.com/auth/drive.readonly',
    'https://www.googleapis.com/auth/spreadsheets',
]
sa_info = json.loads(os.environ['GOOGLE_SA_JSON'])
creds = Credentials.from_service_account_info(sa_info, scopes=SCOPES)

gc            = gspread.Client(auth=creds)
drive_service = build('drive', 'v3', credentials=creds)

# ── 設定値
FOLDER_ID        = os.environ.get('FOLDER_ID',      '1tVQU7ufn_Ob54kspgh88iK03-wUPF7mU')
SPREADSHEET_ID   = os.environ.get('SPREADSHEET_ID', '146fJr4d1TL1PWx_jGwNpznNzqp5Q_BwBH2_jutdjuhs')
GAS_MAIL_URL     = os.environ['GAS_MAIL_URL']
GAS_SECRET_TOKEN = os.environ['GAS_SECRET_TOKEN']
EMAIL_TO         = os.environ.get('EMAIL_TO', 'yukimgidai2020@gmail.com')
FILE_SHEET       = 'ファイル一覧'

# 列番号定数（1始まり）
COL_UPDATED    = 1   # A: 更新日時
COL_FILENAME   = 2   # B: ファイル名
COL_BEFORE     = 3   # C: 変更前ファイル名
COL_FILEURL    = 4   # D: URL
COL_FILEID     = 5   # E: ファイルID
COL_STATUS     = 6   # F: 種別
COL_FOLDERPATH = 7   # G: フォルダパス
TOTAL_COLS     = 7

HEADERS = ['更新日時', 'ファイル名', '変更前ファイル名', 'URL', 'ファイルID', '種別', 'フォルダパス']

# ── 曜日（日本語）
WEEKDAYS_JP = ['月', '火', '水', '木', '金', '土', '日']

def format_datetime_jp(dt):
    wd = WEEKDAYS_JP[dt.weekday()]
    return f"{dt.year}年{dt.month}月{dt.day}日({wd}) {dt.strftime('%H:%M')}"

# ── MIMEタイプ → 拡張子
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

# ── MIMEタイプ → URL生成
def build_file_url(file_id, mime_type):
    urls = {
        'application/vnd.google-apps.document':     f'https://docs.google.com/document/d/{file_id}/edit',
        'application/vnd.google-apps.spreadsheet':  f'https://docs.google.com/spreadsheets/d/{file_id}/edit',
        'application/vnd.google-apps.presentation': f'https://docs.google.com/presentation/d/{file_id}/edit',
        'application/vnd.google-apps.drawing':      f'https://docs.google.com/drawings/d/{file_id}/edit',
    }
    return urls.get(mime_type, f'https://drive.google.com/file/d/{file_id}/view')

# ── Drive API でフォルダ内ファイルを再帰取得
def get_all_files(folder_id, folder_path, result=None):
    if result is None:
        result = []

    page_token = None
    while True:
        response = drive_service.files().list(
            q=f"'{folder_id}' in parents and mimeType != 'application/vnd.google-apps.folder' and trashed = false",
            fields='nextPageToken, files(id, name, mimeType)',
            pageSize=1000,
            pageToken=page_token
        ).execute()
        for f in response.get('files', []):
            mime = f['mimeType']
            ext  = MIME_EXTENSIONS.get(mime, '')
            result.append({
                'fileId':     f['id'],
                'fileName':   f['name'] + ext,
                'mimeType':   mime,
                'fileUrl':    build_file_url(f['id'], mime),
                'folderPath': folder_path,
            })
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

# ── シート取得 or 作成
def get_or_create_sheet(spreadsheet, sheet_name):
    try:
        return spreadsheet.worksheet(sheet_name)
    except gspread.exceptions.WorksheetNotFound:
        sheet = spreadsheet.add_worksheet(title=sheet_name, rows=5000, cols=TOTAL_COLS)
        sheet.append_row(HEADERS)
        return sheet

# ── シートから fileId → 行番号 の辞書を作成（E列＝COL_FILEID=5）
def build_id_to_row_map(sheet):
    all_values = sheet.get_all_values()
    id_to_row = {}
    for i, row in enumerate(all_values):
        if i == 0:
            continue  # ヘッダースキップ
        if len(row) >= COL_FILEID and row[COL_FILEID - 1]:
            id_to_row[row[COL_FILEID - 1]] = i + 1  # 1始まり行番号
    return id_to_row

# ── 行を上書き更新（A:更新日時 B:ファイル名 C:変更前 D:URL E:ファイルID F:種別 G:フォルダパス）
def update_row(sheet, row_num, now_str, file_data, status, before_name=''):
    values = [
        now_str,
        file_data.get('fileName', ''),
        before_name,
        file_data.get('fileUrl', ''),
        file_data.get('fileId', ''),
        status,
        file_data.get('folderPath', ''),
    ]
    range_str = f'A{row_num}:G{row_num}'
    sheet.update(range_name=range_str, values=[values])

    # 削除の場合は行全体を赤文字にする
    if status == '削除':
        sheet.format(range_str, {
            'textFormat': {
                'foregroundColor': {'red': 1.0, 'green': 0.0, 'blue': 0.0}
            }
        })
    else:
        # 削除以外は黒文字に戻す
        sheet.format(range_str, {
            'textFormat': {
                'foregroundColor': {'red': 0.0, 'green': 0.0, 'blue': 0.0}
            }
        })

# ── 新規行を追加
def append_new_row(sheet, now_str, file_data, status):
    row = [
        now_str,
        file_data.get('fileName', ''),
        '',
        file_data.get('fileUrl', ''),
        file_data.get('fileId', ''),
        status,
        file_data.get('folderPath', ''),
    ]
    sheet.append_row(row, value_input_option='USER_ENTERED')

# ── 変更の判定
def detect_changes(previous_files, current_files):
    prev_by_id = {f['fileId']: f for f in previous_files}
    curr_by_id = {f['fileId']: f for f in current_files}

    deleted = []
    renamed = []
    added   = []

    for fid, pf in prev_by_id.items():
        if fid not in curr_by_id:
            deleted.append(pf)
        elif curr_by_id[fid]['fileName'] != pf['fileName']:
            cf = curr_by_id[fid]
            renamed.append({
                'fileId':     fid,
                'fileName':   cf['fileName'],
                'beforeName': pf['fileName'],
                'fileUrl':    cf.get('fileUrl', ''),
                'folderPath': cf.get('folderPath', ''),
                'mimeType':   cf.get('mimeType', ''),
            })

    for fid, cf in curr_by_id.items():
        if fid not in prev_by_id:
            added.append(cf)

    return deleted, renamed, added

# ── メール本文
def build_email_body(deleted, renamed, added):
    now_str = format_datetime_jp(datetime.now(JST))
    lines = [f"検知日時：{now_str}", '']

    if deleted:
        lines += ['=' * 40, '❌ 削除', '=' * 40]
        for f in deleted:
            lines += [
                f"  ファイル名  ：{f.get('fileName', '')}",
                f"  URL        ：{f.get('fileUrl', '')}",
                f"  フォルダパス：{f.get('folderPath', '')}",
                f"  種別       ：削除", ''
            ]

    if renamed:
        lines += ['=' * 40, '■ ファイル名変更', '=' * 40]
        for f in renamed:
            lines += [
                f"  ファイル名      ：{f.get('fileName', '')}",
                f"  変更前ファイル名：{f.get('beforeName', '')}",
                f"  URL            ：{f.get('fileUrl', '')}",
                f"  フォルダパス   ：{f.get('folderPath', '')}",
                f"  種別           ：ファイル名変更", ''
            ]

    if added:
        lines += ['=' * 40, '◉ 新規追加', '=' * 40]
        for f in added:
            lines += [
                f"  ファイル名  ：{f.get('fileName', '')}",
                f"  URL        ：{f.get('fileUrl', '')}",
                f"  フォルダパス：{f.get('folderPath', '')}",
                f"  種別       ：新規", ''
            ]

    return '\n'.join(lines)

# ── メール送信
def send_email(deleted, renamed, added):
    parts = []
    if deleted: parts.append(f'削除{len(deleted)}件')
    if renamed: parts.append(f'名前変更{len(renamed)}件')
    if added:   parts.append(f'新規追加{len(added)}件')

    subject = f'【ドライブ監視】{" / ".join(parts)} が検出されました'
    body    = build_email_body(deleted, renamed, added)

    payload = json.dumps({
        'token':   GAS_SECRET_TOKEN,
        'to':      EMAIL_TO,
        'subject': subject,
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

# ── メイン処理
def monitor_folder():
    print("=" * 50)
    print(f"監視開始：{format_datetime_jp(datetime.now(JST))}")
    print("=" * 50)

    spreadsheet = gc.open_by_key(SPREADSHEET_ID)
    sheet = get_or_create_sheet(spreadsheet, FILE_SHEET)

    print("\n[1] ファイル一覧取得中...")
    root_name = drive_service.files().get(
        fileId=FOLDER_ID, fields='name'
    ).execute().get('name', 'Root')
    current_files = get_all_files(FOLDER_ID, root_name)
    print(f"  現在のファイル数：{len(current_files)} 件")

    # ── シートに既存データがあるか確認
    id_to_row = build_id_to_row_map(sheet)
    is_first_run = len(id_to_row) == 0

    now_str = format_datetime_jp(datetime.now(JST))

    if is_first_run:
        # ────────────────────────────────────────
        # 初回実行：全ファイルを一括登録
        # ────────────────────────────────────────
        print("\n[2] 初回実行 → 全ファイルを登録中...")
        rows = []
        for f in current_files:
            rows.append([
                now_str,
                f.get('fileName', ''),
                '',
                f.get('fileUrl', ''),
                f.get('fileId', ''),
                '正常',
                f.get('folderPath', ''),
            ])
        if rows:
            sheet.append_rows(rows, value_input_option='USER_ENTERED')
        print(f"  {len(rows)} 件を登録しました")

    else:
        # ────────────────────────────────────────
        # 2回目以降：変更検知 → 該当行を上書き
        # ────────────────────────────────────────
        print("\n[2] 前回リストと比較中...")

        # 前回データをシートから再構築（E列のfileIdを使用）
        all_values = sheet.get_all_values()
        previous_files = []
        for i, row in enumerate(all_values):
            if i == 0:
                continue
            if len(row) >= COL_FILEID and row[COL_FILEID - 1]:
                previous_files.append({
                    'fileId':     row[COL_FILEID - 1],
                    'fileName':   row[COL_FILENAME - 1]   if len(row) >= COL_FILENAME   else '',
                    'folderPath': row[COL_FOLDERPATH - 1] if len(row) >= COL_FOLDERPATH else '',
                })

        deleted, renamed, added = detect_changes(previous_files, current_files)

        # 削除
        if deleted:
            print(f"  ❌ 削除：{len(deleted)} 件")
            for f in deleted:
                row_num = id_to_row.get(f['fileId'])
                if row_num:
                    update_row(sheet, row_num, now_str, f, '削除')
                    print(f"    - {f.get('fileName','')} → 行{row_num}を赤文字で上書き")

        # 名前変更
        if renamed:
            print(f"  ■ 名前変更：{len(renamed)} 件")
            for f in renamed:
                row_num = id_to_row.get(f['fileId'])
                if row_num:
                    update_row(sheet, row_num, now_str, f, 'ファイル名変更', before_name=f['beforeName'])
                    print(f"    - {f.get('beforeName','')} → {f.get('fileName','')}（行{row_num}を上書き）")

        # 新規追加
        if added:
            print(f"  ◉ 新規追加：{len(added)} 件")
            for f in added:
                append_new_row(sheet, now_str, f, '新規')
                print(f"    - {f.get('fileName','')} → 新規行追加")

        if not (deleted or renamed or added):
            print("  変更なし（削除・名前変更・追加は検出されませんでした）")
        else:
            print("\n[3] メール通知送信中...")
            send_email(deleted, renamed, added)

    print("\n✅ 監視処理完了")
    print("=" * 50)

# ── 実行
monitor_folder()
