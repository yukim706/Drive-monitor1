# ============================================================
# Googleドライブ ファイル監視スクリプト（GitHub Actions用）
# 【認証方式】サービスアカウント JSON（環境変数 GOOGLE_SA_JSON）
# 【メール送信】GAS WebアプリへのHTTPリクエスト（環境変数 GAS_MAIL_URL）
# 【セキュリティ】GAS_SECRET_TOKEN によるトークン検証付き
# 【検知内容】削除 / 名前変更 / 新規追加
# 【シート構成】1シート（ファイルリスト兼変更履歴）
#   A:更新日時 B:ファイル名 C:変更前ファイル名 D:URL E:種別 F:フォルダパス G:フォルダURL
# 【文字色】新規=黒 / 削除=赤 / 名前変更=青
# 【D列URL】ハイパーリンク付き・URL表示（タップで開く）
# 【G列】フォルダURL（ハイパーリンク付き）
# 【行数】不足時に自動拡張
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
import re
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

# ── 設定値（全てGitHub Secretsから取得）
FOLDER_ID        = os.environ['FOLDER_ID']
SPREADSHEET_ID   = os.environ['SPREADSHEET_ID']
GAS_MAIL_URL     = os.environ['GAS_MAIL_URL']
GAS_SECRET_TOKEN = os.environ['GAS_SECRET_TOKEN']
EMAIL_TO         = os.environ['EMAIL_TO']
FILE_SHEET       = 'ファイル一覧'

# 列番号定数（1始まり）
COL_UPDATED     = 1   # A: 更新日時
COL_FILENAME    = 2   # B: ファイル名
COL_BEFORE      = 3   # C: 変更前ファイル名
COL_FILEURL     = 4   # D: URL
COL_STATUS      = 5   # E: 種別
COL_FOLDERPATH  = 6   # F: フォルダパス
COL_FOLDERURL   = 7   # G: フォルダURL
TOTAL_COLS      = 7

HEADERS = ['更新日時', 'ファイル名', '変更前ファイル名', 'URL', '種別', 'フォルダパス', 'フォルダURL']

# ── 文字色定義
COLOR_BLACK = {'red': 0.0, 'green': 0.0, 'blue': 0.0}
COLOR_RED   = {'red': 1.0, 'green': 0.0, 'blue': 0.0}
COLOR_BLUE  = {'red': 0.0, 'green': 0.0, 'blue': 1.0}

STATUS_COLOR = {
    '正常':          COLOR_BLACK,
    '新規':          COLOR_BLACK,
    '削除':          COLOR_RED,
    'ファイル名変更': COLOR_BLUE,
}

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

# ── MIMEタイプ → ファイルURL生成
def build_file_url(file_id, mime_type):
    urls = {
        'application/vnd.google-apps.document':     f'https://docs.google.com/document/d/{file_id}/edit',
        'application/vnd.google-apps.spreadsheet':  f'https://docs.google.com/spreadsheets/d/{file_id}/edit',
        'application/vnd.google-apps.presentation': f'https://docs.google.com/presentation/d/{file_id}/edit',
        'application/vnd.google-apps.drawing':      f'https://docs.google.com/drawings/d/{file_id}/edit',
    }
    return urls.get(mime_type, f'https://drive.google.com/file/d/{file_id}/view')

# ── フォルダID → フォルダURL生成
def build_folder_url(folder_id):
    return f'https://drive.google.com/drive/folders/{folder_id}'

# ── URLをHYPERLINK数式に変換（URLをそのまま表示・タップで開く）
def make_hyperlink(url):
    if not url:
        return ''
    url_escaped = url.replace('"', '""')
    return f'=HYPERLINK("{url_escaped}","{url_escaped}")'

# ── セルからURLを取り出す（HYPERLINK数式にも対応）
def extract_url_from_cell(cell_value):
    if not cell_value:
        return ''
    if cell_value.startswith('=HYPERLINK('):
        m = re.search(r'=HYPERLINK\("([^"]+)"', cell_value)
        return m.group(1) if m else ''
    return cell_value

# ── URLからfileIdを逆引き
def extract_file_id_from_url(url):
    parts = url.rstrip('/').split('/')
    if 'd' in parts:
        idx = parts.index('d')
        if idx + 1 < len(parts):
            return parts[idx + 1]
    return None

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
                'folderUrl':  build_folder_url(folder_id),
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
        sheet = spreadsheet.add_worksheet(title=sheet_name, rows=1000, cols=TOTAL_COLS)
        sheet.append_row(HEADERS)
        return sheet

# ── 行数が足りなくなったら自動で拡張
EXPAND_ROWS = 1000

def ensure_rows(sheet, rows_needed):
    current_rows = sheet.row_count
    used_rows    = len(sheet.get_all_values())
    if used_rows + rows_needed > current_rows:
        new_total = current_rows + max(EXPAND_ROWS, rows_needed)
        sheet.add_rows(new_total - current_rows)
        print(f"  📋 シートを拡張しました：{current_rows} → {new_total} 行")

# ── シートから fileId → 行番号 の辞書を作成
def build_id_to_row_map(sheet):
    all_values = sheet.get_all_values()
    id_to_row = {}
    for i, row in enumerate(all_values):
        if i == 0:
            continue
        if len(row) >= COL_FILEURL and row[COL_FILEURL - 1]:
            url = extract_url_from_cell(row[COL_FILEURL - 1])
            fid = extract_file_id_from_url(url)
            if fid:
                id_to_row[fid] = i + 1
    return id_to_row

# ── 行に文字色を適用
def apply_row_color(sheet, row_num, status):
    color = STATUS_COLOR.get(status, COLOR_BLACK)
    range_str = f'A{row_num}:G{row_num}'
    sheet.format(range_str, {
        'textFormat': {'foregroundColor': color}
    })

# ── 行を上書き更新
def update_row(sheet, row_num, now_str, file_data, status, before_name=''):
    values = [
        now_str,
        file_data.get('fileName', ''),
        before_name,
        make_hyperlink(file_data.get('fileUrl', '')),
        status,
        file_data.get('folderPath', ''),
        make_hyperlink(file_data.get('folderUrl', '')),
    ]
    range_str = f'A{row_num}:G{row_num}'
    sheet.update(range_name=range_str, values=[values], value_input_option='USER_ENTERED')
    apply_row_color(sheet, row_num, status)

# ── 新規行を追加して文字色を適用
def append_new_row(sheet, now_str, file_data, status):
    row = [
        now_str,
        file_data.get('fileName', ''),
        '',
        make_hyperlink(file_data.get('fileUrl', '')),
        status,
        file_data.get('folderPath', ''),
        make_hyperlink(file_data.get('folderUrl', '')),
    ]
    ensure_rows(sheet, 1)
    sheet.append_row(row, value_input_option='USER_ENTERED')
    row_num = len(sheet.get_all_values())
    apply_row_color(sheet, row_num, status)

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
                'folderUrl':  cf.get('folderUrl', ''),
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
                f"  フォルダURL ：{f.get('folderUrl', '')}",
                "  種別       ：削除", ''
            ]

    if renamed:
        lines += ['=' * 40, '■ ファイル名変更', '=' * 40]
        for f in renamed:
            lines += [
                f"  ファイル名      ：{f.get('fileName', '')}",
                f"  変更前ファイル名：{f.get('beforeName', '')}",
                f"  URL            ：{f.get('fileUrl', '')}",
                f"  フォルダパス   ：{f.get('folderPath', '')}",
                f"  フォルダURL    ：{f.get('folderUrl', '')}",
                "  種別           ：ファイル名変更", ''
            ]

    if added:
        lines += ['=' * 40, '◉ 新規追加', '=' * 40]
        for f in added:
            lines += [
                f"  ファイル名  ：{f.get('fileName', '')}",
                f"  URL        ：{f.get('fileUrl', '')}",
                f"  フォルダパス：{f.get('folderPath', '')}",
                f"  フォルダURL ：{f.get('folderUrl', '')}",
                "  種別       ：新規", ''
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

    id_to_row = build_id_to_row_map(sheet)
    is_first_run = len(id_to_row) == 0
    now_str = format_datetime_jp(datetime.now(JST))

    if is_first_run:
        print("\n[2] 初回実行 → 全ファイルを登録中...")
        rows = []
        for f in current_files:
            rows.append([
                now_str,
                f.get('fileName', ''),
                '',
                make_hyperlink(f.get('fileUrl', '')),
                '正常',
                f.get('folderPath', ''),
                make_hyperlink(f.get('folderUrl', '')),
            ])
        if rows:
            ensure_rows(sheet, len(rows))
            sheet.append_rows(rows, value_input_option='USER_ENTERED')
            last_row = len(rows) + 1
            sheet.format(f'A2:G{last_row}', {
                'textFormat': {'foregroundColor': COLOR_BLACK}
            })
        print(f"  {len(rows)} 件を登録しました")

    else:
        print("\n[2] 前回リストと比較中...")
        all_values = sheet.get_all_values()
        previous_files = []
        for i, row in enumerate(all_values):
            if i == 0:
                continue
            if len(row) >= COL_FILEURL and row[COL_FILEURL - 1]:
                url = extract_url_from_cell(row[COL_FILEURL - 1])
                fid = extract_file_id_from_url(url)
                if fid:
                    folder_url = extract_url_from_cell(row[COL_FOLDERURL - 1]) if len(row) >= COL_FOLDERURL else ''
                    previous_files.append({
                        'fileId':     fid,
                        'fileName':   row[COL_FILENAME - 1]   if len(row) >= COL_FILENAME   else '',
                        'folderPath': row[COL_FOLDERPATH - 1] if len(row) >= COL_FOLDERPATH else '',
                        'fileUrl':    url,
                        'folderUrl':  folder_url,
                    })

        deleted, renamed, added = detect_changes(previous_files, current_files)

        if deleted:
            print(f"  ❌ 削除：{len(deleted)} 件")
            for f in deleted:
                row_num = id_to_row.get(f['fileId'])
                if row_num:
                    update_row(sheet, row_num, now_str, f, '削除')
                    print(f"    - {f.get('fileName','')} → 行{row_num}を赤文字で上書き")

        if renamed:
            print(f"  ■ 名前変更：{len(renamed)} 件")
            for f in renamed:
                row_num = id_to_row.get(f['fileId'])
                if row_num:
                    update_row(sheet, row_num, now_str, f, 'ファイル名変更', before_name=f['beforeName'])
                    print(f"    - {f.get('beforeName','')} → {f.get('fileName','')}（行{row_num}を青文字で上書き）")

        if added:
            print(f"  ◉ 新規追加：{len(added)} 件")
            for f in added:
                append_new_row(sheet, now_str, f, '新規')
                print(f"    - {f.get('fileName','')} → 新規行追加（黒文字）")

        if not (deleted or renamed or added):
            print("  変更なし（削除・名前変更・追加は検出されませんでした）")
        else:
            print("\n[3] メール通知送信中...")
            send_email(deleted, renamed, added)

    print("\n✅ 監視処理完了")
    print("=" * 50)

# ── 実行
monitor_folder()
