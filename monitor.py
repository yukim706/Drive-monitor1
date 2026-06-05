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
HISTORY_SHEET    = '変更履歴'
FILELIST_SHEET   = 'ファイルリスト'

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
    google_apps_url = {
        'application/vnd.google-apps.document':     f'https://docs.google.com/document/d/{file_id}/edit',
        'application/vnd.google-apps.spreadsheet':  f'https://docs.google.com/spreadsheets/d/{file_id}/edit',
        'application/vnd.google-apps.presentation': f'https://docs.google.com/presentation/d/{file_id}/edit',
        'application/vnd.google-apps.drawing':      f'https://docs.google.com/drawings/d/{file_id}/edit',
    }
    return google_apps_url.get(mime_type, f'https://drive.google.com/file/d/{file_id}/view')

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
                'folderPath': folder_path
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
                'folderPath': pf['folderPath'],
            })

    for fid, cf in curr_by_id.items():
        if fid not in prev_by_id:
            added.append(cf)

    return deleted, renamed, added

# ── スプレッドシート操作ヘルパー
def get_or_create_sheet(spreadsheet, sheet_name, headers=None):
    try:
        ws = spreadsheet.worksheet(sheet_name)
        # ヘッダーが1行もない場合は書き込む
        if headers and ws.row_count == 0:
            ws.append_row(headers)
        return ws
    except gspread.exceptions.WorksheetNotFound:
        sheet = spreadsheet.add_worksheet(title=sheet_name, rows=1000, cols=10)
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

def record_history(spreadsheet, deleted, renamed, added):
    """
    A: 年月日・時間　B: ファイル名　C: 変更前ファイル名
    D: URL　E: フォルダ名　F: 種別
    """
    sheet = get_or_create_sheet(
        spreadsheet, HISTORY_SHEET,
        headers=['年月日・時間', 'ファイル名', '変更前ファイル名', 'URL', 'フォルダ名', '種別']
    )

    # ヘッダー行がなければ追加
    existing = sheet.get_all_values()
    if not existing:
        sheet.append_row(['年月日・時間', 'ファイル名', '変更前ファイル名', 'URL', 'フォルダ名', '種別'])

    now_str = format_datetime_jp(datetime.now(JST))
    rows = []

    for f in deleted:
        rows.append([
            now_str,
            f.get('fileName', ''),
            '',
            f.get('fileUrl', ''),
            f.get('folderPath', ''),
            '削除',
        ])

    for f in renamed:
        rows.append([
            now_str,
            f.get('fileName', ''),
            f.get('beforeName', ''),
            f.get('fileUrl', ''),
            f.get('folderPath', ''),
            'ファイル名変更',
        ])

    for f in added:
        rows.append([
            now_str,
            f.get('fileName', ''),
            '',
            f.get('fileUrl', ''),
            f.get('folderPath', ''),
            '新規',
        ])

    if rows:
        sheet.append_rows(rows, value_input_option='USER_ENTERED')
        print(f"  履歴シートに {len(rows)} 件記録しました")

# ── メール本文
def build_email_body(deleted, renamed, added):
    now_str = format_datetime_jp(datetime.now(JST))
    lines = [f"検知日時：{now_str}", '']

    if deleted:
        lines.append('=' * 40)
        lines.append('❌ 削除')
        lines.append('=' * 40)
        for f in deleted:
            lines.append(f"  ファイル名  ：{f.get('fileName', '')}")
            lines.append(f"  URL        ：{f.get('fileUrl', '（なし）')}")
            lines.append(f"  フォルダ名 ：{f.get('folderPath', '')}")
            lines.append(f"  種別       ：削除")
            lines.append('')

    if renamed:
        lines.append('=' * 40)
        lines.append('■ ファイル名変更')
        lines.append('=' * 40)
        for f in renamed:
            lines.append(f"  ファイル名      ：{f.get('fileName', '')}")
            lines.append(f"  変更前ファイル名：{f.get('beforeName', '')}")
            lines.append(f"  URL            ：{f.get('fileUrl', '（なし）')}")
            lines.append(f"  フォルダ名     ：{f.get('folderPath', '')}")
            lines.append(f"  種別           ：ファイル名変更")
            lines.append('')

    if added:
        lines.append('=' * 40)
        lines.append('◉ 新規追加')
        lines.append('=' * 40)
        for f in added:
            lines.append(f"  ファイル名  ：{f.get('fileName', '')}")
            lines.append(f"  URL        ：{f.get('fileUrl', '（なし）')}")
            lines.append(f"  フォルダ名 ：{f.get('folderPath', '')}")
            lines.append(f"  種別       ：新規")
            lines.append('')

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

    print("\n[1] ファイル一覧取得中...")
    root_name = drive_service.files().get(
        fileId=FOLDER_ID, fields='name'
    ).execute().get('name', 'Root')
    current_files = get_all_files(FOLDER_ID, root_name)
    print(f"  現在のファイル数：{len(current_files)} 件")

    print("\n[2] 前回リストと比較中...")
    previous_files = load_previous_files(spreadsheet)

    if previous_files is None:
        print("  前回リストなし → 今回のリストを初回保存します（変更検知はしません）")
    else:
        print(f"  前回のファイル数：{len(previous_files)} 件")
        deleted, renamed, added = detect_changes(previous_files, current_files)

        if deleted:
            print(f"  ❌ 削除：{len(deleted)} 件")
            for f in deleted:
                print(f"    - {f.get('fileName','')}（{f.get('folderPath','')}）")

        if renamed:
            print(f"  ■ 名前変更：{len(renamed)} 件")
            for f in renamed:
                print(f"    - {f.get('beforeName','')} → {f.get('fileName','')}（{f.get('folderPath','')}）")

        if added:
            print(f"  ◉ 新規追加：{len(added)} 件")
            for f in added:
                print(f"    - {f.get('fileName','')}（{f.get('folderPath','')}）")

        if deleted or renamed or added:
            print("\n[3] 履歴シートに記録中...")
            record_history(spreadsheet, deleted, renamed, added)
            print("\n[4] メール通知送信中...")
            send_email(deleted, renamed, added)
        else:
            print("  変更なし（削除・名前変更・追加は検出されませんでした）")

    print("\n[5] 現在のファイルリストを保存中...")
    save_current_files(spreadsheet, current_files)
    print("\n✅ 監視処理完了")
    print("=" * 50)

# ── 実行
monitor_folder()
