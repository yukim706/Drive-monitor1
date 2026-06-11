# ============================================================
# Googleドライブ ファイル監視スクリプト（GitHub Actions用）
# 【認証方式】サービスアカウント JSON（環境変数 GOOGLE_SA_JSON）
# 【メール送信】GAS WebアプリへのHTTPリクエスト（環境変数 GAS_MAIL_URL）
# 【セキュリティ】GAS_SECRET_TOKEN によるトークン検証付き
# 【検知内容】削除 / 名前変更 / 新規追加
# 【シート構成】1シート（ファイルリスト兼変更履歴）
#   A:更新日時 B:ファイル名 C:変更前ファイル名 D:URL E:ファイルID F:種別 G:フォルダパス
# ============================================================

# subprocessというどうぐをよびだす（パソコンにめいれいするためのどうぐ）
import subprocess

# gspread（スプレッドシートをさわるどうぐ）などをインストールする
subprocess.run(['pip', 'install', '--quiet', 'gspread', 'google-auth',
                'google-api-python-client'], check=True)

# gspreadをよびだす（Googleスプレッドシートをよんだりかいたりするどうぐ）
import gspread

# Googleのパスワードかくにんどうぐをよびだす
from google.oauth2.service_account import Credentials

# GoogleドライブをさわるためのどうぐをよびだすAPIをつかうためのもの）
from googleapiclient.discovery import build

# 日付や時刻をあつかうためのどうぐをよびだす
from datetime import datetime, timezone, timedelta

# JSONというデータのかたちをあつかうどうぐをよびだす
import json

# パソコンの環境変数（かくれたせってい）をよみこむどうぐをよびだす
import os

# インターネットにリクエストをおくるどうぐをよびだす
import urllib.request

# ── タイムゾーン（JST）
# 日本時間（UTC+9）をつくる。9時間のずれをあらわしている
JST = timezone(timedelta(hours=9))

# ── 認証（サービスアカウント）
# Googleへのログインにひつような「できること一覧（スコープ）」をきめる
SCOPES = [
    'https://www.googleapis.com/auth/drive.readonly',    # Driveをよむだけのけんり
    'https://www.googleapis.com/auth/spreadsheets',      # スプレッドシートをよんだりかいたりするけんり
]

# 環境変数からサービスアカウントのJSONをよみこんで、Pythonのじしょけいしきにする
sa_info = json.loads(os.environ['GOOGLE_SA_JSON'])

# サービスアカウント情報からGoogleにログインするためのパスポートをつくる
creds = Credentials.from_service_account_info(sa_info, scopes=SCOPES)

# スプレッドシートをさわるためのクライアント（まどぐち）をつくる
gc            = gspread.Client(auth=creds)

# DriveのAPIをつかうためのまどぐちをつくる
drive_service = build('drive', 'v3', credentials=creds)

# ── 設定値
# かんしするフォルダのID（環境変数になければデフォルト値をつかう）
FOLDER_ID        = os.environ.get('FOLDER_ID',      '1tVQU7ufn_Ob54kspgh88iK03-wUPF7mU')

# きろくするスプレッドシートのID（環境変数になければデフォルト値をつかう）
SPREADSHEET_ID   = os.environ.get('SPREADSHEET_ID', '146fJr4d1TL1PWx_jGwNpznNzqp5Q_BwBH2_jutdjuhs')

# メールをおくるGASのURL（かならず環境変数から）
GAS_MAIL_URL     = os.environ['GAS_MAIL_URL']

# GASにおくるひみつのトークン（ただしいリクエストかどうかかくにんするため）
GAS_SECRET_TOKEN = os.environ['GAS_SECRET_TOKEN']

# メールのおくりさき（環境変数になければデフォルトアドレスをつかう）
EMAIL_TO         = os.environ.get('EMAIL_TO', 'yukimgidai2020@gmail.com')

# スプレッドシートのシート名
FILE_SHEET       = 'ファイル一覧'

# 列番号のていぎ（1からはじまる。A列=1、B列=2、…）
COL_UPDATED    = 1   # A列：更新日時
COL_FILENAME   = 2   # B列：ファイル名
COL_BEFORE     = 3   # C列：変更前ファイル名
COL_FILEURL    = 4   # D列：URL
COL_FILEID     = 5   # E列：ファイルID
COL_STATUS     = 6   # F列：種別（正常・削除・名前変更・新規）
COL_FOLDERPATH = 7   # G列：フォルダパス

# ぜんぶで7列つかう
TOTAL_COLS     = 7

# スプレッドシートの1行目（ヘッダー）のみだしをならべたリスト
HEADERS = ['更新日時', 'ファイル名', '変更前ファイル名', 'URL', 'ファイルID', '種別', 'フォルダパス']

# ── 曜日（日本語）
# 月曜=0、火曜=1、…、日曜=6 の順で日本語のよみがなをならべたリスト
WEEKDAYS_JP = ['月', '火', '水', '木', '金', '土', '日']

# 日付をにほんごのひょうじにするかんすう
def format_datetime_jp(dt):
    # 曜日のばんごうをつかって日本語の曜日をとりだす
    wd = WEEKDAYS_JP[dt.weekday()]
    # 「2025年6月11日(水) 09:00」のようなかたちにしてかえす
    return f"{dt.year}年{dt.month}月{dt.day}日({wd}) {dt.strftime('%H:%M')}"

# ── MIMEタイプ → 拡張子
# ファイルのしゅるい（MIMEタイプ）から、ファイルのかくちょうし（.pdf など）にへんかんするじしょ
MIME_EXTENSIONS = {
    'application/vnd.google-apps.document':     '.gdoc',    # Googleドキュメント
    'application/vnd.google-apps.spreadsheet':  '.gsheet',  # Googleスプレッドシート
    'application/vnd.google-apps.presentation': '.gslides', # Googleスライド
    'application/vnd.google-apps.drawing':      '.gdraw',   # Google図形描画
    'application/pdf':                          '',          # PDF（かくちょうしなし）
    'application/vnd.openxmlformats-officedocument.wordprocessingml.document':   '.docx', # Wordファイル
    'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet':         '.xlsx', # Excelファイル
    'application/vnd.openxmlformats-officedocument.presentationml.presentation': '.pptx', # PowerPointファイル
    'image/jpeg': '.jpg',  # JPEG画像
    'image/png':  '.png',  # PNG画像
}

# ── MIMEタイプ → URL生成
# ファイルのしゅるいとIDからひらくためのURLをつくるかんすう
def build_file_url(file_id, mime_type):
    # GoogleアプリのしゅるいごとにひらくURLのかたちがちがうのでじしょでかえす
    urls = {
        'application/vnd.google-apps.document':     f'https://docs.google.com/document/d/{file_id}/edit',
        'application/vnd.google-apps.spreadsheet':  f'https://docs.google.com/spreadsheets/d/{file_id}/edit',
        'application/vnd.google-apps.presentation': f'https://docs.google.com/presentation/d/{file_id}/edit',
        'application/vnd.google-apps.drawing':      f'https://docs.google.com/drawings/d/{file_id}/edit',
    }
    # じしょになければDriveのふつうのURLをかえす
    return urls.get(mime_type, f'https://drive.google.com/file/d/{file_id}/view')

# ── Drive API でフォルダ内ファイルを再帰取得
# していしたフォルダのなかにあるファイルをサブフォルダのおくまですべてリストアップするかんすう
def get_all_files(folder_id, folder_path, result=None):
    # はじめてよばれたときはからのリストをじゅんびする
    if result is None:
        result = []

    # ページングのためのトークン（たくさんあるときにわけてとるため）
    page_token = None
    while True:
        # このフォルダのなかにあるファイル（フォルダ以外）をとりだす
        response = drive_service.files().list(
            q=f"'{folder_id}' in parents and mimeType != 'application/vnd.google-apps.folder' and trashed = false",
            # ↑「このフォルダの子どもで、フォルダじゃなくて、ごみばこにはいっていないもの」
            fields='nextPageToken, files(id, name, mimeType)',  # とりだすじょうほうをしていする
            pageSize=1000,      # いちどにとる最大件数
            pageToken=page_token  # つぎのページがあるときはこのトークンをつかう
        ).execute()

        # とりだしたファイルをひとつずつりすとにくわえる
        for f in response.get('files', []):
            mime = f['mimeType']                        # ファイルのしゅるい
            ext  = MIME_EXTENSIONS.get(mime, '')        # かくちょうしをとりだす
            result.append({
                'fileId':     f['id'],                  # ファイルのID
                'fileName':   f['name'] + ext,          # ファイル名＋かくちょうし
                'mimeType':   mime,                     # ファイルのしゅるい
                'fileUrl':    build_file_url(f['id'], mime),  # ひらくためのURL
                'folderPath': folder_path,              # フォルダのばしょ
            })

        # つぎのページのトークンをとりだす。なければwhileループをぬける
        page_token = response.get('nextPageToken')
        if not page_token:
            break

    # このフォルダのなかにあるサブフォルダをとりだす
    sub_page_token = None
    while True:
        sub_response = drive_service.files().list(
            q=f"'{folder_id}' in parents and mimeType = 'application/vnd.google-apps.folder' and trashed = false",
            # ↑「このフォルダの子どもで、フォルダで、ごみばこにはいっていないもの」
            fields='nextPageToken, files(id, name)',
            pageSize=1000,
            pageToken=sub_page_token
        ).execute()

        # サブフォルダをひとつずつさいきてきによびだして、なかのファイルもとりだす
        for sub in sub_response.get('files', []):
            # フォルダのなかのフォルダも同じかんすうでよびだす（さいきこうぞう）
            get_all_files(sub['id'], folder_path + '/' + sub['name'], result)

        # つぎのページのトークンをとりだす。なければwhileループをぬける
        sub_page_token = sub_response.get('nextPageToken')
        if not sub_page_token:
            break

    # あつめたファイル一覧をかえす
    return result

# ── シート取得 or 作成
# していしたなまえのシートをとりだす。なければあたらしくつくるかんすう
def get_or_create_sheet(spreadsheet, sheet_name):
    try:
        # シートがあればそのままかえす
        return spreadsheet.worksheet(sheet_name)
    except gspread.exceptions.WorksheetNotFound:
        # シートがなければあたらしくつくる
        sheet = spreadsheet.add_worksheet(title=sheet_name, rows=5000, cols=TOTAL_COLS)
        # 1行目にヘッダーをかく
        sheet.append_row(HEADERS)
        return sheet

# ── シートから fileId → 行番号 の辞書を作成（E列＝COL_FILEID=5）
# シートのE列（ファイルID）をよんで、「ファイルID → 行番号」のじしょをつくるかんすう
def build_id_to_row_map(sheet):
    # シートのぜんデータをよみこむ
    all_values = sheet.get_all_values()
    # じしょをじゅんびする（からのじしょ）
    id_to_row = {}
    for i, row in enumerate(all_values):
        # 1行目（ヘッダー）はスキップする
        if i == 0:
            continue
        # E列（ファイルID）があれば、「ファイルID：行番号」のかたちでじしょにくわえる
        if len(row) >= COL_FILEID and row[COL_FILEID - 1]:
            id_to_row[row[COL_FILEID - 1]] = i + 1  # 行番号は1からはじまるので+1する
    # できたじしょをかえす
    return id_to_row

# ── 行を上書き更新
# していした行番号のデータをあたらしい内容でうわがきするかんすう
def update_row(sheet, row_num, now_str, file_data, status, before_name=''):
    # かきこむデータをならべたリストをつくる（A〜G列のじゅん）
    values = [
        now_str,                            # A列：更新日時
        file_data.get('fileName', ''),      # B列：ファイル名
        before_name,                        # C列：変更前ファイル名（名前変更のときだけつかう）
        file_data.get('fileUrl', ''),       # D列：URL
        file_data.get('fileId', ''),        # E列：ファイルID
        status,                             # F列：種別（削除・名前変更・新規など）
        file_data.get('folderPath', ''),    # G列：フォルダパス
    ]
    # どのはんいをうわがきするかをしていする（例：A5:G5）
    range_str = f'A{row_num}:G{row_num}'
    # シートにデータをかきこむ
    sheet.update(range_name=range_str, values=[values])

    # 種別によって文字の色をかえる
    if status == '削除':
        # 削除 → 赤文字
        color = {'red': 1.0, 'green': 0.0, 'blue': 0.0}
    elif status == 'ファイル名変更':
        # 名前変更 → 青文字
        color = {'red': 0.0, 'green': 0.0, 'blue': 1.0}
    else:
        # それ以外（正常・新規など）→ 黒文字
        color = {'red': 0.0, 'green': 0.0, 'blue': 0.0}

    # きめた色でシートの行をいろづけする
    sheet.format(range_str, {
        'textFormat': {
            'foregroundColor': color
        }
    })

# ── 新規行を追加
# あたらしいファイルのじょうほうをシートのさいごについかするかんすう
def append_new_row(sheet, now_str, file_data, status):
    # かきこむデータをならべたリストをつくる
    row = [
        now_str,                            # A列：更新日時
        file_data.get('fileName', ''),      # B列：ファイル名
        '',                                 # C列：変更前ファイル名（新規なのでからっぽ）
        file_data.get('fileUrl', ''),       # D列：URL
        file_data.get('fileId', ''),        # E列：ファイルID
        status,                             # F列：種別
        file_data.get('folderPath', ''),    # G列：フォルダパス
    ]
    # シートのさいごにあたらしい行をくわえる
    sheet.append_row(row, value_input_option='USER_ENTERED')

    # くわえた行のぎょうばんごうをもとめる（さいごの行）
    last_row = len(sheet.get_all_values())
    # 新規追加 → 黒文字（まえに色がついていたばあいにそなえてかならず黒にする）
    sheet.format(f'A{last_row}:G{last_row}', {
        'textFormat': {
            'foregroundColor': {'red': 0.0, 'green': 0.0, 'blue': 0.0}  # 黒色
        }
    })

# ── 変更の判定
# 「まえのファイル一覧」と「いまのファイル一覧」をくらべて、変化をみつけるかんすう
def detect_changes(previous_files, current_files):
    # まえのファイルをファイルIDをキーにしたじしょにする（さがしやすくするため）
    prev_by_id = {f['fileId']: f for f in previous_files}
    # いまのファイルをファイルIDをキーにしたじしょにする
    curr_by_id = {f['fileId']: f for f in current_files}

    # 削除・名前変更・新規追加をいれるリストをじゅんびする
    deleted = []
    renamed = []
    added   = []

    # まえにあったファイルをひとつずつしらべる
    for fid, pf in prev_by_id.items():
        if fid not in curr_by_id:
            # まえはあったのにいまはない → 削除された
            deleted.append(pf)
        elif curr_by_id[fid]['fileName'] != pf['fileName']:
            # IDはおなじなのにファイル名がちがう → 名前が変わった
            cf = curr_by_id[fid]
            renamed.append({
                'fileId':     fid,
                'fileName':   cf['fileName'],       # あたらしいファイル名
                'beforeName': pf['fileName'],       # まえのファイル名
                'fileUrl':    cf.get('fileUrl', ''),
                'folderPath': cf.get('folderPath', ''),
                'mimeType':   cf.get('mimeType', ''),
            })

    # いまあるファイルをひとつずつしらべる
    for fid, cf in curr_by_id.items():
        if fid not in prev_by_id:
            # いまあるのにまえはなかった → あたらしく追加された
            added.append(cf)

    # 削除・名前変更・新規追加のリストをかえす
    return deleted, renamed, added

# ── メール本文
# おくるメールのほんぶんをつくるかんすう
def build_email_body(deleted, renamed, added):
    # いまの日時を日本語でとりだす
    now_str = format_datetime_jp(datetime.now(JST))
    # メールのほんぶんをならべるリストをじゅんびする
    lines = [f"検知日時：{now_str}", '']

    # 削除されたファイルがあればついかする
    if deleted:
        lines += ['=' * 40, '❌ 削除', '=' * 40]
        for f in deleted:
            lines += [
                f"  ファイル名  ：{f.get('fileName', '')}",
                f"  URL        ：{f.get('fileUrl', '')}",
                f"  フォルダパス：{f.get('folderPath', '')}",
                f"  種別       ：削除", ''
            ]

    # 名前が変わったファイルがあればついかする
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

    # あたらしくついかされたファイルがあればついかする
    if added:
        lines += ['=' * 40, '◉ 新規追加', '=' * 40]
        for f in added:
            lines += [
                f"  ファイル名  ：{f.get('fileName', '')}",
                f"  URL        ：{f.get('fileUrl', '')}",
                f"  フォルダパス：{f.get('folderPath', '')}",
                f"  種別       ：新規", ''
            ]

    # リストのかくぎょうをかいぎょうでつないでひとつのもじれつにしてかえす
    return '\n'.join(lines)

# ── メール送信
# GAS Webアプリにリクエストをおくってメールをはっしんするかんすう
def send_email(deleted, renamed, added):
    # メールのけんめいにいれるぶんをつくる（例：「削除1件 / 新規追加2件」）
    parts = []
    if deleted: parts.append(f'削除{len(deleted)}件')
    if renamed: parts.append(f'名前変更{len(renamed)}件')
    if added:   parts.append(f'新規追加{len(added)}件')

    # けんめいをつくる
    subject = f'【ドライブ監視】{" / ".join(parts)} が検出されました'
    # ほんぶんをつくる
    body    = build_email_body(deleted, renamed, added)

    # GASにおくるデータをJSON形式にへんかんする
    payload = json.dumps({
        'token':   GAS_SECRET_TOKEN,    # ひみつのトークン（ほんものかどうかかくにんするため）
        'to':      EMAIL_TO,            # おくりさきのメールアドレス
        'subject': subject,             # けんめい
        'body':    body,                # ほんぶん
    }).encode('utf-8')                  # UTF-8のバイトれつにへんかんする

    # GASのURLにPOSTリクエストをつくる
    req = urllib.request.Request(
        GAS_MAIL_URL,                                   # おくりさきのURL
        data=payload,                                   # おくるデータ
        headers={'Content-Type': 'application/json'},   # データのけいしきをJSON形式とつたえる
        method='POST'                                   # POSTでおくる
    )
    try:
        # リクエストをおくって30秒まつ
        with urllib.request.urlopen(req, timeout=30) as res:
            # GASからのへんじをよみとる
            result = res.read().decode('utf-8')
            # せいこうしたことをひょうじする
            print(f"  メール送信完了 → {EMAIL_TO}（GAS応答: {result}）")
    except Exception as e:
        # しっぱいしたときはエラーないようをひょうじする
        print(f"  ⚠️ メール送信失敗: {e}")

# ── メイン処理
# プログラムのメインのながれをせいりするかんすう
def monitor_folder():
    # かいしのメッセージをひょうじする
    print("=" * 50)
    print(f"監視開始：{format_datetime_jp(datetime.now(JST))}")
    print("=" * 50)

    # スプレッドシートをIDでひらく
    spreadsheet = gc.open_by_key(SPREADSHEET_ID)
    # ファイル一覧シートをとりだす（なければつくる）
    sheet = get_or_create_sheet(spreadsheet, FILE_SHEET)

    print("\n[1] ファイル一覧取得中...")

    # かんしするフォルダのなまえをDrive APIでとりだす
    root_name = drive_service.files().get(
        fileId=FOLDER_ID, fields='name'
    ).execute().get('name', 'Root')

    # フォルダのなかにあるファイルをぜんぶとりだす（サブフォルダのおくまで）
    current_files = get_all_files(FOLDER_ID, root_name)
    print(f"  現在のファイル数：{len(current_files)} 件")

    # シートのE列（ファイルID）から「ファイルID → 行番号」のじしょをつくる
    id_to_row = build_id_to_row_map(sheet)

    # じしょがからっぽ（まだいちどもつかったことがない）かどうかしらべる
    is_first_run = len(id_to_row) == 0

    # いまの日時を日本語でとりだす
    now_str = format_datetime_jp(datetime.now(JST))

    if is_first_run:
        # ────────────────────────────────────────
        # 初回実行：全ファイルを一括登録
        # ────────────────────────────────────────
        print("\n[2] 初回実行 → 全ファイルを登録中...")
        # かきこむデータをいれるリストをじゅんびする
        rows = []
        for f in current_files:
            # ファイルごとに1行ぶんのデータをつくってリストにくわえる
            rows.append([
                now_str,                    # A列：更新日時
                f.get('fileName', ''),      # B列：ファイル名
                '',                         # C列：変更前ファイル名（はじめてなのでからっぽ）
                f.get('fileUrl', ''),       # D列：URL
                f.get('fileId', ''),        # E列：ファイルID
                '正常',                     # F列：種別（はじめてなので「正常」）
                f.get('folderPath', ''),    # G列：フォルダパス
            ])
        # データがあればシートにまとめてかきこむ
        if rows:
            sheet.append_rows(rows, value_input_option='USER_ENTERED')
            # くわえた行のさいごのぎょうばんごうをもとめる
            last_row = len(sheet.get_all_values())
            # くわえたぎょうのはんいをもとめる（2行目からさいごの行まで）
            first_row = last_row - len(rows) + 1
            # 初回登録はぜんぶ黒文字にする
            sheet.format(f'A{first_row}:G{last_row}', {
                'textFormat': {
                    'foregroundColor': {'red': 0.0, 'green': 0.0, 'blue': 0.0}  # 黒色
                }
            })
        print(f"  {len(rows)} 件を登録しました")

    else:
        # ────────────────────────────────────────
        # 2回目以降：変更検知 → 該当行を上書き
        # ────────────────────────────────────────
        print("\n[2] 前回リストと比較中...")

        # シートのぜんデータをよみこむ
        all_values = sheet.get_all_values()

        # まえのファイル一覧をシートからつくりなおす
        previous_files = []
        for i, row in enumerate(all_values):
            # 1行目（ヘッダー）はスキップする
            if i == 0:
                continue
            # ★ F列（種別）が「削除」の行はスキップする
            # （いちどさくじょされたファイルはくらべないようにする。これがないとまいかい「削除」と通知されてしまう）
            if len(row) >= COL_STATUS and row[COL_STATUS - 1] == '削除':
                continue
            # E列（ファイルID）があればまえのファイル一覧にくわえる
            if len(row) >= COL_FILEID and row[COL_FILEID - 1]:
                previous_files.append({
                    'fileId':     row[COL_FILEID - 1],                                          # E列：ファイルID
                    'fileName':   row[COL_FILENAME - 1]   if len(row) >= COL_FILENAME   else '', # B列：ファイル名
                    'folderPath': row[COL_FOLDERPATH - 1] if len(row) >= COL_FOLDERPATH else '', # G列：フォルダパス
                })

        # まえのファイル一覧といまのファイル一覧をくらべて変化をしらべる
        deleted, renamed, added = detect_changes(previous_files, current_files)

        # 削除されたファイルをシートに記録する
        if deleted:
            print(f"  ❌ 削除：{len(deleted)} 件")
            for f in deleted:
                # このファイルがシートのなんぎょうめにあるかさがす
                row_num = id_to_row.get(f['fileId'])
                if row_num:
                    # そのぎょうを「削除」でうわがきする（赤文字になる）
                    update_row(sheet, row_num, now_str, f, '削除')
                    print(f"    - {f.get('fileName','')} → 行{row_num}を赤文字で上書き")

        # 名前が変わったファイルをシートに記録する
        if renamed:
            print(f"  ■ 名前変更：{len(renamed)} 件")
            for f in renamed:
                # このファイルがシートのなんぎょうめにあるかさがす
                row_num = id_to_row.get(f['fileId'])
                if row_num:
                    # そのぎょうを「ファイル名変更」でうわがきする（まえのなまえもきろくする）
                    update_row(sheet, row_num, now_str, f, 'ファイル名変更', before_name=f['beforeName'])
                    print(f"    - {f.get('beforeName','')} → {f.get('fileName','')}（行{row_num}を上書き）")

        # あたらしくついかされたファイルをシートに記録する
        if added:
            print(f"  ◉ 新規追加：{len(added)} 件")
            for f in added:
                # シートのさいごにあたらしい行をついかする
                append_new_row(sheet, now_str, f, '新規')
                print(f"    - {f.get('fileName','')} → 新規行追加")

        # なにも変化がなかったときはそのむねをひょうじする
        if not (deleted or renamed or added):
            print("  変更なし（削除・名前変更・追加は検出されませんでした）")
        else:
            # 変化があったときはメールでおしらせする
            print("\n[3] メール通知送信中...")
            send_email(deleted, renamed, added)

    # かんりょうのメッセージをひょうじする
    print("\n✅ 監視処理完了")
    print("=" * 50)

# ── 実行
# このファイルをちょくせつじっこうしたときだけ monitor_folder() をよびだす
monitor_folder()
