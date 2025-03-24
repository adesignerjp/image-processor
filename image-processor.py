#!/usr/bin/env python3
"""
画像ファイル名からメタデータを抽出し、スプレッドシートに追加するスクリプト（改良版）
- 各ファイルを1行ずつ表示
- 連番順に並べて表示
- Preview URLをThumbnailに自動的に反映
- Edited フラグのあるエントリはスキップ
- 処理の高速化
- Status列の自動管理: 未編集は空欄、編集済みはEdited
- 重複ファイルの効率的なスキップ
"""

import os
import re
import glob
import json
import hashlib
from datetime import datetime
from google.oauth2 import service_account
import gspread
from google.cloud import storage
from gspread.exceptions import GSpreadException, APIError
from gspread import Cell
import time
import logging

# スクリプトのディレクトリを取得（追加）
SCRIPT_DIR = os.path.dirname(os.path.abspath(__file__))

# ログ設定を修正：絶対パスを使用し、コンソール出力も追加
logging.basicConfig(
    level=logging.INFO,
    filename=os.path.join(SCRIPT_DIR, 'image_processing.log'),  # 絶対パスに修正済み
    format='%(asctime)s - %(levelname)s - %(message)s',
    force=True  # 既存のロガー設定を上書き
)
logger = logging.getLogger(__name__)

# 標準出力にもログを出力
console = logging.StreamHandler()
console.setLevel(logging.INFO)
formatter = logging.Formatter('%(asctime)s - %(levelname)s - %(message)s')
console.setFormatter(formatter)
logger.addHandler(console)

# 設定を外部ファイルから読み込み
CONFIG_FILE = os.path.join(SCRIPT_DIR, 'config.json')
try:
    with open(CONFIG_FILE, 'r') as f:
        config = json.load(f)
    SPREADSHEET_ID = config.get('spreadsheet_id', '1dTnuPyLHjYANic63d24fWiugxQi_qA0QKFlE0Jd-nXE')
    SHEET_NAME = config.get('sheet_name', 'sheet1')
    CREDENTIALS_FILE = config.get('credentials_file', '/Users/z/adesignerjp/config/credentials.json')
    BUCKET_NAME = config.get('bucket_name', 'adesignerjp-images')
    LOCAL_IMAGE_DIR = config.get('local_image_dir', 'images')
    # 相対パスを絶対パスに変換
    cache_file_relative = config.get('cache_file', 'image_processing_cache.json')
    CACHE_FILE = os.path.join(SCRIPT_DIR, cache_file_relative)
except FileNotFoundError:
    logger.warning(f"{CONFIG_FILE} が見つかりません。デフォルト設定を使用します")
    SPREADSHEET_ID = '1dTnuPyLHjYANic63d24fWiugxQi_qA0QKFlE0Jd-nXE'
    SHEET_NAME = 'sheet1'
    CREDENTIALS_FILE = '/Users/z/adesignerjp/config/credentials.json'
    BUCKET_NAME = 'adesignerjp-images'
    LOCAL_IMAGE_DIR = 'images'
    CACHE_FILE = os.path.join(SCRIPT_DIR, 'image_processing_cache.json')

# エラー回復用の失敗ファイルリスト
FAILED_FILES_CACHE = 'failed_files.json'

def authenticate_google_apis():
    """Google APIの認証を行います"""
    credentials = service_account.Credentials.from_service_account_file(
        CREDENTIALS_FILE, 
        scopes=[
            'https://www.googleapis.com/auth/spreadsheets',
            'https://www.googleapis.com/auth/drive',
            'https://www.googleapis.com/auth/devstorage.read_write'
        ]
    )
    
    # スプレッドシートAPIクライアント
    gc = gspread.authorize(credentials)
    
    # Google Cloud Storageクライアント
    storage_client = storage.Client.from_service_account_json(CREDENTIALS_FILE)
    
    return gc, storage_client

def parse_filename(filename):
    """ファイル名からメタデータを抽出します（正規表現パターンに基づく）"""
    # 拡張子を除去
    base_name = os.path.splitext(os.path.basename(filename))[0]
    
    # メタデータの初期化
    metadata = {
        'year': None,
        'client': None,
        'title': None,
        'subtitle': None,
        'tags': [],
        'sequence': None,
        'base_name': None  # 連番を除いたベース名（グループ化用）
    }
    
    # アンダースコアで分割
    parts = base_name.split('_')
    
    # 正規表現パターン
    year_pattern = r'^\d{4}$'  # 年（4桁の数字）
    client_pattern = r'^[A-Za-z0-9][A-Za-z0-9\-]*$'  # クライアント名
    title_pattern = r'^[A-Za-z0-9][A-Za-z0-9\-]*$'  # タイトル
    subtitle_pattern = r'^[A-Za-z0-9][A-Za-z0-9\-]*$'  # サブタイトル（オプション）
    tag_pattern = r'^t-[A-Za-z0-9\-]+$'  # タグ（t-で始まる）
    number_pattern = r'^\d{2}$'  # 連番（2桁の数字）
    
    # 連番を検出 (最後の部分が2桁の数字のみかチェック)
    if parts and re.match(number_pattern, parts[-1]):
        metadata['sequence'] = parts[-1]
        parts = parts[:-1]  # 連番部分を削除
    
    # 連番を除いたベース名を保存（グループ化用）
    metadata['base_name'] = '_'.join(parts)
    
    # 年のチェック（最初のパート）
    if parts and re.match(year_pattern, parts[0]):
        metadata['year'] = parts[0]
        parts = parts[1:]
    
    # クライアント名のチェック（2番目のパート）
    if parts and re.match(client_pattern, parts[0]):
        metadata['client'] = parts[0]
        parts = parts[1:]
    
    # タイトルのチェック（3番目のパート）
    if parts and re.match(title_pattern, parts[0]):
        metadata['title'] = parts[0]
        parts = parts[1:]
    
    # 残りのパートをタグとサブタイトルに分類
    for part in parts:
        if part.startswith('t-') and re.match(tag_pattern, part):
            # タグとして処理
            # ハイフンをスペースに変換
            tag_value = part.replace('-', ' ')
            metadata['tags'].append(tag_value)
        elif re.match(subtitle_pattern, part):
            # サブタイトルが未設定なら設定、既に設定済みならタグとして扱う
            if metadata['subtitle'] is None:
                # ハイフンをスペースに変換
                metadata['subtitle'] = part.replace('-', ' ')
            else:
                # タグとして処理し、ハイフンをスペースに変換
                tag_value = part.replace('-', ' ')
                metadata['tags'].append(tag_value)
        else:
            # パターンに一致しない場合でもタグとして追加し、ハイフンをスペースに変換
            tag_value = part.replace('-', ' ')
            metadata['tags'].append(tag_value)
        
            # クライアント名、タイトル、サブタイトルもハイフンをスペースに変換
            if metadata['client']:
               metadata['client'] = metadata['client'].replace('-', ' ')
    
            if metadata['title']:
               metadata['title'] = metadata['title'].replace('-', ' ')
    
            if metadata['subtitle']:
               metadata['subtitle'] = metadata['subtitle'].replace('-', ' ')
        
    return metadata

def load_cache():
    """キャッシュファイルを読み込みます"""
    if os.path.exists(CACHE_FILE):
        try:
            with open(CACHE_FILE, 'r', encoding='utf-8') as f:
                return json.load(f)
        except Exception as e:
            logger.error(f"キャッシュ読み込みエラー: {e}")
    return {
        'processed_files': {},
        'last_processed': None,
        'gcs_urls': {},
        'file_hashes': {}  # ファイル名とハッシュ値のマッピングを追加
    }

def save_cache(cache):
    """キャッシュファイルを保存します"""
    try:
        with open(CACHE_FILE, 'w', encoding='utf-8') as f:
            json.dump(cache, f, ensure_ascii=False, indent=2)
    except Exception as e:
        logger.error(f"キャッシュ保存エラー: {e}")

def load_failed_files():
    """前回の失敗したファイルを読み込みます"""
    if os.path.exists(FAILED_FILES_CACHE):
        try:
            with open(FAILED_FILES_CACHE, 'r', encoding='utf-8') as f:
                return json.load(f)
        except Exception as e:
            logger.error(f"失敗ファイルキャッシュ読み込みエラー: {e}")
    return []

def save_failed_files(failed_files):
    """失敗したファイルを保存します"""
    try:
        with open(FAILED_FILES_CACHE, 'w', encoding='utf-8') as f:
            json.dump(failed_files, f, ensure_ascii=False, indent=2)
    except Exception as e:
        logger.error(f"失敗ファイルキャッシュ保存エラー: {e}")

def get_file_hash(file_path):
    """ファイルのハッシュ値を計算します"""
    try:
        hasher = hashlib.md5()
        with open(file_path, 'rb') as f:
            buf = f.read(65536)
            while len(buf) > 0:
                hasher.update(buf)
                buf = f.read(65536)
        return hasher.hexdigest()
    except Exception as e:
        logger.error(f"ハッシュ計算エラー: {file_path} - {e}")
        return None

def upload_to_gcs(local_path, bucket_name, cache):
    file_hash = get_file_hash(local_path)
    if file_hash and file_hash in cache['gcs_urls']:
        logger.info(f"キャッシュヒット: {os.path.basename(local_path)}")
        # キャッシュされたURLが公開されているか確認
        cached_url = cache['gcs_urls'][file_hash]
        storage_client = storage.Client.from_service_account_json(CREDENTIALS_FILE)
        bucket = storage_client.bucket(bucket_name)
        blob_name = os.path.basename(local_path)
        blob = bucket.blob(blob_name)
        if blob.exists():
            blob.make_public()  # 既存画像でも公開設定を再適用
            updated_url = blob.public_url.strip()
            if cached_url != updated_url:
                cache['gcs_urls'][file_hash] = updated_url
                save_cache(cache)
        return cache['gcs_urls'][file_hash]
    
    if file_hash:
        cache['processed_files'][file_hash] = 'pending'
        save_cache(cache)
    
    retries = 3
    for attempt in range(retries):
        try:
            storage_client = storage.Client.from_service_account_json(CREDENTIALS_FILE)
            bucket = storage_client.bucket(bucket_name)
            blob_name = os.path.basename(local_path)
            blob = bucket.blob(blob_name)
            
            if not blob.exists():
                blob.upload_from_filename(local_path)
            blob.make_public()  # 既存でも新規でも公開に設定
            
            url = blob.public_url.strip()
            logger.info(f"生成されたURL: {url}")
            if file_hash:
                cache['gcs_urls'][file_hash] = url
                cache['processed_files'][file_hash] = time.time()
                save_cache(cache)
            
            return url
        except Exception as e:
            logger.error(f"アップロードエラー（試行 {attempt+1}/{retries}）: {blob_name} - {e}")
            time.sleep(2 ** attempt)
            if attempt == retries - 1 and file_hash and file_hash in cache['processed_files']:
                del cache['processed_files'][file_hash]
                save_cache(cache)
                raise Exception(f"アップロード失敗: {os.path.basename(local_path)}")
    
def generate_detail_text(year, client, title, subtitle, sequence=None):
    """詳細ページ用のテキストを自動生成します"""
    if not year and not client and not title:
        return ""
    
    # 各値を使う前に安全のためにNoneでないことを確認
    title = title or ""
    subtitle = subtitle or ""
    client = client or ""
    year = year or ""
    sequence = sequence or ""
    
    detail = f"{title}"
    if subtitle:
        detail += f" - {subtitle}"
    if sequence:
        detail += f" ({sequence})"
    if client:
        detail += f" for {client}"
    if year:
        detail += f" ({year})"
    
    return detail

def get_existing_data(worksheet):
    """スプレッドシートから既存のデータを取得します"""
    try:
        # まずヘッダー行を取得し検証
        headers = worksheet.row_values(1)
        required_headers = ['Year', 'Client', 'Title', 'Subtitle', 'Detail', 'Tags', 'Preview URL', 'Thumbnail', 'File ID', 'Status']
        
        # ヘッダーを検証し、不足している場合は修正
        if len(headers) < len(required_headers) or any(header not in headers for header in required_headers):
            logger.info("ヘッダーを修正します")
            worksheet.update(values=[required_headers], range_name='A1:J1')  # 非推奨警告を修正
            
        # 全データを取得
        all_values = worksheet.get_all_values()
        if len(all_values) <= 1:  # ヘッダーのみの場合
            return {}, []
        
        data_rows = all_values[1:]  # ヘッダー以外の行
        
        # インデックスを取得
        header_dict = {header: idx for idx, header in enumerate(headers)}
        
        # ファイルIDベースのマッピングを作成
        file_map = {}
        
        # 既存のNewを空欄に変更するためのセル
        new_to_empty_cells = []
        
        for row_idx, row in enumerate(data_rows, start=2):  # 2行目からスタート（ヘッダー行を考慮）
            if len(row) <= header_dict.get('File ID', -1):
                continue  # 行が短すぎる場合はスキップ
                
            file_id = row[header_dict.get('File ID')]
            status_idx = header_dict.get('Status')
            
            # 既存のNewを空欄に変更
            if status_idx is not None and status_idx < len(row) and row[status_idx] == 'New':
                new_to_empty_cells.append(Cell(row_idx, status_idx + 1, ''))
            
            if file_id:
                row_data = {headers[i]: row[i] if i < len(row) else "" for i in range(len(headers))}
                row_data['row_num'] = row_idx
                file_map[file_id] = row_data
        
        return file_map, new_to_empty_cells
        
    except Exception as e:
        logger.error(f"データ取得エラー: {e}")
        return {}, []

def group_files_by_sequence(image_files):
    """画像ファイルをベース名でグループ化し、各グループ内で連番順にソートします"""
    groups = {}
    
    for file_path in image_files:
        metadata = parse_filename(file_path)
        base_name = metadata['base_name']
        
        if base_name not in groups:
            groups[base_name] = []
        
        groups[base_name].append({
            'path': file_path,
            'name': os.path.basename(file_path),
            'metadata': metadata
        })
    
    # 各グループ内で連番順にソート
    for base_name, files in groups.items():
        files.sort(key=lambda x: int(x['metadata']['sequence'] or 0))
    
    # 結果をフラット化し、グループごとに連続するように整理
    result = []
    for base_name, files in groups.items():
        result.extend(files)
    
    return result

def safe_update_cells(worksheet, cells, retries=3, value_input_option='USER_ENTERED'):
    """セルを安全に更新（リトライ付き）"""
    for attempt in range(retries):
        try:
            worksheet.update_cells(cells, value_input_option=value_input_option)
            return True
        except APIError as e:
            logger.error(f"APIエラー（試行 {attempt+1}/{retries}）: {e}")
            time.sleep(3 ** attempt)
        except Exception as e:
            logger.error(f"その他のエラー（試行 {attempt+1}/{retries}）: {e}")
            time.sleep(3 ** attempt)
    logger.error("セル更新に失敗しました")
    return False

def safe_append_rows(worksheet, rows, retries=3):
    """行を安全に追加（リトライ付き）"""
    for attempt in range(retries):
        try:
            worksheet.append_rows(rows)
            return True
        except APIError as e:
            logger.error(f"APIエラー（試行 {attempt+1}/{retries}）: {e}")
            time.sleep(2 ** attempt)
    logger.error("行追加に失敗しました")
    return False

def update_thumbnail_formula(worksheet, rows_count):
    try:
        logger.info("サムネイル列に数式を設定します...")
        headers = worksheet.row_values(1)
        preview_url_idx = headers.index('Preview URL') + 1
        thumbnail_idx = headers.index('Thumbnail') + 1
        
        preview_url_col = chr(64 + preview_url_idx)
        thumbnail_col = chr(64 + thumbnail_idx)
        
        cells_to_update = []
        for row in range(2, rows_count + 2):
            formula = f'=IMAGE({preview_url_col}{row})'
            cells_to_update.append(Cell(row, thumbnail_idx, formula))
            logger.debug(f"設定する数式: {formula} (行 {row})")
        
        batch_size = 10
        for i in range(0, len(cells_to_update), batch_size):
            batch = cells_to_update[i:i+batch_size]
            if safe_update_cells(worksheet, batch, value_input_option='USER_ENTERED'):
                logger.info(f"サムネイル数式を設定: {i+1}～{min(i+batch_size, len(cells_to_update))}行目")
            else:
                logger.error(f"サムネイル数式設定失敗: {i+1}～{min(i+batch_size, len(cells_to_update))}")
            time.sleep(3)
        
        logger.info("サムネイル列の数式を設定しました")
    except Exception as e:
        logger.error(f"サムネイル列の数式設定エラー: {e}")
        import traceback
        logger.error(f"詳細: {traceback.format_exc()}")

def process_files(worksheet, organized_files, file_map, cache, failed_files):
    """画像ファイルを処理し、スプレッドシートに追加または更新するデータを準備します"""
    rows_to_add = []
    cells_to_update = []
    processed_files = set()
    
    for file_info in organized_files:
        file_path = file_info['path']
        file_name = file_info['name']
        metadata = file_info['metadata']
        
        file_hash = get_file_hash(file_path)
        if not file_hash:
            logger.warning(f"スキップ: {file_name} - ハッシュ計算に失敗")
            failed_files.append(file_name)
            continue
            
        if file_name in file_map:
            row_data = file_map[file_name]
            status = row_data.get('Status', '')
            if status == 'Edited':
                logger.info(f"スキップ: {file_name} - 編集済みのため上書きしません")
                processed_files.add(file_hash)
                continue
            
            try:
                gcs_url = upload_to_gcs(file_path, BUCKET_NAME, cache)
                gcs_url = gcs_url.strip().replace('\n', '').replace('\r', '')
                processed_files.add(file_hash)
            except Exception as e:
                logger.error(f"アップロードエラー: {file_name} - {e}")
                failed_files.append(file_name)
                continue
            
            detail = generate_detail_text(
                metadata.get('year', ''),
                metadata.get('client', ''),
                metadata.get('title', ''),
                metadata.get('subtitle', ''),
                metadata.get('sequence', '')
            )
            
            row_num = row_data['row_num']
            cell_updates = [
                Cell(row_num, 1, metadata.get('year', '')),
                Cell(row_num, 2, metadata.get('client', '')),
                Cell(row_num, 3, metadata.get('title', '')),
                Cell(row_num, 4, metadata.get('subtitle', '')),
                Cell(row_num, 5, detail),
                Cell(row_num, 6, ', '.join(metadata.get('tags', []))),
                Cell(row_num, 7, gcs_url),
                Cell(row_num, 9, file_name),
                Cell(row_num, 10, '')
            ]
            cells_to_update.extend(cell_updates)
            logger.info(f"更新予定: {file_name}")
            
        else:
            try:
                gcs_url = upload_to_gcs(file_path, BUCKET_NAME, cache)
                gcs_url = gcs_url.strip().replace('\n', '').replace('\r', '')
                processed_files.add(file_hash)
            except Exception as e:
                logger.error(f"アップロードエラー: {file_name} - {e}")
                failed_files.append(file_name)
                continue
            
            detail = generate_detail_text(
                metadata.get('year', ''),
                metadata.get('client', ''),
                metadata.get('title', ''),
                metadata.get('subtitle', ''),
                metadata.get('sequence', '')
            )
            
            row_data = [
                metadata.get('year', ''),
                metadata.get('client', ''),
                metadata.get('title', ''),
                metadata.get('subtitle', ''),
                detail,
                ', '.join(metadata.get('tags', [])),
                gcs_url,
                '',
                file_name,
                ''
            ]
            rows_to_add.append(row_data)
            logger.info(f"追加予定: {file_name}")
    
    # 50ファイルごとにキャッシュを保存
    if len(processed_files) % 50 == 0 and processed_files:
        save_cache(cache)
        logger.info(f"キャッシュを中間保存しました（{len(processed_files)}件処理済み）")
    
    return rows_to_add, cells_to_update, processed_files

def update_spreadsheet(worksheet, cells_to_update, rows_to_add):
    """スプレッドシートを更新します（バッチサイズと待機時間を調整）"""
    if cells_to_update:
        batch_size = 500  # バッチサイズを小さくして安定性を向上
        for i in range(0, len(cells_to_update), batch_size):
            batch = cells_to_update[i:i+batch_size]
            if safe_update_cells(worksheet, batch):
                logger.info(f"セルバッチ更新: {len(batch)} セル")
            else:
                logger.error(f"セル更新失敗: {i+1}～{i+len(batch)}")
            time.sleep(2)  # 待機時間を長く
    
    if rows_to_add:
        batch_size = 25  # バッチサイズを小さくして安定性を向上
        for i in range(0, len(rows_to_add), batch_size):
            batch = rows_to_add[i:i+batch_size]
            if safe_append_rows(worksheet, batch):
                logger.info(f"行バッチ追加: {len(batch)} 行")
            else:
                logger.error(f"行追加失敗: {i+1}～{i+len(batch)}")
                for row_data in batch:
                    try:
                        worksheet.append_row(row_data)
                        logger.info(f"個別追加成功: {row_data[8]}")
                        time.sleep(0.5)
                    except Exception as inner_e:
                        logger.error(f"個別行追加エラー: {row_data[8]} - {inner_e}")
            time.sleep(2)  # 待機時間を長く

def collect_statistics(image_files, unique_files, cells_to_update, rows_to_add, processed_files, duplicate_count, execution_time):
    """処理統計を収集して表示します"""
    logger.info(f"処理が完了しました。実行時間: {execution_time:.2f}秒")
    logger.info(f"重複スキップ: {duplicate_count}件のファイルをスキップしました")
    
    logger.info("===== 処理統計 =====")
    logger.info(f"スキャンしたファイル数: {len(image_files)}件")
    logger.info(f"一意のファイル数: {len(unique_files)}件")
    logger.info(f"更新したエントリ数: {len(cells_to_update) // 9}件")
    logger.info(f"新規追加したエントリ数: {len(rows_to_add)}件")
    logger.info(f"処理済みキャッシュに保存したファイル数: {len(processed_files)}件")
    logger.info("===================")

def main():
    """メイン処理（エラー処理を強化）"""
    try:
        start_time = time.time()
        logger.info("処理を開始します...")
        
        cache = load_cache()
        logger.info(f"前回の処理: {cache['last_processed'] or '初回実行'}")
        
        failed_files = load_failed_files()
        
        logger.info("Google APIの認証中...")
        gc, storage_client = authenticate_google_apis()
        
        spreadsheet = gc.open_by_key(SPREADSHEET_ID)
        
        try:
            worksheet = spreadsheet.worksheet(SHEET_NAME)
        except:
            logger.info(f"シート '{SHEET_NAME}' が見つかりません。新しく作成します。")
            worksheet = spreadsheet.add_worksheet(title=SHEET_NAME, rows=1000, cols=10)
            headers = ['Year', 'Client', 'Title', 'Subtitle', 'Detail', 'Tags', 'Preview URL', 'Thumbnail', 'File ID', 'Status']
            worksheet.update(values=[headers], range_name='A1:J1')  # 非推奨警告を修正
        
        logger.info("スプレッドシートから既存データを取得中...")
        file_map, new_to_empty_cells = get_existing_data(worksheet)
        
        if new_to_empty_cells:
            try:
                batch_size = 1000
                for i in range(0, len(new_to_empty_cells), batch_size):
                    batch = new_to_empty_cells[i:i+batch_size]
                    if safe_update_cells(worksheet, batch):
                        logger.info(f"Newステータスを空欄に変更: {len(batch)} セル")
                    else:
                        logger.error(f"Newステータス更新失敗: {i+1}～{i+len(batch)}")
                    time.sleep(1)
            except Exception as e:
                logger.error(f"ステータス更新エラー: {e}")
        
        logger.info("ローカル画像を探索中...")
        image_files = glob.glob(f"{LOCAL_IMAGE_DIR}/**/*.jpg", recursive=True) + \
                      glob.glob(f"{LOCAL_IMAGE_DIR}/**/*.jpeg", recursive=True) + \
                      glob.glob(f"{LOCAL_IMAGE_DIR}/**/*.png", recursive=True) + \
                      glob.glob(f"{LOCAL_IMAGE_DIR}/**/*.gif", recursive=True)
        
        logger.info(f"{len(image_files)}枚の画像が見つかりました。")
        
        logger.info("画像ファイルのハッシュ値を計算中...")
        file_hashes = {}
        unique_files = []
        duplicate_count = 0
        
        all_image_files = failed_files + image_files
        
        for file_path in all_image_files:
            file_hash = get_file_hash(file_path)
            if not file_hash:
                continue
                
            file_name = os.path.basename(file_path)
            
            if file_hash in cache['processed_files'] and cache['processed_files'][file_hash] != 'pending':
                logger.info(f"スキップ: {file_name} - 前回の実行で処理済み")
                continue
            
            if file_hash in file_hashes:
                logger.info(f"スキップ: {file_name} - 重複ファイル（元ファイル: {file_hashes[file_hash]}）")
                duplicate_count += 1
                continue
                
            cache['file_hashes'][file_name] = file_hash
            file_hashes[file_hash] = file_name
            unique_files.append(file_path)
        
        logger.info(f"{len(unique_files)}枚のユニーク画像を処理します。{duplicate_count}枚の重複をスキップしました。")
        
        organized_files = group_files_by_sequence(unique_files)
        logger.info(f"{len(organized_files)}件のファイルを処理します。")
        
        rows_to_add, cells_to_update, processed_files = process_files(worksheet, organized_files, file_map, cache, failed_files)
        
        cache_limit = 1000
        if len(cache['processed_files']) > cache_limit:
            oldest_files = sorted(
                [(k, v) for k, v in cache['processed_files'].items() if v != 'pending'],
                key=lambda x: x[1]
            )
            cache['processed_files'] = dict(oldest_files[-cache_limit:] + [(k, v) for k, v in cache['processed_files'].items() if v == 'pending'])
        
        cache['processed_files'] = {k: v for k, v in cache['processed_files'].items() if k in processed_files and v != 'pending'}
        cache['last_processed'] = datetime.now().isoformat()
        save_cache(cache)
        save_failed_files(failed_files)
        
        update_spreadsheet(worksheet, cells_to_update, rows_to_add)
        
        try:
            all_values = worksheet.get_all_values()
            rows_count = len(all_values) - 1
            logger.info(f"実際の行数: {rows_count}")
        except Exception as e:
            logger.error(f"行数取得エラー: {e}")
            rows_count = len(file_map) + len(rows_to_add)
            logger.info(f"推定行数: {rows_count}")
        
        time.sleep(5)
        update_thumbnail_formula(worksheet, rows_count)
        
        try:
            worksheet = spreadsheet.worksheet(SHEET_NAME)
            test_values = worksheet.row_values(1)
            logger.info(f"スプレッドシートの再検証完了: {len(test_values)}列")
        except Exception as e:
            logger.error(f"スプレッドシート再検証エラー: {e}")
        
        end_time = time.time()
        execution_time = end_time - start_time
        collect_statistics(image_files, unique_files, cells_to_update, rows_to_add, processed_files, duplicate_count, execution_time)
    
    except Exception as e:
        logger.error(f"予期せぬエラーが発生しました: {e}")
        import traceback
        logger.error(f"詳細: {traceback.format_exc()}")
    finally:
        logger.info("処理を終了します")

if __name__ == "__main__":
    main()

# オプション: 並列処理の例（必要に応じて有効化）
# from concurrent.futures import ThreadPoolExecutor
# def process_files_parallel(worksheet, organized_files, file_map, cache, failed_files):
#     with ThreadPoolExecutor(max_workers=4) as executor:
#         futures = [executor.submit(process_file_single, f, worksheet, file_map, cache, failed_files) for f in organized_files]
#         results = [f.result() for f in futures]
#     rows_to_add = [r[0] for r in results if r[0] is not None]
#     cells_to_update = [r[1] for r in results if r[1] is not None]
#     processed_files = set(r[2] for r in results if r[2] is not None)
#     return rows_to_add, cells_to_update, processed_files