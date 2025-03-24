#!/usr/bin/env python3
"""
スプレッドシートからギャラリーデータのJSONを生成するスクリプト
"""

import os
import json
from google.oauth2 import service_account
import gspread

# 設定
SPREADSHEET_ID = 'あなたのスプレッドシートID'
SHEET_NAME = 'Gallery'
CREDENTIALS_FILE = 'your-credentials.json'
OUTPUT_FILE = 'data/gallery_data.json'

def authenticate_google_apis():
    """Google APIの認証を行います"""
    credentials = service_account.Credentials.from_service_account_file(
        CREDENTIALS_FILE, 
        scopes=[
            'https://www.googleapis.com/auth/spreadsheets',
            'https://www.googleapis.com/auth/drive'
        ]
    )
    
    # スプレッドシートAPIクライアント
    gc = gspread.authorize(credentials)
    
    return gc

def process_tags(tags_str):
    """タグ文字列を処理して正規化します"""
    if not tags_str:
        return []
        
    # カンマで分割して各タグをトリム
    tags = [tag.strip() for tag in tags_str.split(',')]
    
    # 空のタグを除去
    return [tag for tag in tags if tag]

def load_category_mapping():
    """カテゴリ設定を読み込んでタグからメインカテゴリへのマッピングを作成"""
    try:
        with open('config/categories.json', 'r', encoding='utf-8') as f:
            categories = json.load(f)
        
        # タグからメインカテゴリへのマッピング
        tag_to_category = {}
        
        for category in categories:
            # 自身のIDをメインカテゴリとして登録
            tag_to_category[category['id']] = category['id']
            
            # サブカテゴリがある場合
            if 'subcategories' in category and category['subcategories']:
                for subcategory in category['subcategories']:
                    # サブカテゴリのタグを親カテゴリにマッピング
                    tag_to_category[subcategory['tag']] = category['id']
        
        return tag_to_category
    except Exception as e:
        print(f"カテゴリ設定の読み込みに失敗しました: {e}")
        return {}

def main():
    """メイン処理"""
    print("Google APIの認証中...")
    gc = authenticate_google_apis()
    
    print("カテゴリマッピングを読み込み中...")
    tag_to_category = load_category_mapping()
    
    print("スプレッドシートからデータを取得中...")
    try:
        spreadsheet = gc.open_by_key(SPREADSHEET_ID)
        worksheet = spreadsheet.worksheet(SHEET_NAME)
        
        # ヘッダー付きでデータを取得
        data = worksheet.get_all_records()
        
        # JSONに変換
        gallery_data = []
        
        for row in data:
            # 必須フィールドがない場合はスキップ
            if not row.get('Title') or not row.get('PreviewURL'):
                continue
            
            # タグを処理
            tags = process_tags(row.get('Tags', ''))
            
            # メインカテゴリを判定
            main_category = None
            for tag in tags:
                if tag in tag_to_category:
                    main_category = tag_to_category[tag]
                    break
            
            # アイテムを作成
            item = {
                'title': row.get('Title', ''),
                'subtitle': row.get('Subtitle', ''),
                'year': row.get('Year', ''),
                'client': row.get('Client', ''),
                'detail': row.get('Detail', ''),
                'tags': tags,
                'previewURL': row.get('PreviewURL', ''),
                'thumbnail': row.get('Thumbnail', '') or row.get('PreviewURL', ''),
                'fileId': row.get('File ID', ''),
                'mainCategory': main_category
            }
            
            # オプションフィールド
            if 'Order' in row and row['Order']:
                try:
                    item['order'] = int(row['Order'])
                except ValueError:
                    pass
            
            gallery_data.append(item)
        
        # 出力ディレクトリを確認
        os.makedirs(os.path.dirname(OUTPUT_FILE), exist_ok=True)
        
        # JSONファイルに書き出し
        with open(OUTPUT_FILE, 'w', encoding='utf-8') as f:
            json.dump(gallery_data, f, ensure_ascii=False, indent=2)
        
        print(f"{len(gallery_data)}件のデータを{OUTPUT_FILE}に書き出しました。")
        
    except Exception as e:
        print(f"処理中にエラーが発生しました: {e}")

if __name__ == "__main__":
    main()
