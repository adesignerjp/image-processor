import os
import re
import json
import argparse
from pathlib import Path

def load_tag_categories(config_path):
    """
    タグ設定JSONファイルを読み込んで有効なタグのリストを作成する
    
    Args:
        config_path (str): タグ設定JSONファイルのパス
    
    Returns:
        list: 有効なタグのリスト
    """
    try:
        with open(config_path, 'r', encoding='utf-8') as f:
            categories = json.load(f)
        
        valid_tags = []
        for category in categories:
            if 'subcategories' in category:
                for subcategory in category['subcategories']:
                    valid_tags.append(subcategory['tag'])
        
        return valid_tags
    except Exception as e:
        print(f"タグ設定ファイルの読み込みエラー: {e}")
        return []

def analyze_filename(filename, valid_tags):
    """
    ファイル名を分析して、その構造とエラーを詳細に返す
    
    Args:
        filename (str): 分析するファイル名
        valid_tags (list): 有効なタグのリスト
    
    Returns:
        dict: 分析結果
    """
    # 拡張子を除去
    name_without_ext = os.path.splitext(filename)[0]
    parts = name_without_ext.split('_')
    
    # 正規表現パターン
    year_pattern = r'^\d{4}$'  # 年（4桁の数字）
    client_pattern = r'^[A-Za-z0-9][A-Za-z0-9\-]*$'  # クライアント名
    title_pattern = r'^[A-Za-z0-9][A-Za-z0-9\-]*$'  # タイトル
    subtitle_pattern = r'^[A-Za-z0-9][A-Za-z0-9\-]*$'  # サブタイトル（オプション）
    tag_pattern = r'^t-[A-Za-z0-9\-]+$'  # タグ（t-で始まる）
    number_pattern = r'^\d{2}$'  # 連番（2桁の数字）
    
    # 分析結果
    result = {
        'filename': filename,
        'parts': parts,
        'errors': [],
        'parsed': {
            'year': parts[0] if len(parts) > 0 else None,
            'client': parts[1] if len(parts) > 1 else None,
            'title': parts[2] if len(parts) > 2 else None,
            'subtitles': [],
            'tags': [],
            'tag_part': None,
            'number': parts[-1] if len(parts) > 3 else None
        }
    }
    
    # 基本構造チェック
    if len(parts) < 4:
        result['errors'].append(f"パート不足: 最低4つのパートが必要ですが、{len(parts)}つしかありません")
        return result
    
    # 年のチェック
    if not re.match(year_pattern, parts[0]):
        result['errors'].append(f"年の形式が不正: {parts[0]}")
    
    # クライアント名のチェック
    if not re.match(client_pattern, parts[1]):
        result['errors'].append(f"クライアント名の形式が不正: {parts[1]}")
    
    # タイトルのチェック
    if not re.match(title_pattern, parts[2]):
        result['errors'].append(f"タイトルの形式が不正: {parts[2]}")
    
    # 連番のチェック
    if not re.match(number_pattern, parts[-1]):
        result['errors'].append(f"連番の形式が不正: {parts[-1]}")
    
    # タグとサブタイトルのチェック
    has_tag_part = False
    for i in range(3, len(parts) - 1):
        if parts[i].startswith('t-'):
            has_tag_part = True
            result['parsed']['tag_part'] = parts[i]
            
            # タグの形式チェック
            if not re.match(tag_pattern, parts[i]):
                result['errors'].append(f"タグの形式が不正: {parts[i]}")
            
            # タグの存在確認
            sub_tags = parts[i][2:].split('-')  # t- を除外してタグを分割
            for tag in sub_tags:
                if tag:
                    result['parsed']['tags'].append(tag)
                    if tag not in valid_tags:
                        result['errors'].append(f"無効なタグが使用されています: {tag}")
        else:
            # サブタイトルの処理
            result['parsed']['subtitles'].append(parts[i])
            
            # サブタイトルの形式チェック
            if not re.match(subtitle_pattern, parts[i]):
                result['errors'].append(f"サブタイトルの形式が不正: {parts[i]}")
            
            # サブタイトルが有効なタグ名と一致するかチェック
            if parts[i] in valid_tags:
                result['errors'].append(f"有効なタグがサブタイトルとして使用されています。't-{parts[i]}' の形式で指定してください: {parts[i]}")
    
    # タグ部分のチェック
    if not has_tag_part and len(parts) > 4:
        result['errors'].append("タグ部分がありません。タグは 't-タグ1-タグ2' の形式で指定してください")
    
    return result

def find_non_matching_files(folder_path, valid_tags, image_extensions=None, verbose=False):
    """
    指定されたフォルダ内の画像ファイルから、命名規則に違反するファイル名を抽出する
    
    Args:
        folder_path (str): 検索対象のフォルダパス
        valid_tags (list): 有効なタグのリスト
        image_extensions (list, optional): 画像ファイルの拡張子リスト
        verbose (bool): 詳細情報を出力するかどうか
    
    Returns:
        list: パターンに違反するファイルパスと分析結果のタプルのリスト
    """
    if image_extensions is None:
        image_extensions = ['.jpg', '.jpeg', '.png', '.gif', '.bmp', '.tiff', '.webp']
    
    non_matching_files = []
    
    # フォルダ内のファイルを走査
    for root, _, files in os.walk(folder_path):
        for file in files:
            # 拡張子が画像ファイル形式かチェック
            file_ext = os.path.splitext(file)[1].lower()
            if file_ext not in image_extensions:
                continue
            
            # ファイル名を分析
            analysis = analyze_filename(file, valid_tags)
            
            # エラーがあればリストに追加
            if analysis['errors']:
                file_path = os.path.join(root, file)
                non_matching_files.append((file_path, analysis))
    
    return non_matching_files

def main():
    parser = argparse.ArgumentParser(description='命名規則に違反する画像ファイル名を検出するツール')
    parser.add_argument('--folder', default="/Users/z/adesignerjp/images", help='検索対象のフォルダパス')
    parser.add_argument('--config', default="/Users/z/adesignerjp/config/categories.json", help='タグ設定JSONファイルのパス')
    parser.add_argument('--extensions', default='.jpg,.jpeg,.png,.gif,.bmp,.tiff,.webp', help='画像ファイルの拡張子（カンマ区切り）')
    parser.add_argument('--output', help='結果を出力するファイルパス（指定しない場合は標準出力）')
    parser.add_argument('--verbose', action='store_true', help='詳細情報を出力する')
    parser.add_argument('--analyze', help='特定のファイル名を分析する')
    
    args = parser.parse_args()
    
    # タグ設定を読み込む
    valid_tags = load_tag_categories(args.config)
    if not valid_tags:
        print("警告: 有効なタグが読み込めませんでした。タグの検証はスキップされます。")
    else:
        print(f"読み込まれた有効なタグ: {', '.join(valid_tags)}")
    
    # 特定のファイル名を分析する場合
    if args.analyze:
        analysis = analyze_filename(args.analyze, valid_tags)
        print(f"ファイル名分析: {args.analyze}")
        print(f"分解されたパーツ: {analysis['parts']}")
        
        if analysis['errors']:
            print(f"検出されたエラー:")
            for error in analysis['errors']:
                print(f"  - {error}")
        else:
            print("エラーなし - 有効なファイル名です")
        
        print("\n構造分析:")
        print(f"  年: {analysis['parsed']['year']}")
        print(f"  クライアント: {analysis['parsed']['client']}")
        print(f"  タイトル: {analysis['parsed']['title']}")
        print(f"  サブタイトル: {', '.join(analysis['parsed']['subtitles'])}")
        print(f"  タグ部分: {analysis['parsed']['tag_part']}")
        print(f"  タグ: {', '.join(analysis['parsed']['tags'])}")
        print(f"  連番: {analysis['parsed']['number']}")
        
        return
    
    extensions = [ext.strip() if ext.strip().startswith('.') else f'.{ext.strip()}' for ext in args.extensions.split(',')]
    
    # 違反ファイルを抽出
    non_matching_files = find_non_matching_files(args.folder, valid_tags, extensions, args.verbose)
    
    # 結果を出力
    output_text = f"命名規則に違反するファイル: {len(non_matching_files)}件\n\n"
    
    for file_path, analysis in non_matching_files:
        output_text += f"{file_path}\n"
        if args.verbose:
            output_text += f"  パート: {analysis['parts']}\n"
            output_text += f"  エラー:\n"
            for error in analysis['errors']:
                output_text += f"    - {error}\n"
            output_text += "\n"
        else:
            output_text += f"  エラー: {', '.join(analysis['errors'])}\n"
    
    if args.output:
        with open(args.output, 'w', encoding='utf-8') as f:
            f.write(output_text)
        print(f"結果を {args.output} に保存しました")
    else:
        print(output_text)

if __name__ == "__main__":
    main()