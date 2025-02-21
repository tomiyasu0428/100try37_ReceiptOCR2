# 領収書OCRシステム v2.0

## 概要
PDFまたは画像形式の領収書から情報を自動抽出し、Excelファイルに整理するWebアプリケーション。
Tesseract OCRとGoogle Gemini APIを組み合わせることで、高精度な情報抽出を実現しています。

## 主な機能
- PDFおよび画像ファイルからの情報抽出
- Gemini APIを活用した高精度なデータ認識
- 自動ID採番機能付きExcel出力
- 既存Excelファイルへのデータ追加
- モダンなWebインターフェース

## 技術スタック
- Python 3.9+
- Flask
- Tesseract OCR
- Google Gemini API
- OpenCV
- openpyxl

## セットアップ
1. 必要なパッケージのインストール
```bash
pip install -r requirements.txt
```

2. Tesseract OCRのインストール
```bash
# macOS
brew install tesseract

# Ubuntu
sudo apt-get install tesseract-ocr
```

3. 環境変数の設定
```bash
# .envファイルを作成
GEMINI_API_KEY=your_api_key_here
```

4. アプリケーションの起動
```bash
python app.py
```

## 使用方法
1. ブラウザで`http://localhost:5000`にアクセス
2. PDFまたは画像ファイルをアップロード
3. 新規作成または既存のExcelファイルを選択
4. 処理完了後、Excelファイルをダウンロード

## ディレクトリ構成
```
.
├── app.py              # メインアプリケーション
├── config.py           # 設定ファイル
├── requirements.txt    # 依存パッケージ
├── static/            # 静的ファイル
│   └── css/
├── templates/         # HTMLテンプレート
├── uploads/          # アップロードファイル一時保存
├── excel_files/      # 生成されたExcelファイル
└── utils/            # ユーティリティ
    ├── ocr.py       # OCR処理
    └── excel.py     # Excel処理
```

## 注意事項
- アップロードできるファイルサイズは最大16MBまで
- 対応ファイル形式: PDF, PNG, JPG, JPEG
- OCRの精度は画像の品質に依存します
- Gemini APIの使用には別途APIキーが必要です

## ライセンス
MIT License

## 作者
Hiroki Tomiyasu
