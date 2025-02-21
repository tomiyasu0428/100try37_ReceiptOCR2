from flask import Flask, request, render_template, flash, redirect, url_for, send_from_directory
import os
from werkzeug.utils import secure_filename
from utils import ocr, excel
from config import UPLOAD_FOLDER, EXCEL_FOLDER, ALLOWED_EXTENSIONS, MAX_CONTENT_LENGTH
from datetime import datetime

app = Flask(__name__)
app.config["UPLOAD_FOLDER"] = UPLOAD_FOLDER
app.config["MAX_CONTENT_LENGTH"] = MAX_CONTENT_LENGTH
app.secret_key = "your_secret_key"  # セッションやflash用のキー（適宜変更してください）

# アップロード可能なファイル形式をチェック
def allowed_file(filename):
    return "." in filename and filename.rsplit(".", 1)[1].lower() in ALLOWED_EXTENSIONS

# ファイル名を生成（タイムスタンプ付き）
def generate_filename(original_filename):
    timestamp = datetime.now().strftime("%Y%m%d%H%M%S")
    ext = os.path.splitext(original_filename)[1]
    return f"{timestamp}_{secure_filename(original_filename)}"

# トップページを表示
@app.route("/")
def index():
    # アップロードディレクトリ内のExcelファイルを取得
    excel_files = []
    if os.path.exists(EXCEL_FOLDER):
        excel_files = [f for f in os.listdir(EXCEL_FOLDER) if f.endswith('.xlsx')]
    return render_template("index.html", excel_files=excel_files)

# ファイルアップロード処理
@app.route("/upload", methods=["POST"])
def upload_file():
    """
    1. ファイルの受信・検証
    2. OCR処理
    3. データ抽出
    4. Excel生成
    5. 結果表示（ダウンロードボタン付き）
    """
    try:
        # アップロードフォルダが存在しない場合は作成
        os.makedirs(UPLOAD_FOLDER, exist_ok=True)
        os.makedirs(EXCEL_FOLDER, exist_ok=True)

        # 1. ファイルの受信・検証
        if "receipt" not in request.files:
            flash("ファイルが選択されていません")
            return redirect(url_for("index"))

        file = request.files["receipt"]
        if file.filename == "":
            flash("ファイルが選択されていません")
            return redirect(url_for("index"))

        if not allowed_file(file.filename):
            flash("許可されていないファイル形式です")
            return redirect(url_for("index"))

        # 既存のExcelファイルの選択を確認
        excel_file = request.form.get('excel_file', '')
        if not excel_file:
            # 新規Excelファイルの場合、既存のファイルを全て削除
            for f in os.listdir(EXCEL_FOLDER):
                if f.endswith('.xlsx'):
                    os.remove(os.path.join(EXCEL_FOLDER, f))
            # 新しいファイル名を生成
            excel_file = f"領収書データ_{datetime.now().strftime('%Y%m%d_%H%M%S')}.xlsx"

        excel_path = os.path.join(EXCEL_FOLDER, excel_file)

        # ファイルの保存
        filename = generate_filename(file.filename)
        filepath = os.path.join(app.config["UPLOAD_FOLDER"], filename)
        print(f"ファイルを保存しました: {filepath}")
        file.save(filepath)

        # OCR処理とデータ抽出
        result = ocr.main(filepath)
        if result is None:
            flash("OCR処理に失敗しました")
            return redirect(url_for("index"))

        # Excel生成
        if excel.create_excel_receipt(result, excel_path):
            flash(f"OCR処理が完了しました。Excelファイルをダウンロードできます。")
        else:
            flash("Excel作成中にエラーが発生しました")
        
        return redirect(url_for("index"))

    except Exception as e:
        print(f"エラーが発生しました: {str(e)}")
        flash("処理中にエラーが発生しました")
        return redirect(url_for("index"))

# Excelファイルのダウンロード
@app.route("/download/<filename>")
def download_file(filename):
    try:
        return send_from_directory(EXCEL_FOLDER, filename, as_attachment=True)
    except Exception as e:
        flash(f"ダウンロードエラー: {str(e)}")
        return redirect(url_for("index"))

if __name__ == "__main__":
    # アップロードディレクトリの作成
    if not os.path.exists(UPLOAD_FOLDER):
        os.makedirs(UPLOAD_FOLDER)
    app.run(debug=True)
