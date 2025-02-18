from flask import Flask, render_template, request, redirect, url_for, send_file, flash
import os
from werkzeug.utils import secure_filename
from utils import ocr, excel
from config import UPLOAD_FOLDER, ALLOWED_EXTENSIONS, MAX_CONTENT_LENGTH

app = Flask(__name__)
app.config["UPLOAD_FOLDER"] = UPLOAD_FOLDER
app.config["MAX_CONTENT_LENGTH"] = MAX_CONTENT_LENGTH
app.secret_key = "your_secret_key"  # セッションやflash用のキー（適宜変更してください）

# アップロード可能なファイル形式をチェック
def allowed_file(filename):
    return "." in filename and filename.rsplit(".", 1)[1].lower() in ALLOWED_EXTENSIONS


# トップページを表示
@app.route("/")
def index():
    return render_template("index.html")


# ファイルアップロード処理
@app.route("/upload", methods=["POST"])
def upload_file():
    """
    1. ファイルの受信・検証
    2. OCR処理
    3. データ抽出
    4. Excel生成
    5. 結果表示
    """
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

    try:
        # ファイルの保存
        filename = secure_filename(file.filename)
        filepath = os.path.join(app.config["UPLOAD_FOLDER"], filename)
        file.save(filepath)
        print(f"ファイルを保存しました: {filepath}")

        # 2. OCR処理
        ocr_text = ocr.process_image(filepath)
        if not ocr_text:
            raise ValueError("OCR処理に失敗しました")
        print(f"OCRテキスト: {ocr_text}")

        # 3. データ抽出
        ocr_result = ocr.parse_receipt_text(ocr_text)
        if not ocr_result:
            raise ValueError("テキストの解析に失敗しました")
        print(f"解析結果: {ocr_result}")

        # 必須項目のチェック
        required_fields = ["発行日", "発行者", "金額"]
        missing_fields = [field for field in required_fields if not ocr_result.get(field)]
        if missing_fields:
            raise ValueError(f"必須項目が見つかりません: {', '.join(missing_fields)}")

        # 4. Excel生成
        excel_filename = os.path.splitext(filename)[0] + ".xlsx"
        excel_filepath = os.path.join(app.config["UPLOAD_FOLDER"], excel_filename)
        excel.create_excel_receipt(ocr_result, excel_filepath)

        # 5. 結果表示
        return render_template("result.html", excel_file=excel_filename)

    except ValueError as e:
        flash(str(e))
        return redirect(url_for("index"))
    except Exception as e:
        flash(f"エラーが発生しました: {str(e)}")
        return redirect(url_for("index"))
    finally:
        # 一時ファイルの削除（画像ファイル）
        if os.path.exists(filepath):
            os.remove(filepath)


# Excelファイルのダウンロード
@app.route("/download/<filename>")
def download_file(filename):
    try:
        filepath = os.path.join(app.config["UPLOAD_FOLDER"], filename)
        return send_file(filepath, as_attachment=True)
    except Exception as e:
        flash(f"ダウンロードエラー: {str(e)}")
        return redirect(url_for("index"))


if __name__ == "__main__":
    # アップロードディレクトリの作成
    if not os.path.exists(UPLOAD_FOLDER):
        os.makedirs(UPLOAD_FOLDER)
    app.run(debug=True)
