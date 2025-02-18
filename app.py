from flask import Flask, render_template, request, redirect, url_for, send_file, flash
import os
from werkzeug.utils import secure_filename
from utils import ocr, excel
from config import UPLOAD_FOLDER, ALLOWED_EXTENSIONS, MAX_CONTENT_LENGTH

app = Flask(__name__)
app.config["UPLOAD_FOLDER"] = UPLOAD_FOLDER
app.config["MAX_CONTENT_LENGTH"] = MAX_CONTENT_LENGTH
app.secret_key = "your_secret_key"  # セッションやflash用のキー（適宜変更してください）

# アップロード可能なファイル形式
def allowed_file(filename):
    return "." in filename and filename.rsplit(".", 1)[1].lower() in ALLOWED_EXTENSIONS


@app.route("/")
def index():
    return render_template("index.html")


@app.route("/upload", methods=["GET", "POST"])
def upload_file():
    if request.method == "GET":
        return render_template("index.html")
        
    if "receipt" not in request.files:
        flash("ファイルが選択されていません。")
        return redirect(request.url)
    file = request.files["receipt"]
    if file.filename == "":
        flash("ファイルが選択されていません。")
        return redirect(request.url)
    if file and allowed_file(file.filename):
        filename = secure_filename(file.filename)
        filepath = os.path.join(app.config["UPLOAD_FOLDER"], filename)
        file.save(filepath)
        print(f"ファイルを保存しました: {filepath}")

        # OCR処理
        ocr_text = ocr.process_image(filepath)
        print(f"OCRテキスト: {ocr_text}")
        
        ocr_result = ocr.parse_receipt_text(ocr_text)
        print(f"解析結果: {ocr_result}")

        # Excel出力
        excel_filename = os.path.splitext(filename)[0] + ".xlsx"
        excel_filepath = os.path.join(app.config["UPLOAD_FOLDER"], excel_filename)
        excel.create_excel_receipt(ocr_result, excel_filepath)

        # 必要に応じてアップロードされた画像ファイルを削除することも可能
        # os.remove(filepath)

        return render_template("result.html", excel_file=excel_filename)
    else:
        flash("許可されていないファイル形式です。")
        return redirect(request.url)


@app.route("/download/<filename>")
def download_file(filename):
    filepath = os.path.join(app.config["UPLOAD_FOLDER"], filename)
    return send_file(filepath, as_attachment=True)


if __name__ == "__main__":
    if not os.path.exists(UPLOAD_FOLDER):
        os.makedirs(UPLOAD_FOLDER)
    app.run(debug=True)
