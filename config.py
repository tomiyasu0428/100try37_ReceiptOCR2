import os

# プロジェクトのベースディレクトリを取得
BASE_DIR = os.path.abspath(os.path.dirname(__file__))

# アップロードされたファイルの保存先
UPLOAD_FOLDER = os.path.join(BASE_DIR, "uploads")
EXCEL_FOLDER = os.path.join(BASE_DIR, "excel_files")

# アップロードを許可する拡張子
ALLOWED_EXTENSIONS = {"png", "jpg", "jpeg", "gif", "pdf"}

# アップロードファイルの最大サイズ（16MB）
MAX_CONTENT_LENGTH = 16 * 1024 * 1024  # 16MBまで許容
