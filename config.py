import os

# プロジェクトのベースディレクトリを取得
BASE_DIR = os.path.abspath(os.path.dirname(__file__))

UPLOAD_FOLDER = os.path.join(BASE_DIR, "uploads")
ALLOWED_EXTENSIONS = {"png", "jpg", "jpeg", "gif", "pdf"}
MAX_CONTENT_LENGTH = 16 * 1024 * 1024  # 16MBまで許容
