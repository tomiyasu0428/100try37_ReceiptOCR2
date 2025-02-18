import os
import pytesseract
import re
from PIL import Image


def preprocess_image(image):
    """画像の前処理を行い、OCRの精度を向上させる"""
    # グレースケールに変換
    if image.mode != "L":
        image = image.convert("L")

    return image


def process_image(image_path):
    """
    画像ファイルに対してOCR処理を実施します。
    PDFの場合はpdf2imageを用いて画像に変換後、各ページに対してOCR処理を行います。
    """
    print(f"OCR処理開始: {image_path}")
    ext = os.path.splitext(image_path)[1].lower()
    if ext == ".pdf":
        try:
            from pdf2image import convert_from_path

            print("PDFファイルを処理中...")
            # PDFの全ページを画像に変換（popplerのパス指定が必要な場合は、convert_from_pathの引数で設定してください）
            pages = convert_from_path(image_path)
            full_text = ""
            for i, page in enumerate(pages):
                print(f"ページ {i+1} を処理中...")
                # 画像の前処理
                processed_page = preprocess_image(page)
                text = pytesseract.image_to_string(processed_page, lang="jpn")
                full_text += text + "\n"
            print("PDF処理完了")
            return full_text
        except Exception as e:
            print(f"PDFエラー: {str(e)}")
            return f"PDFファイルのOCR処理中にエラーが発生しました: {str(e)}"
    else:
        try:
            print("画像ファイルを処理中...")
            img = Image.open(image_path)
            print(f"画像サイズ: {img.size}")
            print(f"画像モード: {img.mode}")
            # 画像の前処理
            processed_img = preprocess_image(img)
            text = pytesseract.image_to_string(processed_img, lang="jpn")
            print("画像処理完了")
            return text
        except Exception as e:
            print(f"画像エラー: {str(e)}")
            return f"OCR処理中にエラーが発生しました: {str(e)}"


def parse_receipt_text(ocr_text):
    """
    OCR結果のテキストから各項目を抽出し、辞書形式で返す。
    項目が見つからなければNoneをセットする。
    """
    print("OCRテキストの解析を開始:")
    print(ocr_text)

    # OCRの誤認識を修正
    ocr_text = ocr_text.replace("P", "2")  # 2が誤ってPと認識される場合がある
    ocr_text = ocr_text.replace("O", "0")  # 0が誤ってOと認識される場合がある

    patterns = {
        "発行日": r"(?:発行日|日付|date)[\s\:：]*([12]\d{3}(?:[\/\-年\.]\d{1,2}[\/\-月\.]\d{1,2}日?|\d{2}\/\d{2}))",
        "領収書番号": r"(?:領収書番号|No|NO|番号)[\s\:：]*([A-Za-z0-9\-]+)",
        "宛名": r"(?:宛名|支払人|お名前)[\s\:：]*(.+?)(?:\n|$)",
        "金額": r"(?:金額|合計|￥|\¥)[\s\:：]*[\¥\$\\]?\s*([0-9,]+)(?:円|$)",
        "支払方法": r"(?:支払方法|お支払|支払)[\s\:：]*(.+?)(?:\n|$)",
        "但し書き": r"(?:但し書き|但書|内容|品目)[\s\:：]*(.+?)(?:\n|$)",
        "発行者": r"(?:発行者|会社名)[\s\:：]*(?!.*(?:発行日|日付))(.+?)(?:\n|$)",  # 発行日や日付を含まない行のみマッチ
    }

    data = {}
    print("\nデバッグ情報:")
    print("OCRテキスト（行ごと）:")
    for i, line in enumerate(ocr_text.split('\n')):
        print(f"行 {i+1}: {line}")
    
    print("\nパターンマッチング結果:")
    for field, pattern in patterns.items():
        matches = re.finditer(pattern, ocr_text, re.IGNORECASE | re.MULTILINE)
        print(f"\n{field}のマッチング:")
        found = False
        for match in matches:
            found = True
            print(f"  マッチした文字列: {match.group(0)}")
            print(f"  抽出された値: {match.group(1)}")
            print(f"  位置: {match.start()}-{match.end()}")
        if not found:
            print("  マッチなし")

        match = re.search(pattern, ocr_text, re.IGNORECASE | re.MULTILINE)
        if match:
            value = match.group(1).strip()
            print(f"{field}: {value}")
            data[field] = value
        else:
            print(f"{field}: 見つかりませんでした")
            data[field] = None

    # 金額のカンマと通貨記号を削除
    if data["金額"]:
        data["金額"] = (
            data["金額"].replace(",", "").replace("\\", "").replace("¥", "").replace("￥", "").strip()
        )

    return data
