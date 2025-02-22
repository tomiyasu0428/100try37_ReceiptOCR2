import os
from dotenv import load_dotenv
import cv2
import numpy as np
import pytesseract
from PIL import Image
import google.generativeai as genai
import re
import sys
import base64
import json
import pandas as pd
from pathlib import Path

# .envファイルから環境変数を読み込む
load_dotenv()


def preprocess_image(image_path):
    """画像の前処理を行う"""
    try:
        # 画像を読み込む
        image = cv2.imread(image_path)
        if image is None:
            print(f"画像の読み込みに失敗: {image_path}")
            return None

        # チャンネル数を確認
        if len(image.shape) == 2:  # グレースケール画像の場合
            image = cv2.cvtColor(image, cv2.COLOR_GRAY2BGR)
        elif image.shape[2] == 4:  # RGBA画像の場合
            image = cv2.cvtColor(image, cv2.COLOR_BGRA2BGR)

        # グレースケールに変換
        gray = cv2.cvtColor(image, cv2.COLOR_BGR2GRAY)

        # ノイズ除去（メディアンフィルタ）
        denoised = cv2.medianBlur(gray, 3)

        # 傾き補正
        coords = np.column_stack(np.where(denoised > 0))
        angle = cv2.minAreaRect(coords)[-1]
        if angle < -45:
            angle = 90 + angle
        center = tuple(np.array(denoised.shape[1::-1]) / 2)
        rot_mat = cv2.getRotationMatrix2D(center, angle, 1.0)
        rotated = cv2.warpAffine(
            denoised, rot_mat, denoised.shape[1::-1], flags=cv2.INTER_CUBIC, borderMode=cv2.BORDER_REPLICATE
        )

        # コントラスト強調
        clahe = cv2.createCLAHE(clipLimit=2.0, tileGridSize=(8, 8))
        enhanced = clahe.apply(rotated)

        # アダプティブ閾値処理
        binary = cv2.adaptiveThreshold(
            enhanced, 255, cv2.ADAPTIVE_THRESH_GAUSSIAN_C, cv2.THRESH_BINARY, 11, 2
        )

        # モルフォロジー演算（ノイズ除去）
        kernel = np.ones((2, 2), np.uint8)
        cleaned = cv2.morphologyEx(binary, cv2.MORPH_CLOSE, kernel)

        # デバッグ用に前処理後の画像を保存
        debug_path = os.path.join(os.path.dirname(image_path), "debug_" + os.path.basename(image_path))
        cv2.imwrite(debug_path, cleaned)
        print(f"前処理後の画像を保存: {debug_path}")

        return cleaned

    except Exception as e:
        print(f"画像前処理エラー: {str(e)}")
        return None


def use_gemini_api(image_path, field_name=None):
    """
    Gemini APIを使用して特定のフィールドを抽出
    """
    try:
        api_key = os.getenv("GEMINI_API_KEY")
        if not api_key:
            print("Gemini APIキーが設定されていません")
            return None

        genai.configure(api_key=api_key)
        model = genai.GenerativeModel("gemini-1.5-flash")

        image = Image.open(image_path)

        if field_name:
            # 特定のフィールドの抽出
            prompts = {
                "発行日": """
                この領収書から発行日を抽出してください。
                - YYYY/MM/DD形式で返してください
                - 日付のみを返してください（余計な文字は不要）
                - 電話番号や注文番号などは無視してください
                """,
                "支払先名": """
                この領収書から店舗・会社名を抽出してください。
                - 正式名称を返してください
                - 支店名や店舗名も含めてください
                - 住所は含めないでください
                - 電話番号は含めないでください
                """,
                "金額": """
                この領収書から合計金額を抽出してください。
                - 数値のみを返してください（カンマや円記号は不要）
                - 税込の最終合計金額を返してください
                - 小計や税額は無視してください
                """,
                "インボイス番号": """
                この領収書からインボイス番号（登録番号）を抽出してください。
                - 数値のみを返してください
                - T+13桁の数字、または13桁以上の登録番号を探してください
                - 通常「登録番号」という文字列の後に記載されています
                """,
            }
            prompt = prompts.get(field_name, f"この領収書から「{field_name}」を抽出してください。")
        else:
            # 全フィールドの抽出
            prompt = """
            この領収書から以下の情報を抽出し、正確にJSON形式で返してください。
            必ず以下のフォーマットで返してください：

            {
                "発行日": "YYYY/MM/DD形式で。日付のみを抽出。電話番号は無視",
                "支払先名": "店舗・会社の正式名称。支店名も含める。住所や電話番号は含めない",
                "金額": "税込の最終合計金額。数値のみ（カンマや円記号は不要）",
                "インボイス番号": "T+13桁の数字、または登録番号。数値のみ"
            }

            - 各フィールドは必ず指定された形式で返してください
            - 特定の情報が見つからない場合は、空文字列を設定してください
            - 余計な説明は不要です。JSONのみを返してください
            """

        response = model.generate_content([prompt, image])
        print(f"Gemini API レスポンス: {response.text}")

        if field_name:
            # 数値のクリーンアップ
            value = response.text.strip()
            if field_name in ["金額", "インボイス番号"]:
                value = re.sub(r"[^\d]", "", value)
            elif field_name == "発行日":
                # YYYY/MM/DD形式に標準化
                date_match = re.search(r"\d{4}[-/年]\d{1,2}[-/月]\d{1,2}", value)
                if date_match:
                    date_str = date_match.group()
                    date_str = date_str.replace("年", "/").replace("月", "/").replace("日", "")
                    value = date_str
            return value
        else:
            # JSON形式の応答をパース
            try:
                # レスポンスから{...}の部分を抽出
                json_str = re.search(r"\{[^{}]*\}", response.text)
                if not json_str:
                    print(f"JSONが見つかりませんでした。レスポンス全文:\n{response.text}")
                    return None

                # JSONをパース
                result = json.loads(json_str.group())

                # 必須キーの存在確認と初期化
                required_keys = ["発行日", "支払先名", "金額", "インボイス番号"]
                for key in required_keys:
                    if key not in result:
                        result[key] = ""

                # 数値のクリーンアップ
                if result.get("金額"):
                    result["金額"] = re.sub(r"[^\d]", "", str(result["金額"]))
                if result.get("インボイス番号"):
                    result["インボイス番号"] = re.sub(r"[^\dT]", "", str(result["インボイス番号"]))

                return result

            except json.JSONDecodeError as e:
                print(f"JSONパースエラー: {str(e)}\nレスポンス全文:\n{response.text}")
                return None
            except Exception as e:
                print(f"Gemini API結果の処理中にエラー: {str(e)}")
                return None

    except Exception as e:
        print(f"Gemini API error: {str(e)}")
        return None


def process_image_with_gemini(image_path):
    """GeminiでOCR結果を解析"""
    try:
        # まず全項目を一括で取得
        api_result = use_gemini_api(image_path)
        print("Gemini API 全項目抽出結果:")
        print(api_result)

        if api_result:
            # 欠けている項目をGeminiの結果で補完
            for field in api_result.keys():
                if not api_result[field]:
                    value = use_gemini_api(image_path, field)
                    if value:
                        api_result[field] = value
                        print(f"{field}: GeminiAPIの結果で補完 -> {value}")

        return api_result

    except Exception as e:
        print(f"Gemini API処理エラー: {str(e)}")
        return None


def process_image(image_path):
    """画像ファイルに対してOCR処理を実施"""
    try:
        print("=== OCR処理開始 ===")

        # 画像の前処理
        processed_image = preprocess_image(image_path)
        if processed_image is None:
            print("画像の前処理に失敗しました")
            return None

        # Tesseractでテキスト抽出
        ocr_text = pytesseract.image_to_string(processed_image, lang="jpn")
        result = process_ocr_result(ocr_text)

        # OCRの結果が不十分な場合、Geminiを使用
        if not result or not all([result.get("発行日"), result.get("支払先名"), result.get("金額")]):
            print("Tesseract OCRの結果が不十分です。Geminiを使用して再試行します。")
            gemini_result = use_gemini_api(image_path)
            if gemini_result:
                return gemini_result

        return result

    except Exception as e:
        print(f"OCR処理エラー: {str(e)}")
        return None


def process_pdf(pdf_path):
    """PDFファイルに対してOCR処理を実施"""
    try:
        from pdf2image import convert_from_path

        # PDFを画像に変換
        print("=== PDF変換開始 ===")
        pages = convert_from_path(pdf_path)

        results = []
        for i, page in enumerate(pages):
            print(f"\nページ {i+1} の処理を開始")

            # 一時的に画像を保存
            temp_image_path = os.path.join(os.path.dirname(pdf_path), f"temp_page_{i+1}.png")
            page.save(temp_image_path, "PNG")

            # 画像に対してOCR処理を実行
            result = process_image(temp_image_path)
            if result:
                results.append(result)

            # 一時ファイルを削除
            try:
                os.remove(temp_image_path)
            except:
                pass

        print("\n=== 全ページの処理が完了しました ===")

        if not results:
            print("データを抽出できませんでした")
            return None

        # PDFの場合は全ての結果をリストとして返す
        return results

    except Exception as e:
        print(f"PDF処理エラー: {str(e)}")
        return None


def process_multiple_files(file_paths):
    """複数のファイルを処理してExcelに出力"""
    results = []

    for file_path in file_paths:
        path = Path(file_path)
        if path.suffix.lower() == ".pdf":
            # PDFの場合
            pdf_results = process_pdf(str(path))
            if pdf_results:
                results.extend(pdf_results)
        else:
            # 画像ファイルの場合
            result = process_image(str(path))
            if result:
                results.append(result)

    if results:
        # 結果をDataFrameに変換
        df = pd.DataFrame(results)

        # Excelファイルとして保存
        output_path = "receipt_results.xlsx"
        df.to_excel(output_path, index=False)
        print(f"結果を{output_path}に保存しました。")
        return output_path

    return None


def main(image_path):
    """
    画像ファイルに対してOCR処理を実施します。
    PDFの場合はpdf2imageを用いて画像に変換後、各ページに対してOCR処理を行います。
    """
    try:
        # ファイルの拡張子を取得
        file_ext = os.path.splitext(image_path)[1].lower()

        # PDFファイルの場合
        if file_ext == ".pdf":
            return process_pdf(image_path)
        # 画像ファイルの場合
        elif file_ext in [".jpg", ".jpeg", ".png", ".tiff", ".bmp"]:
            return process_image(image_path)
        else:
            raise ValueError(f"サポートされていないファイル形式です: {file_ext}")

    except Exception as e:
        print(f"処理エラー: {str(e)}")
        return None


def combine_results(ocr_result, gemini_result):
    """OCRとGemini APIの結果を統合して最適な結果を返す"""
    final_result = {"発行日": "", "支払先名": "", "金額": "", "インボイス番号": ""}

    # 各フィールドについて、より信頼できる値を選択
    for field in final_result.keys():
        ocr_value = ocr_result.get(field, "")
        gemini_value = gemini_result.get(field, "")

        # 金額の場合、数値のみを抽出して比較
        if field == "金額":
            ocr_amount = (
                int(re.sub(r"[^\d]", "", ocr_value)) if re.sub(r"[^\d]", "", ocr_value).isdigit() else 0
            )
            gemini_amount = (
                int(re.sub(r"[^\d]", "", gemini_value)) if re.sub(r"[^\d]", "", gemini_value).isdigit() else 0
            )

            # より大きい金額を採用（小額の値引きなどは無視）
            final_result[field] = str(max(ocr_amount, gemini_amount))

        # インボイス番号の場合、形式をチェック
        elif field == "インボイス番号":
            ocr_invoice = re.sub(r"[^\dT]", "", ocr_value)
            gemini_invoice = re.sub(r"[^\dT]", "", gemini_value)

            if re.match(r"T\d{13}", ocr_invoice):
                final_result[field] = ocr_invoice
            elif re.match(r"T\d{13}", gemini_invoice):
                final_result[field] = gemini_invoice
            else:
                final_result[field] = (
                    ocr_invoice if len(ocr_invoice) > len(gemini_invoice) else gemini_invoice
                )

        # 日付の場合、形式をチェック
        elif field == "発行日":
            # 日付形式の正規化
            ocr_date = re.sub(r"[年月]", "/", ocr_value).replace("日", "")
            gemini_date = re.sub(r"[年月]", "/", gemini_value).replace("日", "")

            # より完全な日付形式を採用
            if re.match(r"\d{4}/\d{1,2}/\d{1,2}", ocr_date):
                final_result[field] = ocr_date
            elif re.match(r"\d{4}/\d{1,2}/\d{1,2}", gemini_date):
                final_result[field] = gemini_date
            else:
                final_result[field] = ocr_date if len(ocr_date) > len(gemini_date) else gemini_date

        # その他のフィールドは、より長い値を採用
        else:
            final_result[field] = ocr_value if len(ocr_value) > len(gemini_value) else gemini_value

    return final_result


def extract_text_from_image(image):
    """画像からテキストを抽出"""
    try:
        if isinstance(image, str):
            # 画像パスが渡された場合は読み込む
            image = cv2.imread(image)
            if image is None:
                raise ValueError("画像の読み込みに失敗しました")

        # OCR設定
        custom_config = r"--oem 3 --psm 6 -l jpn+eng"

        # OCRを実行し、信頼度スコアを取得
        data = pytesseract.image_to_data(image, config=custom_config, output_type=pytesseract.Output.DICT)

        # 信頼度の高い文字列のみを抽出
        text_lines = []
        n_boxes = len(data["text"])
        for i in range(n_boxes):
            if float(data["conf"][i]) > 60:  # 信頼度60%以上の文字列のみを使用
                text_lines.append(data["text"][i])

        text = " ".join(text_lines)

        # デバッグ出力
        print("\n=== OCR抽出テキスト（信頼度60%以上） ===")
        print(text)
        print("=====================================\n")

        return text

    except Exception as e:
        print(f"OCR処理エラー: {str(e)}")
        return None


def validate_amount(amount_str):
    """金額の妥当性を検証"""
    try:
        # 数値以外の文字を除去
        amount = int(re.sub(r"[^\d]", "", amount_str))

        # 基本的な検証ルール
        if amount <= 0:
            return False
        if amount > 10000000:  # 1000万円以上は不正解の可能性が高い
            return False
        if len(str(amount)) > 8:  # 8桁以上は不正解の可能性が高い
            return False

        # 金額の一般的なパターン
        if amount % 10 != 0:  # 1の位が0でない場合は不正解の可能性が高い
            return False

        return True
    except:
        return False


def validate_date(date_str):
    """日付の妥当性を検証"""
    try:
        # YYYY/MM/DD形式に変換
        date_str = date_str.replace("年", "/").replace("月", "/").replace("日", "")
        date_parts = date_str.split("/")

        if len(date_parts) != 3:
            return False

        year = int(date_parts[0])
        month = int(date_parts[1])
        day = int(date_parts[2])

        # 基本的な検証ルール
        if not (1900 <= year <= 2100):  # 年の範囲チェック
            return False
        if not (1 <= month <= 12):
            return False
        if not (1 <= day <= 31):
            return False

        return True
    except:
        return False


def extract_info_from_text(text):
    """テキストから情報を抽出（信頼度評価付き）"""
    result = {"発行日": "", "支払先名": "", "金額": "", "インボイス番号": ""}

    lines = text.split("\n")

    # 金額の候補を収集
    amount_candidates = []
    amount_patterns = [
        (
            r"(?:領収金額|お支払金額|ご利用金額|合計金額|利用金額|お会計|小計)[\s:：]*[¥\\]?[\s]*([0-9,\.]+)",
            3,
        ),  # 重み付け
        (r"(?:金額)[\s:：]*[¥\\]?[\s]*([0-9,\.]+)", 2),
        (r"[¥\\][\s]*([0-9,\.]+)", 1),
        (r"([0-9]+[\s\.]+[0-9]+)", 1),
        (r"(?:税込|税込み|税込金額)[\s:：]*[¥\\]?[\s]*([0-9,\.]+)", 2),
    ]

    for line in lines:
        for pattern, weight in amount_patterns:
            matches = re.finditer(pattern, line)
            for match in matches:
                amount_str = match.group(1).strip()
                amount = normalize_amount(amount_str)
                if amount > 0 and validate_amount(str(amount)):
                    amount_candidates.append(
                        {"amount": amount, "weight": weight, "line": line, "pattern": pattern}
                    )
                    print(f"金額候補を検出: {amount} (パターン: {pattern}, 重み: {weight})")

    # 重み付けと出現頻度を考慮して最適な金額を選択
    if amount_candidates:
        amount_scores = {}
        for candidate in amount_candidates:
            amount = candidate["amount"]
            weight = candidate["weight"]
            amount_scores[amount] = amount_scores.get(amount, 0) + weight

        # 最も高いスコアの金額を採用
        best_amount = max(amount_scores.items(), key=lambda x: x[1])[0]
        result["金額"] = str(best_amount)
        print(f"最終的な金額を決定: {best_amount} (スコア: {amount_scores[best_amount]})")

    # 日付の検出（信頼度評価付き）
    date_patterns = [
        (r"(\d{4}[-/年]\d{1,2}[-/月]\d{1,2})", 3),  # YYYY/MM/DD形式
        (r"(?:令和|R)(\d{1,2})年\d{1,2}月\d{1,2}日", 2),  # 令和表記
        (r"(\d{2,4}[-/年]\d{1,2}[-/月]\d{1,2})", 1),  # その他の日付形式
    ]

    date_candidates = []
    for line in lines:
        for pattern, weight in date_patterns:
            match = re.search(pattern, line)
            if match:
                date_str = match.group(1)
                if "令和" in line or "R" in line:
                    reiwa_year = int(re.search(r"(?:令和|R)(\d{1,2})", line).group(1))
                    date_str = f"{2018 + reiwa_year}{date_str[2:]}"
                elif len(date_str.split("/")[0]) == 2:
                    year = int(date_str.split("/")[0])
                    if year < 50:
                        date_str = f"20{date_str}"
                    else:
                        date_str = f"19{date_str}"

                date_str = date_str.replace("年", "/").replace("月", "/").replace("日", "")
                if validate_date(date_str):
                    date_candidates.append({"date": date_str, "weight": weight, "line": line})
                    print(f"日付候補を検出: {date_str} (重み: {weight})")

    # 最も信頼度の高い日付を選択
    if date_candidates:
        best_date = max(date_candidates, key=lambda x: x["weight"])
        result["発行日"] = best_date["date"]
        print(f"最終的な日付を決定: {best_date['date']} (重み: {best_date['weight']})")

    # 支払先名の検出（優先度付き）
    company_candidates = []
    for line in lines:
        if "様" in line or "御中" in line:
            company = re.sub(r"[\s　]+", " ", line).strip()
            weight = 3 if "様" in line and "御中" in line else 2
            company_candidates.append({"name": company, "weight": weight, "line": line})
            print(f"支払先候補を検出: {company} (重み: {weight})")

    if company_candidates:
        best_company = max(company_candidates, key=lambda x: x["weight"])
        result["支払先名"] = best_company["name"]
        print(f"最終的な支払先を決定: {best_company['name']} (重み: {best_company['weight']})")

    # インボイス番号の検出（検証付き）
    invoice_patterns = [
        (r"(?:登録番号|登録得号)[\s:：]*T?([0-9]{13})", 3),
        (r"T([0-9]{13})", 2),
    ]

    invoice_candidates = []
    for line in lines:
        for pattern, weight in invoice_patterns:
            match = re.search(pattern, line)
            if match:
                invoice_num = match.group(1)
                if len(invoice_num) == 13:  # 13桁であることを確認
                    invoice_candidates.append(
                        {
                            "number": "T" + invoice_num if not invoice_num.startswith("T") else invoice_num,
                            "weight": weight,
                            "line": line,
                        }
                    )
                    print(f"インボイス番号候補を検出: {invoice_num} (重み: {weight})")

    if invoice_candidates:
        best_invoice = max(invoice_candidates, key=lambda x: x["weight"])
        result["インボイス番号"] = best_invoice["number"]
        print(f"最終的なインボイス番号を決定: {best_invoice['number']} (重み: {best_invoice['weight']})")

    return result


def normalize_amount(amount_str):
    """金額文字列を正規化して数値に変換"""
    try:
        # 空白とカンマを除去
        amount_str = amount_str.replace(" ", "").replace(",", "")

        # ドット区切りの処理
        if "." in amount_str:
            parts = amount_str.split(".")
            if len(parts) == 2:
                # 8.800 -> 8800のように変換
                amount_str = parts[0] + parts[1].ljust(3, "0")

        # 数値以外の文字を除去
        amount_str = re.sub(r"[^\d]", "", amount_str)

        if amount_str:
            amount = int(amount_str)
            # 異常に小さい値や大きい値は除外
            if amount < 10 or amount > 10000000:  # 1000万円を超える金額は誤認識の可能性が高い
                return 0
            return amount
    except (ValueError, TypeError):
        return 0
    return 0


def extract_date(text):
    """テキストから日付を抽出する"""
    try:
        # 和暦と西暦の両方に対応
        date_patterns = [
            r"(令和|R|㎶|H)\s*(\d{1,2}|\元)年\s*(\d{1,2})月\s*(\d{1,2})日",
            r"(\d{4}|\d{2})年\s*(\d{1,2})月\s*(\d{1,2})日",
            r"(\d{4}|\d{2})/(\d{1,2})/(\d{1,2})",
            r"(\d{4}|\d{2})-(\d{1,2})-(\d{1,2})",
        ]

        for pattern in date_patterns:
            match = re.search(pattern, text)
            if match:
                return format_date(match)
        return ""
    except Exception as e:
        print(f"日付抽出エラー: {str(e)}")
        return ""


def extract_amount(text):
    """テキストから金額を抽出する"""
    try:
        # 金額のパターンを定義
        amount_patterns = [
            (r"¥\s*(\d[\d,]*)", 2),  # ¥マークで始まる金額
            (r"\\s*(\d[\d,]*)", 2),  # \マークで始まる金額
            (r"合計\s*[:：]?\s*(\d[\d,]*)", 2),  # 合計の後の金額
            (r"金額\s*[:：]?\s*(\d[\d,]*)", 2),  # 金額の後の数字
            (r"([0-9]+[\s\.]+[0-9]+)", 1),  # 数字のみの場合（最も優先度低）
            (r"([0-9]+)円", 1),  # 円で終わる金額
        ]

        candidates = []
        for pattern, weight in amount_patterns:
            matches = re.finditer(pattern, text)
            for match in matches:
                amount_str = match.group(1).replace(",", "").replace(" ", "")
                if amount_str:
                    try:
                        amount = int(amount_str)
                        if 100 <= amount <= 10000000:  # 妥当な金額範囲
                            print(f"金額候補を検出: {amount} (パターン: {pattern}, 重み: {weight})")
                            candidates.append((amount, weight))
                    except ValueError:
                        continue

        if candidates:
            # 重みが最も高い金額を選択
            amount, score = max(candidates, key=lambda x: x[1])
            print(f"最終的な金額を決定: {amount} (スコア: {score})")
            return str(amount)
        return ""
    except Exception as e:
        print(f"金額抽出エラー: {str(e)}")
        return ""


def process_ocr_result(ocr_text, confidence_threshold=60):
    """OCR結果を処理し、必要な情報を抽出する"""
    try:
        # 信頼度の高いテキストのみを使用
        filtered_text = "\n".join(
            [
                word
                for word in ocr_text.split("\n")
                if any(char.strip() for char in word)  # 空白文字のみの行を除外
            ]
        )

        # 各項目を抽出
        extracted_data = {
            "発行日": extract_date(filtered_text),
            "支払先名": extract_company_name(filtered_text),
            "金額": extract_amount(filtered_text),
            "インボイス番号": extract_invoice_number(filtered_text),
        }

        # 結果の検証
        if not any(extracted_data.values()):
            print("警告: 全ての項目の抽出に失敗しました")
            return None

        return extracted_data
    except Exception as e:
        print(f"OCR結果の処理中にエラーが発生しました: {str(e)}")
        return None


def extract_company_name(text):
    """テキストから会社名を抽出する"""
    try:
        # 会社名のパターン
        company_patterns = [
            r"(.+)(?:株式会社|有限会社|合同会社|事務所)",  # 会社形態が後ろにある場合
            r"(?:株式会社|有限会社|合同会社)(.+)",  # 会社形態が前にある場合
            r"(.+)(?:様|御中)",  # 様や御中で終わる場合
        ]

        for pattern in company_patterns:
            match = re.search(pattern, text)
            if match:
                company_name = match.group(1).strip()
                if company_name:
                    return company_name

        return ""
    except Exception as e:
        print(f"会社名抽出エラー: {str(e)}")
        return ""


if __name__ == "__main__":
    if len(sys.argv) > 1:
        files = sys.argv[1:]
        if len(files) > 1:
            # 複数ファイルの処理
            result = process_multiple_files(files)
        else:
            # 単一ファイルの処理
            result = main(files[0])
        print(result)
    else:
        print("使用方法: python ocr.py <画像ファイルまたはPDFファイルのパス>")
