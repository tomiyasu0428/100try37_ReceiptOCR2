from openpyxl import Workbook


def create_excel_receipt(data, output_path):
    """
    data: 領収書情報の辞書。キーは「発行日」「領収書番号」「宛名」など。
    output_path: 出力先のExcelファイルパス
    """
    wb = Workbook()
    ws = wb.active
    ws.title = "Receipt Data"

    # ヘッダー作成（抽出した項目名順）
    headers = ["発行者", "発行日", "領収書番号", "宛名", "金額", "支払方法", "但し書き"]
    ws.append(headers)

    # 各項目の値を並べる
    row = [data.get(header, "") for header in headers]
    ws.append(row)

    wb.save(output_path)
