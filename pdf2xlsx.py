import pdfplumber
import pandas as pd
import sys
import os

# セルの内容を処理する関数
def process_cell(cell):
    # 文字列でない場合はそのまま返す
    if not isinstance(cell, str):
        return cell
    # 文字列である場合、空白で分割
    parts = cell.split()
    
    converted_parts = []
    for p in parts:
        # 数値変換を試みる
        try:
            val = float(p) if '.' in p else int(p)
        except ValueError:
            val = p
        converted_parts.append(val)
    # 分割結果をカンマ区切りにまとめる
    return ", ".join(map(str, converted_parts))

if __name__ == "__main__":
    # 標準入力からPDFファイル名を取得
    input_pdf = input("PDFファイル名を入力してください: ").strip()
    input_pdf = os.path.join("input", input_pdf)

    if not os.path.exists(input_pdf):
        print(f"Error: File {input_pdf} not found.")
        sys.exit(1)

    # 出力ファイル名をベースネームから生成
    base_name = os.path.splitext(os.path.basename(input_pdf))[0]
    output_dir = "output"
    os.makedirs(output_dir, exist_ok=True)
    output_excel = os.path.join(output_dir, base_name + ".xlsx")
    
    with pdfplumber.open(input_pdf) as pdf:
        all_rows = []
        for page in pdf.pages:
            # ページごとにテーブルを抽出
            tables = page.extract_tables()
            for table in tables:
                # 各行を処理
                for row in table:
                    # 各セルを処理
                    processed_row = [process_cell(cell) for cell in row]
                    all_rows.append(processed_row)

    # 各行を最大長に合わせてNoneで埋める
    max_length = max((len(r) for r in all_rows), default=0)
    normalized_rows = [r + [None]*(max_length-len(r)) for r in all_rows]

    df = pd.DataFrame(normalized_rows)

    # ここからカンマ区切りになったセルを行方向に展開する処理を追加
    # ただし ", "（カンマ＋スペース）がある場合のみ分割
    target_cols = df.columns.tolist()

    # ", "を含む場合のみリスト化
    for col in target_cols:
        df[col] = df[col].apply(lambda x: x.split(', ') if isinstance(x, str) and ', ' in x else [x])

    # 各行を、target_colsの中で最大のリスト要素数に応じて複製し展開
    expanded_rows = []
    for i, row in df.iterrows():
        # 各ターゲット列のリスト長を取得
        lengths = [len(row[c]) for c in target_cols if isinstance(row[c], list)]
        max_len = max(lengths) if lengths else 1
        # 最大長に合わせて複製
        for j in range(max_len):
            new_row = row.copy()
            for c in target_cols:
                vals = row[c]
                # j番目の要素があればそれを使用、なければNone
                val = vals[j] if j < len(vals) else None
                new_row[c] = val
            expanded_rows.append(new_row)

    df_expanded = pd.DataFrame(expanded_rows)

    # 数値に変換可能なものは変換する（必要なら）
    for c in df_expanded.columns:
        # 数値変換を試みる
        try:
            df_expanded[c] = pd.to_numeric(df_expanded[c])
        except:
            # 数値変換できない列はそのまま
            pass

    # Excelとして保存
    df_expanded.to_excel(output_excel, index=False, header=False)
    print(f"pdfをExcelに変換して保存しました: {output_excel}")