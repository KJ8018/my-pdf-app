import streamlit as st
import easyocr
import pandas as pd
from pdf2image import convert_from_bytes
import io
import numpy as np
from PIL import Image
from openpyxl.styles import Border, Side, Alignment

st.set_page_config(page_title="AI PDF", layout="wide")
st.title("AI PDF 文字起こしアプリ")

# 誤字脱字を自動修正する辞書
def final_fix(text):
    d = {"汗fx": "if x", "ifx": "if x", "拓孤": "括弧", "閉じ志れ": "閉じ忘れ", "付け志れ": "付け忘れ", "說明": "説明", "工ラー": "エラー", "エラ一名": "エラー名"}
    for k, v in d.items():
        text = text.replace(k, v)
    return text

uploaded_file = st.file_uploader("PDFを選択してください", type="pdf")

if uploaded_file is not None:
    st.info("表を再構築中... (約1分)")
    reader = easyocr.Reader(['ja', 'en'], gpu=False)
    pdf_bytes = uploaded_file.getvalue()
    
    # メモリと精度のバランスをとったDPI=200
    images = convert_from_bytes(pdf_bytes, dpi=200)
    
    raw_texts = []
    for img in images:
        img_array = np.array(img)
        # テキストとその位置情報を取得
        results = reader.readtext(img_array, detail=0)
        raw_texts.extend(results)

    # --- 表を組み立てるロジック ---
    table_data = []
    current_item = {"エラー名": "", "説明": "", "発生例": ""}
    active_key = None

    for text in raw_texts:
        clean_text = final_fix(text.strip())
        
        # キーワードを検知して項目を振り分ける
        if "エラー名" in clean_text:
            if current_item["エラー名"]: # 次のエラー名が来たら保存
                table_data.append(current_item)
                current_item = {"エラー名": "", "説明": "", "発生例": ""}
            active_key = "エラー名"
            # 「エラー名: SyntaxError」のように1行に入っている場合
            val = clean_text.replace("エラー名", "").replace(":", "").replace("・", "").strip()
            current_item["エラー名"] = val
        elif "説明" in clean_text:
            active_key = "説明"
            val = clean_text.replace("説明", "").replace(":", "").strip()
            current_item["説明"] += val
        elif "発生例" in clean_text:
            active_key = "発生例"
            val = clean_text.replace("発生例", "").replace(":", "").strip()
            current_item["発生例"] += val
        elif active_key:
            # キーワードがない行は、直前の項目に付け足す
            current_item[active_key] += " " + clean_text

    if current_item["エラー名"]: # 最後の1件を追加
        table_data.append(current_item)

    # --- 表示とExcel保存 ---
    if table_data:
        df = pd.DataFrame(table_data)
        st.success(f"解析完了！ {len(df)} 件の項目を見つけました。")
        st.dataframe(df, use_container_width=True)

        # Excel作成（見た目を整える）
        excel_io = io.BytesIO()
        with pd.ExcelWriter(excel_io, engine='openpyxl') as writer:
            df.to_excel(writer, index=False, sheet_name='Result')
            ws = writer.sheets['Result']
            # 枠線と折り返し設定
            side = Side(style='thin')
            border = Border(top=side, bottom=side, left=side, right=side)
            for row in ws.iter_rows(min_row=1, max_row=len(df)+1, max_col=3):
                for cell in row:
                    cell.border = border
                    cell.alignment = Alignment(wrap_text=True, vertical='top')
            ws.column_dimensions['A'].width = 25
            ws.column_dimensions['B'].width = 40
            ws.column_dimensions['C'].width = 40

        st.download_button("Excelを保存 (発表用)", data=excel_io.getvalue(), 
                           file_name="AI_OCR_Result.xlsx", 
                           mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
