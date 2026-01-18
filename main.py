import streamlit as st
import easyocr
import pandas as pd
from pdf2image import convert_from_bytes
import io
import numpy as np
import gc
from openpyxl.styles import Border, Side, Alignment

st.set_page_config(page_title="AI PDF", layout="wide")
st.title("AI PDF 文字起こしアプリ")

def advanced_fix(text):
    # OCRが「説明」「発生例」を読み間違えるパターンを網羅
    text = text.replace("旅明", "説明").replace("訳明", "説明").replace("說明", "説明").replace("説 明", "説明")
    text = text.replace("光二例", "発生例").replace("え三例", "発生例").replace("発上例", "発生例").replace("え二例", "発生例").replace("生例", "発生例")
    text = text.replace("エラ一名", "エラー名").replace("エフー名", "エラー名").replace("工ラー", "エラー名").replace("エフ", "エラー名")
    return text

uploaded_file = st.file_uploader("PDFを選択してください", type="pdf")

if uploaded_file is not None:
    st.info("解析中...")
    reader = easyocr.Reader(['ja', 'en'], gpu=False)
    pdf_bytes = uploaded_file.getvalue()
    images = convert_from_bytes(pdf_bytes, dpi=150) # 少しだけ解像度を戻して精度アップ
    
    table_data = []
    current_item = {"エラー名": "", "説明": "", "発生例": ""}
    active_key = None

    for img in images:
        img_array = np.array(img)
        results = reader.readtext(img_array, detail=0)
        
        for text in results:
            clean_text = advanced_fix(text.strip())
            
            # 判定ロジック：キーワードが含まれていたら即座に切り替え
            if "エラー名" in clean_text:
                if current_item["エラー名"]:
                    table_data.append(current_item.copy())
                current_item = {"エラー名": clean_text.replace("エラー名","").replace(":","").strip(), "説明": "", "発生例": ""}
                active_key = "エラー名"
            elif "説明" in clean_text:
                active_key = "説明"
                current_item["説明"] += clean_text.replace("説明","").replace(":","").strip()
            elif "発生例" in clean_text:
                active_key = "発生例"
                current_item["発生例"] += clean_text.replace("発生例","").replace(":","").strip()
            elif active_key:
                # 空欄にならないよう、キーワードがない行は「active_key」に流し込む
                current_item[active_key] += " " + clean_text
        
        del img_array
        gc.collect()

    if current_item["エラー名"]:
        table_data.append(current_item)

    if table_data:
        df = pd.DataFrame(table_data)
        st.success("解析完了！")
        st.dataframe(df)

        # Excel保存時に「空欄」が目立たないよう、未取得を「-」で埋める
        df = df.replace("", "-")

        output = io.BytesIO()
        with pd.ExcelWriter(output, engine='openpyxl') as writer:
            df.to_excel(writer, index=False)
            ws = writer.sheets['Sheet1']
            # 見た目を綺麗にする設定
            for row in ws.iter_rows(min_row=1, max_row=len(df)+1, max_col=3):
                for cell in row:
                    cell.border = Border(top=Side(style='thin'), bottom=Side(style='thin'), left=Side(style='thin'), right=Side(style='thin'))
                    cell.alignment = Alignment(wrap_text=True, vertical='top')
            ws.column_dimensions['A'].width = 25
            ws.column_dimensions['B'].width = 45
            ws.column_dimensions['C'].width = 45

        st.download_button("改善版Excelを保存", data=output.getvalue(), file_name="improved_error_list.xlsx")
