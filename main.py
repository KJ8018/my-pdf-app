import streamlit as st
import easyocr
import pandas as pd
from pdf2image import convert_from_bytes
import io
import numpy as np
import re
from PIL import Image, ImageEnhance, ImageOps
from openpyxl.styles import Border, Side, Alignment

st.set_page_config(page_title="AI PDF", layout="wide")
st.title("AI PDF 文字起こしアプリ")

def final_fix(text):
    d = {
        "汗fx": "if x", "ifx": "if x", "deffunc": "def func", "Noneappend": "None.append",
        "メ=": "x =", "X=": "x =", "付け志れ": "付け忘れ", "閉じ志れ": "閉じ忘れ",
        "指孤": "括弧", "報おう": "扱おう", "1OOO": "1000", "mathexp": "math.exp",
        "インボート": "インポート", "説明": "説明", "エラ一名": "エラー名", "工ラー": "エラー"
    }
    for k, v in d.items():
        text = text.replace(k, v)
    return text

uploaded_file = st.file_uploader("PDFを選択してください", type="pdf")

if uploaded_file is not None:
    st.info("解析を開始します。しばらくお待ちください...")
    
    # 1. OCR準備
    reader = easyocr.Reader(['ja', 'en'])
    pdf_bytes = uploaded_file.getvalue()
    
    # 2. PDFを画像に変換 (DPIを200に下げてメモリ節約)
    images = convert_from_bytes(pdf_bytes, dpi=200)
    
    table_data = []
    current_item = {"エラー名": "", "説明": "", "発生例": ""}
    active_key = None

    for img in images:
        img = ImageOps.grayscale(img)
        img_array = np.array(img)
        results = reader.readtext(img_array, detail=1)
        
        for res in results:
            clean_text = final_fix(res[1].strip())
            
            if any(k in clean_text for k in ["エラー名", "エラ一名"]):
                if current_item["エラー名"]: table_data.append(current_item)
                current_item = {"エラー名": "", "説明": "", "発生例": ""}
                val = re.sub(r'^.*?名', '', clean_text).replace(":", "").strip()
                current_item["エラー名"] = val
                active_key = "エラー名"
            elif any(k in clean_text for k in ["説明"]):
                active_key = "説明"
                val = re.sub(r'^.*?明', '', clean_text).replace(":", "").strip()
                current_item["説明"] += val
            elif any(k in clean_text for k in ["発生例"]):
                active_key = "発生例"
                val = re.sub(r'^.*?例', '', clean_text).replace(":", "").strip()
                current_item["発生例"] += val
            elif active_key:
                current_item[active_key] += " " + clean_text

    if current_item["エラー名"]: table_data.append(current_item)

    # 3. 表示と保存
    if table_data:
        df = pd.DataFrame(table_data)
        st.success("解析が完了しました！")
        st.dataframe(df)

        excel_io = io.BytesIO()
        with pd.ExcelWriter(excel_io, engine='openpyxl') as writer:
            df.to_excel(writer, index=False)
        st.download_button("Excelを保存", data=excel_io.getvalue(), file_name="error_list.xlsx")
