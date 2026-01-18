import streamlit as st
import easyocr
from docx import Document
import pandas as pd
from pdf2image import convert_from_bytes
import io
import numpy as np
import re
from PIL import Image, ImageEnhance, ImageOps
from openpyxl.styles import Border, Side, Alignment

st.set_page_config(page_title="AI PDF ", layout="wide")
st.title("AI PDF 文字起こしアプリ")

def final_fix(text):
    """Python専門用語や読み間違いを徹底補正"""
    d = {
        "汗fx": "if x", "ifx": "if x", "deffunc": "def func", "Noneappend": "None.append",
        "メ=": "x =", "X=": "x =", "付け志れ": "付け忘れ", "閉じ志れ": "閉じ忘れ",
        "指孤": "括弧", "報おう": "扱おう", "1OOO": "1000", "mathexp": "math.exp",
        "Nlemory": "Memory", "VVorld": "World", "インボート": "インポート",
        "說明": "説明", "エラ一名": "エラー名", "工ラー": "エラー", "説 明": "説明", "発 生 例": "発生例"
    }
    for k, v in d.items():
        text = text.replace(k, v)
    return text

uploaded_file = st.file_uploader("PDFを選択してください", type="pdf")

if uploaded_file is not None:
    st.info("PDF解析中...（約1〜2分かかります）")
    reader = easyocr.Reader(['ja', 'en'])
    pdf_bytes = uploaded_file.getvalue()
    
    # DPIを400に上げ、さらに鮮明に
    images = convert_from_bytes(pdf_bytes, dpi=400)
    
    table_data = []
    current_item = {"エラー名": "", "説明": "", "発生例": ""}
    active_key = None

    for img in images:
        # 画像処理：グレースケール化とコントラスト最大化
        img = ImageOps.grayscale(img)
        img = ImageEnhance.Contrast(img).enhance(2.5)
        img_array = np.array(img)
        
        results = reader.readtext(img_array, detail=1)
        
        for res in results:
            raw_text = res[1].strip()
            clean_text = final_fix(raw_text)
            
            # 「エラー名」というキーワードを見つけたら、強制的に「新しい行」にする
            if any(k in clean_text for k in ["エラー名", "工ラー名", "エラ一名"]):
                if current_item["エラー名"]:
                    table_data.append(current_item)
                    current_item = {"エラー名": "", "説明": "", "発生例": ""}
                
                val = re.sub(r'^.*?名', '', clean_text).replace(":", "").replace("・", "").strip()
                current_item["エラー名"] = val
                active_key = "エラー名"
                
            elif any(k in clean_text for k in ["説明", "說明"]):
                active_key = "説明"
                val = re.sub(r'^.*?明', '', clean_text).replace(":", "").replace("・", "").strip()
                current_item["説明"] += val
                
            elif any(k in clean_text for k in ["発生例", "生例"]):
                active_key = "発生例"
                val = re.sub(r'^.*?例', '', clean_text).replace(":", "").replace("・", "").strip()
                current_item["発生例"] += val
            
            elif active_key:
                # 英語のエラー名(例: SyntaxError)の直後に日本語が来たら説明文へ飛ばす
                if active_key == "エラー名" and (len(clean_text) > 10 or any(c in clean_text for c in "あいうえお")):
                    current_item["説明"] += " " + clean_text
                    active_key = "説明"
                else:
                    current_item[active_key] += " " + clean_text

    # 最後のデータを保存
    if current_item["エラー名"]:
        table_data.append(current_item)

    df = pd.DataFrame(table_data)
    
    # 最終クリーンアップ（英単語と日本語の分離を再徹底）
    def post_process(row):
        name = str(row['エラー名'])
        match = re.search(r'([A-Za-z]+Error)(.*)', name)
        if match:
            row['エラー名'] = match.group(1)
            if match.group(2):
                row['説明'] = match.group(2).strip() + " " + str(row['説明'])
        
        # 全ての列に補正をかけ、nan（空データ）を除去
        row['説明'] = final_fix(str(row['説明']).replace("nan", ""))
        row['発生例'] = final_fix(str(row['発生例']).replace("nan", ""))
        return row

    if not df.empty:
        df = df.apply(post_process, axis=1)
        st.success(f"解析完了！ {len(df)} 件の項目を整理しました。")
        st.dataframe(df)

        # Excel保存（枠線付き）
        excel_io = io.BytesIO()
        with pd.ExcelWriter(excel_io, engine='openpyxl') as writer:
            df.to_excel(writer, index=False, sheet_name='ErrorList')
            ws = writer.sheets['ErrorList']
            side = Side(style='thin')
            border = Border(top=side, bottom=side, left=side, right=side)
            for r in ws.iter_rows(min_row=1, max_col=3, max_row=len(df)+1):
                for c in r:
                    c.border = border
                    c.alignment = Alignment(wrap_text=True, vertical='top')
            ws.column_dimensions['A'].width = 20
            ws.column_dimensions['B'].width = 45
            ws.column_dimensions['C'].width = 45

        st.download_button("Excelを保存", data=excel_io.getvalue(), file_name="final_error_table.xlsx")
