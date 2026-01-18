import streamlit as st
import easyocr
import pandas as pd
from pdf2image import convert_from_bytes
import io
import numpy as np
import re
from PIL import Image
from openpyxl.styles import Border, Side, Alignment

st.set_page_config(page_title="AI PDF", layout="wide")
st.title("AI PDF 文字起こしアプリ")

# --- OCRの読み間違いを徹底的に直す辞書 ---
def super_fix(text):
    d = {
        "エフ": "エラー名", "エラ一名": "エラー名", "工ラー": "エラー", "エフー名": "エラー名",
        "說明": "説明", "訳明": "説明", "峻的な": "一般的な", "俊的な": "一般的な",
        "光二例": "発生例", "え三例": "発生例", "え二例": "発生例", "発上例": "発生例", "光上例": "発生例",
        "毒照": "参照", "蓋照": "参照", "報おう": "扱おう", "坂おう": "扱おう", "汲う": "扱う",
        "拓孤": "括弧", "持孤": "括弧", "亮れ": "忘れ", "志れ": "忘れ",
        "子期せず": "予期せず", "子期しない": "予期しない",
        "detfuncU": "def func()", "ixっ": "if x", "VNorld": "World", "ipurt": "import"
    }
    for k, v in d.items():
        text = text.replace(k, v)
    return text

uploaded_file = st.file_uploader("PDFを選択してください", type="pdf")

if uploaded_file is not None:
    st.info("AIが表を解析して組み立てています。1〜2分ほどお待ちください...")
    
    # OCRエンジンの起動
    reader = easyocr.Reader(['ja', 'en'], gpu=False)
    
    # PDFを画像に変換 (精度重視のDPI 200)
    pdf_bytes = uploaded_file.getvalue()
    images = convert_from_bytes(pdf_bytes, dpi=200)
    
    table_data = []
    current_item = {"エラー名": "", "説明": "", "発生例": ""}
    active_key = None

    for img in images:
        img_array = np.array(img)
        # テキスト抽出
        results = reader.readtext(img_array, detail=0)
        
        for text in results:
            # 抽出したテキストを即座に「正しい言葉」に補正
            clean_text = super_fix(text.strip())
            
            # 仕分けロジック（キーワードを広めに設定）
            if "エラー名" in clean_text:
                if current_item["エラー名"]: # 次のデータが来たら保存
                    table_data.append(current_item.copy())
                    current_item = {"エラー名": "", "説明": "", "発生例": ""}
                current_item["エラー名"] = clean_text.replace("エラー名", "").replace(":", "").replace("・", "").strip()
                active_key = "エラー名"
            elif "説明" in clean_text:
                active_key = "説明"
                current_item["説明"] += clean_text.replace("説明", "").replace(":", "").strip()
            elif "発生例" in clean_text or ("例" in clean_text and len(clean_text) < 5):
                active_key = "発生例"
                current_item["発生例"] += clean_text.replace("発生例", "").replace(":", "").strip()
            elif active_key:
                # どの項目にも当てはまらない行は、直前の項目に追記
                current_item[active_key] += " " + clean_text

    # 最後の1件を保存
    if current_item["エラー名"]:
        table_data.append(current_item)

    # --- 画面表示とExcel保存 ---
    if table_data:
        df = pd.DataFrame(table_data)
        st.success(f"解析完了！ {len(df)}件のエラーを整理しました。")
        st.dataframe(df, use_container_width=True)

        # Excelの見た目を整えて保存
        output = io.BytesIO()
        with pd.ExcelWriter(output, engine='openpyxl') as writer:
            df.to_excel(writer, index=False, sheet_name='ErrorList')
            ws = writer.sheets['ErrorList']
            side = Side(style='thin')
            for row in ws.iter_rows(min_row=1, max_row=len(df)+1, max_col=3):
                for cell in row:
                    cell.border = Border(top=side, bottom=side, left=side, right=side)
                    cell.alignment = Alignment(wrap_text=True, vertical='top')
            ws.column_dimensions['A'].width = 25
            ws.column_dimensions['B'].width = 45
            ws.column_dimensions['C'].width = 45

        st.download_button("整理されたExcelを保存", data=output.getvalue(), 
                           file_name="Python_Error_Table.xlsx", 
                           mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
