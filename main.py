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

# --- OCRの読み間違いを補正する最小限の辞書 ---
def light_fix(text):
    d = {
        "エラ一名": "エラー名", "エフー名": "エラー名", "工ラー": "エラー名", "エフ": "エラー名",
        "說明": "説明", "訳明": "説明", "光二例": "発生例", "え三例": "発生例", "発上例": "発生例",
        "え二例": "発生例", "光上例": "発生例", "拓孤": "括弧", "志れ": "忘れ"
    }
    for k, v in d.items():
        text = text.replace(k, v)
    return text

uploaded_file = st.file_uploader("PDFを選択してください", type="pdf")

if uploaded_file is not None:
    st.info("解析中... サーバーのメモリを節約するため、1ページずつゆっくり処理しています。")
    
    # OCRエンジンの起動
    reader = easyocr.Reader(['ja', 'en'], gpu=False)
    
    pdf_bytes = uploaded_file.getvalue()
    # DPIを120まで下げて、メモリパンクを徹底回避
    images = convert_from_bytes(pdf_bytes, dpi=120)
    
    table_data = []
    current_item = {"エラー名": "", "説明": "", "発生例": ""}
    active_key = None

    for img in images:
        img_array = np.array(img)
        # detail=0 で座標情報を捨ててメモリを節約
        results = reader.readtext(img_array, detail=0)
        
        for text in results:
            clean_text = light_fix(text.strip())
            
            # 振り分けロジック
            if "エラー名" in clean_text:
                if current_item["エラー名"]:
                    table_data.append(current_item.copy())
                current_item = {"エラー名": clean_text.replace("エラー名","").replace(":","").strip(), "説明": "", "発生例": ""}
                active_key = "エラー名"
            elif "説明" in clean_text:
                active_key = "説明"
                current_item["説明"] += clean_text.replace("説明","").replace(":","").strip()
            elif any(k in clean_text for k in ["発生例", "生例", "光二例", "え三例"]):
                active_key = "発生例"
                current_item["発生例"] += clean_text.replace("発生例","").replace(":","").strip()
            elif active_key:
                current_item[active_key] += " " + clean_text
        
        # 1ページごとにメモリを強制解放
        del img_array
        gc.collect()

    if current_item["エラー名"]:
        table_data.append(current_item)

    if table_data:
        df = pd.DataFrame(table_data)
        st.success("解析成功！")
        st.dataframe(df)

        output = io.BytesIO()
        with pd.ExcelWriter(output, engine='openpyxl') as writer:
            df.to_excel(writer, index=False)
            ws = writer.sheets['Sheet1']
            side = Side(style='thin')
            border = Border(top=side, bottom=side, left=side, right=side)
            for row in ws.iter_rows(min_row=1, max_row=len(df)+1, max_col=3):
                for cell in row:
                    cell.border = border
                    cell.alignment = Alignment(wrap_text=True, vertical='top')
            ws.column_dimensions['A'].width = 25
            ws.column_dimensions['B'].width = 45
            ws.column_dimensions['C'].width = 45

        st.download_button("Excelを保存", data=output.getvalue(), file_name="error_list.xlsx")
