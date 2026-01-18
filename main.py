import streamlit as st
import easyocr
import pandas as pd
from pdf2image import convert_from_bytes
import io
import numpy as np
from openpyxl.styles import Border, Side, Alignment

st.set_page_config(page_title="AI PDF", layout="wide")
st.title("AI PDF 文字起こしアプリ")

st.warning("⚠️ブラウザの自動翻訳を『オフ』にしてからPDFを入れてください。")

uploaded_file = st.file_uploader("PDFを選択してください", type="pdf")

if uploaded_file is not None:
    st.info("解析を開始しました。約1分ほどお待ちください...")
    
    # OCRエンジンの設定（メモリ節約モード）
    reader = easyocr.Reader(['ja', 'en'], gpu=False)
    
    # PDF読み込み
    pdf_bytes = uploaded_file.getvalue()
    # DPIを130まで下げて、メモリパンクを徹底回避
    images = convert_from_bytes(pdf_bytes, dpi=130)
    
    raw_texts = []
    for img in images:
        img_array = np.array(img)
        # detail=0 にして余計な情報を取得しない
        results = reader.readtext(img_array, detail=0)
        raw_texts.extend(results)

    # --- 整理ロジック（ここが肝心です） ---
    table_data = []
    current_item = {"エラー名": "", "説明": "", "発生例": ""}
    active_key = None

    for text in raw_texts:
        t = text.strip()
        # キーワードが見つかったら切り替える（OCRの誤字も考慮）
        if "エラー名" in t or "エラ一名" in t:
            if current_item["エラー名"]:
                table_data.append(current_item.copy())
            current_item = {"エラー名": t.replace("エラー名","").replace("・","").replace(":","").strip(), "説明": "", "発生例": ""}
            active_key = "エラー名"
        elif "説明" in t or "說明" in t:
            active_key = "説明"
            current_item["説明"] += t.replace("説明","").replace(":","").strip()
        elif "発生例" in t or "生例" in t:
            active_key = "発生例"
            current_item["発生例"] += t.replace("発生例","").replace(":","").strip()
        elif active_key:
            current_item[active_key] += " " + t

    if current_item["エラー名"]:
        table_data.append(current_item)

    # --- 画面表示とExcel保存 ---
    if table_data:
        df = pd.DataFrame(table_data)
        st.success(f"解析成功！ {len(df)}件のエラーを抽出しました。")
        st.dataframe(df)

        # Excelの作成
        output = io.BytesIO()
        with pd.ExcelWriter(output, engine='openpyxl') as writer:
            df.to_excel(writer, index=False, sheet_name='ErrorList')
            ws = writer.sheets['ErrorList']
            # 枠線と自動折り返し
            side = Side(style='thin')
            for row in ws.iter_rows(min_row=1, max_row=len(df)+1, max_col=3):
                for cell in row:
                    cell.border = Border(top=side, bottom=side, left=side, right=side)
                    cell.alignment = Alignment(wrap_text=True, vertical='top')
            ws.column_dimensions['A'].width = 20
            ws.column_dimensions['B'].width = 40
            ws.column_dimensions['C'].width = 40

        st.download_button("Excelをダウンロード", data=output.getvalue(), 
                           file_name="AI_OCR_Result.xlsx", 
                           mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
