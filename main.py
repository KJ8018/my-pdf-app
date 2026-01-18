import streamlit as st
import easyocr
import pandas as pd
from pdf2image import convert_from_bytes
import io
import numpy as np
from PIL import Image

st.set_page_config(page_title="AI PDF", layout="wide")
st.title("AI PDF 文字起こしアプリ")

# 翻訳エラー対策：ブラウザ翻訳がオンだと落ちるため警告
st.warning("⚠️ブラウザの『自動翻訳機能』は必ずオフにして使用してください。")

uploaded_file = st.file_uploader("PDFを選択してください", type="pdf")

if uploaded_file is not None:
    st.info("解析中... (軽量モードで実行中)")
    
    # OCRエンジンの読み込み（ここでメモリを節約）
    reader = easyocr.Reader(['ja', 'en'], gpu=False)
    
    pdf_bytes = uploaded_file.getvalue()
    
    # DPIを150に下げて、サーバーが落ちないようにする
    images = convert_from_bytes(pdf_bytes, dpi=150)
    
    all_results = []
    for img in images:
        img_array = np.array(img)
        # 詳細設定をオフにして高速化
        results = reader.readtext(img_array, detail=0)
        all_results.extend(results)
    
    if all_results:
        st.success("解析完了！")
        # 取得したテキストをシンプルに表示
        df = pd.DataFrame(all_results, columns=["抽出テキスト"])
        st.dataframe(df)
        
        # CSVとしてダウンロード（Excelより軽い）
        csv = df.to_csv(index=False).encode('utf_8_sig')
        st.download_button("結果を保存 (CSV)", data=csv, file_name="result.csv", mime='text/csv')
    else:
        st.error("テキストが検出できませんでした。")
