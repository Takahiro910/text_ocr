from azure.core.credentials import AzureKeyCredential
from azure.ai.formrecognizer import DocumentAnalysisClient
from dotenv import load_dotenv
import fitz
import io, os
import numpy as np
import pandas as pd
from PIL import Image
import streamlit as st
import xlsxwriter # Excelファイル出力


# ---Environment---
try:
    # Streamlit Cloud
    KEY = st.secrets["KEY"]
    ENDPOINT = st.secrets["ENDPOINT"]
except:
    # ローカル環境
    load_dotenv(verbose=True, dotenv_path='.env')
    KEY = os.environ.get("KEY")
    ENDPOINT = os.environ.get("ENDPOINT")

# ---Function---
def pdf_to_images(pdf_file):
    doc = fitz.open(pdf_file)
    images = []
    for page in doc:
        pix = page.get_pixmap()
        img = Image.frombytes("RGB", [pix.width, pix.height], pix.samples)
        images.append(img)
    return images

# ---Main Page---
st.title("手書きの表（画像） → CSV")
st.write("無料につき500枚/月まで。Created by 工藤")

df_l = []
PATH = "."

st.markdown("## PDF to JPEG Converter")
st.write("PDFファイルはこちらでまず画像に変換してください。")
st.write("PDFをむりやり画像になおすので粗くて読み取り精度が落ちるかもしれません。")
uploaded_pdfs = st.file_uploader("Upload a PDF file", type="pdf", accept_multiple_files=True)

if uploaded_pdfs:
    for pdf in uploaded_pdfs:
        pdfPath = os.path.join(PATH, pdf.name)
        with open(pdfPath, "wb") as f:
            f.write(pdf.read()) # PDFファイルを一時保存

        images = pdf_to_images(pdfPath)
        os.remove(pdfPath)

    for i, image in enumerate(images):
        st.image(image, caption=f"Page {i+1}")

        # ダウンロードボタン
        img_byte_arr = io.BytesIO()
        image.save(img_byte_arr, format='JPEG')
        img_byte_arr = img_byte_arr.getvalue()
        st.download_button(
            label=f"Download Page {i+1} as JPEG",
            data=img_byte_arr,
            file_name=f"page_{i+1}.jpg",
            mime="image/jpeg",
        )


# ---Main Function---
st.markdown("## 画像認識")
st.write("画像の中の表を読み取ってxlsxファイルにして吐き出します。")

uploaded_files = st.file_uploader("Upload files", type=["jpg", "png", "tif"], accept_multiple_files=True)
if uploaded_files:
    for i, file in enumerate(uploaded_files):
        img_path = os.path.join(PATH, file.name)
        with open(img_path, "wb") as f:
            f.write(file.read())
        img = Image.open(img_path)
        img = img.convert('RGB')

        img_byte_arr = io.BytesIO()
        img.save(img_byte_arr, format='JPEG')
        img_byte_arr = img_byte_arr.getvalue()
        
        document_analysis_client = DocumentAnalysisClient(endpoint=ENDPOINT, credential=AzureKeyCredential(KEY))
        poller = document_analysis_client.begin_analyze_document("prebuilt-layout", img_byte_arr)
        response = poller.result()

        if response.tables: # 表が存在する場合
            for j, table_data in enumerate(response.tables): # 複数の表をループ処理
                df = pd.DataFrame(columns=list(range(table_data.column_count)), index=list(range(table_data.row_count)))
                for cell in table_data.cells:
                    r, c, text = cell.row_index, cell.column_index, cell.content
                    df.loc[r, c] = text
                # NaN/INF値を空文字に変換 (例)
                df = df.replace([np.inf, -np.inf], "")
                df = df.fillna("")

                # Excelファイルにシートを追加 (シート名はPDFページ番号と表番号を組み合わせる)
                if not 'workbook' in locals(): # 最初の表の場合、Excelブックを作成
                    workbook = xlsxwriter.Workbook("extracted_tables.xlsx")
                worksheet = workbook.add_worksheet(f"Page_{i+1}_Table_{j+1}") # シート名変更

                # DataFrameをExcelシートに書き込む
                for row_num, row_data in df.iterrows():
                    for col_num, cell_data in enumerate(row_data):
                        worksheet.write(row_num, col_num, cell_data)
        os.remove(img_path)

    if 'workbook' in locals(): # Excelブックが存在する場合
        workbook.close() # Excelファイルを保存
        with open("extracted_tables.xlsx", "rb") as f:
            file_data = f.read()

        st.download_button(
            label="Download Tables as XLSX",
            data=file_data,
            file_name="extracted_tables.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )