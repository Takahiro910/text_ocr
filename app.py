from azure.core.credentials import AzureKeyCredential
from azure.ai.formrecognizer import DocumentAnalysisClient
from dotenv import load_dotenv
import io, os
import pandas as pd
from PIL import Image
import streamlit as st


# ---Environment---
load_dotenv(verbose=True, dotenv_path='.env')
KEY = os.environ.get("KEY") # Or write your API KEY directly
ENDPOINT = os.environ.get("ENDPOINT")

st.title("手書きの表（画像） → CSV")
st.write("無料につき500枚/月まで。Created by 工藤")

df_l = []
PATH = "."

uploaded_files = st.file_uploader("Upload files", type=["jpg", "png", "tif"], accept_multiple_files=True)
if uploaded_files:
    for file in uploaded_files:
        img_path = os.path.join(PATH, file.name)
        with open(img_path, "wb") as f:
            f.write(file.read())
        img = Image.open(img_path)

        img_byte_arr = io.BytesIO()
        img.save(img_byte_arr, format='JPEG')
        img_byte_arr = img_byte_arr.getvalue()
        
        document_analysis_client = DocumentAnalysisClient(endpoint=ENDPOINT, credential=AzureKeyCredential(KEY))
        poller = document_analysis_client.begin_analyze_document("prebuilt-layout", img_byte_arr)
        response = poller.result()

        table_data = response.tables[0]
        df = pd.DataFrame(columns=list(range(table_data.column_count)), index=list(range(table_data.row_count)))
        for cell in table_data.cells:
            r, c, text = cell.row_index, cell.column_index, cell.content
            df.loc[r, c] = text
        # df = df[1:]

        df_l.append(df)
        os.remove(img_path)

if df_l != []:
    df = pd.concat(df_l)

    st.write("プレビュー")
    col1, col2 = st.columns(2)
    with col1:
        st.header("Uploaded file")
        st.image(uploaded_files[0])
    with col2:
        st.header("Results")
        st.dataframe(df.head(10))

    st.header("Download Here!")
    # csvファイルをダウンロードするためのUIを作成
    if st.download_button(label='Download CSV', data=df.to_csv(header=False, index=False).encode("utf-8"), file_name="data.csv", mime='text/csv'):
        st.write("Thank you!")
        os.remove("data.csv")
        df_l = []
