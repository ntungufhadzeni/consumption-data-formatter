import pandas as pd
import streamlit as st
import io

def format_excel(data):
    df = pd.read_excel(data)
    df['TRANS Date'] = df['TRANS Date'].astype(str)
    df['TRANS Time'] = df['TRANS Time'].astype(str)
    df['TRANS DT'] = df['TRANS Date'] + df['TRANS Time']
    df['TRANS DT'] = pd.to_datetime(df['TRANS DT'], format='%Y%m%d%H%M%S')
    df['Card No.'] = df['Card No.'].astype(str)
    
    return df


st.set_page_config(page_title="Formatter", page_icon=":robot_face:")
st.markdown("<h1 style='text-align: center;'>Formatter</h1>", unsafe_allow_html=True)

st.write("Please upload your Excel file below.")
data = st.file_uploader("Upload a Excel")
df = pd.DataFrame()

if data:
    generate = st.button(label='Generate')
    if generate:
        try:
            df = format_excel(data)
            st.write("Excel formatted successfully. Click 'Download'")
        except:
            st.write('Error formatting excel. Make sure there is column row.')
            
    if not df.empty:
        def create_xlsx(data_pd):
            buffer = io.BytesIO()
            with pd.ExcelWriter(buffer) as writer:
                data_pd.to_excel(writer)
            return buffer

        excel_file_name = data.name

        if st.download_button(
                label="Download Excel",
                data=create_xlsx(df),
                file_name=excel_file_name,
                mime='application/vnd.ms-excel'
            ):
            st.write("thank you for downloading!")
