from datetime import date

import pandas as pd
import streamlit as st
import sys
from io import BytesIO

import Functions.Google_Builder as GB

st.set_page_config(page_title='SF-DJ Campaign Builder Tool')
st.header('Welcome To Campaign Builder!')
st.write('V1 - 11.11.2022')
st.write('Please follow below instructions to create your outputs')

Builder_type = st.selectbox(label='Please Choose which builder to work with:',
                            options=['Google Builder', 'Bing Builder', 'Desjardins Builder', 'Google Spanish Builder'])
st.write('You are working on: ', Builder_type)

## Create Base Variables
today_date = date.today().strftime("%m.%d.%y")

### -- Upload File to App ---
Data_df = st.file_uploader(f'Please upload {Builder_type} Data file')
Ref_df = st.file_uploader(f'Please upload {Builder_type} Reference file')

if ((Data_df is None) or (Ref_df is None)):
    st.write('Please upload Data and Ref Files')
    st.stop()

df = pd.read_excel(Data_df)
st.write(df)

### --- Downloading File ---
# Convert DF to a Streamlit excel downloadable version
def to_excel(df, Sheet_Name):
    output = BytesIO()
    writer = pd.ExcelWriter(output)  #, engine='xlsxwriter')
    df.to_excel(writer, index=False, sheet_name=Sheet_Name)
    workbook = writer.book
    worksheet = writer.sheets[Sheet_Name]
#    format1 = workbook.add_format({'num_format': '0.00'}) 
#    worksheet.set_column('A:A', None, format1)  
    writer.save()
    processed_data = output.getvalue()
    return processed_data

def st_download_button(df_xlsx, name, key_name):
    st.download_button(
        label = f'''📥 Press to Download {name.split(' ')[2]} File''',
        data = df_xlsx,
        file_name = name,
        key = key_name
    )

# Create and download files as per the input
if Builder_type == 'Google Builder':
    Google_data_set = GB.main(df, Ref_df)
    df_xlsx = to_excel(Google_data_set[0], 'Structured Snippet Upload')
    st_download_button(df_xlsx, f'New Agent Snippet Upload - {today_date}.xlsx', 'Snip')

    df_xlsx = to_excel(Google_data_set[1], 'Sitelink Upload')
    st_download_button(df_xlsx, f'New Agent Sitelink Upload - {today_date}.xlsx', 'Site')

    df_xlsx = to_excel(Google_data_set[2], 'Radius Location Upload')
    st_download_button(df_xlsx, f'New Agent Radius-Target Upload - {today_date}.xlsx', 'Radius')

    df_xlsx = to_excel(Google_data_set[3], 'Call Upload')
    st_download_button(df_xlsx, f'New Agent Call Upload - {today_date}.xlsx', 'Call')

    st.write('''Wait for a bit more. 
    If it says "Running" on Top-Right corner, it is coming''')
    df_xlsx = to_excel(Google_data_set[4], 'Bulk Upload - {today_date}')
    st_download_button(df_xlsx, f'New Agent Bulk Upload - {today_date}.xlsx', 'Bulk')

    st.stop()

elif Builder_type == 'Bing Builder':
    Bing = GB.main(df, Ref_df)

elif Builder_type == 'Desjardins Builder':
    Desjardins = GB.main(df, Ref_df)