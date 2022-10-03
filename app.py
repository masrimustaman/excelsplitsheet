# %%
import streamlit as st
import pandas as pd

st.title("Test")

# %%
# variable
input_excel_filename = st.file_uploader("upload a file")
output_excel_filename = "output.xlsx"

# %%
df = pd.read_excel(input_excel_filename)

# %%
# preview dataframe
df.head()

# %%
# list all columns 
col_names = df.columns

# %%
# col to remove 
col_to_remove = []

# %%
df.drop(columns=col_to_remove)

# %%
col_to_split = "KELAS (SESI 2022/2023)2"

# %%
df[col_to_split] = df[col_to_split].str.strip()

# %%
# sheet name 
sheet_name = df[col_to_split].unique()

# %%
with pd.ExcelWriter(output_excel_filename) as writer:
    for sheet in sheet_name:
        tempdf = df[df[col_to_split] == sheet]
        tempdf.to_excel(writer, sheet_name=sheet, index=False)

with open('output.xlsx', 'rb') as f:
    st.download_button('Download Zip', f, file_name='output.xlsx') 