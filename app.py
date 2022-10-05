# %%
import streamlit as st
import pandas as pd
from io import BytesIO


# %%
# Functions to output to excel
def to_excel(df):
    output = BytesIO()
    writer = pd.ExcelWriter(output, engine='xlsxwriter')
    for sheet in sheet_name:
        tempdf = df[df[col_to_split] == sheet]
        tempdf.to_excel(writer, sheet_name=sheet[:20], index=False)
    workbook = writer.book
    writer.save()
    processed_data = output.getvalue()
    return processed_data

st.title("Split Excel")


# %%

# streamlit uploader 
st.write("Choose an excel file:")
uploaded_file = st.file_uploader("The excel file must only contains 1 table only")

# Checking if files is uploaded
if uploaded_file is not None:
    df = pd.read_excel(uploaded_file)

    st.info('Your selected excel file upload is successful', icon="ℹ️")
    # Listing all column names for user to select
    column_names = list(df.columns)

    # Getting input and to remove column that user selected
    col_to_remove = st.multiselect(
    'Please select the columns you would like to remove (if any)',column_names)

    df = df.drop(columns=col_to_remove)

    # Getting input on the column that user wanted to split
    col_to_split = st.selectbox(
    'Please select the one column you would like to split',
    column_names[1:], index=1)

    # Data cleaning
    # Eliminate whitespace
    df[col_to_split] = df[col_to_split].str.strip()

    #listing the sheet name generated from the unique value of user selected columns
    sheet_name = df[col_to_split].unique()

    #Preview of the 1st sheet
    st.write("Preview of the 1st sheet:")
    first_df = df[df[col_to_split] == sheet_name[0]] #specify the 1st sheet name here
    st.dataframe(first_df)

    # Calling the function to generate excel file
    df_xlsx = to_excel(df)

    # Finalizing the output file name (append _split.xlsx at the back)
    output_excel_filename = uploaded_file.name.replace(".xlsx", "_split.xlsx")
    st.info('Output will be downloaded with the name ' + output_excel_filename , icon="ℹ️")
    st.download_button(label='📥 Download Current Result',
                                    data=df_xlsx ,
                                    file_name= output_excel_filename)
else:
    st.warning("You need to upload an excel file.")