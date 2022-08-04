from tkinter.messagebox import NO
import streamlit as st
import plotly_express as px
import pandas as pd

st.set_option('deprecation.showfileUploaderEncoding', False)

st.title("Data visualization App")

#Add the sidebar
st.sidebar.subheader("Visualization Settings")

uploaded_file = st.sidebar.file_uploader(label = "Upload your CSV or Excel file."
                                , type = ['csv' , 'xlsx' , 'xlsm'])

global df 
if uploaded_file is not None: 
    df = pd.read_excel(uploaded_file, sheet_name='hat5',skiprows=3,usecols='B:F')
    df['Precedence'] = df['Precedence'].astype(str)

st.write(df)