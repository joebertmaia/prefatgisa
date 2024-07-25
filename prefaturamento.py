import io
import streamlit as st
import pandas as pd
import datetime
import numpy as np
import matplotlib.pyplot as plt
from io import StringIO

# Import statistics Library
#import statistics
from openpyxl import reader,load_workbook,Workbook
#import xlsxwriter
#from openpyxl.utils.dataframe import dataframe_to_rows

PAGE_TITLE = "App de Pré-Faturamento"
PAGE_ICON = "https://cse.energisa.com.br/img/grupo_energisa.png"

st.set_page_config(page_title=PAGE_TITLE, page_icon=PAGE_ICON)

st.title('Pré-faturamento')
st.write("Esse webapp auxilia no envio das listas de unidades para manutenção do grupo A e grupo B.")

uploaded_file = st.file_uploader("Escolha uma planilha", type = 'xlsx')

wb = Workbook()

if uploaded_file is not None:
    wb = load_workbook(uploaded_file, read_only=False)
    st.write(f"Nome do arquivo: {uploaded_file.name}")
    st.write(f"Nome das planilhas: {wb.sheetnames}")

    prefat = pd.read_excel(uploaded_file, engine='openpyxl',skiprows=2)
    st.write(prefat.head())

    filtro_desconectada = prefat[prefat['SITUAÇÃO_COMUNICAÇÃO'] == 'DESCONECTADA']
    st.write('UCs desconectadas')
    st.write(filtro_desconectada)

buffer = io.BytesIO()
wb.save(buffer)

st.download_button(
label="Download da planilha do Excel",
data=buffer,
file_name="fk.xlsx",
)
