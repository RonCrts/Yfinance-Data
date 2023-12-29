import yfinance as yf 
import streamlit as st
import datetime
import pandas as pd
from openpyxl.styles import numbers

tickers = yf.Tickers(['AAPL', 'MSFT', 'GOOGL', 'AMZN', 'FB', 'TSLA', 'NVDA', 'JPM', 'V', 'JNJ', 'WMT', 'UNH', 'BAC', 'MA', 'PYPL', 'HD', 'DIS', 'PG', 'VZ', 'CMCSA', 'ADBE', 'NFLX', 'INTC', 'KO', 'PFE', 'T', 'PEP', 'MRK', 'ABT', 'NKE', 'CRM', 'ABBV', 'CSCO', 'XOM', 'CVX', 'MCD', 'TMO', 'ACN', 'IBM', 'QCOM', 'MDT', 'HON', 'TXN', 'AMGN', 'ORCL', 'COST', 'AVGO', 'NEE', 'UNP', 'LIN'])

# Crear un buscador para que el usuario elija un símbolo
selected_symbol = st.selectbox('Seleccione un símbolo', tickers.tickers)

# Obtener los datos del símbolo seleccionado
tickerData = yf.Ticker(selected_symbol)

# Obtener las fechas de inicio y fin
start_date = st.date_input('Fecha de inicio', datetime.date(2010, 5, 31))
end_date = st.date_input('Fecha de fin', datetime.date(2020, 5, 31))

# Convertir las fechas a formato de cadena
start_date_str = start_date.strftime('%Y-%m-%d')
end_date_str = end_date.strftime('%Y-%m-%d')

# Obtener los datos históricos sin dividens y stock splits
tickerDf = tickerData.history(period='1d', start=start_date_str, end=end_date_str, actions=False)

# Crear un archivo de Excel y guardar los datos en una tabla
excel_file = 'data.xlsx'
with pd.ExcelWriter(excel_file, engine='openpyxl') as writer:
    tickerDf.to_excel(writer, index=False, sheet_name='Data')
    workbook = writer.book
    worksheet = writer.sheets['Data']
    
    # Formatear las celdas de precios
    number_format = numbers.FORMAT_NUMBER_COMMA_SEPARATED1
    worksheet.column_dimensions['E'].number_format = number_format
    worksheet.column_dimensions['F'].number_format = number_format
    
    # Ajustar el ancho de las columnas
    worksheet.column_dimensions['A'].width = 12
    worksheet.column_dimensions['B'].width = 12
    worksheet.column_dimensions['C'].width = 12
    worksheet.column_dimensions['D'].width = 12
    worksheet.column_dimensions['E'].width = 15
    worksheet.column_dimensions['F'].width = 15

# Agregar un botón de descarga
st.download_button(
    label="Descargar archivo",
    data=open(excel_file, 'rb').read(),
    file_name=excel_file,
    mime='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'
)
