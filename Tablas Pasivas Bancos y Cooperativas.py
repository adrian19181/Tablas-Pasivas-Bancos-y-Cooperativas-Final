import io
import os
import sys
import pandas as pd
import requests
import streamlit as st

st.set_page_config(layout="wide", initial_sidebar_state="collapsed")

st.markdown(
    """
    <style>
    #MainMenu {visibility: hidden;}
    footer {visibility: hidden;}
    header {visibility: hidden;}
    .block-container {
        padding-top: 0rem;
        padding-bottom: 0rem;
        padding-left: 0.5rem;
        padding-right: 0.5rem;
    }
    </style>
""",
    unsafe_allow_html=True,
)

st.title("Tablas Pasivas Bancos y coopertivas")

if st.button("🔄 Refrescar"):
  st.rerun()

URL_EXCEL_DRIVE = "https://docs.google.com/spreadsheets/d/1kiwAirOeaVEYxOdRP3FWB2xH-IewYOcx/export?format=xlsx"


@st.cache_data(ttl=300)
def cargar_datos_excel(url):
  response = requests.get(url)
  excel_file = io.BytesIO(response.content)
  return pd.read_excel(excel_file, sheet_name=None)


datos_excel = cargar_datos_excel(URL_EXCEL_DRIVE)
lista_instituciones = list(datos_excel.keys())

institucion_seleccionada = st.selectbox(
    "Selecciona o escribe la Institución:",
    options=lista_instituciones,
    index=0 if lista_instituciones else None,
)

producto_seleccionado = st.selectbox(
    "Selecciona el Tipo de Producto:",
    options=["Todos", "Plazos Fijos", "Ahorros Programados", "Cuentas de Ahorros"],
)

if __name__ == "__main__":
  if "streamlit" not in sys.argv[0]:
    os.system(f'streamlit run "{__file__}" --server.port 8501')