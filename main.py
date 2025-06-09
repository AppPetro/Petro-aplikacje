# main.py
import streamlit as st

# importy Twoich trzech aplikacji:
from excel_to_epp.excel_to_epp_app import run as run_e2epp
from excel_vs_wz.excel_vs_wz_app import run as run_evsz
from pdf_to_excel.pdf_to_excel_app import run as run_p2x

st.set_page_config(page_title="Petro – 3 Aplikacje", layout="wide")

st.sidebar.title("Wybierz aplikację")
choice = st.sidebar.radio("", [
    "Excel → EPP",
    "Excel vs WZ",
    "PDF → Excel",
])

if choice == "Excel → EPP":
    run_e2epp()
elif choice == "Excel vs WZ":
    run_evsz()
else:
    run_p2x()
