import streamlit as st
import runpy
import os

BASE_DIR = os.path.dirname(os.path.abspath(__file__))

APP_FILE = os.path.join(BASE_DIR, "1.py")
#HDFC_FILE = os.path.join(BASE_DIR, "final2.py")
HDFC_FILE = os.path.join(BASE_DIR, "hdfc_escrow_mid_mapping_processor_irctc_pa_pg.py")
#HDFC_FILE = os.path.join(BASE_DIR, "New_HDFC.py")

CROSSCHECK_FILE = os.path.join(BASE_DIR, "test_hdfc_upi_fixed.py")

st.set_page_config(page_title="Ops Automation", layout="wide")

st.title("Ops Automation")

# SIMPLE SELECTOR ONLY (NO EXTRA UI)
option = st.selectbox(
    "Select your process",
    [
        "Statement Processor",
        "HDFC ESCROW MID MAPPING",
        "SP Cross Check"
    ]
)

# ================================
# RUN ORIGINAL FILES DIRECTLY
# ================================
if option == "Statement Processor":
    runpy.run_path(APP_FILE, run_name="__main__")

elif option == "HDFC ESCROW MID MAPPING":
    runpy.run_path(HDFC_FILE, run_name="__main__")

elif option == "SP Cross Check":
    runpy.run_path(CROSSCHECK_FILE, run_name="__main__")
