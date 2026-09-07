# import streamlit as st
# import runpy
# import os

# BASE_DIR = os.path.dirname(os.path.abspath(__file__))

# APP_FILE = os.path.join(BASE_DIR, "1.py")
# #HDFC_FILE = os.path.join(BASE_DIR, "final2.py")
# #HDFC_FILE = os.path.join(BASE_DIR, "hdfc_escrow_mid_mapping_processor_irctc_pa_pg.py")


# #main_file
# HDFC_FILE = os.path.join(BASE_DIR, "New_HDFC.py")

# CROSSCHECK_FILE = os.path.join(BASE_DIR, "test_hdfc_upi_fixed.py")

# st.set_page_config(page_title="Ops Automation", layout="wide")

# st.title("Ops Automation")

# # SIMPLE SELECTOR ONLY (NO EXTRA UI)
# option = st.selectbox(
#     "Select your process",
#     [
#         "Statement Processor",
#         "HDFC ESCROW MID MAPPING",
#         "SP Cross Check"
#     ]
# )

# # ================================
# # RUN ORIGINAL FILES DIRECTLY
# # ================================
# if option == "Statement Processor":
#     runpy.run_path(APP_FILE, run_name="__main__")

# elif option == "HDFC ESCROW MID MAPPING":
#     runpy.run_path(HDFC_FILE, run_name="__main__")

# elif option == "SP Cross Check":
#     runpy.run_path(CROSSCHECK_FILE, run_name="__main__")













# import streamlit as st
# import runpy
# import os

# BASE_DIR = os.path.dirname(os.path.abspath(__file__))

# APP_FILE = os.path.join(BASE_DIR, "1.py")
# HDFC_FILE = os.path.join(BASE_DIR, "New_HDFC.py")
# CROSSCHECK_FILE = os.path.join(BASE_DIR, "kotak_added_in_mid.py")

# #CROSSCHECK_FILE = os.path.join(BASE_DIR, "kotak_added_in_mid_added_workbook.py")

# # Main app config
# st.set_page_config(page_title="Ops Automation", layout="wide")
# st.title("Ops Automation")

# # Simple selector only (no extra UI)
# option = st.selectbox(
#     "Select your process",
#     [
#         "Statement Processor",
#         "HDFC ESCROW MID MAPPING",
#         "SP Cross Check"
#     ]
# )

# def run_child_script(file_path):
#     """
#     Runs a child Streamlit script safely.
#     Temporarily disables st.set_page_config inside the child file
#     so duplicate page_config errors do not happen.
#     """
#     original_set_page_config = st.set_page_config

#     try:
#         # Prevent child file from calling set_page_config again
#         st.set_page_config = lambda *args, **kwargs: None
#         runpy.run_path(file_path, run_name="__main__")
#     finally:
#         # Restore original function
#         st.set_page_config = original_set_page_config

# # Run original files directly
# if option == "Statement Processor":
#     run_child_script(APP_FILE)

# elif option == "HDFC ESCROW MID MAPPING":
#     run_child_script(HDFC_FILE)

# elif option == "SP Cross Check":
#     run_child_script(CROSSCHECK_FILE)












import streamlit as st


# ============================================================
# PAGE CONFIGURATION
# ============================================================

st.set_page_config(
    page_title="Ops Automation",
    page_icon="⚙️",
    layout="wide",
    initial_sidebar_state="expanded",
)


# ============================================================
# IMPORT THE UI FUNCTIONS FROM EACH APPLICATION
# ============================================================
#
# IMPORTANT:
# Each of these files must contain a function with the
# corresponding name:
#
# 1.py                  -> statement_processor_app()
# New_HDFC.py           -> hdfc_escrow_app()
# kotak_added_in_mid.py -> sp_crosscheck_app()
#
# See the example structure below.
#
# ============================================================

from New_HDFC import hdfc_escrow_app
from kotak_added_in_mid import sp_crosscheck_app

# Python cannot normally import "1.py" with:
#
#     from 1 import ...
#
# because the filename starts with a number.
#
# Therefore, either:
#
# A) Rename 1.py to something like:
#       statement_processor.py
#
# OR
#
# B) use importlib as shown below.
#
# The recommended approach is A.
#
# Assuming you rename:
#
#       1.py
#       ↓
#       statement_processor.py
#
#from statement_processor import statement_processor_app


# ============================================================
# HEADER
# ============================================================

st.title("Ops Automation")
st.caption("Select an operation from the menu below.")


# ============================================================
# PROCESS SELECTOR
# ============================================================

option = st.selectbox(
    "Select your process",
    [
        "HDFC ESCROW MID MAPPING",
        "SP Cross Check",
    ],
    index=0,
)


# ============================================================
# RUN SELECTED APPLICATION
# ============================================================


if option == "HDFC ESCROW MID MAPPING":

    hdfc_escrow_app()


elif option == "SP Cross Check":

    sp_crosscheck_app()






