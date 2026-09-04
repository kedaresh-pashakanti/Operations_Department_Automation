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
# #CROSSCHECK_FILE = os.path.join(BASE_DIR, "kotak_added_in_mid.py")

# CROSSCHECK_FILE = os.path.join(BASE_DIR, "kotak_added_in_mid_added_workbook.py")

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
import runpy
import os
import sys


# =========================================================
# BASE DIRECTORY
# =========================================================

BASE_DIR = os.path.dirname(
    os.path.abspath(__file__)
)


# =========================================================
# APPLICATION FILES
# =========================================================

APP_FILE = os.path.join(
    BASE_DIR,
    "1.py"
)

HDFC_FILE = os.path.join(
    BASE_DIR,
    "New_HDFC.py"
)

CROSSCHECK_FILE = os.path.join(
    BASE_DIR,
    "kotak_added_in_mid_added_workbook.py"
)


# =========================================================
# MAKE PROJECT DIRECTORY AVAILABLE FOR IMPORTS
# =========================================================
#
# This is important because:
#
# kotak_added_in_mid_added_workbook.py
# imports:
#
#     from New_HDFC import ...
#
# =========================================================

if BASE_DIR not in sys.path:
    sys.path.insert(0, BASE_DIR)


# =========================================================
# MAIN APP CONFIG
# =========================================================

st.set_page_config(
    page_title="Ops Automation",
    layout="wide"
)

st.title("Ops Automation")


# =========================================================
# PROCESS SELECTOR
# =========================================================

option = st.selectbox(
    "Select your process",
    [
        "Statement Processor",
        "HDFC ESCROW MID MAPPING",
        "SP Cross Check",
    ]
)


# =========================================================
# RUN CHILD STREAMLIT SCRIPT
# =========================================================

def run_child_script(file_path):
    """
    Execute a child Streamlit application.

    The child application's st.set_page_config()
    is temporarily disabled because the parent app
    has already configured the Streamlit page.
    """

    # -----------------------------------------------------
    # Validate file exists
    # -----------------------------------------------------

    if not os.path.exists(file_path):
        st.error(
            f"Application file not found:\n{file_path}"
        )
        return

    # -----------------------------------------------------
    # Keep project directory available for imports
    # -----------------------------------------------------

    if BASE_DIR not in sys.path:
        sys.path.insert(0, BASE_DIR)

    # -----------------------------------------------------
    # Temporarily disable child page config
    # -----------------------------------------------------

    original_set_page_config = st.set_page_config

    try:

        st.set_page_config = (
            lambda *args, **kwargs: None
        )

        runpy.run_path(
            file_path,
            run_name="__main__"
        )

    except Exception as e:

        st.error(
            f"Failed to run application:\n{e}"
        )

    finally:

        st.set_page_config = (
            original_set_page_config
        )


# =========================================================
# PROCESS ROUTING
# =========================================================

if option == "Statement Processor":

    run_child_script(
        APP_FILE
    )


elif option == "HDFC ESCROW MID MAPPING":

    run_child_script(
        HDFC_FILE
    )


elif option == "SP Cross Check":

    run_child_script(
        CROSSCHECK_FILE
    )




