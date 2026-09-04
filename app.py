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







import ast
import os
import sys
import traceback

import streamlit as st


# ============================================================
# BASE DIRECTORY
# ============================================================

BASE_DIR = os.path.dirname(os.path.abspath(__file__))

if BASE_DIR not in sys.path:
    sys.path.insert(0, BASE_DIR)


# ============================================================
# CHILD APPLICATION FILES
# ============================================================

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


# ============================================================
# MAIN APP CONFIG
# ============================================================

st.set_page_config(
    page_title="Ops Automation",
    layout="wide"
)


# ============================================================
# HEADER
# ============================================================

st.title("Ops Automation")


# ============================================================
# SELECT PROCESS
# ============================================================

option = st.selectbox(
    "Select your process",
    [
        "Statement Processor",
        "HDFC ESCROW MID MAPPING",
        "SP Cross Check"
    ]
)


# ============================================================
# AST TRANSFORMER
# ============================================================

class RemoveStreamlitPageConfig(ast.NodeTransformer):

    def visit_Call(self, node):
        """
        Remove every form of:

            st.set_page_config(...)

        from the child script.
        """

        self.generic_visit(node)

        # st.set_page_config(...)
        if (
            isinstance(node.func, ast.Attribute)
            and node.func.attr == "set_page_config"
            and isinstance(node.func.value, ast.Name)
            and node.func.value.id == "st"
        ):
            return None

        return node


# ============================================================
# RUN CHILD SCRIPT
# ============================================================

def run_child_script(file_path):

    # --------------------------------------------------------
    # File existence
    # --------------------------------------------------------

    if not os.path.exists(file_path):

        st.error(
            f"File not found:\n{file_path}"
        )

        return

    try:

        # ----------------------------------------------------
        # Read child source
        # ----------------------------------------------------

        with open(
            file_path,
            "r",
            encoding="utf-8"
        ) as f:

            source = f.read()


        # ----------------------------------------------------
        # Parse Python
        # ----------------------------------------------------

        tree = ast.parse(
            source,
            filename=file_path
        )


        # ----------------------------------------------------
        # Remove ALL st.set_page_config() calls
        # ----------------------------------------------------

        tree = RemoveStreamlitPageConfig().visit(tree)

        ast.fix_missing_locations(tree)


        # ----------------------------------------------------
        # Execution namespace
        # ----------------------------------------------------

        child_globals = {
            "__name__": "__main__",
            "__file__": file_path,
            "__package__": None,
            "__cached__": None,
        }


        # ----------------------------------------------------
        # Execute child program
        # ----------------------------------------------------

        compiled = compile(
            tree,
            file_path,
            "exec"
        )

        exec(
            compiled,
            child_globals,
            child_globals
        )


    except Exception as e:

        st.error(
            f"Error running {os.path.basename(file_path)}"
        )

        st.exception(e)

        with st.expander("Technical traceback"):

            st.code(
                traceback.format_exc(),
                language="text"
            )


# ============================================================
# PROCESS ROUTING
# ============================================================

if option == "Statement Processor":

    run_child_script(APP_FILE)


elif option == "HDFC ESCROW MID MAPPING":

    run_child_script(HDFC_FILE)


elif option == "SP Cross Check":

    run_child_script(CROSSCHECK_FILE)








