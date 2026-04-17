
import os
import sys
import io
import re
import zipfile
import tempfile
import importlib.util
from datetime import date
from contextlib import contextmanager

import pandas as pd
import streamlit as st


# =============================================================================
# Paths
# =============================================================================
BASE_DIR = os.path.dirname(os.path.abspath(__file__))

APP_FILE = os.path.join(BASE_DIR, "1.py")
HDFC_FILE = os.path.join(BASE_DIR, "final2.py")
CROSSCHECK_FILE = os.path.join(BASE_DIR, "test_hdfc_upi_fixed.py")


# =============================================================================
# Dummy Streamlit used only while importing the old scripts
# =============================================================================
class _DummySpinner:
    def __enter__(self):
        return self

    def __exit__(self, exc_type, exc, tb):
        return False


class _DummySessionState(dict):
    def __getattr__(self, name):
        return self.get(name)

    def __setattr__(self, name, value):
        self[name] = value


class _DummyStreamlit:
    def __init__(self):
        self.sidebar = self
        self.session_state = _DummySessionState()

    def __getattr__(self, name):
        if name in {"spinner"}:
            return lambda *args, **kwargs: _DummySpinner()
        if name in {"columns", "tabs"}:
            def _builder(*args, **kwargs):
                count = args[0] if args else 0
                return [self for _ in range(count)]
            return _builder
        if name in {"radio", "selectbox"}:
            def _pick(*args, **kwargs):
                if "options" in kwargs and kwargs["options"]:
                    return kwargs["options"][0]
                if len(args) >= 2 and args[1]:
                    return args[1][0]
                return None
            return _pick
        if name in {"file_uploader"}:
            return lambda *args, **kwargs: None
        if name in {"button"}:
            return lambda *args, **kwargs: False
        if name in {"date_input"}:
            return lambda *args, **kwargs: date.today()
        if name in {"number_input"}:
            return lambda *args, **kwargs: 0
        if name in {"text_input", "text_area"}:
            return lambda *args, **kwargs: ""
        if name in {"checkbox"}:
            return lambda *args, **kwargs: False
        if name in {"download_button", "dataframe", "table", "write", "markdown", "caption", "title", "header", "subheader", "success", "warning", "error", "info", "code", "metric", "divider", "set_page_config"}:
            return lambda *args, **kwargs: None
        return lambda *args, **kwargs: None


def _load_module_from_path(module_name: str, file_path: str):
    """
    Import an old Streamlit script safely by temporarily swapping streamlit
    with a dummy module so its top-level UI code does not run.
    """
    if not os.path.exists(file_path):
        return None

    original_streamlit = sys.modules.get("streamlit")
    dummy = _DummyStreamlit()
    sys.modules["streamlit"] = dummy  # type: ignore

    try:
        spec = importlib.util.spec_from_file_location(module_name, file_path)
        if spec is None or spec.loader is None:
            raise ImportError(f"Unable to load module from {file_path}")
        module = importlib.util.module_from_spec(spec)
        sys.modules[module_name] = module
        spec.loader.exec_module(module)
        return module
    finally:
        if original_streamlit is not None:
            sys.modules["streamlit"] = original_streamlit
        else:
            sys.modules.pop("streamlit", None)


# Load the three systems from your folder.
app_mod = _load_module_from_path("app_system", APP_FILE)
hdfc_mod = _load_module_from_path("hdfc_system", HDFC_FILE)
cross_mod = _load_module_from_path("crosscheck_system", CROSSCHECK_FILE)


# =============================================================================
# Shared UI helpers
# =============================================================================
def _temp_save_uploaded(uploaded_file):
    """
    Save a Streamlit uploaded file to a temporary path.
    Returns the temp file path.
    """
    suffix = os.path.splitext(uploaded_file.name)[1] or ".tmp"
    tmp = tempfile.NamedTemporaryFile(delete=False, suffix=suffix)
    try:
        tmp.write(uploaded_file.getbuffer())
        tmp.flush()
        return tmp.name
    finally:
        tmp.close()


def _cleanup_paths(paths):
    for p in paths:
        if p and os.path.exists(p):
            try:
                os.remove(p)
            except Exception:
                pass


def _download_excel_bytes(file_name, bytes_data):
    st.download_button(
        label=f"Download {file_name}",
        data=bytes_data,
        file_name=file_name,
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
    )


# =============================================================================
# Page 1 — Statement Processor
# =============================================================================
def page_statement_processor():
    st.header("1) Statement Processor")
    st.caption("Process raw files and append them into a target workbook.")

    if app_mod is None:
        st.error("app.py not found in the same folder as this launcher.")
        return

    raw_files = st.file_uploader(
        "Upload raw statement file(s)",
        type=["csv", "xlsx", "xlsm"],
        accept_multiple_files=True,
        key="raw_files",
    )

    target_file = st.file_uploader(
        "Upload target Excel file",
        type=["xlsx", "xlsm"],
        key="target_file",
    )

    sheet_name = None
    if target_file is not None:
        try:
            target_bytes = target_file.getvalue()
            wb = app_mod.load_workbook(io.BytesIO(target_bytes))
            sheet_name = st.selectbox("Select target sheet", wb.sheetnames, index=0)
        except Exception as e:
            st.error(f"Unable to read workbook sheets: {e}")

    if st.button("Process Statement", key="process_statement"):
        if not raw_files:
            st.error("Please upload at least one raw file.")
            return
        if target_file is None:
            st.error("Please upload the target Excel file.")
            return

        temp_paths = []
        try:
            processed_parts = []
            with st.spinner("Processing raw file(s)..."):
                for file in raw_files:
                    df = app_mod.read_raw_statement(file)
                    std_df = app_mod.standardize_statement_df(df)
                    processed_parts.append(std_df)

            combined_df = pd.concat(processed_parts, ignore_index=True) if processed_parts else pd.DataFrame()

            if combined_df.empty:
                st.warning("No rows found after processing raw files.")
                return

            st.success(f"Processed successfully. Total rows: {len(combined_df)}")
            st.dataframe(combined_df)

            preview_buf = io.BytesIO()
            with pd.ExcelWriter(preview_buf, engine="openpyxl") as writer:
                combined_df.to_excel(writer, index=False, sheet_name="Processed_Data")
            preview_buf.seek(0)
            _download_excel_bytes("processed_statement.xlsx", preview_buf.getvalue())

            target_bytes = target_file.getvalue()
            updated_bytes = app_mod.append_to_workbook(
                target_bytes=target_bytes,
                append_df=combined_df,
                sheet_name=sheet_name,
            )
            _download_excel_bytes(f"updated_{target_file.name}", updated_bytes)

        except Exception as e:
            st.error(f"Processing failed: {e}")


# =============================================================================
# Page 2 — HDFC ESCROW MID MAPPING
# =============================================================================
def page_hdfc_escrow():
    st.header("2) HDFC ESCROW MID MAPPING")
    st.caption("Upload input file + template file. Optional HDFC UPI refund file can also be uploaded.")

    if hdfc_mod is None:
        st.error("final.py not found in the same folder as this launcher.")
        return

    input_upload = st.file_uploader(
        "Upload Input Excel / CSV File",
        type=["xlsx", "xlsm", "xls", "csv"],
        key="hdfc_input",
    )
    template_upload = st.file_uploader(
        "Upload Template Excel File",
        type=["xlsx", "xlsm"],
        key="hdfc_template",
    )
    refund_upload = st.file_uploader(
        "Upload HDFC UPI Refund File",
        type=["xlsx", "xlsm", "xls", "csv"],
        key="hdfc_refund",
    )

    if input_upload is not None:
        st.success("Input file uploaded successfully")
    if template_upload is not None:
        st.success("Template file uploaded successfully")
    if refund_upload is not None:
        st.info("HDFC UPI Refund file uploaded successfully")

    if input_upload is not None and template_upload is not None:
        if st.button("Process HDFC ESCROW", key="process_hdfc"):
            temp_input = temp_template = temp_refund = None
            try:
                with st.spinner("Processing..."):
                    temp_input = _temp_save_uploaded(input_upload)
                    temp_template = _temp_save_uploaded(template_upload)
                    if refund_upload is not None:
                        temp_refund = _temp_save_uploaded(refund_upload)

                    output_bytes, processed_df = hdfc_mod.process_workbook_to_bytes(
                        temp_input,
                        temp_template,
                        refund_file_path=temp_refund,
                    )

                st.success("Processing Done ✅")

                blank_df = processed_df[
                    processed_df["Tranaction Tag"].astype(str).str.strip().eq("")
                ] if "Tranaction Tag" in processed_df.columns else pd.DataFrame()

                st.subheader("Rows with blank Transaction Tag")
                if blank_df.empty:
                    st.info("No blank Transaction Tag rows found.")
                else:
                    st.dataframe(blank_df)

                download_name = os.path.splitext(input_upload.name)[0] + "_processed.xlsx"
                st.download_button(
                    label="Download Output File",
                    data=output_bytes,
                    file_name=download_name,
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                )

            except Exception as e:
                st.error(f"Processing failed: {e}")
            finally:
                _cleanup_paths([temp_input, temp_template, temp_refund])


# =============================================================================
# Page 3 — SP Cross Check
# =============================================================================
def page_sp_cross_check():
    st.header("3) SP Cross Check")
    st.caption("Upload a ZIP file, select MPR date, and compare the vendor totals.")

    if cross_mod is None:
        st.error("test_hdfc_upi_fixed.py not found in the same folder as this launcher.")
        return

    zip_file = st.file_uploader("Upload ZIP file", type=["zip"], key="zip_file")
    mpr_date = st.date_input("Select MPR Date", value=date.today(), key="mpr_date")

    mid_mapping_upload = st.file_uploader(
        "Optional: Upload MID Mapping Excel file",
        type=["xlsx", "xlsm"],
        key="mid_mapping_upload",
    )

    if mid_mapping_upload is not None:
        temp_mid = None
        try:
            temp_mid = _temp_save_uploaded(mid_mapping_upload)
            if hasattr(cross_mod, "MID_MAPPING_FILE_PATH"):
                cross_mod.MID_MAPPING_FILE_PATH = temp_mid
            st.info("MID mapping file loaded for this session.")
        except Exception as e:
            st.warning(f"Unable to set MID mapping file: {e}")

    if zip_file is not None and st.button("Process ZIP", key="process_zip"):
        try:
            zip_bytes = zip_file.getvalue()
            zip_ref = zipfile.ZipFile(io.BytesIO(zip_bytes))
            zip_names = zip_ref.namelist()

            excel_data = {}
            summary_results = {}

            with st.spinner("Processing vendor files..."):
                for vendor in cross_mod.vendors:
                    status_text, total, df_out, error_text = cross_mod.process_vendor_files(
                        zip_ref, zip_names, vendor
                    )
                    if total is not None and df_out is not None:
                        summary_results[vendor["name"]] = total
                        excel_data[vendor["name"]] = df_out
                        st.success(status_text)
                    else:
                        st.error(status_text)

            if summary_results:
                st.subheader("Summary Table")
                summary_df = pd.DataFrame(
                    list(summary_results.items()),
                    columns=["Vendor", "Total"],
                )
                st.dataframe(summary_df)

                try:
                    main_xlsx = cross_mod.build_workbook_bytes(excel_data, summary_results, mpr_date)
                    summary_xlsx = cross_mod.build_summary_workbook_bytes(summary_df)

                    st.download_button(
                        label="Download Combined Workbook",
                        data=main_xlsx,
                        file_name=f"combined_{mpr_date.strftime('%Y-%m-%d')}.xlsx",
                        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                    )
                    st.download_button(
                        label="Download Summary Workbook",
                        data=summary_xlsx,
                        file_name=f"summary_{mpr_date.strftime('%Y-%m-%d')}.xlsx",
                        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                    )
                except Exception as e:
                    st.warning(f"Could not build download workbook(s): {e}")

                # MID mapping compare if master file is available
                if hasattr(cross_mod, "build_mid_mapping_comparison"):
                    try:
                        comparison_df, filtered_mid_df, mid_error = cross_mod.build_mid_mapping_comparison(
                            summary_df, mpr_date
                        )
                        st.subheader("MID Mapping Compare")
                        if mid_error:
                            st.warning(mid_error)
                        else:
                            st.dataframe(comparison_df)

                            if hasattr(cross_mod, "build_mid_mapping_workbook_bytes"):
                                mid_xlsx = cross_mod.build_mid_mapping_workbook_bytes(
                                    comparison_df, filtered_mid_df
                                )
                                st.download_button(
                                    label="Download MID Mapping Compare Excel",
                                    data=mid_xlsx,
                                    file_name=f"MID_Mapping_Compare_{mpr_date.strftime('%d %b %Y')}.xlsx",
                                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                                )
                    except Exception as e:
                        st.warning(f"MID mapping comparison failed: {e}")
            else:
                st.warning("No vendor totals were produced from the ZIP file.")
        except Exception as e:
            st.error(f"ZIP processing failed: {e}")


# =============================================================================
# Main navigation
# =============================================================================
st.set_page_config(page_title="Ops Automation", layout="wide")
st.title("3-in-1 Operations Launcher")
st.caption("Choose the system from the navigation bar on the left.")

page = st.sidebar.radio(
    "Navigation",
    [
        "Statement Processor",
        "HDFC ESCROW MID MAPPING",
        "SP Cross Check",
    ],
)

if page == "Statement Processor":
    page_statement_processor()
elif page == "HDFC ESCROW MID MAPPING":
    page_hdfc_escrow()
else:
    page_sp_cross_check()
