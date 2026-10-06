import contextlib
import io
import pandas as pd
import re
import tempfile
from datetime import datetime
from pathlib import Path

import streamlit as st

from chdn_clean import run_clean


def _parse_clean_output(clean_output: str) -> dict[str, int]:
    metrics: dict[str, int] = {}
    patterns = {
        "children_code_missing": r"children_code missing:\s*(\d+)",
        "children_code_duplicate_values": r"duplicate children_code values:\s*(\d+)",
        "dob_after_first_visit": r"DOB later than first_visit_date:\s*(\d+)",
        "later_before_prior": r"later dose earlier than prior dose:\s*(\d+)",
        "later_while_prior_not_received": r"later dose with prior dose not received yet:\s*(\d+)",
        "interval_under_min_days": r"interval between doses < 28 days[^:\n]*:\s*(\d+)",
        "pw_code_missing": r"pw_code missing:\s*(\d+)",
        "pw_code_duplicate_values": r"duplicate pw_code values:\s*(\d+)",
        "td2_earlier_than_td1": r"Td2 earlier than Td1:\s*(\d+)",
        "td2_while_td1_not_received": r"Td2 received while Td1 not received yet:\s*(\d+)",
    }
    for key, pattern in patterns.items():
        match = re.search(pattern, clean_output)
        metrics[key] = int(match.group(1)) if match else 0
    return metrics


def _status_label(count: int) -> str:
    return "Issue" if count > 0 else "OK"


def _style_status_table(df: pd.DataFrame):
    def style_row(row):
        issue = row["Status"] == "Issue"
        color = "#f8d7da" if issue else "#d1e7dd"
        return [f"background-color: {color}" if col == "Status" else "" for col in df.columns]

    return df.style.apply(style_row, axis=1)


def _show_verification_dashboard(metrics: dict[str, int]) -> None:
    st.subheader("Verification Dashboard")

    child_df = pd.DataFrame([
        {"Check": "children_code missing", "Count": metrics.get("children_code_missing", 0), "Status": _status_label(metrics.get("children_code_missing", 0))},
        {"Check": "duplicate children_code", "Count": metrics.get("children_code_duplicate_values", 0), "Status": _status_label(metrics.get("children_code_duplicate_values", 0))},
        {"Check": "DOB later than first visit", "Count": metrics.get("dob_after_first_visit", 0), "Status": _status_label(metrics.get("dob_after_first_visit", 0))},
        {"Check": "later dose earlier than prior", "Count": metrics.get("later_before_prior", 0), "Status": _status_label(metrics.get("later_before_prior", 0))},
        {"Check": "later dose with prior not received", "Count": metrics.get("later_while_prior_not_received", 0), "Status": _status_label(metrics.get("later_while_prior_not_received", 0))},
        {"Check": "dose interval < 28 days (OPV/Penta/MMR)", "Count": metrics.get("interval_under_min_days", 0), "Status": _status_label(metrics.get("interval_under_min_days", 0))},
    ])

    pw_df = pd.DataFrame([
        {"Check": "pw_code missing", "Count": metrics.get("pw_code_missing", 0), "Status": _status_label(metrics.get("pw_code_missing", 0))},
        {"Check": "duplicate pw_code", "Count": metrics.get("pw_code_duplicate_values", 0), "Status": _status_label(metrics.get("pw_code_duplicate_values", 0))},
        {"Check": "Td2 earlier than Td1", "Count": metrics.get("td2_earlier_than_td1", 0), "Status": _status_label(metrics.get("td2_earlier_than_td1", 0))},
        {"Check": "Td2 received while Td1 not received", "Count": metrics.get("td2_while_td1_not_received", 0), "Status": _status_label(metrics.get("td2_while_td1_not_received", 0))},
    ])

    left, right = st.columns(2)
    with left:
        st.caption("Child sheet")
        st.table(_style_status_table(child_df))
    with right:
        st.caption("PW sheet")
        st.table(_style_status_table(pw_df))

st.set_page_config(page_title="CHDN EPI verification and clean file export", layout="wide")
st.title("CHDN EPI verification and clean file export")

st.info("This app runs the EPI verification and cleaning step, then exports the cleaned workbook.")

uploaded_file = st.file_uploader("Upload the source EPI workbook", type=["xlsx", "xlsm"])

if uploaded_file is not None:
    file_id = uploaded_file.file_id

    if st.session_state.get("selected_file_id") != file_id:
        st.session_state["selected_file_id"] = file_id
        st.session_state.pop("processed_file_id", None)
        st.session_state.pop("result_status", None)
        st.session_state.pop("result_data", None)
        st.session_state.pop("export_name", None)
        st.session_state.pop("verification_metrics", None)
        st.session_state.pop("clean_output", None)

    run_report = st.button("Run report", type="primary")

    if run_report:
        export_stamp = datetime.now().strftime("%Y%m%d_%H-%M")
        export_name = f"CHDN_Cleaned_{export_stamp}.xlsx"
        output_path = Path(__file__).resolve().with_name(export_name)
        if output_path.exists():
            output_path.unlink()

        with tempfile.NamedTemporaryFile(delete=False, suffix=".xlsx") as tmp_file:
            tmp_file.write(uploaded_file.getvalue())
            temp_input = Path(tmp_file.name)

        with st.spinner("Running the cleaning step..."):
            stdout_buffer = io.StringIO()
            with contextlib.redirect_stdout(stdout_buffer):
                status = run_clean(temp_input, output_path)
            clean_output = stdout_buffer.getvalue()

        result_data = None
        if status == 0 and output_path.exists():
            with open(output_path, "rb") as fh:
                result_data = fh.read()

        st.session_state["processed_file_id"] = file_id
        st.session_state["result_status"] = status
        st.session_state["result_data"] = result_data
        st.session_state["export_name"] = export_name
        st.session_state["clean_output"] = clean_output
        st.session_state["verification_metrics"] = _parse_clean_output(clean_output)

    result_status = st.session_state.get("result_status")
    result_data = st.session_state.get("result_data")
    export_name = st.session_state.get("export_name", "CHDN_Cleaned.xlsx")
    verification_metrics = st.session_state.get("verification_metrics")
    clean_output = st.session_state.get("clean_output", "")

    if result_status == 0 and result_data is not None:
        st.success("Cleaning step completed successfully.")
        if verification_metrics is not None:
            _show_verification_dashboard(verification_metrics)
        st.download_button(
            label="Download cleaned workbook",
            data=result_data,
            file_name=export_name,
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        )
        if clean_output:
            with st.expander("Verification details"):
                st.text(clean_output)
    elif st.session_state.get("processed_file_id") == file_id:
        st.error("The cleaning step did not finish successfully.")
        if clean_output:
            with st.expander("Run details"):
                st.text(clean_output)
    else:
        st.info("Upload a workbook, then click Run report to start the cleaning step.")
else:
    # Clear cached results when no file is present so a re-upload is processed fresh.
    st.session_state.pop("selected_file_id", None)
    st.session_state.pop("processed_file_id", None)
    st.session_state.pop("result_status", None)
    st.session_state.pop("result_data", None)
    st.session_state.pop("export_name", None)
    st.session_state.pop("verification_metrics", None)
    st.session_state.pop("clean_output", None)
    st.warning("Please upload a workbook to begin.")
