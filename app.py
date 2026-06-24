from __future__ import annotations

import json
from dataclasses import asdict
from datetime import datetime
from pathlib import Path

import pandas as pd
import streamlit as st

from process_sprinklr import (
    DEFAULT_DATA_SHEET,
    load_config,
    safe_stem,
)
from google_sheet_uploader import (
    periodical_direct_update,
    weekly_direct_append,
)


BASE_DIR = Path(__file__).resolve().parent
CONFIG_PATH = BASE_DIR / "config" / "settings.json"
UPLOAD_DIR = BASE_DIR / "app_uploads"
OUTPUT_DIR = BASE_DIR / "outputs"
LOG_DIR = BASE_DIR / "logs"


def save_uploaded_file(uploaded_file) -> Path:
    UPLOAD_DIR.mkdir(parents=True, exist_ok=True)
    timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
    safe_name = f"{timestamp}_{safe_stem(Path(uploaded_file.name))}{Path(uploaded_file.name).suffix}"
    destination = UPLOAD_DIR / safe_name
    destination.write_bytes(uploaded_file.getbuffer())
    return destination


def run_google_flow(mode: str, uploaded_path: Path, config: dict, dry_run: bool) -> dict:
    LOG_DIR.mkdir(parents=True, exist_ok=True)
    timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
    run_type = "dry_run" if dry_run else "write"
    log_path = LOG_DIR / f"app_{mode}_{run_type}_{safe_stem(uploaded_path)}_{timestamp}.json"

    try:
        if mode == "weekly":
            payload = weekly_direct_append(
                sprinklr_path=uploaded_path,
                sheet_url=config["sheet_url"],
                worksheet_name=config.get("worksheet_name", DEFAULT_DATA_SHEET),
                config=config,
                dry_run=dry_run,
            )
        else:
            payload = periodical_direct_update(
                sprinklr_path=uploaded_path,
                sheet_url=config["sheet_url"],
                worksheet_name=config.get("worksheet_name", DEFAULT_DATA_SHEET),
                config=config,
                dry_run=dry_run,
            )
        payload["status"] = "success"
    except Exception as exc:
        payload = {
            "mode": mode,
            "status": "failed",
            "sprinklr_path": str(uploaded_path),
            "dry_run": dry_run,
            "error": str(exc),
        }

    log_path.write_text(json.dumps(payload, indent=2), encoding="utf-8")
    payload["log_path"] = str(log_path)
    return payload


def metric_card(label: str, value) -> None:
    st.metric(label, value if value is not None else 0)


def audit_dataframe(payload: dict) -> pd.DataFrame:
    rows = payload.get("row_audit") or []
    if not rows:
        return pd.DataFrame(columns=["source_row", "action", "permalink", "reason", "target_row"])
    return pd.DataFrame(rows)


def main() -> None:
    st.set_page_config(
        page_title="Data Workflow",
        page_icon="📊",
        layout="wide",
    )

    st.title("📊Data Workflow")
    st.caption("Upload Sprinklr exports, Preview Changes, and update the reporting Google Sheet!")
    st.divider()

    config = load_config(CONFIG_PATH)

    with st.sidebar.expander("⚙️ Configuration", expanded=False):     
        st.write("Google Sheet ")
        st.success("Connected")
        st.write("Worksheet")
        st.code(config.get("worksheet_name", DEFAULT_DATA_SHEET), language="text")
        st.divider()
        st.write("Current mode")
        st.caption("Preview the changes first. Google Sheet is updated only after Confirm Write.")
        
        
    st.subheader("Select Mode")
    mode_label = st.radio(
        "",
        ["📥 Weekly Append", "📈 Periodical Update"],
        horizontal=True,
    )
    mode = "weekly" if mode_label == "📥 Weekly Append" else "periodical"
    if mode == "weekly":
        st.info("Adds new posts to Google Sheet. Existing permalinks are skipped.")
    else:
        st.info("Updates metrics for existing posts by matching permalinks.")

    uploaded_file = st.file_uploader(
        "Upload Sprinklr dump",
        type=["xlsx", "xlsm"],
        help="Upload the Sprinklr export file here.",
    )

    run_clicked = st.button("🔍 Preview Changes", type="primary", disabled=uploaded_file is None)

    if run_clicked and uploaded_file is not None:
        uploaded_path = save_uploaded_file(uploaded_file)
        with st.spinner("Processing Sprinklr dump..."):
            payload = run_google_flow(mode, uploaded_path, config, dry_run=True)
        st.session_state["last_payload"] = payload
        st.session_state["last_uploaded_path"] = str(uploaded_path)
        st.session_state["last_mode"] = mode

    payload = st.session_state.get("last_payload")

    if not payload:
        st.info("Upload a Sprinklr dump, choose a workflow, and click Preview Changes")
        return

    status = payload.get("status")
    if status == "failed":
        st.error("Preview failed.")
        st.write(payload.get("error", "Unknown error"))
        return

    st.success("Preview completed. Check the audit before writing to Google Sheets.")

    st.subheader("Summary")
    c1, c2, c3, c4, c5, c6 = st.columns(6)
    with c1:
        metric_card("Rows seen", payload.get("sprinklr_rows_seen"))
    with c2:
        metric_card("To append", payload.get("rows_to_append", payload.get("rows_appended")))
    with c3:
        metric_card("To update", payload.get("rows_to_update", payload.get("rows_updated")))
    with c4:
        metric_card("Deleted removed", payload.get("deleted_posts_removed"))
    with c5:
        metric_card("Duplicates skipped", payload.get("duplicate_rows_skipped"))
    with c6:
        metric_card("Unmatched", payload.get("unmatched_periodical_rows"))

    missing_source = payload.get("missing_source_columns") or []
    missing_master = payload.get("missing_master_columns") or []
    missing_google = payload.get("missing_google_headers") or []
    if missing_source or missing_master or missing_google:
        st.warning("Missing columns were detected.")
        if missing_source:
            st.write("Missing Sprinklr columns:", ", ".join(missing_source))
        if missing_master:
            st.write("Missing master columns:", ", ".join(missing_master))
        if missing_google:
            st.write("Missing Google Sheet columns:", ", ".join(missing_google))

    with st.expander("🔍 Row Audit",expanded=False):
        df = audit_dataframe(payload)
        if not df.empty:
            action_filter = st.multiselect(
                "Filter actions",
                sorted(df["action"]
                    .dropna()
                    .unique()
                    .tolist()),
                default=sorted(
                    df["action"]
                        .dropna()
                        .unique()
                        .tolist()),)
            if action_filter:
                df = df[
                    df["action"].isin(action_filter)
                    ]
                st.dataframe(
                    df,
                    use_container_width=True,
                    hide_index=True,)
            else:
                st.info("No audit rows available.")

    st.divider()
    st.subheader("⚠️Confirm Write")
    st.warning("If you are satisfied with the preview, click the button below. This action will write Google Sheet.")
    confirm = st.button("Update Google Sheet", type="primary")
    if confirm:
        uploaded_path = Path(st.session_state["last_uploaded_path"])
        last_mode = st.session_state["last_mode"]
        with st.spinner("Writing to Google Sheet..."):
            write_payload = run_google_flow(last_mode, uploaded_path, config, dry_run=False)
        st.session_state["last_write_payload"] = write_payload
        if write_payload.get("status") == "success":
            st.success("✅ Append/Update in Google Sheet Succcessful.")
            st.subheader("Write Summary")
            c1, c2, c3 = st.columns(3)
            if write_payload.get("mode") == "weekly_direct_append":
                with c1:
                    st.metric("Posts Found",write_payload.get("sprinklr_rows_seen",0))
                with c2:
                    st.metric("Already In Sheet",write_payload.get("duplicate_rows_skipped",0))
                with c3:
                    st.metric("New Posts Added",write_payload.get("rows_appended",0))
            else:
                with c1:
                    st.metric("Posts Found",write_payload.get("sprinklr_rows_seen",0))
                with c2:
                    st.metric("Posts To Update",write_payload.get("rows_to_update",0))
                with c3:
                    st.metric("Posts Updated",write_payload.get("rows_updated",0))
            with st.expander("📄 Technical Details",expanded=False):
                st.json(
                    {
                        k: v
                        for k, v in write_payload.items()
                        if k != "row_audit"
                    }
                )
        else:
            st.error("Google Sheet write failed.")
            st.write(write_payload.get("error", "Unknown error"))

    st.subheader("📋Run Log")
    st.code(payload.get("log_path", ""), language="text")


if __name__ == "__main__":
    main()
