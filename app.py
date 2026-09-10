import io
import os
import subprocess
from datetime import datetime
from typing import Dict, List, Optional, Tuple

import pandas as pd
import plotly.express as px
import streamlit as st


st.set_page_config(
    page_title="TAT Merge Dashboard",
    page_icon="📦",
    layout="wide",
    initial_sidebar_state="expanded",
)


SUPPORTED_EXTENSIONS = ["xlsx", "xls", "csv"]


def read_uploaded_file(uploaded_file) -> pd.DataFrame:
    """Read CSV/XLS/XLSX uploads into a DataFrame."""
    if uploaded_file is None:
        return pd.DataFrame()

    file_name = uploaded_file.name.lower()
    try:
        if file_name.endswith(".csv"):
            return pd.read_csv(uploaded_file)
        if file_name.endswith(".xlsx"):
            return pd.read_excel(uploaded_file, engine="openpyxl")
        if file_name.endswith(".xls"):
            try:
                return pd.read_excel(uploaded_file, engine="xlrd")
            except Exception as exc:
                raise ValueError(
                    "'.xls' files require the optional 'xlrd' package. Install it with: pip install xlrd"
                ) from exc
        raise ValueError("Unsupported file type. Please upload a .csv, .xls, or .xlsx file.")
    except Exception as exc:
        raise ValueError(f"Failed to read {uploaded_file.name}: {exc}") from exc


def normalize_identifier(value: object) -> str:
    if pd.isna(value):
        return ""
    text = str(value).strip().lower()
    text = " ".join(text.split())
    return text


def humanize_tat(hours_value: Optional[float]) -> str:
    if pd.isna(hours_value):
        return ""
    total_minutes = int(round(float(hours_value) * 60))
    sign = "-" if total_minutes < 0 else ""
    total_minutes = abs(total_minutes)
    days, remainder = divmod(total_minutes, 24 * 60)
    hours, minutes = divmod(remainder, 60)
    return f"{sign}{days} Days, {hours} Hours, {minutes} Mins"


def get_available_columns(df: pd.DataFrame) -> List[str]:
    return [str(col) for col in df.columns.tolist()]


def safe_selectbox(label: str, options: List[str], key: str) -> Optional[str]:
    if not options:
        st.error(f"No columns available for {label}.")
        return None
    return st.selectbox(label, options, key=key)


def detect_column_by_keywords(columns: List[str], keywords: List[str]) -> Optional[str]:
    """Return the first column that contains any of the keywords (case-insensitive)."""
    low_cols = [c.lower() for c in columns]
    for kw in keywords:
        kw_low = kw.lower()
        for orig, low in zip(columns, low_cols):
            if kw_low in low:
                return orig
    return None


def build_standardized_frame(
    df: pd.DataFrame,
    product_col: str,
    timestamp_col: str,
    source_label: str,
) -> pd.DataFrame:
    frame = df.copy()
    frame = frame.reset_index(drop=True)
    frame["__row_id"] = frame.index
    frame["__normalized_product"] = frame[product_col].apply(normalize_identifier)
    frame["__source"] = source_label
    frame["__product_name"] = frame[product_col]
    frame["__timestamp_raw"] = frame[timestamp_col]
    frame["__timestamp"] = pd.to_datetime(frame[timestamp_col], errors="coerce")
    return frame


def assign_product_ids(*frames: pd.DataFrame) -> Tuple[List[pd.DataFrame], Dict[str, str]]:
    combined_keys: List[str] = []
    for frame in frames:
        if not frame.empty:
            combined_keys.extend([key for key in frame["__normalized_product"].dropna().tolist() if key != ""])

    unique_keys = sorted(set(combined_keys))
    product_map = {key: f"PID-{1000 + idx}" for idx, key in enumerate(unique_keys, start=1)}

    updated_frames: List[pd.DataFrame] = []
    for frame in frames:
        current = frame.copy()
        current["Product_ID"] = current["__normalized_product"].map(product_map)
        updated_frames.append(current)

    return updated_frames, product_map


def merge_frames(
    ready_df: pd.DataFrame,
    dispatch_df: pd.DataFrame,
    join_type: str,
) -> pd.DataFrame:
    ready_cols = {
        "__product_name": "Ready_Product_Name",
        "__timestamp": "Ready_Timestamp",
        "__timestamp_raw": "Ready_Timestamp_Raw",
        "__normalized_product": "__normalized_product",
        "Product_ID": "Product_ID",
    }
    dispatch_cols = {
        "__product_name": "Dispatch_Product_Name",
        "__timestamp": "Dispatch_Timestamp",
        "__timestamp_raw": "Dispatch_Timestamp_Raw",
        "__normalized_product": "__normalized_product",
        "Product_ID": "Product_ID",
    }

    ready_view = ready_df[list(ready_cols.keys())].rename(columns=ready_cols)
    dispatch_view = dispatch_df[list(dispatch_cols.keys())].rename(columns=dispatch_cols)

    merged = pd.merge(
        ready_view,
        dispatch_view,
        on=["Product_ID", "__normalized_product"],
        how=join_type,
        suffixes=("_ready", "_dispatch"),
    )
    return merged


def calculate_tat(df: pd.DataFrame) -> pd.DataFrame:
    output = df.copy()
    output["Time_Gap_Timedelta"] = output["Dispatch_Timestamp"] - output["Ready_Timestamp"]
    output["TAT_Hours"] = output["Time_Gap_Timedelta"].dt.total_seconds() / 3600

    def classify_row(row: pd.Series) -> str:
        ready_ts = row.get("Ready_Timestamp")
        dispatch_ts = row.get("Dispatch_Timestamp")
        tat_hours = row.get("TAT_Hours")

        if pd.isna(ready_ts):
            return "Missing Ready"
        if pd.isna(dispatch_ts):
            return "Pending Dispatch"
        if pd.notna(tat_hours) and tat_hours < 0:
            return "Data Error"
        return "OK"

    output["Status"] = output.apply(classify_row, axis=1)
    output["Time_Gap"] = output["TAT_Hours"].apply(humanize_tat)
    output.loc[output["Status"] == "Pending Dispatch", "Time_Gap"] = "Pending Dispatch"
    output.loc[output["Status"] == "Data Error", "Time_Gap"] = "Data Error"
    output["Time_Lag"] = output["Time_Gap"]
    return output


def prepare_export_frame(df: pd.DataFrame) -> pd.DataFrame:
    export_df = df.copy()
    export_df["Time_Gap"] = export_df["Time_Gap"].fillna("")
    columns_order = [
        "Product_ID",
        "Ready_Product_Name",
        "Dispatch_Product_Name",
        "Ready_Timestamp",
        "Dispatch_Timestamp",
        "Time_Lag",
        "Time_Gap",
        "TAT_Hours",
        "Status",
        "Ready_Timestamp_Raw",
        "Dispatch_Timestamp_Raw",
    ]
    existing_columns = [col for col in columns_order if col in export_df.columns]
    remaining_columns = [col for col in export_df.columns if col not in existing_columns]
    return export_df[existing_columns + remaining_columns]


def to_excel_bytes(df: pd.DataFrame) -> bytes:
    buffer = io.BytesIO()
    with pd.ExcelWriter(buffer, engine="openpyxl") as writer:
        df.to_excel(writer, index=False, sheet_name="Merged_TAT_Report")
    buffer.seek(0)
    return buffer.read()


def render_kpis(df: pd.DataFrame) -> None:
    matched_rows = df[(df["Status"] == "OK") & df["Ready_Timestamp"].notna() & df["Dispatch_Timestamp"].notna()]
    total_matched = int(matched_rows["Product_ID"].nunique()) if not matched_rows.empty else 0
    valid_tat = df.loc[df["Status"] == "OK", "TAT_Hours"] if not df.empty else pd.Series(dtype=float)
    average_tat = float(valid_tat.mean()) if not valid_tat.empty else 0.0
    max_tat = float(valid_tat.max()) if not valid_tat.empty else 0.0
    error_pending = int(df["Status"].isin(["Data Error", "Pending Dispatch", "Missing Ready"]).sum()) if not df.empty else 0

    col1, col2, col3, col4 = st.columns(4)
    col1.metric("Total Matched Products", f"{total_matched}")
    col2.metric("Average TAT (Hours)", f"{average_tat:.2f}")
    col3.metric("Max TAT (Hours)", f"{max_tat:.2f}")
    col4.metric("Error / Pending Count", f"{error_pending}")


def main() -> None:
    st.title("Production vs Dispatch TAT Dashboard")
    st.caption("Upload two files, map the relevant columns, and generate a merged turnaround-time report.")

    with st.sidebar:
        st.header("File Upload & Mapping")
        ready_upload = st.file_uploader(
            "Upload Ready/Production Sheet",
            type=SUPPORTED_EXTENSIONS,
            accept_multiple_files=False,
        )
        dispatch_upload = st.file_uploader(
            "Upload Dispatch Sheet",
            type=SUPPORTED_EXTENSIONS,
            accept_multiple_files=False,
        )

        # Option to save uploaded files into the repository and push to remote
        st.markdown("---")
        st.write("You can save uploaded files into the repo (folder: `uploaded_files/`) and attempt to push to remote.")
        if st.button("Save uploaded files to repo"):
            def save_upload_to_repo(uploaded_file, subdir="uploaded_files") -> str:
                if uploaded_file is None:
                    return "no_file"
                try:
                    os.makedirs(subdir, exist_ok=True)
                    # ensure stream is at start
                    try:
                        uploaded_file.seek(0)
                    except Exception:
                        pass
                    data = uploaded_file.read()
                    timestamp = datetime.utcnow().strftime("%Y%m%d_%H%M%S")
                    safe_name = os.path.basename(uploaded_file.name)
                    dest_name = f"{timestamp}_{safe_name}"
                    dest_path = os.path.join(subdir, dest_name)
                    with open(dest_path, "wb") as f:
                        if isinstance(data, str):
                            f.write(data.encode())
                        else:
                            f.write(data)
                    return dest_path
                except Exception as exc:
                    return f"error:{exc}"

            saved_paths = []
            for f in (ready_upload, dispatch_upload):
                if f is not None:
                    res = save_upload_to_repo(f)
                    saved_paths.append(res)

            if not saved_paths:
                st.warning("No uploaded files to save. Upload files first.")
            else:
                st.success("Saved files: " + ", ".join(saved_paths))

                # attempt git add/commit/push
                try:
                    # add
                    subprocess.run(["git", "add"] + saved_paths, check=True)
                    commit_msg = f"Add uploaded files: {' ,'.join([os.path.basename(p) for p in saved_paths])}"
                    subprocess.run(["git", "commit", "-m", commit_msg], check=True)
                    push_proc = subprocess.run(["git", "push"], check=False, capture_output=True, text=True)
                    if push_proc.returncode == 0:
                        st.success("Pushed committed uploads to remote.")
                    else:
                        st.error("Commit created locally but push failed. See output below.")
                        st.code(push_proc.stderr)
                except FileNotFoundError:
                    st.error("`git` not found in the server environment. Cannot commit/push from this app instance.")
                except subprocess.CalledProcessError as cpe:
                    st.error(f"Git command failed: {cpe}")
                except Exception as exc:
                    st.error(f"Unexpected error while committing/pushing: {exc}")

        join_type = st.radio(
            "Merge Type",
            options=["inner", "outer"],
            index=0,
            horizontal=True,
        )

    if ready_upload is None or dispatch_upload is None:
        st.info("Upload both files to begin.")
        st.stop()

    try:
        ready_df = read_uploaded_file(ready_upload)
        dispatch_df = read_uploaded_file(dispatch_upload)
    except ValueError as exc:
        st.error(str(exc))
        st.stop()

    if ready_df.empty or dispatch_df.empty:
        st.error("One or both uploaded files are empty.")
        st.stop()

    with st.expander("Preview Uploaded Ready/Production Sheet", expanded=False):
        st.dataframe(ready_df, use_container_width=True, height=260)

    with st.expander("Preview Uploaded Dispatch Sheet", expanded=False):
        st.dataframe(dispatch_df, use_container_width=True, height=260)

    ready_columns = get_available_columns(ready_df)
    dispatch_columns = get_available_columns(dispatch_df)

    st.sidebar.subheader("Column Mapping")
    # Attempt to auto-detect likely columns based on common header keywords
    ready_product_default = detect_column_by_keywords(
        ready_columns,
        [
            "product",
            "product id",
            "material",
            "item",
            "goods",
            "original line item",
            "item automatically created",
            "user name",
        ],
    )
    ready_timestamp_default = detect_column_by_keywords(
        ready_columns, ["ready", "ready timestamp", "production", "date", "time", "timestamp", "posting date"]
    )

    dispatch_product_default = detect_column_by_keywords(
        dispatch_columns, ["product", "material", "item", "goods", "dispatch", "delivery"]
    )
    dispatch_timestamp_default = detect_column_by_keywords(
        dispatch_columns, ["dispatch", "dispatch timestamp", "delivery date", "date", "time", "timestamp"]
    )

    if ready_product_default and ready_product_default in ready_columns:
        ready_product_col = st.selectbox(
            "Sheet 1: Product Identifier / Name",
            ready_columns,
            index=ready_columns.index(ready_product_default),
            key="ready_product_col",
        )
    else:
        ready_product_col = safe_selectbox("Sheet 1: Product Identifier / Name", ready_columns, "ready_product_col")

    if ready_timestamp_default and ready_timestamp_default in ready_columns:
        ready_timestamp_col = st.selectbox(
            "Sheet 1: Ready Timestamp",
            ready_columns,
            index=ready_columns.index(ready_timestamp_default),
            key="ready_timestamp_col",
        )
    else:
        ready_timestamp_col = safe_selectbox("Sheet 1: Ready Timestamp", ready_columns, "ready_timestamp_col")

    if dispatch_product_default and dispatch_product_default in dispatch_columns:
        dispatch_product_col = st.selectbox(
            "Sheet 2: Product Identifier / Name",
            dispatch_columns,
            index=dispatch_columns.index(dispatch_product_default),
            key="dispatch_product_col",
        )
    else:
        dispatch_product_col = safe_selectbox("Sheet 2: Product Identifier / Name", dispatch_columns, "dispatch_product_col")

    if dispatch_timestamp_default and dispatch_timestamp_default in dispatch_columns:
        dispatch_timestamp_col = st.selectbox(
            "Sheet 2: Dispatch Timestamp",
            dispatch_columns,
            index=dispatch_columns.index(dispatch_timestamp_default),
            key="dispatch_timestamp_col",
        )
    else:
        dispatch_timestamp_col = safe_selectbox("Sheet 2: Dispatch Timestamp", dispatch_columns, "dispatch_timestamp_col")

    if not all([ready_product_col, ready_timestamp_col, dispatch_product_col, dispatch_timestamp_col]):
        st.stop()

    try:
        ready_std = build_standardized_frame(ready_df, ready_product_col, ready_timestamp_col, "ready")
        dispatch_std = build_standardized_frame(dispatch_df, dispatch_product_col, dispatch_timestamp_col, "dispatch")

        ready_std, dispatch_std = assign_product_ids(ready_std, dispatch_std)
        merged_df = merge_frames(ready_std, dispatch_std, join_type=join_type)
        processed_df = calculate_tat(merged_df)
    except Exception as exc:
        st.error(f"Unable to process data: {exc}")
        st.stop()

    render_kpis(processed_df)

    st.divider()
    st.subheader("Visual Analysis")

    chart_col1, chart_col2 = st.columns([2, 1])
    with chart_col1:
        plot_kind = st.radio("Chart Type", options=["Histogram", "Boxplot"], horizontal=True)
        chart_data = processed_df.loc[processed_df["Status"] == "OK", ["Product_ID", "TAT_Hours", "Ready_Product_Name"]].copy()
        if chart_data.empty:
            st.warning("No valid TAT data available for plotting.")
        else:
            if plot_kind == "Histogram":
                fig = px.histogram(
                    chart_data,
                    x="TAT_Hours",
                    nbins=20,
                    title="Distribution of Turnaround Times",
                    labels={"TAT_Hours": "TAT Hours"},
                )
            else:
                fig = px.box(
                    chart_data,
                    y="TAT_Hours",
                    points="all",
                    title="Turnaround Time Boxplot",
                    labels={"TAT_Hours": "TAT Hours"},
                )
            fig.update_layout(height=450, margin=dict(l=20, r=20, t=50, b=20))
            st.plotly_chart(fig, use_container_width=True)

    with chart_col2:
        st.subheader("Search & Filter")
        search_term = st.text_input("Search by Product_ID or Product Name")
        tat_threshold = st.slider("Filter TAT greater than (hours)", min_value=0, max_value=240, value=24, step=1)
        status_filter = st.multiselect(
            "Status Filter",
            options=["OK", "Data Error", "Pending Dispatch", "Missing Ready"],
            default=["OK", "Data Error", "Pending Dispatch", "Missing Ready"],
        )

    display_df = processed_df.copy()
    if search_term.strip():
        search_value = search_term.strip().lower()
        display_df = display_df[
            display_df["Product_ID"].fillna("").astype(str).str.lower().str.contains(search_value, na=False, regex=False)
            | display_df["Ready_Product_Name"].fillna("").astype(str).str.lower().str.contains(search_value, na=False, regex=False)
            | display_df["Dispatch_Product_Name"].fillna("").astype(str).str.lower().str.contains(search_value, na=False, regex=False)
        ]

    display_df = display_df[display_df["Status"].isin(status_filter)]
    display_df = display_df[(display_df["TAT_Hours"].isna()) | (display_df["TAT_Hours"] > tat_threshold) | (display_df["Status"] != "OK")]

    st.subheader("Merged Data")

    table_columns = [
        "Product_ID",
        "Ready_Product_Name",
        "Dispatch_Product_Name",
        "Ready_Timestamp",
        "Dispatch_Timestamp",
        "Time_Lag",
        "Time_Gap",
        "TAT_Hours",
        "Status",
        "Ready_Timestamp_Raw",
        "Dispatch_Timestamp_Raw",
    ]
    table_columns = [col for col in table_columns if col in display_df.columns]

    st.dataframe(
        display_df[table_columns],
        use_container_width=True,
        height=500,
    )

    export_df = prepare_export_frame(processed_df)
    excel_bytes = to_excel_bytes(export_df)

    st.download_button(
        label="Download final_complete_sheet.xlsx",
        data=excel_bytes,
        file_name="final_complete_sheet.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
    )

    with st.expander("Processing Details", expanded=False):
        st.write("Join type:", join_type)
        st.write("Ready rows:", len(ready_df))
        st.write("Dispatch rows:", len(dispatch_df))
        st.write("Merged rows:", len(processed_df))
        st.write("Unique normalized products:", len(set(processed_df["Product_ID"].dropna().tolist())))


if __name__ == "__main__":
    main()