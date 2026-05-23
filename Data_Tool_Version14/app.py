import os
import io
import json
import time
import threading
import uuid
import traceback
from typing import Dict, List, Tuple
from datetime import datetime

import pandas as pd
import streamlit as st
from streamlit.runtime.scriptrunner import add_script_run_ctx

from engine import (
    preview_file,
    get_headers_from_file,
    list_lookup_files,
    run_engine,
    safe_file_exists,
    build_output_headers,
    run_pdf_merger,
    run_excel_csv_merger,
    run_excel_search,
    run_name_match,
    run_file_zipper,
    run_file_downloader,
    run_mail_merge_tool,
    dataframe_summary,
    apply_all_transformations,
    run_name_cleaner,
    run_data_cleaner,       
    run_advanced_profiler,   # <-- Fixed Import
    run_base64_extractor,    # <-- Fixed Import
    run_education_ocr        # <-- Fixed Import
)

# Set wider layout and a clean initial state
st.set_page_config(page_title="High Scale Data Suite", page_icon="🚀", layout="wide", initial_sidebar_state="expanded")

# =========================================================
# UI COLOR THEME & CUSTOM CSS
# =========================================================
st.markdown("""
<style>
    /* Gradient Background for Headers */
    h1, h2, h3 {
        background: -webkit-linear-gradient(45deg, #1E88E5, #00ACC1);
        -webkit-background-clip: text;
        -webkit-text-fill-color: transparent;
        font-weight: 700;
    }
    /* Subtle tinted background for containers to make them pop */
    [data-testid="stVerticalBlockBorderWrapper"] {
        background-color: #f8fbff;
        border-radius: 10px;
        box-shadow: 0 4px 6px rgba(0,0,0,0.05);
        border: 1px solid #e3f2fd !important;
    }
    /* Accent buttons */
    div.stButton > button:first-child {
        border-radius: 6px;
        transition: all 0.3s ease;
    }
    div.stButton > button:first-child:hover {
        transform: translateY(-2px);
        box-shadow: 0 4px 12px rgba(0,0,0,0.15);
    }
    /* Sidebar restyling */
    [data-testid="stSidebar"] {
        background-color: 0d1117;
        border-right: 1px solid #d0e3ff;
    }
</style>
""", unsafe_allow_html=True)


# =========================================================
# SESSION STATE DEFAULTS
# =========================================================
DEFAULT_PROGRESS = {
    "phase": "IDLE",
    "lookup_percent": 0,
    "input_percent": 0,
    "lookup_rows_scanned": 0,
    "lookup_files_done": 0,
    "lookup_files_total": 0,
    "input_chunk_count": 0,
    "input_rows_processed": 0,
    "throughput_rps": 0,
    "eta_sec": 0,
    "final_total": 0,
    "unique_count": 0,
    "duplicate_count": 0,
    "review_count": 0,
    "value_counts": {},
    "status": "IDLE",
    "error": "",
}

DEFAULT_METADATA = {
    "input_headers": [],
    "main_lookup_headers": [],
    "special_lookup_headers": [],
    "main_lookup_files": [],
    "input_preview": pd.DataFrame(),
    "main_lookup_preview": pd.DataFrame(),
    "special_lookup_preview": pd.DataFrame(),
    "loaded": False,
}

if "logs" not in st.session_state: st.session_state.logs = []
if "progress_state" not in st.session_state: st.session_state.progress_state = DEFAULT_PROGRESS.copy()
if "run_result" not in st.session_state: st.session_state.run_result = None
if "is_running" not in st.session_state: st.session_state.is_running = False
if "metadata" not in st.session_state: st.session_state.metadata = DEFAULT_METADATA.copy()
if "saved_config" not in st.session_state: st.session_state.saved_config = None
if "main_row_ids" not in st.session_state: st.session_state.main_row_ids = [str(uuid.uuid4())]
if "special_row_ids" not in st.session_state: st.session_state.special_row_ids = [str(uuid.uuid4())]

# Lookup Engine Memory
for key, default_val in [
    ("le_input_file", ""), 
    ("le_lookup_folders", ""), 
    ("le_special_file", ""), 
    ("le_output_folder", ""), 
    ("le_db_path", "lookup_index.duckdb"), 
    ("le_preview_n", 500),
    ("le_key_type", "text")
]:
    if key not in st.session_state: st.session_state[key] = default_val

# Tool States 
for tool in ["pdf", "em", "es", "nm", "zip", "dl", "mm", "nc", "dc", "sp", "b64", "ocr"]:
    if f"{tool}_status" not in st.session_state: st.session_state[f"{tool}_status"] = "IDLE"
    if f"{tool}_logs" not in st.session_state: st.session_state[f"{tool}_logs"] = []
    if f"{tool}_start_time" not in st.session_state: st.session_state[f"{tool}_start_time"] = None

if "lookup_step" not in st.session_state: st.session_state.lookup_step = "1. Setup & Previews"
if "nm_headers" not in st.session_state: st.session_state.nm_headers = []
if "dl_headers" not in st.session_state: st.session_state.dl_headers = []
if "nc_headers" not in st.session_state: st.session_state.nc_headers = []
if "dc_headers" not in st.session_state: st.session_state.dc_headers = []
if "sp_headers" not in st.session_state: st.session_state.sp_headers = []


# =========================================================
# UI STYLING & HELPERS
# =========================================================
def format_eta(seconds: int) -> str:
    if seconds < 0: return "Calculating..."
    m, s = divmod(seconds, 60)
    h, m = divmod(m, 60)
    if h > 0: return f"{h}h {m}m {s}s"
    elif m > 0: return f"{m}m {s}s"
    return f"{s}s"

def render_status_badge(status: str, phase: str = ""):
    if status == "COMPLETED":
        st.success("🟢 **STATUS: COMPLETED** | All tasks finished successfully.")
    elif status == "FAILED":
        st.error(f"🔴 **STATUS: FAILED** | The job crashed. Review the telemetry logs below.")
    elif status == "RUNNING":
        phase_str = f" | Phase: {phase}" if phase else ""
        st.info(f"🔵 **STATUS: RUNNING**{phase_str} | Working in background...")
    else:
        st.markdown("⚪ **STATUS: IDLE** | Ready to start.")

def render_structured_logs(logs: List[Dict]):
    display_logs = logs[-50:] if len(logs) > 50 else logs
    for log in display_logs:
        color = "blue" if log["level"] == "INFO" else "green" if log["level"] == "SUCCESS" else "red" if log["level"] == "ERROR" else "orange"
        st.markdown(f"**:{color}[{log['timestamp']}]** | **{log['level']}** | {log['message']}")

@st.cache_data(show_spinner=False)
def cached_get_headers(file_path: str) -> List[str]: return get_headers_from_file(file_path)

@st.cache_data(show_spinner=False)
def cached_preview(file_path: str, nrows: int) -> pd.DataFrame: return preview_file(file_path, nrows=nrows)

@st.cache_data(show_spinner=False)
def cached_list_lookup_files(lookup_folders_tuple: tuple) -> List[str]: return list_lookup_files(list(lookup_folders_tuple))

def parse_lookup_folders(folder_text: str) -> List[str]: return [line.strip() for line in folder_text.splitlines() if line.strip()]

def reset_loaded_metadata():
    st.session_state.metadata = DEFAULT_METADATA.copy()
    st.session_state.saved_config = None
    st.session_state.main_row_ids = [str(uuid.uuid4())]
    st.session_state.special_row_ids = [str(uuid.uuid4())]

def load_metadata(input_file: str, lookup_folders: List[str], special_lookup_file: str, preview_n: int):
    metadata = DEFAULT_METADATA.copy()
    if input_file and safe_file_exists(input_file):
        metadata["input_headers"] = cached_get_headers(input_file)
        metadata["input_preview"] = cached_preview(input_file, int(preview_n))

    main_lookup_files = cached_list_lookup_files(tuple(lookup_folders))
    metadata["main_lookup_files"] = main_lookup_files

    first_main_lookup = main_lookup_files[0] if main_lookup_files else ""
    if first_main_lookup and safe_file_exists(first_main_lookup):
        metadata["main_lookup_headers"] = cached_get_headers(first_main_lookup)
        metadata["main_lookup_preview"] = cached_preview(first_main_lookup, int(preview_n))

    if special_lookup_file and safe_file_exists(special_lookup_file):
        metadata["special_lookup_headers"] = cached_get_headers(special_lookup_file)
        metadata["special_lookup_preview"] = cached_preview(special_lookup_file, int(preview_n))

    metadata["loaded"] = True
    st.session_state.metadata = metadata

def start_processing(config: Dict):
    st.session_state.logs = []
    st.session_state.progress_state = DEFAULT_PROGRESS.copy()
    st.session_state.run_result = None
    st.session_state.is_running = True
    st.session_state.le_start_time = time.time()

    def target():
        try:
            result = run_engine(config=config, log_queue=st.session_state.logs, progress_state=st.session_state.progress_state)
            st.session_state.run_result = result
        except Exception as e:
            st.session_state.progress_state["status"] = "FAILED"
            st.session_state.progress_state["phase"] = "CRASHED"
            st.session_state.progress_state["error"] = str(e)
        finally:
            st.session_state.is_running = False

    t = threading.Thread(target=target, daemon=True)
    add_script_run_ctx(t)
    t.start()

def sync_col_name(pref: str, r_id: str):
    st.session_state[f"{pref}_new_column_{r_id}"] = st.session_state[f"{pref}_lookup_column_{r_id}"]

def build_mapping_rows_form(
    prefix: str, row_ids: List[str], available_lookup_cols: List[str], existing_output_choices: List[str], disabled: bool,
) -> Tuple[List[Dict], List[str]]:
    mappings, cols_to_delete = [], []
    if not available_lookup_cols or not row_ids: return mappings, cols_to_delete

    for i, row_id in enumerate(row_ids):
        st.markdown(f"**Row {i + 1}**")
        c1, c2, c3, c4, c5, c6 = st.columns([2, 2, 2, 1.5, 1, 0.5])
        
        if f"{prefix}_lookup_column_{row_id}" not in st.session_state:
            st.session_state[f"{prefix}_lookup_column_{row_id}"] = available_lookup_cols[0] if available_lookup_cols else ""
        if f"{prefix}_new_column_{row_id}" not in st.session_state:
            st.session_state[f"{prefix}_new_column_{row_id}"] = st.session_state[f"{prefix}_lookup_column_{row_id}"]
        if f"{prefix}_destination_type_{row_id}" not in st.session_state:
            st.session_state[f"{prefix}_destination_type_{row_id}"] = "new"

        with c1: lookup_column = st.selectbox(f"Fetch Column", options=available_lookup_cols, key=f"{prefix}_lookup_column_{row_id}", disabled=disabled, on_change=sync_col_name, args=(prefix, row_id))
        with c2: destination_type = st.selectbox(f"Action", options=["new", "existing"], format_func=lambda x: "Add to Output" if x == "new" else "Update Existing", key=f"{prefix}_destination_type_{row_id}", disabled=disabled)
        with c3:
            target_existing_column, new_column_name = "", ""
            if destination_type == "existing": target_existing_column = st.selectbox(f"Target Column", options=existing_output_choices if existing_output_choices else [""], key=f"{prefix}_target_existing_{row_id}", disabled=disabled or not existing_output_choices)
            else: new_column_name = st.text_input(f"New Name", key=f"{prefix}_new_column_{row_id}", disabled=disabled)
        with c4: write_mode = st.selectbox(f"Logic", options=["fill_if_blank", "overwrite"], index=0, key=f"{prefix}_write_mode_{row_id}", disabled=disabled)
        with c5: 
            st.write("Mandatory")
            mandatory = st.checkbox(f"Yes", key=f"{prefix}_mandatory_{row_id}", disabled=disabled)
        with c6:
            st.write("")
            if st.button("❌", key=f"{prefix}_del_{row_id}", disabled=disabled, help="Remove this mapping row"): cols_to_delete.append(row_id)
        
        mappings.append({"lookup_column": lookup_column, "destination_type": destination_type, "target_existing_column": target_existing_column, "new_column_name": new_column_name, "write_mode": write_mode, "mandatory": mandatory})
    return mappings, cols_to_delete

@st.cache_data(show_spinner=False)
def convert_df_to_csv_bytes(df: pd.DataFrame) -> bytes:
    return df.to_csv(index=False).encode("utf-8")

@st.cache_data(show_spinner=False)
def convert_df_to_excel_bytes(df: pd.DataFrame) -> bytes:
    if df.empty and len(df.columns) == 0:
        df = pd.DataFrame({"Empty": ["No data available"]})
        
    output = io.BytesIO()
    with pd.ExcelWriter(output, engine="openpyxl") as writer:
        df.to_excel(writer, index=False, sheet_name="cleaned_data")
    output.seek(0)
    return output.read()


# =========================================================
# UI SIDEBAR NAVIGATION
# =========================================================
with st.sidebar:
    st.title("🧰 High Scale Suite")
    st.markdown("Navigate through production-grade tools for scaling your data workflows.")
    st.markdown("---")
    
    app_mode = st.radio("Select Application Tool:", [
        "🔍 Lookup Engine",
        "📄 PDF Folder Merger",
        "📊 Excel/CSV Merger",
        "🔎 Excel Multi-Search",
        "🤝 Name Matcher",
        "🗜️ File Zipper",
        "⬇️ Bulk File Downloader",
        "📝 Mail Merge to PDF",
        "🧹 Data Cleaning Utility",
        "🧼 Name Cleaner & Categorizer",
        "📊 Advanced Status Profiler",
        "🗃️ Base64 to File Extractor",
        "🎓 Education Documents OCR"
    ], label_visibility="collapsed")
    
    st.markdown("---")
    st.caption("Version 14.1 | Enterprise Edition")


# =========================================================
# 🔍 LOOKUP ENGINE (DASHBOARD WORKFLOW)
# =========================================================
if app_mode == "🔍 Lookup Engine":
    st.title("🔍 High Scale Lookup Engine")
    st.markdown("*Map, join, and resolve massive datasets locally using the power of DuckDB.*")
    
    # Styled Workflow Stepper
    steps = ["1. Setup & Previews", "2. Map Columns", "3. Process Dashboard", "4. Summary & Files"]
    st.session_state.lookup_step = st.radio("Pipeline Navigation", steps, horizontal=True, label_visibility="collapsed", index=steps.index(st.session_state.lookup_step))
    st.markdown("<hr style='margin-top: 5px; margin-bottom: 20px; border: 1px solid #1E88E5;'>", unsafe_allow_html=True)

    if st.session_state.lookup_step == "1. Setup & Previews":
        col1, col2 = st.columns([1.2, 2])
        
        with col1:
            with st.container(border=True):
                st.markdown("<h4 style='color: #1565C0;'>⚙️ Configuration Paths</h4>", unsafe_allow_html=True)
                
                # raw_input = st.text_input("Input Data upload (CSV/Excel)", value=st.session_state.le_input_file, placeholder=r"D:\data\input.csv", disabled=st.session_state.is_running, on_change=reset_loaded_metadata)
                # st.session_state.le_input_file = raw_input.strip('\"').strip("\'") if raw_input else ""

                # USI JAGAH PAR YE NAYA CODE PASTE KAR DO:
                uploaded_file = st.file_uploader(
                  "Input Data upload (CSV/Excel)", 
                   type=["csv", "xlsx", "xls"], 
                   disabled=st.session_state.is_running,
                   on_change=reset_loaded_metadata
                )
                 if uploaded_file is not None:
                 temp_dir = "temp_uploads"
                 if not os.path.exists(temp_dir):
                 os.makedirs(temp_dir)
        
                temp_file_path = os.path.join(temp_dir, uploaded_file.name)
    
                with open(temp_file_path, "wb") as f:
                f.write(uploaded_file.getbuffer())
        
                st.session_state.le_input_file = temp_file_path  # The old logic found its way back here!
                else:
                st.session_state.le_input_file = ""
                
                raw_lookup = st.text_area("Primary Lookup Folders (One per line)", value=st.session_state.le_lookup_folders, height=100, disabled=st.session_state.is_running, on_change=reset_loaded_metadata)
                st.session_state.le_lookup_folders = raw_lookup.strip('\"').strip("\'") if raw_lookup else ""
                
                raw_special = st.text_input("Special Priority upload (Optional)", value=st.session_state.le_special_file, placeholder=r"D:\special_lookup.xlsx", disabled=st.session_state.is_running, on_change=reset_loaded_metadata)
                st.session_state.le_special_file = raw_special.strip('\"').strip("\'") if raw_special else ""
                
                st.markdown("<h5 style='color: #1565C0;'>Output Rules</h5>", unsafe_allow_html=True)
                raw_out = st.text_input("Save Outputs To", value=st.session_state.le_output_folder, placeholder=r"D:\output", disabled=st.session_state.is_running)
                st.session_state.le_output_folder = raw_out.strip('\"').strip("\'") if raw_out else ""
                
                c_a, c_b = st.columns(2)
                with c_a: 
                    raw_db = st.text_input("DuckDB Storage", value=st.session_state.le_db_path, disabled=st.session_state.is_running)
                    st.session_state.le_db_path = raw_db.strip('\"').strip("\'") if raw_db else ""
                with c_b: 
                    st.session_state.le_preview_n = st.number_input("Preview Rows", min_value=1, value=st.session_state.le_preview_n, step=100, disabled=st.session_state.is_running, on_change=reset_loaded_metadata)
                
                st.markdown("<br>", unsafe_allow_html=True)
                if st.button("Load Metadata & Previews", type="primary", use_container_width=True, disabled=st.session_state.is_running):
                    try:
                        load_metadata(st.session_state.le_input_file, parse_lookup_folders(st.session_state.le_lookup_folders), st.session_state.le_special_file, int(st.session_state.le_preview_n))
                        st.success("Metadata loaded successfully.")
                    except Exception as e:
                        st.error(f"Failed to load metadata: {e}")

        with col2:
            with st.container(border=True):
                st.markdown("<h4 style='color: #00838F;'>👁️ Data Profiler & Previews</h4>", unsafe_allow_html=True)
                metadata = st.session_state.metadata
                if not metadata["loaded"]:
                    st.info("Configure paths on the left and click **'Load Metadata'** to profile your data.")
                else:
                    st.markdown("**1. Input File Target**")
                    if not metadata["input_preview"].empty:
                        cols = len(metadata["input_preview"].columns)
                        st.caption(f"✅ Headers Detected: {cols} | Found: `{[c for c in metadata['input_headers']]}`")
                        st.dataframe(metadata["input_preview"], height=150, use_container_width=True)
                    
                    st.divider()
                    
                    st.markdown(f"**2. Main Database Files** *(Files Discovered: {len(metadata['main_lookup_files'])})*")
                    if not metadata["main_lookup_preview"].empty:
                        cols = len(metadata["main_lookup_preview"].columns)
                        st.caption(f"✅ Headers Detected: {cols} | Found: `{[c for c in metadata['main_lookup_headers']]}`")
                        st.dataframe(metadata["main_lookup_preview"], height=150, use_container_width=True)
                        
                    if metadata["special_lookup_headers"]:
                        st.divider()
                        st.markdown("**3. Special Database File**")
                        st.caption(f"✅ Headers Detected: `{[c for c in metadata['special_lookup_headers']]}`")
                        st.dataframe(metadata["special_lookup_preview"], height=150, use_container_width=True)

    elif st.session_state.lookup_step == "2. Map Columns":
        metadata = st.session_state.metadata
        if not metadata["loaded"]:
            st.warning("⚠️ Please complete Step 1 (Load Metadata) to generate mapping schemas.")
        else:
            with st.container(border=True):
                st.markdown("<h4 style='color: #1565C0;'>🎯 Baseline Target Rules</h4>", unsafe_allow_html=True)
                a1, a2, a3 = st.columns([1, 1, 1])
                with a1: input_key_col = st.selectbox("Input Search Key (The column to look up)", options=metadata["input_headers"], disabled=st.session_state.is_running)
                with a2: st.session_state.le_key_type = st.selectbox("Lookup Key Type (Cleaning logic)", ["text", "mobile"], index=["text", "mobile"].index(st.session_state.le_key_type), disabled=st.session_state.is_running)
                with a3: selected_input_output_cols = st.multiselect("Input Retention (Keep these columns in final output)", options=metadata["input_headers"], default=metadata["input_headers"][:min(10, len(metadata["input_headers"]))] if metadata["input_headers"] else [], disabled=st.session_state.is_running)

            st.markdown("<br>", unsafe_allow_html=True)

            with st.container(border=True):
                st.markdown("<h4 style='color: #00838F;'>🗄️ Main Database Engine Mapping</h4>", unsafe_allow_html=True)
                st.caption("Map columns from your primary lookup folders into the output data.")
                b1, _ = st.columns(2)
                with b1: main_lookup_key_col = st.selectbox("Database Match Key", options=metadata["main_lookup_headers"] if metadata["main_lookup_headers"] else [""], disabled=st.session_state.is_running or not metadata["main_lookup_headers"])
                available_main_fetch_cols = [c for c in metadata["main_lookup_headers"] if c != main_lookup_key_col] if metadata["main_lookup_headers"] else []
                
                tm1, tm2, tm3 = st.columns([1, 1, 4])
                with tm1:
                    if st.button("➕ Add Row", key="add_m"):
                        st.session_state.main_row_ids.append(str(uuid.uuid4()))
                        st.rerun()
                with tm2:
                    if st.button("⚡ Map All", key="map_m", type="secondary"):
                        st.session_state.main_row_ids = [str(uuid.uuid4()) for _ in available_main_fetch_cols]
                        for idx, col in enumerate(available_main_fetch_cols):
                            r_id = st.session_state.main_row_ids[idx]
                            st.session_state[f"main_lookup_lookup_column_{r_id}"] = col
                            st.session_state[f"main_lookup_new_column_{r_id}"] = col
                            st.session_state[f"main_lookup_destination_type_{r_id}"] = "new"
                        st.rerun()
                with tm3:
                     if st.button("🗑️ Clear", key="clr_m"):
                         st.session_state.main_row_ids = []
                         st.rerun()

                st.markdown("---")
                base_existing_choices = list(selected_input_output_cols)

                main_lookup_mappings, main_cols_to_delete = build_mapping_rows_form("main_lookup", st.session_state.main_row_ids, available_main_fetch_cols, base_existing_choices, st.session_state.is_running)

                if main_cols_to_delete:
                    for del_id in main_cols_to_delete: st.session_state.main_row_ids.remove(del_id)
                    st.rerun()

                main_cols = [m["lookup_column"] for m in main_lookup_mappings if m["lookup_column"]]
                main_dupes = set([col for col in main_cols if main_cols.count(col) > 1])
                if main_dupes: st.error(f"⚠️ **Duplicate Extraction:** `{', '.join(main_dupes)}`")

            if metadata["special_lookup_headers"]:
                st.markdown("<br>", unsafe_allow_html=True)
                with st.container(border=True):
                    st.markdown("<h4 style='color: #6A1B9A;'>⭐ Special Database Engine Mapping</h4>", unsafe_allow_html=True)
                    st.caption("Map priority override columns from your special lookup file.")
                    
                    c1, _ = st.columns(2)
                    with c1: special_lookup_key_col = st.selectbox("Special Match Key", options=metadata["special_lookup_headers"] if metadata["special_lookup_headers"] else [""], disabled=st.session_state.is_running or not metadata["special_lookup_headers"])
                    available_special_fetch_cols = [c for c in metadata["special_lookup_headers"] if c != special_lookup_key_col] if metadata["special_lookup_headers"] else []

                    ts1, ts2, ts3 = st.columns([1, 1, 4])
                    with ts1:
                        if st.button("➕ Add Row", key="add_s"):
                            st.session_state.special_row_ids.append(str(uuid.uuid4()))
                            st.rerun()
                    with ts2:
                        if st.button("⚡ Map All", key="map_s", type="secondary"):
                            st.session_state.special_row_ids = [str(uuid.uuid4()) for _ in available_special_fetch_cols]
                            for idx, col in enumerate(available_special_fetch_cols):
                                r_id = st.session_state.special_row_ids[idx]
                                st.session_state[f"special_lookup_lookup_column_{r_id}"] = col
                                st.session_state[f"special_lookup_new_column_{r_id}"] = col
                                st.session_state[f"special_lookup_destination_type_{r_id}"] = "new"
                            st.rerun()
                    with ts3:
                         if st.button("🗑️ Clear", key="clr_s"):
                             st.session_state.special_row_ids = []
                             st.rerun()
                    
                    st.markdown("---")
                    all_choices_after_main = list(dict.fromkeys(base_existing_choices + [m["new_column_name"].strip() for m in main_lookup_mappings if m["destination_type"] == "new" and m["new_column_name"].strip()]))

                    special_lookup_mappings, special_cols_to_delete = build_mapping_rows_form("special_lookup", st.session_state.special_row_ids, available_special_fetch_cols, all_choices_after_main, st.session_state.is_running)

                    if special_cols_to_delete:
                        for del_id in special_cols_to_delete: st.session_state.special_row_ids.remove(del_id)
                        st.rerun()

                    special_cols = [m["lookup_column"] for m in special_lookup_mappings if m["lookup_column"]]
                    special_dupes = set([col for col in special_cols if special_cols.count(col) > 1])
                    if special_dupes: st.error(f"⚠️ **Duplicate Extraction:** `{', '.join(special_dupes)}`")
            else:
                special_lookup_mappings, special_lookup_key_col, special_dupes, special_cols = [], "", set(), []

            disable_save = bool(main_dupes) or bool(special_dupes)
            st.markdown("---")
            if st.button("💾 Save Schema & Finalize", type="primary", disabled=disable_save, use_container_width=True):
                st.session_state.saved_config = {
                    "input_file": st.session_state.le_input_file, 
                    "lookup_folders": parse_lookup_folders(st.session_state.le_lookup_folders),
                    "special_lookup_file": st.session_state.le_special_file if st.session_state.le_special_file else None, 
                    "output_folder": st.session_state.le_output_folder,
                    "db_path": st.session_state.le_db_path, 
                    "input_key_col": input_key_col, 
                    "key_type": st.session_state.le_key_type,
                    "selected_input_output_cols": selected_input_output_cols,
                    "main_lookup_key_col": main_lookup_key_col, 
                    "main_fetch_cols": list(dict.fromkeys(main_cols)), 
                    "main_lookup_mappings": main_lookup_mappings,
                    "special_lookup_key_col": special_lookup_key_col, 
                    "special_fetch_cols": list(dict.fromkeys(special_cols)), 
                    "special_lookup_mappings": special_lookup_mappings,
                }
                st.success("Schema Locked! Proceed to Step 3 to execute the job.")

    elif st.session_state.lookup_step == "3. Process Dashboard":
        if st.session_state.saved_config is None:
            st.warning("⚠️ Please complete and save mappings in Step 2 first.")
        else:
            progress = st.session_state.progress_state
            
            with st.container(border=True):
                st.markdown("<h4 style='color: #E65100;'>🕹️ Execution Controls</h4>", unsafe_allow_html=True)
                c1, c2, c3, c4 = st.columns([1, 1, 1, 2.5])
                with c1:
                    if st.button("🚀 Start Engine", type="primary", disabled=st.session_state.is_running, use_container_width=True):
                        start_processing(st.session_state.saved_config)
                with c2:
                    st.button("🔄 Check Status", use_container_width=True)
                with c3:
                    if st.button("🗑️ Reset Engine", disabled=st.session_state.is_running, use_container_width=True):
                        st.session_state.progress_state = DEFAULT_PROGRESS.copy()
                        st.session_state.logs = []
                        st.rerun()
                with c4:
                    render_status_badge(progress["status"], progress["phase"])

            with st.container(border=True):
                st.markdown("<h4 style='color: #0277BD;'>📊 Live Telemetry Metrics</h4>", unsafe_allow_html=True)
                m1, m2, m3, m4, m5 = st.columns(5)
                m1.metric("Rows Scanned (DB)", f"{progress['lookup_rows_scanned']:,}")
                m2.metric("Files Indexed", f"{progress['lookup_files_done']} / {progress['lookup_files_total']}")
                m3.metric("Processed (Input)", f"{progress['input_rows_processed']:,}")
                
                speed = f"{progress['throughput_rps']:,} /s" if progress['status'] == "RUNNING" else "-"
                eta = format_eta(progress['eta_sec']) if progress['status'] == "RUNNING" else "-"
                
                elapsed_time = "-"
                if st.session_state.is_running and hasattr(st.session_state, 'le_start_time'):
                    elapsed_time = format_eta(int(time.time() - st.session_state.le_start_time))
                
                m4.metric("Throughput Speed", speed)
                m5.metric("Elapsed | ETA", f"{elapsed_time} | {eta}")

                st.markdown("<hr style='margin-top: 10px; margin-bottom: 10px;'>", unsafe_allow_html=True)
                st.markdown("**DuckDB Phase:** Indexing Vector Space")
                st.progress(progress["lookup_percent"] / 100 if progress["lookup_percent"] else 0)
                
                st.markdown("**Batch Phase:** Input Resolution")
                st.progress(progress["input_percent"] / 100 if progress["input_percent"] else 0)

            with st.container(border=True):
                st.markdown("<h4 style='color: #2E7D32;'>🖥️ Structured Event Logs</h4>", unsafe_allow_html=True)
                with st.container(height=300):
                    render_structured_logs(st.session_state.logs)

            if st.session_state.is_running:
                time.sleep(0.5)
                st.rerun()

    elif st.session_state.lookup_step == "4. Summary & Files":
        result = st.session_state.run_result
        if not result:
            st.info("⚠️ Execute the engine in Step 3 to generate reports.")
        else:
            with st.container(border=True):
                st.markdown("<h4 style='color: #1565C0;'>📈 Job Summary & Audit</h4>", unsafe_allow_html=True)
                input_stats = result["input_stats"]
                lookup_stats = result["lookup_stats"]

                s1, s2, s3, s4 = st.columns(4)
                s1.metric("🎯 Matched Rows", f"{input_stats['matched_rows']:,}")
                s2.metric("⚠️ Unmatched Rows", f"{input_stats['unmatched_rows']:,}")
                s3.metric("📝 Total Processed", f"{input_stats['processed_rows']:,}")
                s4.metric("🗄️ DuckDB Indexed", f"{lookup_stats['unique_lookup_keys_indexed']:,}")

            with st.container(border=True):
                st.markdown("<h4 style='color: #00838F;'>📁 Artifacts Generated</h4>", unsafe_allow_html=True)
                output_files = [result["matched_output_path"], result["unmatched_output_path"], result["consolidated_output_path"], result["file_level_summary_path"]]
                output_df = pd.DataFrame({
                    "System Path": output_files,
                    "Size (MB)": [round(os.path.getsize(p) / (1024 * 1024), 2) if os.path.exists(p) else 0 for p in output_files],
                    "Status": ["✅ Available" if os.path.exists(p) else "❌ Missing" for p in output_files]
                })
                st.dataframe(output_df, use_container_width=True)


# =========================================================
# STANDALONE TOOLS (DASHBOARD UI)
# =========================================================
else:
    tool_map = {
        "📄 PDF Folder Merger": ("pdf", "Combine PDFs and Images recursively using strict folder mapping."),
        "📊 Excel/CSV Merger": ("em", "Smart-join multiple spreadsheets into a single unified master file."),
        "🔎 Excel Multi-Search": ("es", "Bulk search thousands of rows across workbooks for exact/alphanumeric hits."),
        "🤝 Name Matcher": ("nm", "RapidFuzz-powered Jaro-Winkler analysis to score and flag fuzzy name matches."),
        "🗜️ File Zipper": ("zip", "Compress massive directories into MB-capped batch zip files automatically."),
        "⬇️ Bulk File Downloader": ("dl", "Extract and locally mirror direct/Drive links listed in an Excel manifest."),
        "📝 Mail Merge to PDF": ("mm", "Automate Microsoft Word COM object generation to build and export merge PDFs."),
        "🧹 Data Cleaning Utility": ("dc", "Stream, Clean, Filter, and Export massive datasets using memory-safe batch processing."),
        "🧼 Name Cleaner & Categorizer": ("nc", "Clean, deduplicate, and categorize names based on address/business heuristics."),
        "📊 Advanced Status Profiler": ("sp", "Safely read giant CSVs to profile column values and calculate duplicates."),
        "🗃️ Base64 to File Extractor": ("b64", "Scan folders for .txt files containing Base64 strings and decode them back into binary files (PDF/Images)."),
        "🎓 Education Documents OCR": ("ocr", "Extract fields from certificates/marksheets using docTR and Ollama AI.")
    }
    
    t_name = app_mode.split(" ", 1)[1]
    t_key, t_desc = tool_map[app_mode]
    status = st.session_state[f"{t_key}_status"]
    
    st.title(f"{app_mode.split(' ')[0]} {t_name}")
    st.markdown(f"*{t_desc}*")
    st.markdown("<hr style='border: 1px solid #1E88E5;'>", unsafe_allow_html=True)
    
    can_run = True

    # --- Configuration Card ---
    with st.container(border=True):
        st.markdown("<h4 style='color: #1565C0;'>⚙️ Configuration</h4>", unsafe_allow_html=True)
        
        if t_key == "pdf":
            raw_val = st.text_input("Main Target Directory", disabled=(status == "RUNNING"))
            folder = raw_val.strip('\"').strip("\'") if raw_val else ""
            def run_task(): return run_pdf_merger(folder, st.session_state.pdf_logs)
            
        elif t_key == "em":
            raw_val = st.text_input("Source Directory", disabled=(status == "RUNNING"))
            folder = raw_val.strip('\"').strip("\'") if raw_val else ""
            
            out_name = st.text_input("Output File Name", value="master_merge_output.csv", disabled=(status == "RUNNING"))
            def run_task(): return run_excel_csv_merger(folder, out_name, st.session_state.em_logs)
            
        elif t_key == "es":
            raw_val = st.text_input("Target Directory to Scan", disabled=(status == "RUNNING"))
            folder = raw_val.strip('\"').strip("\'") if raw_val else ""
            
            terms = st.text_area("Search Manifest (One per line)", height=150, placeholder="8884064564\njohn.doe@email.com\nINV-4021", disabled=(status == "RUNNING"))
            def run_task(): return run_excel_search(folder, [t.strip() for t in terms.splitlines() if t.strip()], st.session_state.es_logs)

        elif t_key == "nm":
            raw_val = st.text_input("Dataset Path (Excel/CSV)", disabled=(status == "RUNNING"))
            file_path = raw_val.strip('\"').strip("\'") if raw_val else ""
            
            if st.button("Read Columns", disabled=(status == "RUNNING" or not file_path)):
                if os.path.isfile(file_path):
                    try:
                        st.session_state.nm_headers = cached_get_headers(file_path)
                        st.success("Schema captured successfully.")
                    except Exception as e: st.error(f"Schema load failed: {e}")
                else: st.error("File not found on system.")
                    
            if st.session_state.nm_headers:
                nc1, nc2, nc3 = st.columns(3)
                with nc1: col1 = st.selectbox("Primary Name Column", options=st.session_state.nm_headers)
                with nc2: col2 = st.selectbox("Secondary Name Column", options=st.session_state.nm_headers)
                out_options = ["-- Auto Append to End --"] + st.session_state.nm_headers
                with nc3: out_col = st.selectbox("Score Output Target", options=out_options)

                def run_task(): 
                    selected_out = None if out_col == "-- Auto Append to End --" else out_col
                    return run_name_match(file_path, col1, col2, selected_out, st.session_state.nm_logs)
            else:
                can_run = False
                st.info("Input a path and click **'Read Columns'** to configure analysis metrics.")
                def run_task(): return False

        elif t_key == "zip":
            raw_val = st.text_input("Directory to Compress", disabled=(status == "RUNNING"))
            folder = raw_val.strip('\"').strip("\'") if raw_val else ""
            
            max_mb = st.number_input("Maximum Batch Size (MB)", min_value=1.0, value=30.0, step=5.0, disabled=(status == "RUNNING"))
            def run_task(): return run_file_zipper(folder, max_mb, st.session_state.zip_logs)

        elif t_key == "dl":
            raw_val = st.text_input("URL Manifest Path (Excel/CSV)", disabled=(status == "RUNNING"))
            file_path = raw_val.strip('\"').strip("\'") if raw_val else ""
            
            if st.button("Read Columns", disabled=(status == "RUNNING" or not file_path)):
                if os.path.isfile(file_path):
                    try:
                        st.session_state.dl_headers = cached_get_headers(file_path)
                        st.success("Schema captured successfully.")
                    except Exception as e: st.error(f"Schema load failed: {e}")
                else: st.error("File not found on system.")
                    
            if st.session_state.dl_headers:
                dc1, dc2 = st.columns(2)
                with dc1: url_col = st.selectbox("Target URL Column", options=st.session_state.dl_headers)
                with dc2: rename_col = st.selectbox("Target Save-Name Column", options=st.session_state.dl_headers)
                
                raw_out = st.text_input("Local Download Directory", placeholder=r"C:\downloads", disabled=(status == "RUNNING"))
                out_dir = raw_out.strip('\"').strip("\'") if raw_out else ""
                
                file_source_type = st.radio("Server Architecture", options=["Auto-Detect", "Google Drive", "Direct Link / S3"], horizontal=True, disabled=(status == "RUNNING"))

                def run_task(): 
                    if not out_dir: raise ValueError("Download Directory cannot be empty.")
                    return run_file_downloader(file_path, url_col, rename_col, out_dir, file_source_type, st.session_state.dl_logs)
            else:
                can_run = False
                st.info("Input a path and click **'Read Columns'** to configure the mirror engine.")
                def run_task(): return False

        elif t_key == "mm":
            raw_1 = st.text_input("Excel Database Path", disabled=(status == "RUNNING"))
            mm_excel = raw_1.strip('\"').strip("\'") if raw_1 else ""
            
            raw_2 = st.text_input("Word Template (Macro) Path", disabled=(status == "RUNNING"))
            mm_word = raw_2.strip('\"').strip("\'") if raw_2 else ""
            
            raw_3 = st.text_input("Compiled PDF Save Directory", disabled=(status == "RUNNING"))
            mm_output = raw_3.strip('\"').strip("\'") if raw_3 else ""

            st.markdown("##### Initiation Timestamps")
            col_d, col_t = st.columns(2)
            with col_d: mm_date = st.date_input("Start Date", disabled=(status == "RUNNING"))
            with col_t: mm_time = st.time_input("Start Time", disabled=(status == "RUNNING"))

            if not mm_excel or not mm_word or not mm_output:
                can_run = False
                st.info("Ensure all three core paths are provided to unlock execution.")

            def run_task():
                start_datetime = datetime.combine(mm_date, mm_time)
                return run_mail_merge_tool(mm_excel, mm_word, mm_output, start_datetime, st.session_state.mm_logs)

        elif t_key == "nc":
            raw_val = st.text_input("Dataset Path (CSV/Excel)", disabled=(status == "RUNNING"))
            file_path = raw_val.strip('\"').strip("\'") if raw_val else ""
            
            if st.button("Read Columns", disabled=(status == "RUNNING" or not file_path)):
                if os.path.isfile(file_path):
                    try:
                        st.session_state.nc_headers = cached_get_headers(file_path)
                        st.success("Schema captured successfully.")
                    except Exception as e: st.error(f"Schema load failed: {e}")
                else: st.error("File not found on system.")
                    
            if st.session_state.nc_headers:
                c1, c2, c3 = st.columns(3)
                with c1: col_mob = st.selectbox("Mobile Column", options=st.session_state.nc_headers)
                with c2: col_name = st.selectbox("Name/Response Column", options=st.session_state.nc_headers)
                with c3: col_status = st.selectbox("Match Status Column", options=st.session_state.nc_headers)
                
                raw_out = st.text_input("Output Directory", placeholder=r"C:\cleaned_output", disabled=(status == "RUNNING"))
                out_dir = raw_out.strip('\"').strip("\'") if raw_out else ""

                def run_task(): 
                    if not out_dir: raise ValueError("Output Directory cannot be empty.")
                    return run_name_cleaner(file_path, col_mob, col_name, col_status, out_dir, st.session_state.nc_logs)
            else:
                can_run = False
                st.info("Input a path and click **'Read Columns'** to configure cleaning mapping.")
                def run_task(): return False

        elif t_key == "dc":
            col_f1, col_f2 = st.columns(2)
            with col_f1: 
                raw_in = st.text_input("Input Dataset Path (CSV/Excel)", disabled=(status == "RUNNING"))
                dc_input_file = raw_in.strip('\"').strip("\'") if raw_in else ""
            with col_f2: 
                raw_out = st.text_input("Output Directory for Cleaned File", disabled=(status == "RUNNING"))
                dc_output_folder = raw_out.strip('\"').strip("\'") if raw_out else ""

            if st.button("Read Schema & Generate Preview", disabled=(status == "RUNNING" or not dc_input_file)):
                if os.path.isfile(dc_input_file):
                    try:
                        st.session_state.dc_preview_df = cached_preview(dc_input_file, 500)
                        st.session_state.dc_headers = list(st.session_state.dc_preview_df.columns)
                        st.success("Schema captured successfully.")
                    except Exception as e: st.error(f"Schema load failed: {e}")
                else: st.error("File not found on system.")

            if not st.session_state.dc_headers:
                can_run = False
                st.info("👆 Input paths and generate a preview to unlock cleaning rules.")
                def run_task(): return False
            else:
                preview_df = st.session_state.dc_preview_df
                
                with st.sidebar:
                    st.markdown("---")
                    st.markdown("### 🧹 Cleaning Actions")
                    
                    st.subheader("1. Rename Headers")
                    rename_map = {}
                    rename_cols = st.multiselect("Select columns to rename", options=st.session_state.dc_headers, disabled=(status == "RUNNING"))
                    for col in rename_cols:
                        rename_map[col] = st.text_input(f"New name for '{col}'", value=col, key=f"rename_{col}", disabled=(status == "RUNNING"))
                    
                    temp_columns_after_rename = [rename_map.get(c, c) for c in st.session_state.dc_headers]
                    
                    st.subheader("2. Replace Values")
                    num_replace_rules = st.number_input("How many replacement rules?", min_value=0, max_value=50, value=0, step=1, disabled=(status == "RUNNING"))
                    replacement_rules = []
                    for i in range(num_replace_rules):
                        with st.expander(f"Rule {i+1}", expanded=True):
                            col = st.selectbox(f"Column #{i+1}", options=temp_columns_after_rename, key=f"replace_col_{i}", disabled=(status == "RUNNING"))
                            old_value = st.text_input(f"Old value #{i+1}", key=f"old_val_{i}", disabled=(status == "RUNNING"))
                            new_value = st.text_input(f"New value #{i+1}", key=f"new_val_{i}", disabled=(status == "RUNNING"))
                            match_type = st.selectbox(f"Match type #{i+1}", options=["exact", "contains"], key=f"match_type_{i}", disabled=(status == "RUNNING"))
                            case_sensitive = st.checkbox(f"Case sensitive #{i+1}", value=False, key=f"case_sensitive_{i}", disabled=(status == "RUNNING"))
                            if old_value != "":
                                replacement_rules.append({"column": col, "old_value": old_value, "new_value": new_value, "match_type": match_type, "case_sensitive": case_sensitive})

                    st.subheader("3. Add Columns")
                    num_add_rules = st.number_input("How many columns to add?", min_value=0, max_value=30, value=0, step=1, disabled=(status == "RUNNING"))
                    add_rules = []
                    for i in range(num_add_rules):
                        with st.expander(f"Add Column Rule {i+1}", expanded=True):
                            new_col = st.text_input(f"New column name #{i+1}", key=f"new_col_{i}", disabled=(status == "RUNNING"))
                            logic_type = st.selectbox(f"Logic type #{i+1}", options=["constant", "from_existing_column", "concat", "math", "conditional_if_else"], key=f"logic_type_{i}", disabled=(status == "RUNNING"))
                            rule = {"new_column": new_col, "logic_type": logic_type}
                            
                            if logic_type == "constant":
                                rule["value"] = st.text_input(f"Constant value #{i+1}", key=f"constant_{i}", disabled=(status == "RUNNING"))
                            elif logic_type == "from_existing_column":
                                rule["source_column"] = st.selectbox(f"Source column #{i+1}", options=temp_columns_after_rename, key=f"src_col_{i}", disabled=(status == "RUNNING"))
                            elif logic_type == "concat":
                                rule["source_columns"] = st.multiselect(f"Columns to concat #{i+1}", options=temp_columns_after_rename, key=f"concat_cols_{i}", disabled=(status == "RUNNING"))
                                rule["separator"] = st.text_input(f"Separator #{i+1}", value=" ", key=f"separator_{i}", disabled=(status == "RUNNING"))
                            elif logic_type == "math":
                                rule["left_column"] = st.selectbox(f"Left column #{i+1}", options=temp_columns_after_rename, key=f"left_col_{i}", disabled=(status == "RUNNING"))
                                rule["right_column"] = st.selectbox(f"Right column #{i+1}", options=temp_columns_after_rename, key=f"right_col_{i}", disabled=(status == "RUNNING"))
                                rule["operation"] = st.selectbox(f"Operation #{i+1}", options=["add", "subtract", "multiply", "divide"], key=f"math_op_{i}", disabled=(status == "RUNNING"))
                            elif logic_type == "conditional_if_else":
                                rule["source_column"] = st.selectbox(f"Condition source column #{i+1}", options=temp_columns_after_rename, key=f"cond_src_{i}", disabled=(status == "RUNNING"))
                                rule["operator"] = st.selectbox(f"Operator #{i+1}", options=["==", "!=", ">", "<", ">=", "<=", "contains"], key=f"cond_op_{i}", disabled=(status == "RUNNING"))
                                rule["compare_value"] = st.text_input(f"Compare value #{i+1}", key=f"compare_value_{i}", disabled=(status == "RUNNING"))
                                rule["true_value"] = st.text_input(f"True value #{i+1}", key=f"true_value_{i}", disabled=(status == "RUNNING"))
                                rule["false_value"] = st.text_input(f"False value #{i+1}", key=f"false_value_{i}", disabled=(status == "RUNNING"))
                            if new_col.strip() != "": add_rules.append(rule)

                    st.subheader("4. Remove Columns")
                    cols_to_drop = st.multiselect("Select columns to drop entirely", options=temp_columns_after_rename, disabled=(status == "RUNNING"))

                    st.subheader("5. Filter Rows")
                    num_filter_rules = st.number_input("How many filter rules?", min_value=0, max_value=20, value=0, step=1, disabled=(status == "RUNNING"))
                    filter_rules = []
                    
                    temp_available_cols = [c for c in temp_columns_after_rename if c not in cols_to_drop]
                    for r in add_rules:
                        if r.get("new_column"): temp_available_cols.append(r["new_column"])

                    for i in range(num_filter_rules):
                        with st.expander(f"Filter Rule {i+1}", expanded=True):
                            f_col = st.selectbox(f"Column to filter #{i+1}", options=temp_available_cols, key=f"f_col_{i}", disabled=(status == "RUNNING"))
                            f_op = st.selectbox(f"Condition #{i+1}", options=["==", "!=", ">", "<", ">=", "<=", "contains", "is_null", "not_null"], key=f"f_op_{i}", disabled=(status == "RUNNING"))
                            f_val = ""
                            if f_op not in ["is_null", "not_null"]:
                                f_val = st.text_input(f"Value #{i+1}", key=f"f_val_{i}", disabled=(status == "RUNNING"))
                            filter_rules.append({"column": f_col, "operator": f_op, "value": f_val})

                    st.subheader("6. Select Final Columns")
                    final_columns = st.multiselect("Select and reorder output columns", options=temp_available_cols, default=temp_available_cols, disabled=(status == "RUNNING"))

                # In-Memory Preview Generation
                preview_cleaned = apply_all_transformations(
                    preview_df, rename_map=rename_map, replacement_rules=replacement_rules,
                    cols_to_drop=cols_to_drop, add_rules=add_rules, filter_rules=filter_rules, final_columns=final_columns
                )

                st.markdown("#### 👁️ Live Data Preview (First 500 rows)")
                st.dataframe(preview_cleaned, use_container_width=True, height=300)
                
                with st.expander("Show Transformation Config JSON"):
                    config_json = {
                        "rename_map": rename_map, "replacement_rules": replacement_rules,
                        "cols_to_drop": cols_to_drop, "add_rules": add_rules,
                        "filter_rules": filter_rules, "final_columns": final_columns
                    }
                    st.code(json.dumps(config_json, indent=2), language="json")

                def run_task():
                    if not dc_input_file or not dc_output_folder: raise ValueError("Both Input File and Output Folder paths are required.")
                    config = {
                        "input_file": dc_input_file,
                        "output_folder": dc_output_folder,
                        "rename_map": rename_map,
                        "replacement_rules": replacement_rules,
                        "cols_to_drop": cols_to_drop,
                        "add_rules": add_rules,
                        "filter_rules": filter_rules,
                        "final_columns": final_columns
                    }
                    return run_data_cleaner(config, st.session_state.dc_logs, st.session_state.progress_state)

        elif t_key == "sp":
            raw_val = st.text_input("Target File Path (CSV/Excel)", disabled=(status == "RUNNING"))
            file_path = raw_val.strip('\"').strip("\'") if raw_val else ""
            
            if st.button("Read Columns", disabled=(status == "RUNNING" or not file_path)):
                if os.path.isfile(file_path):
                    try:
                        st.session_state.sp_headers = cached_get_headers(file_path)
                        st.success("Schema captured successfully.")
                    except Exception as e: st.error(f"Schema load failed: {e}")
                else: st.error("File not found on system.")
            
            if st.session_state.sp_headers:
                c1, c2 = st.columns(2)
                with c1: 
                    profile_col = st.selectbox("Column to Profile (e.g., Status)", options=st.session_state.sp_headers)
                with c2: 
                    dedupe_opts = ["-- Skip Duplicate Check --", "-- Entire Row (All Columns) --"] + st.session_state.sp_headers
                    dedupe_col = st.selectbox("Check Duplicates Using", options=dedupe_opts)

                def run_task(): 
                    st.session_state.progress_state["final_total"] = 0
                    st.session_state.progress_state["value_counts"] = {}
                    st.session_state.progress_state["unique_count"] = 0
                    st.session_state.progress_state["duplicate_count"] = 0
                    return run_advanced_profiler(file_path, profile_col, dedupe_col, st.session_state.sp_logs, st.session_state.progress_state)
            else:
                can_run = False
                st.info("Input a file path and click **'Read Columns'** to configure profiling.")
                def run_task(): return False

        elif t_key == "b64":
            col_b1, col_b2 = st.columns(2)
            with col_b1:
                raw_in = st.text_input("Master Directory (Containing .txt files)", disabled=(status == "RUNNING"))
                master_dir = raw_in.strip('\"').strip("\'") if raw_in else ""
            with col_b2:
                raw_out = st.text_input("Output Directory for Decoded Files", disabled=(status == "RUNNING"))
                out_dir = raw_out.strip('\"').strip("\'") if raw_out else ""
                
            if not master_dir or not out_dir:
                can_run = False
                st.info("Input both Master and Output directory paths to execute decoding.")
                
            def run_task(): 
                return run_base64_extractor(master_dir, out_dir, st.session_state.b64_logs, st.session_state.progress_state)

        elif t_key == "ocr":
            col_o1, col_o2 = st.columns(2)
            with col_o1:
                raw_in = st.text_input("Target Folder (Contains PDFs/Images)", disabled=(status == "RUNNING"))
                folder_path = raw_in.strip('\"').strip("\'") if raw_in else ""
            with col_o2:
                enable_llm = st.checkbox("Enable Ollama LLM Fallback", value=False, disabled=(status == "RUNNING"), help="Must have Ollama running locally.")
                ollama_model = st.text_input("Ollama Model Name", value="llama3.1", disabled=(status == "RUNNING" or not enable_llm))

            c_o1, c_o2, c_o3 = st.columns(3)
            with c_o1: preprocess = st.checkbox("Enable OpenCV Preprocessing", value=False, disabled=(status == "RUNNING"))
            with c_o2: max_side = st.number_input("Max Image Size", value=1600, disabled=(status == "RUNNING"))
            with c_o3: keep_raw_ocr = st.checkbox("Save Raw OCR Text", value=False, disabled=(status == "RUNNING"))

            if not folder_path:
                can_run = False
                st.info("Provide a target directory to begin processing.")

            def run_task():
                st.session_state.progress_state["input_rows_processed"] = 0
                st.session_state.progress_state["input_percent"] = 0
                st.session_state.progress_state["throughput_rps"] = 0
                st.session_state.progress_state["eta_sec"] = 0
                st.session_state.progress_state["review_count"] = 0
                return run_education_ocr(
                    folder_path, enable_llm, ollama_model, preprocess, max_side, keep_raw_ocr, 
                    st.session_state.ocr_logs, st.session_state.progress_state
                )

    # --- Control & Dash Card ---
    with st.container(border=True):
        st.markdown("<h4 style='color: #E65100;'>🕹️ Control Panel</h4>", unsafe_allow_html=True)
        c1, c2, c3, c4 = st.columns([1, 1, 1, 2.5])
        with c1:
            if st.button("🚀 Execute Job", type="primary", disabled=(status in ["RUNNING", "COMPLETED"] or not can_run), use_container_width=True):
                st.session_state[f"{t_key}_logs"] = []
                st.session_state[f"{t_key}_status"] = "RUNNING"
                st.session_state[f"{t_key}_start_time"] = time.time()
                
                def thread_target():
                    try:
                        success = run_task()
                        st.session_state[f"{t_key}_status"] = "COMPLETED" if success else "FAILED"
                    except Exception as e:
                        st.session_state[f"{t_key}_logs"].append({"timestamp": time.strftime("%H:%M:%S"), "level": "ERROR", "message": f"Exception: {str(e)}\n\n{traceback.format_exc()}"})
                        st.session_state[f"{t_key}_status"] = "FAILED"

                t = threading.Thread(target=thread_target, daemon=True)
                add_script_run_ctx(t)
                t.start()
        with c2:
            st.button("🔄 Check Status", key=f"{t_key}_check", use_container_width=True)
        with c3:
            if st.button("🗑️ Discard Run", disabled=(status == "RUNNING"), use_container_width=True):
                st.session_state[f"{t_key}_status"] = "IDLE"
                st.session_state[f"{t_key}_logs"] = []
                st.session_state[f"{t_key}_start_time"] = None
                if t_key == "nm": st.session_state.nm_headers = []
                if t_key == "dl": st.session_state.dl_headers = []
                if t_key == "nc": st.session_state.nc_headers = []
                if t_key == "dc": st.session_state.dc_headers = []
                if t_key == "sp": st.session_state.sp_headers = []
                st.rerun()
        with c4:
            phase_text = ""
            if status == "RUNNING" and st.session_state[f"{t_key}_start_time"] is not None:
                elapsed = int(time.time() - st.session_state[f"{t_key}_start_time"])
                phase_text = f"(Elapsed: {format_eta(elapsed)})"
            render_status_badge(status, phase_text)

    # --- Telemetry Logs ---
    if t_key == "sp" and status in ["COMPLETED", "RUNNING"]:
        with st.container(border=True):
            st.markdown("<h4 style='color: #0277BD;'>📊 Data Profiling Results</h4>", unsafe_allow_html=True)
            pr = st.session_state.progress_state
            
            t_count = pr.get('final_total', 0)
            v_counts = pr.get('value_counts', {})
            u_count = pr.get('unique_count', 0)
            d_count = pr.get('duplicate_count', 0)
            
            if status == "RUNNING":
                st.info("Streaming file blocks into memory...")
                st.metric("Rows Scanned So Far", f"{pr.get('input_rows_processed', 0):,}")
            else:
                m1, m2, m3 = st.columns(3)
                m1.metric("Total Data Rows", f"{t_count:,}")
                if u_count > 0 or d_count > 0:
                    m2.metric("Unique Entries", f"{u_count:,}")
                    m3.metric("Duplicate Entries", f"{d_count:,}")
                
                if v_counts:
                    st.markdown("##### 📌 Column Value Breakdown")
                    df_counts = pd.DataFrame(list(v_counts.items()), columns=["Value", "Count"]).sort_values(by="Count", ascending=False)
                    df_counts["% of Total"] = (df_counts["Count"] / t_count * 100).round(2).astype(str) + "%" if t_count > 0 else "0%"
                    st.dataframe(df_counts, use_container_width=True)

    elif t_key == "dc" and status in ["RUNNING", "COMPLETED", "FAILED"]:
        with st.container(border=True):
            st.markdown("<h4 style='color: #0277BD;'>📊 Live Telemetry Metrics</h4>", unsafe_allow_html=True)
            pr = st.session_state.progress_state
            m1, m2, m3, m4 = st.columns(4)
            m1.metric("Rows Read", f"{pr.get('input_rows_processed', 0):,}")
            m2.metric("Throughput Speed", f"{pr.get('throughput_rps', 0):,} rps" if status == "RUNNING" else "-")
            m3.metric("ETA", format_eta(pr.get('eta_sec', 0)) if status == "RUNNING" else "-")
            
            if status == "COMPLETED":
                m4.metric("Total Rows Exported", f"{pr.get('final_total', 0):,}")
            else:
                m4.metric("Rows Exported", "Calculating...")
            
            st.progress(pr.get("input_percent", 0) / 100 if pr.get("input_percent") else 0)
            
    elif t_key == "ocr" and status in ["RUNNING", "COMPLETED", "FAILED"]:
        with st.container(border=True):
            st.markdown("<h4 style='color: #0277BD;'>📊 OCR Telemetry</h4>", unsafe_allow_html=True)
            pr = st.session_state.progress_state
            m1, m2, m3, m4 = st.columns(4)
            m1.metric("Files Processed", f"{pr.get('input_rows_processed', 0):,}")
            m2.metric("Processing Speed", f"{pr.get('throughput_rps', 0)} sec/file" if status == "RUNNING" else "-")
            m3.metric("ETA", format_eta(pr.get('eta_sec', 0)) if status == "RUNNING" else "-")
            m4.metric("Reviews Required", f"{pr.get('review_count', 0):,}")
            st.progress(pr.get("input_percent", 0) / 100 if pr.get("input_percent") else 0)

    with st.container(border=True):
        st.markdown("<h4 style='color: #2E7D32;'>🖥️ Operation Telemetry</h4>", unsafe_allow_html=True)
        with st.container(height=350):
            render_structured_logs(st.session_state[f"{t_key}_logs"])

    if status == "RUNNING":
        time.sleep(0.5)
        st.rerun()
