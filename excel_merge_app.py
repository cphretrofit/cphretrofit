import streamlit as st
import pandas as pd
import io
import os
import re

st.set_page_config(page_title="Excel Merge Tool", page_icon="📊")

def safe_str(val):
    return "" if pd.isna(val) else str(val).strip()

def excel_col_to_index(col):
    col = col.upper()
    idx = 0
    for c in col:
        idx = idx * 26 + (ord(c) - 64)
    return idx - 1

def extract_uprn(notes_value):
    """Extract UPRN number from a Notes cell. Returns empty string if not found."""
    if pd.isna(notes_value):
        return ""
    match = re.search(r'UPRN:\s*(\d+)', str(notes_value))
    return match.group(1) if match else ""

st.title("Excel Merge Tool")

uploaded_file = st.file_uploader("Load Excel or CSV File", type=["xlsx", "csv"])
output_dir = st.text_input("Output folder path (or leave blank to download)", value="")

# Auto-generate output filename based on uploaded file
if uploaded_file:
    original_name = uploaded_file.name
    default_output_name = f"Merged - {original_name}"
else:
    default_output_name = "merged_output.xlsx"

output_name = st.text_input("Output file name", value=default_output_name)

st.subheader("Merge Rules")

if "rules" not in st.session_state:
    st.session_state.rules = []

col1, col2, col3, col4 = st.columns([2, 2, 1, 1])
with col1:
    st.caption("Output column name")
with col2:
    st.caption("Columns to merge (e.g., B,C,D)")
with col3:
    st.caption("Separator")
with col4:
    st.caption("")

for i, rule in enumerate(st.session_state.rules):
    col1, col2, col3, col4 = st.columns([2, 2, 1, 1])
    with col1:
        rule["out_name"] = st.text_input("Output column name", value=rule.get("out_name", ""), key=f"name_{i}")
    with col2:
        rule["cols"] = st.text_input("Columns to merge", value=rule.get("cols", "B,C,D"), key=f"cols_{i}")
    with col3:
        rule["sep"] = st.text_input("Separator", value=rule.get("sep", ", "), key=f"sep_{i}")
    with col4:
        if st.button("Remove", key=f"del_{i}"):
            st.session_state.rules.pop(i)
            st.rerun()

if st.button("Add Merge Rule"):
    st.session_state.rules.append({"out_name": "Merged Column", "cols": "B,C,D", "sep": ", "})
    st.rerun()

# ── UPRN Extraction ──────────────────────────────────────────────────────────
st.subheader("UPRN Extraction")
extract_uprn_enabled = st.checkbox("Extract UPRN Number from Notes column")

raw_file = None
uprn_notes_col = "L"
uprn_address_col = "E"
uprn_insert_after = ""

if extract_uprn_enabled:
    st.markdown(
        "Extracts the UPRN number from the Notes column and adds it as a "
        "**UPRN Number** column. Everything else in Notes is discarded."
    )
    raw_file = st.file_uploader(
        "Upload raw file for UPRN lookup (if UPRNs are in a different file)",
        type=["xlsx", "csv"],
        key="raw_uprn_file",
    )
    if raw_file:
        st.caption(
            "The tool will match addresses between the raw file and your main "
            "file to look up each UPRN."
        )
    ucol1, ucol2 = st.columns(2)
    with ucol1:
        uprn_notes_col = st.text_input(
            "Notes column letter (in raw/main file)", value="L"
        )
    with ucol2:
        uprn_address_col = st.text_input(
            "Address column letter (in raw file, for lookup matching)", value="E"
        )
    uprn_insert_after = st.text_input(
        "Insert UPRN Number after this output column name (leave blank for first column)",
        value="",
    )

if st.button("Run Merge"):
    if not uploaded_file:
        st.error("Please load an Excel file")
    elif not st.session_state.rules:
        st.error("Please add at least one merge rule")
    else:
        if uploaded_file.name.endswith(".csv"):
            df = pd.read_csv(uploaded_file, dtype=str)
        else:
            df = pd.read_excel(uploaded_file, dtype=str)
        
        # The dataframe preserves the row order from the uploaded Excel file
        # No sorting needed - we respect the existing order
        
        # Perform merging
        out_df = pd.DataFrame()
        for rule in st.session_state.rules:
            col_indexes = [
                excel_col_to_index(c.strip())
                for c in rule["cols"].split(",")
                if c.strip()
            ]
            if not col_indexes:
                continue
            
            # Check if all column indexes are valid
            valid_indexes = [idx for idx in col_indexes if idx < len(df.columns)]
            if not valid_indexes:
                st.warning(f"No valid columns found for rule: {rule['out_name']}")
                continue
            
            merged = df.iloc[:, valid_indexes[0]].apply(safe_str)
            for idx in valid_indexes[1:]:
                merged += rule["sep"] + df.iloc[:, idx].apply(safe_str)
            out_df[rule["out_name"]] = merged.str.strip(rule["sep"])
        
        # ── UPRN Extraction Step ─────────────────────────────────────────
        if extract_uprn_enabled:
            notes_idx = excel_col_to_index(uprn_notes_col.strip())

            if raw_file:
                # Lookup mode: build UPRN map from a separate raw file
                raw_file.seek(0)
                if raw_file.name.endswith(".csv"):
                    raw_df = pd.read_csv(raw_file, dtype=str)
                else:
                    raw_df = pd.read_excel(raw_file, dtype=str)

                addr_idx = excel_col_to_index(uprn_address_col.strip())

                # Build map: address -> UPRN (from rows that have a Notes value)
                uprn_map = {}
                if notes_idx < len(raw_df.columns) and addr_idx < len(raw_df.columns):
                    for _, row in raw_df.iterrows():
                        addr = safe_str(row.iloc[addr_idx])
                        uprn_val = extract_uprn(row.iloc[notes_idx])
                        if addr and uprn_val:
                            uprn_map[addr] = uprn_val

                    st.info(f"Built UPRN lookup with {len(uprn_map)} entries from raw file")

                    # Match against the output address column
                    if uprn_insert_after and uprn_insert_after in out_df.columns:
                        match_col = uprn_insert_after
                    elif len(out_df.columns) > 0:
                        match_col = out_df.columns[0]
                    else:
                        match_col = None

                    if match_col:
                        out_df["UPRN Number"] = out_df[match_col].apply(
                            lambda x: uprn_map.get(safe_str(x), "")
                        )
                    else:
                        out_df["UPRN Number"] = ""
                else:
                    st.warning("Notes or Address column letter is out of range in raw file")
                    out_df["UPRN Number"] = ""
            else:
                # Direct mode: extract UPRN from the Notes column of the main file
                if notes_idx < len(df.columns):
                    out_df["UPRN Number"] = df.iloc[:, notes_idx].apply(extract_uprn)
                else:
                    st.warning(f"Column {uprn_notes_col} is out of range")
                    out_df["UPRN Number"] = ""

            # Reorder: place UPRN Number after the chosen column
            cols = list(out_df.columns)
            cols.remove("UPRN Number")
            if uprn_insert_after and uprn_insert_after in cols:
                insert_pos = cols.index(uprn_insert_after) + 1
            else:
                insert_pos = 0
            cols.insert(insert_pos, "UPRN Number")
            out_df = out_df[cols]

            matched = (out_df["UPRN Number"] != "").sum()
            st.success(f"UPRN extraction: {matched} of {len(out_df)} rows matched")
        
        # Save or download
        if output_dir:
            out_path = os.path.join(output_dir, output_name)
            out_df.to_excel(out_path, index=False)
            st.success(f"File saved to: {out_path}")
        else:
            buffer = io.BytesIO()
            out_df.to_excel(buffer, index=False)
            buffer.seek(0)
            st.download_button(
                label="Download Merged Excel",
                data=buffer,
                file_name=output_name,
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )
