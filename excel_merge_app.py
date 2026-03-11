import streamlit as st
import pandas as pd
import io
import os

st.set_page_config(page_title="Excel Merge Tool", page_icon="📊")

def safe_str(val):
    return "" if pd.isna(val) else str(val).strip()

def excel_col_to_index(col):
    col = col.upper()
    idx = 0
    for c in col:
        idx = idx * 26 + (ord(c) - 64)
    return idx - 1

st.title("Excel Merge Tool")

uploaded_file = st.file_uploader("Load Excel File", type=["xlsx"])
output_dir = st.text_input("Output folder path (or leave blank to download)", value="")
output_name = st.text_input("Output file name", value="merged_output.xlsx")

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

if st.button("Run Merge"):
    if not uploaded_file:
        st.error("Please load an Excel file")
    elif not st.session_state.rules:
        st.error("Please add at least one merge rule")
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
