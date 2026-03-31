import streamlit as st
import pandas as pd
import io
import re

st.set_page_config(page_title="VMS Report Generator", layout="centered")

# --- 1. SMART HEADER DETECTION ---
def clean_col(name):
    return re.sub(r'[^A-Z0-9]', '', str(name).upper())

st.title("📊 VMS Report Generator")
uploaded_file = st.file_uploader("Upload Excel File", type=["xlsx"])

if uploaded_file:
    # Read first 20 rows to find the header
    df_preview = pd.read_excel(uploaded_file, header=None, nrows=20)
    header_idx = None
    
    for i, row in df_preview.iterrows():
        row_values = [clean_col(x) for x in row.values if pd.notnull(x)]
        # Look for the row that contains both Roll and Name
        if any("ROLL" in x for x in row_values) and any("NAME" in x for x in row_values):
            header_idx = i
            break
            
    if header_idx is None:
        st.error("Could not find the header row. Please ensure 'Roll No' and 'Student Name' are present.")
    else:
        df = pd.read_excel(uploaded_file, header=header_idx)
        df.columns = [str(c).strip() for c in df.columns]
        
        # Mapping with high flexibility
        c_map = {}
        for c in df.columns:
            c_clean = clean_col(c)
            if "ROLL" in c_clean: c_map['roll'] = c
            elif "STUDENTNAME" in c_clean or "NAME" in c_clean: c_map['name'] = c
            elif "BATCH" in c_clean: c_map['batch'] = c
            elif "SUBJECT" in c_clean or "COURSE" in c_clean: c_map['subject'] = c
            elif "ATTENDEDHOURSWITHAPPROVEDLEAVEPERCENTAGE" in c_clean: c_map['attendance'] = c

        # Verification
        required = ['roll', 'name', 'batch', 'subject', 'attendance']
        missing = [r for r in required if r not in c_map]
        
        if missing:
            st.error(f"Still missing: {missing}")
            st.write("Current Headers found:", list(df.columns))
        else:
            st.success("All columns mapped successfully!")
            # Proceed with report generation logic...
            if st.button("Generate Final Report"):
                # (Your processing logic here)
                st.write("Processing data...")
