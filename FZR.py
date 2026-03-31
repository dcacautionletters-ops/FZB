import streamlit as st
import pandas as pd
import io

# --- 1. LOGIN ---
st.set_page_config(page_title="VMS Utility", layout="wide")
if 'auth' not in st.session_state: st.session_state.auth = False
if not st.session_state.auth:
    p = st.text_input("Password", type="password")
    if st.button("Login"):
        if p == "VMS@123": 
            st.session_state.auth = True
            st.rerun()
    st.stop()

# --- 2. THE GENERATOR ---
st.title("📊 VMS Alpha-Numeric Report Generator")
st.info("Scanner Active: Detecting headers from Row 3, Column B...")

uploaded_file = st.file_uploader("Upload Universal Attendance Excel", type=["xlsx"])

if uploaded_file:
    # Step 1: Read raw data (no types assumed yet)
    df_raw = pd.read_excel(uploaded_file, header=None)
    
    start_row = None
    start_col = None
    
    # Step 2: Scan for the "Roll No" anchor
    for r_idx, row in df_raw.iterrows():
        for c_idx, value in enumerate(row):
            if str(value).strip().upper() == "ROLL NO":
                start_row = r_idx
                start_col = c_idx
                break
        if start_row is not None: break
            
    if start_row is None:
        st.error("❌ Could not find 'Roll No'. Please check Row 3, Column B.")
        st.dataframe(df_raw.head(10)) 
    else:
        # Step 3: Slice the data from the anchor point
        df_sliced = df_raw.iloc[start_row:].copy()
        df_sliced = df_sliced.iloc[:, start_col:]
        
        # Set headers and drop the header row from data
        df_sliced.columns = df_sliced.iloc[0]
        df = df_sliced[1:].reset_index(drop=True)
        
        # Clean column names
        df.columns = [str(c).strip() for c in df.columns]
        
        # Define exact column keys
        COL_ROLL = "Roll No"
        COL_NAME = "Student Name"
        COL_BATCH = "Batch"
        COL_SUB = "Subject Name"
        COL_ATT = "Attended Hours with Approved Leave Percentage"

        if COL_ROLL not in df.columns:
            st.error(f"❌ Found anchor, but column '{COL_ROLL}' is missing. Check spelling.")
        else:
            # --- CRITICAL FIX FOR ALPHA-NUMERIC ---
            # Force Roll No and Batch to be strings to prevent data loss
            df[COL_ROLL] = df[COL_ROLL].astype(str).str.strip()
            df[COL_BATCH] = df[COL_BATCH].astype(str).str.strip()
            
            st.success(f"✅ Table recognized at Row {start_row + 1}!")

            if st.button("Generate Reports"):
                output = io.BytesIO()
                writer = pd.ExcelWriter(output, engine='openpyxl')
                
                # Summary Sheet
                pd.DataFrame({"Status": ["Alpha-Numeric Processing Complete"]}).to_excel(writer, sheet_name="Summary", index=False)
                
                # Unique batches
                batches = df[COL_BATCH].unique()
                
                for b in batches:
                    if pd.isna(b) or str(b).lower() == 'nan': continue
                    
                    b_df = df[df[COL_BATCH] == b].copy()
                    
                    # Ensure attendance is treated as a number for the pivot
                    b_df[COL_ATT] = pd.to_numeric(b_df[COL_ATT], errors='coerce')
                    
                    # Create the Grid
                    # Using 'first' ensures alpha-numeric Roll Nos aren't summed or averaged
                    grid = b_df.pivot_table(
                        index=[COL_ROLL, COL_NAME],
                        columns=COL_SUB,
                        values=COL_ATT,
                        aggfunc='first'
                    ).reset_index()
                    
                    # Format Sheet Name
                    sheet_name = str(b)[:30].replace("/", "-").replace("\\", "-")
                    grid.to_excel(writer, sheet_name=sheet_name, index=False)
                
                writer.close()
                st.download_button("📥 Download Excel Report", output.getvalue(), "VMS_Formatted_Report.xlsx")
