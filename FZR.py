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
st.title("📊 VMS Report Generator")
st.info("Scanner Active: Searching for 'Roll No' in Row 2, Column B...")

uploaded_file = st.file_uploader("Upload Universal Attendance Excel", type=["xlsx"])

if uploaded_file:
    # Read the raw file without headers to find the coordinates
    df_raw = pd.read_excel(uploaded_file, header=None)
    
    start_row = None
    start_col = None
    
    # SCANNING LOGIC: Find "Roll No" anywhere in the sheet
    for r_idx, row in df_raw.iterrows():
        for c_idx, value in enumerate(row):
            if str(value).strip().upper() == "ROLL NO":
                start_row = r_idx
                start_col = c_idx
                break
        if start_row is not None: break
            
    if start_row is None:
        st.error("❌ Still could not find 'Roll No'.")
        st.write("Found instead in Row 2, Col B:", df_raw.iloc[1, 1] if df_raw.shape[0] > 1 else "Empty")
        st.dataframe(df_raw.head(10)) # Show what the app sees
    else:
        # RELOAD: Start reading from the identified row and skip the empty columns
        df = pd.read_excel(uploaded_file, header=start_row)
        # Drop columns to the left of "Roll No"
        df = df.iloc[:, start_col:] 
        
        # Clean column names
        df.columns = [str(c).strip() for c in df.columns]
        
        # Define the exact names we need
        COL_ROLL = "Roll No"
        COL_NAME = "Student Name"
        COL_BATCH = "Batch"
        COL_SUB = "Subject Name"
        COL_ATT = "Attended Hours with Approved Leave Percentage"

        # Check if they exist
        actual_cols = list(df.columns)
        missing = [c for c in [COL_ROLL, COL_NAME, COL_BATCH, COL_SUB, COL_ATT] if c not in actual_cols]

        if missing:
            st.error(f"❌ Found 'Roll No' but missing other columns: {missing}")
            st.write("Headers detected at that location:", actual_cols)
        else:
            st.success(f"✅ Data found starting at Row {start_row + 1}!")
            
            if st.button("Generate Plain Reports"):
                output = io.BytesIO()
                writer = pd.ExcelWriter(output, engine='openpyxl')
                
                # Safety sheet
                pd.DataFrame({"Status": ["Processed"]}).to_excel(writer, sheet_name="Summary", index=False)
                
                # Unique batches
                batches = df[COL_BATCH].dropna().unique()
                
                for b in batches:
                    b_df = df[df[COL_BATCH] == b].copy()
                    
                    # Convert attendance to numeric for math
                    b_df[COL_ATT] = pd.to_numeric(b_df[COL_ATT], errors='coerce')
                    
                    # Create the Grid
                    grid = b_df.pivot_table(
                        index=[COL_ROLL, COL_NAME],
                        columns=COL_SUB,
                        values=COL_ATT,
                        aggfunc='first'
                    ).reset_index()
                    
                    # Excel sheet name limits
                    sheet_name = str(b)[:30].replace("/", "-")
                    grid.to_excel(writer, sheet_name=sheet_name, index=False)
                
                writer.close()
                st.download_button("📥 Download Report", output.getvalue(), "VMS_Final_Report.xlsx")
