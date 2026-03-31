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

# --- 2. THE RIGID LOGIC ---
st.title("VMS Report Generator (Plain Mode)")

uploaded_file = st.file_uploader("Upload Universal Attendance Excel", type=["xlsx"])

if uploaded_file:
    # Scan for the header row containing "Roll No"
    df_temp = pd.read_excel(uploaded_file, header=None)
    header_row = 0
    found_header = False
    for i, row in df_temp.iterrows():
        if "Roll No" in row.values:
            header_row = i
            found_header = True
            break
            
    if not found_header:
        st.error("Could not find 'Roll No' in any row. Please check your Excel headers.")
    else:
        df = pd.read_excel(uploaded_file, header=header_row)
        df = df.dropna(how='all', axis=1).dropna(how='all', axis=0)
        
        # EXACT Column Names
        COL_ROLL = "Roll No"
        COL_NAME = "Student Name"
        COL_BATCH = "Batch"
        COL_SUB = "Subject Name"
        COL_ATT = "Attended Hours with Approved Leave Percentage"

        existing_cols = list(df.columns)
        missing = [c for c in [COL_ROLL, COL_NAME, COL_BATCH, COL_SUB, COL_ATT] if c not in existing_cols]

        if missing:
            st.error(f"Missing columns: {missing}")
            st.write("Columns found in your file:", existing_cols)
        else:
            threshold = st.number_input("Shortage %", 0, 100, 75)
            
            if st.button("Generate Plain Reports"):
                output = io.BytesIO()
                # Create the writer
                writer = pd.ExcelWriter(output, engine='openpyxl')
                
                # --- THE FIX: Create a sheet IMMEDIATELY ---
                # This prevents the IndexError even if the loop finds nothing.
                pd.DataFrame({"Report Status": ["Generated Successfully"], 
                             "Threshold Used": [f"{threshold}%"]}).to_excel(writer, sheet_name="Summary", index=False)
                
                # Grouping and Processing
                batches = df[COL_BATCH].dropna().unique()
                
                for b in batches:
                    b_df = df[df[COL_BATCH] == b].copy()
                    
                    if not b_df.empty:
                        # Convert attendance to numeric
                        b_df[COL_ATT] = pd.to_numeric(b_df[COL_ATT], errors='coerce')
                        
                        # Create the Grid
                        grid = b_df.pivot_table(
                            index=[COL_ROLL, COL_NAME, COL_BATCH],
                            columns=COL_SUB,
                            values=COL_ATT
                        ).reset_index()
                        
                        # Clean sheet name
                        sn = str(b)[:30].replace("/", "-").replace("\\", "-")
                        grid.to_excel(writer, sheet_name=sn, index=False)
                
                # Close and save
                writer.close()
                
                st.success("Reports Generated!")
                st.download_button("Download Excel", output.getvalue(), "VMS_Plain_Report.xlsx")
