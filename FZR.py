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
st.info("Direct Mode: Identifying data starting from Column B...")

uploaded_file = st.file_uploader("Upload Universal Attendance Excel", type=["xlsx"])

if uploaded_file:
    try:
        # Load the whole sheet, starting from Column B (index 1)
        # 'usecols' tells pandas to ignore Column A entirely
        df_raw = pd.read_excel(uploaded_file, header=None, usecols="B:Z")
        
        # FIND THE HEADER: Look for "Roll No" in Column B
        start_row_idx = None
        for i, val in enumerate(df_raw.iloc[:, 0]): # Check Column B (now index 0)
            if str(val).strip().upper() == "ROLL NO":
                start_row_idx = i
                break
        
        if start_row_idx is None:
            st.error("❌ Could not find 'Roll No' in Column B. Check your headers.")
            st.dataframe(df_raw.head(10)) # Show preview
        else:
            # Slice the data: Headers are at start_row_idx, Data is below it
            df = df_raw.iloc[start_row_idx:].copy()
            df.columns = df.iloc[0] # Set the row with "Roll No" as header
            df = df[1:].reset_index(drop=True) # Remove the header row from data
            
            # Clean Column Names
            df.columns = [str(c).strip() for c in df.columns]
            
            # Exact Column Names to map
            COL_ROLL = "Roll No"
            COL_NAME = "Student Name"
            COL_BATCH = "Batch"
            COL_SUB = "Subject Name"
            COL_ATT = "Attended Hours with Approved Leave Percentage"

            # Force Alphanumeric Roll Nos to String
            df[COL_ROLL] = df[COL_ROLL].astype(str).str.strip()

            if COL_ROLL not in df.columns:
                st.error(f"Missing column: {COL_ROLL}")
            else:
                st.success(f"✅ Table found! First Roll No detected: {df[COL_ROLL].iloc[0]}")

                if st.button("Generate Plain Reports"):
                    output = io.BytesIO()
                    writer = pd.ExcelWriter(output, engine='openpyxl')
                    
                    # Safety Summary Sheet
                    pd.DataFrame({"Status": ["Processed"]}).to_excel(writer, sheet_name="Summary", index=False)
                    
                    # Grouping by Batch
                    for b in df[COL_BATCH].dropna().unique():
                        b_df = df[df[COL_BATCH] == b].copy()
                        
                        # Convert attendance percentage to number
                        b_df[COL_ATT] = pd.to_numeric(b_df[COL_ATT], errors='coerce')
                        
                        # Create the Pivot Grid
                        grid = b_df.pivot_table(
                            index=[COL_ROLL, COL_NAME],
                            columns=COL_SUB,
                            values=COL_ATT,
                            aggfunc='first'
                        ).reset_index()
                        
                        # Clean batch name for Excel sheet tab
                        sheet_name = str(b)[:30].replace("/", "-").replace("\\", "-")
                        grid.to_excel(writer, sheet_name=sheet_name, index=False)
                    
                    writer.close()
                    st.download_button("📥 Download Report", output.getvalue(), "VMS_Final_Report.xlsx")
                    
    except Exception as e:
        st.error(f"Critical Error: {e}")
