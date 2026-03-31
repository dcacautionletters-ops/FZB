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

uploaded_file = st.file_uploader("Upload Universal Attendance Excel", type=["xlsx"])

if uploaded_file:
    try:
        # Read the raw sheet
        df_raw = pd.read_excel(uploaded_file, header=None)
        
        start_row = None
        start_col = None
        
        # Scan for the word "Roll No" in your Column B (or anywhere else)
        for r_idx, row in df_raw.iterrows():
            for c_idx, value in enumerate(row):
                if str(value).strip().upper() == "ROLL NO":
                    start_row = r_idx
                    start_col = c_idx
                    break
            if start_row is not None: break
                
        if start_row is None:
            st.error("❌ Could not find the header 'Roll No'. Please check your Excel file.")
        else:
            # Crop the data to start exactly where the Roll No (e.g., 25CG001) starts
            df_sliced = df_raw.iloc[start_row:].copy()
            df_sliced = df_sliced.iloc[:, start_col:] 
            
            # Use the first row as the Header
            df_sliced.columns = df_sliced.iloc[0]
            df = df_sliced[1:].reset_index(drop=True)
            
            # Clean spaces from column names
            df.columns = [str(c).strip() for c in df.columns]
            
            # Map the exact columns you use
            COL_ROLL = "Roll No"
            COL_NAME = "Student Name"
            COL_BATCH = "Batch"
            COL_SUB = "Subject Name"
            COL_ATT = "Attended Hours with Approved Leave Percentage"

            # --- THE FIX FOR 25CG001 ---
            # This line ensures alpha-numeric IDs are NEVER treated as pure numbers
            df[COL_ROLL] = df[COL_ROLL].astype(str).str.strip()
            
            st.success(f"✅ Table found! Example ID detected: {df[COL_ROLL].iloc[0]}")

            if st.button("Generate Reports"):
                output = io.BytesIO()
                writer = pd.ExcelWriter(output, engine='openpyxl')
                
                # Summary Page
                pd.DataFrame({"Status": ["Processed Successfully"]}).to_excel(writer, sheet_name="Summary", index=False)
                
                # Grouping by Batch
                for b in df[COL_BATCH].dropna().unique():
                    b_df = df[df[COL_BATCH] == b].copy()
                    
                    # Convert attendance to numeric for calculation
                    b_df[COL_ATT] = pd.to_numeric(b_df[COL_ATT], errors='coerce')
                    
                    # Create the Grid (Pivot)
                    grid = b_df.pivot_table(
                        index=[COL_ROLL, COL_NAME],
                        columns=COL_SUB,
                        values=COL_ATT,
                        aggfunc='first' # Keeps the alpha-numeric ID exactly as it is
                    ).reset_index()
                    
                    # Save to Batch Sheet
                    sn = str(b)[:30].replace("/", "-")
                    grid.to_excel(writer, sheet_name=sn, index=False)
                
                writer.close()
                st.download_button("📥 Download Final Report", output.getvalue(), "VMS_Final_Report.xlsx")
                
    except Exception as e:
        st.error(f"Error: {e}")
