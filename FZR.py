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
        # Step 1: Read the entire raw sheet (no bounds)
        df_raw = pd.read_excel(uploaded_file, header=None)
        
        start_row = None
        start_col = None
        
        # Step 2: Scan for "Roll No" (The Anchor)
        # This scans every cell until it finds the header.
        for r_idx, row in df_raw.iterrows():
            for c_idx, value in enumerate(row):
                if str(value).strip().upper() == "ROLL NO":
                    start_row = r_idx
                    start_col = c_idx
                    break
            if start_row is not None: break
                
        if start_row is None:
            st.error("❌ Could not find 'Roll No' in your Excel file.")
            st.write("First 5 rows of your file:")
            st.dataframe(df_raw.head(5))
        else:
            # Step 3: Slice the data from the anchor point
            # This ignores Column A and Rows 1-2 automatically.
            df_sliced = df_raw.iloc[start_row:].copy()
            df_sliced = df_sliced.iloc[:, start_col:] # Crop everything left of Roll No
            
            # Set the first row of the slice as headers
            df_sliced.columns = df_sliced.iloc[0]
            df = df_sliced[1:].reset_index(drop=True)
            
            # Clean Column Names
            df.columns = [str(c).strip() for c in df.columns]
            
            # Column mapping
            COL_ROLL = "Roll No"
            COL_NAME = "Student Name"
            COL_BATCH = "Batch"
            COL_SUB = "Subject Name"
            COL_ATT = "Attended Hours with Approved Leave Percentage"

            # Check if headers exist
            if COL_ROLL not in df.columns:
                st.error(f"Found anchor, but column '{COL_ROLL}' is missing. Check spelling.")
            else:
                # Force IDs (like 22MCA01) and Batches to be text
                df[COL_ROLL] = df[COL_ROLL].astype(str).str.strip()
                df[COL_BATCH] = df[COL_BATCH].astype(str).str.strip()
                
                st.success(f"✅ Data detected starting at Row {start_row + 1}, Column {start_col + 1}")

                if st.button("Generate Reports"):
                    output = io.BytesIO()
                    writer = pd.ExcelWriter(output, engine='openpyxl')
                    
                    # Create Summary sheet (prevents IndexError)
                    pd.DataFrame({"Status": ["Processing Complete"]}).to_excel(writer, sheet_name="Summary", index=False)
                    
                    # Group by Batch
                    batches = df[COL_BATCH].unique()
                    
                    for b in batches:
                        if pd.isna(b) or str(b).lower() == 'nan': continue
                        
                        b_df = df[df[COL_BATCH] == b].copy()
                        
                        # Convert attendance to number
                        b_df[COL_ATT] = pd.to_numeric(b_df[COL_ATT], errors='coerce')
                        
                        # Create the Pivot Grid (Alpha-Numeric proof)
                        grid = b_df.pivot_table(
                            index=[COL_ROLL, COL_NAME],
                            columns=COL_SUB,
                            values=COL_ATT,
                            aggfunc='first'
                        ).reset_index()
                        
                        # Clean sheet name for Excel
                        sheet_name = str(b)[:30].replace("/", "-").replace("\\", "-")
                        grid.to_excel(writer, sheet_name=sheet_name, index=False)
                    
                    writer.close()
                    st.download_button("📥 Download Excel Report", output.getvalue(), "VMS_Final_Report.xlsx")
                
    except Exception as e:
        st.error(f"Error: {e}")
