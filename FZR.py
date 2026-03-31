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
        # Load raw data
        df_raw = pd.read_excel(uploaded_file, header=None)
        
        # DEFAULT FALLBACK: Row 3, Column B (Python index 2, 1)
        # We try to find "Roll No" first, but if we can't, we use your coordinates.
        start_row, start_col = 2, 1 
        
        for r in range(min(10, len(df_raw))):
            for c in range(min(10, len(df_raw.columns))):
                val = str(df_raw.iloc[r, c]).strip().upper()
                if "ROLL" in val or "REG" in val:
                    start_row, start_col = r, c
                    break
            if start_row != 2: break

        # Slice the data
        df_sliced = df_raw.iloc[start_row:].copy()
        df_sliced = df_sliced.iloc[:, start_col:] 
        
        # Set Headers
        df_sliced.columns = df_sliced.iloc[0]
        df = df_sliced[1:].reset_index(drop=True)
        
        # Clean Column Names
        df.columns = [str(c).strip() for c in df.columns]
        
        # Map columns by finding keywords (even more flexible)
        c_map = {}
        for col in df.columns:
            c_up = str(col).upper()
            if "ROLL" in c_up or "REG" in c_up: c_map['roll'] = col
            elif "NAME" in c_up: c_map['name'] = col
            elif "BATCH" in c_up: c_map['batch'] = col
            elif "SUBJECT" in c_up or "COURSE" in c_up: c_map['sub'] = col
            elif "ATTENDED" in c_up and "PERCENTAGE" in c_up: c_map['att'] = col

        # Verify minimum requirements
        required = ['roll', 'name', 'batch', 'sub', 'att']
        missing = [r for r in required if r not in c_map]

        if missing:
            st.error(f"❌ Guru, still missing columns: {missing}")
            st.write("Headers I found at your location:", list(df.columns))
        else:
            # Force Alpha-Numeric (25CG001) to String
            df[c_map['roll']] = df[c_map['roll']].astype(str).str.strip()
            
            st.success(f"✅ Data found! Starting with: {df[c_map['roll']].iloc[0]}")

            if st.button("Generate Reports"):
                output = io.BytesIO()
                writer = pd.ExcelWriter(output, engine='openpyxl')
                
                pd.DataFrame({"Status": ["Processed"]}).to_excel(writer, sheet_name="Summary", index=False)
                
                for b in df[c_map['batch']].dropna().unique():
                    b_df = df[df[c_map['batch']] == b].copy()
                    b_df[c_map['att']] = pd.to_numeric(b_df[c_map['att']], errors='coerce')
                    
                    grid = b_df.pivot_table(
                        index=[c_map['roll'], c_map['name']],
                        columns=c_map['sub'],
                        values=c_map['att'],
                        aggfunc='first'
                    ).reset_index()
                    
                    sn = str(b)[:30].replace("/", "-")
                    grid.to_excel(writer, sheet_name=sn, index=False)
                
                writer.close()
                st.download_button("📥 Download Final Report", output.getvalue(), "VMS_Report.xlsx")
                
    except Exception as e:
        st.error(f"Error: {e}")
