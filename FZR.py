import streamlit as st
import pandas as pd
import io

# --- 1. BASIC UI ---
st.set_page_config(page_title="VMS Report Generator", layout="centered")
MASTER_PASSWORD = "VMS@123"

if 'auth' not in st.session_state: st.session_state.auth = False
if not st.session_state.auth:
    st.title("📊 VMS Login")
    p = st.text_input("Password", type="password")
    if st.button("Login"):
        if p == MASTER_PASSWORD: 
            st.session_state.auth = True
            st.rerun()
        else: st.error("Incorrect password")
    st.stop()

# --- 2. HELPER FUNCTIONS ---
KEYWORDS_TO_IGNORE = ["BADMINTON", "BASKETBALL", "CROSS FITNESS", "SWIMMING", "ZUMBA", "TABLE TENNIS", "SOFT SKILL", "ATOM"]
ATT_COL_KEY = "ATTENDED HOURS WITH APPROVED LEAVE PERCENTAGE"

def is_valid_subject(subject_name):
    s_upper = str(subject_name).upper()
    return not any(bad in s_upper for bad in KEYWORDS_TO_IGNORE)

def process_grid(data_df, cols, batch_subjects, threshold, show_all=False):
    if data_df.empty: return None
    df_local = data_df.copy()
    df_local[cols['attendance']] = pd.to_numeric(df_local[cols['attendance']], errors='coerce')
    
    grid = df_local.pivot_table(index=[cols['roll'], cols['name'], cols['batch']],
                                columns=cols['subject'], values=cols['attendance'], sort=False).reset_index()
    
    subs = [s for s in batch_subjects if is_valid_subject(s)]
    for s in subs:
        if s not in grid.columns: grid[s] = None
        grid[s] = pd.to_numeric(grid[s], errors='coerce')

    theory = [c for c in subs if not any(x in str(c).upper() for x in ["LAB", "PRACTICAL", "WORKSHOP"])]
    grid['Theory Avg'] = grid[theory].mean(axis=1).round(2)
    grid['Final Avg'] = grid[subs].mean(axis=1).round(2)
    
    if not show_all:
        mask = (grid[subs] < threshold).any(axis=1)
        grid = grid[mask].copy()
        if grid.empty: return None
        for s in subs:
            grid[s] = grid[s].apply(lambda x: x if (pd.notnull(x) and x < threshold) else "")

    grid.insert(0, 'Sl No.', range(1, len(grid) + 1))
    return grid

# --- 3. DATA PROCESSING ---
st.title("📊 VMS Report Generator")
uploaded_file = st.file_uploader("Upload Universal Attendance Excel", type=["xlsx"])

if uploaded_file:
    # 1. Detect Header Row
    df_raw = pd.read_excel(uploaded_file)
    h_row = 0
    for i, row in df_raw.head(20).iterrows():
        row_str = " ".join([str(x).upper() for x in row.values])
        if "ROLL NO" in row_str and "STUDENT NAME" in row_str:
            h_row = i
            break
    
    df = pd.read_excel(uploaded_file, header=h_row)
    df.columns = [str(c).strip() for c in df.columns] # Clean spaces
    
    # 2. Smart Column Mapping
    c_map = {}
    for c in df.columns:
        c_up = c.upper()
        if "ROLL NO" in c_up: c_map['roll'] = c
        elif "STUDENT NAME" in c_up: c_map['name'] = c
        elif "BATCH" in c_up: c_map['batch'] = c
        elif "SUBJECT" in c_up or "COURSE" in c_up: c_map['subject'] = c
        elif ATT_COL_KEY in c_up: c_map['attendance'] = c

    # Check for missing columns
    required = ['roll', 'name', 'batch', 'subject', 'attendance']
    missing = [r for r in required if r not in c_map]
    
    if missing:
        st.error(f"Missing columns in Excel: {missing}")
        st.info("Ensure your Excel has: Roll No, Student Name, Batch, Subject Name, and Attended Hours %")
    else:
        threshold = st.number_input("Shortage Threshold (%)", 50, 100, 75)
        
        if st.button("Generate Reports", use_container_width=True):
            output = io.BytesIO()
            writer = pd.ExcelWriter(output, engine='openpyxl')
            
            # Initial safety sheet
            pd.DataFrame({"Status": ["Report Generated"]}).to_excel(writer, sheet_name="Summary", index=False)
            
            try:
                df['Dept'] = df[c_map['batch']].astype(str).apply(lambda x: x.split()[0].upper())
                
                # Grouping Logic
                unique_batches = df[c_map['batch']].astype(str).unique()
                series_list = set()
                for b in unique_batches:
                    b_parts = b.split()
                    if "BCA" in b.upper() or "MCA" in b.upper():
                        year = next((p for p in b_parts if p.isdigit()), "2025")
                        series_list.add(f"{b_parts[0].upper()} {year}")
                    elif "MBA" in b.upper() or "MCOM" in b.upper():
                        series_list.add(" ".join(b_parts[:3])) # MBA BU 2025
                    else:
                        series_list.add(" ".join(b_parts[:2]))

                for series in sorted(list(series_list)):
                    s_parts = series.split()
                    # Filter for BCA/MCA sub-streams together
                    s_df = df[df[c_map['batch']].astype(str).str.contains(s_parts[0], case=False) & 
                              df[c_map['batch']].astype(str).str.contains(s_parts[-1], case=False)]
                    
                    subs = sorted([s for s in s_df[c_map['subject']].unique() if is_valid_subject(s)])
                    
                    # Shortage Sheet
                    gen = process_grid(s_df, c_map, subs, threshold, False)
                    if gen is not None: gen.to_excel(writer, sheet_name=f"{series} GEN"[:31], index=False)
                    
                    # Full Sheet
                    all_at = process_grid(s_df, c_map, subs, threshold, True)
                    if all_at is not None: all_at.to_excel(writer, sheet_name=f"{series} GEN ALL"[:31], index=False)

                writer.close()
                st.success("Reports created!")
                st.download_button("📥 Download Report", output.getvalue(), "VMS_Report.xlsx", use_container_width=True)
            
            except Exception as e:
                st.error(f"Error during grouping: {e}")
                writer.close()
