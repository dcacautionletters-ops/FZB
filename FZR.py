import streamlit as st
import pandas as pd
import io

# --- 1. MINIMAL UI ---
st.set_page_config(page_title="VMS Report Generator", layout="centered")
MASTER_PASSWORD = "VMS@123"

st.title("VMS Report Generator")
st.write("Upload a file to generate plain Excel reports.")

# --- 2. SIMPLE AUTHENTICATION ---
if 'auth' not in st.session_state: st.session_state.auth = False
if not st.session_state.auth:
    p = st.text_input("Password", type="password")
    if st.button("Login"):
        if p == MASTER_PASSWORD: 
            st.session_state.auth = True
            st.rerun()
        else: st.error("Incorrect password")
    st.stop()

# --- 3. HELPER FUNCTIONS ---
KEYWORDS_TO_IGNORE = ["BADMINTON", "BASKETBALL", "CROSS FITNESS", "SWIMMING", "ZUMBA", "TABLE TENNIS", 
                      "FREESLOT", "FREE SLOT", "SOFT SKILL", "ATOM", "DSA"]
ATT_COL_NAME = "Attended Hours with Approved Leave Percentage"

def is_valid_subject(subject_name):
    s_upper = str(subject_name).upper()
    return not any(bad in s_upper for bad in KEYWORDS_TO_IGNORE)

def get_bracket_summary(data_df, cols, subjects):
    summary_data = []
    for sub in subjects:
        sub_vals = pd.to_numeric(data_df[data_df[cols['subject']] == sub][cols['attendance']], errors='coerce').dropna()
        b1 = len(sub_vals[(sub_vals >= 0) & (sub_vals < 50)])
        b2 = len(sub_vals[(sub_vals >= 50) & (sub_vals < 60)])
        b3 = len(sub_vals[(sub_vals >= 60) & (sub_vals < 70)])
        b4 = len(sub_vals[(sub_vals >= 70) & (sub_vals < 75)])
        summary_data.append({
            "Subject": sub, "0.00-49.99": b1, "50.00-59.99": b2, 
            "60.00-69.99": b3, "70.00-74.99": b4, "Total": b1+b2+b3+b4
        })
    return pd.DataFrame(summary_data)

def process_grid(data_df, cols, batch_subjects, threshold, show_all=False):
    if data_df.empty: return None
    df_local = data_df.copy()
    df_local[cols['attendance']] = pd.to_numeric(df_local[cols['attendance']], errors='coerce')
    
    grid = df_local.pivot_table(index=[cols['roll'], cols['name'], cols['batch'], cols['sem']],
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

# --- 4. EXECUTION ---
uploaded_file = st.file_uploader("Choose Excel File", type=["xlsx"])

if uploaded_file:
    df_raw = pd.read_excel(uploaded_file)
    h_row = 0
    for i, row in df_raw.head(15).iterrows():
        if any("ROLL NO" in str(x).upper() for x in row.values):
            h_row = i
            break
    df = pd.read_excel(uploaded_file, header=h_row)
    
    c_map = {'sem': df.columns[5]}
    for c in df.columns:
        cs = str(c).strip()
        if "Roll No" in cs: c_map['roll'] = c
        elif "Student Name" in cs: c_map['name'] = c
        elif "Batch" in cs: c_map['batch'] = c
        elif any(x in cs for x in ["Course", "Subject"]): c_map['subject'] = c
        elif ATT_COL_NAME in cs: c_map['attendance'] = c

    threshold = st.number_input("Shortage Threshold (%)", 50, 100, 75)
    
    if st.button("Generate Reports"):
        output = io.BytesIO()
        sheets_created = 0 # TRACKER TO PREVENT THE ERROR
        
        with pd.ExcelWriter(output, engine='openpyxl') as writer:
            df['Dept'] = df[c_map['batch']].astype(str).apply(lambda x: x.split()[0].upper())
            
            for dept in sorted(df['Dept'].unique()):
                d_df = df[df['Dept'] == dept]
                unique_batches = d_df[c_map['batch']].astype(str).unique()
                series_list = set()
                
                for b in unique_batches:
                    b_parts = b.split()
                    if "BCA" in b.upper() or "MCA" in b.upper():
                        year = next((p for p in b_parts if p.isdigit()), "Series")
                        series_list.add(f"{b_parts[0].upper()} {year}")
                    elif "MBA" in b.upper() or "MCOM" in b.upper():
                        series_list.add(' '.join(b_parts[:3]))
                    else:
                        series_list.add(' '.join(b_parts[:2]))
                
                for series in sorted(list(series_list)):
                    s_parts = series.split()
                    s_df = d_df[d_df[c_map['batch']].astype(str).str.contains(s_parts[0]) & 
                                d_df[c_map['batch']].astype(str).str.contains(s_parts[-1])]
                    
                    subs = sorted([s for s in s_df[c_map['subject']].unique() if is_valid_subject(s)])
                    
                    # 1. GEN
                    gen_data = process_grid(s_df, c_map, subs, threshold, False)
                    if gen_data is not None:
                        sn = f"{series} GEN"[:31]
                        gen_data.to_excel(writer, sheet_name=sn, index=False)
                        get_bracket_summary(s_df, c_map, subs).to_excel(writer, sheet_name=sn, startrow=len(gen_data)+2, index=False)
                        sheets_created += 1

                    # 2. GEN ALL
                    all_data = process_grid(s_df, c_map, subs, threshold, True)
                    if all_data is not None:
                        sn_all = f"{series} GEN ALL"[:31]
                        all_data.to_excel(writer, sheet_name=sn_all, index=False)
                        get_bracket_summary(s_df, c_map, subs).to_excel(writer, sheet_name=sn_all, startrow=len(all_at if 'all_at' in locals() else all_data)+2, index=False)
                        sheets_created += 1
            
            # THE FIX: If no sheets were made, create a blank one so openpyxl doesn't crash
            if sheets_created == 0:
                pd.DataFrame({"Message": ["No data matched your filtering criteria."]}).to_excel(writer, sheet_name="No Data Found")

        st.success("Report Ready!")
        st.download_button("Download Excel", output.getvalue(), "VMS_Plain_Reports.xlsx")
