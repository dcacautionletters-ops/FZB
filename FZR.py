import streamlit as st
import pandas as pd
import io
from openpyxl.styles import PatternFill, Font, Alignment, Border, Side

# --- 1. BASIC UI CONFIGURATION ---
st.set_page_config(page_title="VMS Reporting Utility", layout="wide")
MASTER_PASSWORD = "VMS@123"

# Minimalist Header
st.title("📊 VMS Reporting Utility")
st.caption("Standard Report Generation System")

# --- 2. AUTHENTICATION (BASIC) ---
if 'authenticated' not in st.session_state: st.session_state.authenticated = False
if not st.session_state.authenticated:
    p = st.text_input("Enter Password", type="password")
    if st.button("Login"):
        if p == MASTER_PASSWORD: 
            st.session_state.authenticated = True
            st.rerun()
        else: 
            st.error("Invalid Password")
    st.stop()

# --- 3. CORE LOGIC ---
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

def apply_styles(ws, threshold):
    thin = Side(style='thin', color="000000")
    border = Border(left=thin, right=thin, top=thin, bottom=thin)
    h_fill = PatternFill(start_color="D3D3D3", end_color="D3D3D3", fill_type="solid")
    crit_fill = PatternFill(start_color="FF0000", end_color="FF0000", fill_type="solid") # Red
    warn_fill = PatternFill(start_color="CCFFCC", end_color="CCFFCC", fill_type="solid") # Pale Green
    
    for col in range(1, ws.max_column + 1):
        cell = ws.cell(row=1, column=col)
        cell.font, cell.fill, cell.border = Font(bold=True), h_fill, border
        ws.column_dimensions[cell.column_letter].width = 18

    for row in ws.iter_rows(min_row=2, max_row=ws.max_row):
        for cell in row:
            cell.border, cell.alignment = border, Alignment(horizontal="center")
            if not isinstance(cell.value, str):
                try:
                    val = float(cell.value)
                    if val < 70: cell.fill, cell.font = crit_fill, Font(color="FFFFFF", bold=True)
                    elif 70 <= val < threshold: cell.fill = warn_fill
                except: pass

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

# --- 4. UPLOADER & PROCESSING ---
uploaded_file = st.file_uploader("Upload Excel File", type=["xlsx"])

if uploaded_file:
    df = pd.read_excel(uploaded_file)
    # Automatic Header Detection
    h_row = 0
    for i, row in df.head(15).iterrows():
        if any("ROLL NO" in str(x).upper() for x in row.values):
            h_row = i
            break
    df = pd.read_excel(uploaded_file, header=h_row)
    
    # Column Mapping
    c_map = {'sem': df.columns[5]}
    for c in df.columns:
        cs = str(c).strip()
        if "Roll No" in cs: c_map['roll'] = c
        elif "Student Name" in cs: c_map['name'] = c
        elif "Batch" in cs: c_map['batch'] = c
        elif any(x in cs for x in ["Course", "Subject"]): c_map['subject'] = c
        elif ATT_COL_NAME in cs: c_map['attendance'] = c

    threshold = st.number_input("Shortage Threshold (%)", 50, 100, 75)
    
    if st.button("Generate & Download Report", use_container_width=True):
        output = io.BytesIO()
        with pd.ExcelWriter(output, engine='openpyxl') as writer:
            
            # Grouping Logic
            df['Dept'] = df[c_map['batch']].astype(str).apply(lambda x: x.split()[0].upper())
            depts = sorted(df['Dept'].unique())
            
            for dept in depts:
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
                    s_df = d_df[d_df[c_map['batch']].astype(str).str.contains(series.split()[0]) & 
                                d_df[c_map['batch']].astype(str).str.contains(series.split()[-1])]
                    
                    subs = sorted([s for s in s_df[c_map['subject']].unique() if is_valid_subject(s)])
                    
                    # 1. GEN (Shortage)
                    gen = process_grid(s_df, c_map, subs, threshold, False)
                    if gen is not None:
                        sn = f"{series} GEN"[:31]
                        gen.to_excel(writer, sheet_name=sn, index=False)
                        get_bracket_summary(s_df, c_map, subs).to_excel(writer, sheet_name=sn, startrow=len(gen)+2, index=False)
                        apply_styles(writer.sheets[sn], threshold)

                    # 2. GEN ALL (Full)
                    all_at = process_grid(s_df, c_map, subs, threshold, True)
                    if all_at is not None:
                        sn_all = f"{series} GEN ALL"[:31]
                        all_at.to_excel(writer, sheet_name=sn_all, index=False)
                        get_bracket_summary(s_df, c_map, subs).to_excel(writer, sheet_name=sn_all, startrow=len(all_at)+2, index=False)
                        apply_styles(writer.sheets[sn_all], threshold)

        st.success("Reports Generated Successfully!")
        st.download_button("📥 Click here to Download", output.getvalue(), "VMS_Reports.xlsx", use_container_width=True)
