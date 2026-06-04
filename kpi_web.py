import streamlit as st
import pandas as pd
import io
import re
import numpy as np
from datetime import datetime
import plotly.express as px

# ================= PAGE CONFIGURATION =================
st.set_page_config(page_title="Stop-In-Out KPI Dashboard", page_icon="🚌", layout="wide")

# ================= STATE MANAGEMENT ===================
if 'analysis_complete' not in st.session_state:
    st.session_state.analysis_complete = False
    st.session_state.main_excel = None
    st.session_state.raw_excel = None
    st.session_state.summary_data = []
    st.session_state.all_accepted = []
    st.session_state.bunched_trips = []
    st.session_state.auto_rush = []
    st.session_state.run_time = None
    st.session_state.loaded_from_file = False

# ================= HELPERS =================
def extract_path_info(file_obj):
    file_obj.seek(0)
    if file_obj.name.lower().endswith('.csv'):
        try: raw = pd.read_csv(file_obj, header=None, nrows=1, encoding='utf-8')
        except UnicodeDecodeError: 
            file_obj.seek(0)
            raw = pd.read_csv(file_obj, header=None, nrows=1, encoding='latin1')
    else:
        raw = pd.read_excel(file_obj, header=None, nrows=1)

    file_obj.seek(0) 
    row_str = " ".join(raw.iloc[0].dropna().astype(str).tolist())
    stops = re.findall(r'\d+\s*-\s*(.+?)\s*\(\d+\)', row_str)
    
    if len(stops) >= 2:
        return f"{stops[0].strip().title()} → {stops[-1].strip().title()}"
    return "Unknown → Unknown"

def parse_datetime(date_series, time_series=None):
    if time_series is not None: dt_str = date_series.astype(str).str[:10] + ' ' + time_series.astype(str)
    else: dt_str = date_series.astype(str)
    dt_str = dt_str.replace(r'.*nan$', pd.NaT, regex=True)
    return pd.to_datetime(dt_str, dayfirst=True, errors='coerce')

def normalize_path(p):
    return re.sub(r'[^a-zA-Z0-9]', '', str(p)).lower()

def make_tracking_url(bus_id, start_dt, end_dt, display_text="Track Trip", raw_url_only=False):
    try:
        if pd.isna(start_dt) or pd.isna(end_dt) or pd.isna(bus_id): return "N/A"
        bus = str(bus_id).strip().upper().replace('EV', '20').replace('IB', '21')
        bus = re.sub(r'[^\d]', '', bus)
        start_str = start_dt.strftime('%Y%m%d%H%M%S')
        end_str = end_dt.strftime('%Y%m%d%H%M%S')
        url = f"https://cdaui.kentkart.pk/vts/#/?busid={bus}&offline=1&startdate={start_str}&enddate={end_str}"
        if raw_url_only: return url
        safe_text = str(display_text).replace('"', '""')
        return f'=HYPERLINK("{url}", "{safe_text}")'
    except: return "N/A"

def extract_raw_url(hyperlink_formula):
    try:
        if pd.isna(hyperlink_formula): return None
        match = re.search(r'=HYPERLINK\("([^"]+)"', str(hyperlink_formula))
        return match.group(1) if match else None
    except: return None

# ================= EXISTING DASHBOARD LOADER =================
def load_existing_dashboard(uploaded_file):
    try:
        xls = pd.ExcelFile(uploaded_file)
        
        if "KPI_SUMMARY" in xls.sheet_names:
            df_kpi_raw = pd.read_excel(xls, sheet_name="KPI_SUMMARY", header=None)
            df_kpi = df_kpi_raw.set_index(0).T
            df_kpi = df_kpi.drop(columns=[c for c in df_kpi.columns if str(c).startswith("Go To ") or c == " " or c == "-- QUICK NAVIGATION --"], errors='ignore')
            
            for col in df_kpi.columns:
                df_kpi[col] = pd.to_numeric(df_kpi[col], errors='ignore')
                
            st.session_state.summary_data = df_kpi.to_dict('records')
            
        if "ACCEPTED_TRIPS" in xls.sheet_names:
            st.session_state.all_accepted = pd.read_excel(xls, sheet_name="ACCEPTED_TRIPS").to_dict('records')
        if "BUNCHED_TRIPS" in xls.sheet_names:
            st.session_state.bunched_trips = pd.read_excel(xls, sheet_name="BUNCHED_TRIPS").to_dict('records')
        if "AUTO_DETECTED_RUSH" in xls.sheet_names:
            st.session_state.auto_rush = pd.read_excel(xls, sheet_name="AUTO_DETECTED_RUSH").to_dict('records')
            
        st.session_state.analysis_complete = True
        st.session_state.loaded_from_file = True
        st.session_state.run_time = datetime.now().strftime("%Y%m%d_%H%M")
        return True
    except Exception as e:
        st.error(f"Failed to read dashboard file: {str(e)}")
        return False

# ================= DATA LOADING =================
def load_master_data(master_file, status_log):
    if not master_file: return None, None, None, None, None, None, None, None
    try:
        master_dict, master_times, master_times_fast = {}, {}, {}
        intimations_dict, corrections_dict, rush_hours_dict, vts_dict = {}, {}, {}, {}
        path_distances = {}
        
        master_file.seek(0)
        is_csv = master_file.name.lower().endswith('.csv')
        
        if is_csv: 
            df_sched = pd.read_csv(master_file, header=None)
        else: 
            xls = pd.ExcelFile(master_file)
            df_sched = pd.read_excel(xls, sheet_name='Scheduled Times' if 'Scheduled Times' in xls.sheet_names else 0, header=None)
            
        # Parse schedules with date dependencies
        for col_idx in range(df_sched.shape[1]):
            col_data = df_sched.iloc[:, col_idx]
            path_val = str(col_data.iloc[0]).strip()
            if not path_val or "nan" in path_val.lower() or "unnamed" in path_val.lower(): continue
            
            try:
                s_raw, e_raw = col_data.iloc[1], col_data.iloc[2]
                s_date = pd.to_datetime(s_raw, errors='coerce').date() if pd.notna(s_raw) else None
                e_date = pd.to_datetime(e_raw, errors='coerce').date() if pd.notna(e_raw) else None
                
                if s_date and e_date:
                    times_series = col_data.iloc[3:]
                else:
                    s_date, e_date = pd.to_datetime('2000-01-01').date(), pd.to_datetime('2099-12-31').date()
                    times_series = col_data.iloc[1:]
                    
                times = times_series.astype(str).str.extract(r'(\d{2}:\d{2}:\d{2}|\d{2}:\d{2})')[0].dropna().tolist()
                if times:
                    if path_val not in master_dict: master_dict[path_val] = []
                    master_dict[path_val].append({'start': s_date, 'end': e_date, 'times': sorted(times)})
            except Exception as e:
                pass
        
        if not is_csv:
            if 'Total Time' in xls.sheet_names:
                df_times = pd.read_excel(xls, sheet_name='Total Time')
                for col_name in df_times.columns:
                    path_val = str(col_name).strip()
                    if "unnamed" in path_val.lower(): continue
                    try:
                        val0 = pd.to_numeric(df_times[col_name].iloc[0], errors='coerce')
                        if pd.notna(val0): master_times[path_val] = float(val0)
                        if len(df_times) > 1:
                            val1 = pd.to_numeric(df_times[col_name].iloc[1], errors='coerce')
                            if pd.notna(val1): master_times_fast[path_val] = float(val1)
                    except: pass

            if 'Distances' in xls.sheet_names:
                df_dist = pd.read_excel(xls, sheet_name='Distances')
                for col_name in df_dist.columns:
                    path_val = str(col_name).strip()
                    if "unnamed" in path_val.lower(): continue
                    try:
                        val0 = pd.to_numeric(df_dist[col_name].iloc[0], errors='coerce')
                        if pd.notna(val0): path_distances[normalize_path(path_val)] = float(val0)
                    except: pass

            if 'Intimations' in xls.sheet_names:
                df_int = pd.read_excel(xls, sheet_name='Intimations')
                df_int.columns = [str(c).strip() for c in df_int.columns]
                if set(['Date', 'Path']).issubset(df_int.columns):
                    for _, r in df_int.iterrows():
                        if pd.notna(r['Date']):
                            try:
                                d = pd.to_datetime(r['Date']).date()
                                p = normalize_path(r['Path'])
                                if 'Intimation' in df_int.columns and pd.notna(r['Intimation']):
                                    intimations_dict[(d, p)] = int(float(r['Intimation']))
                                if 'Correction' in df_int.columns and pd.notna(r['Correction']):
                                    corrections_dict[(d, p)] = int(float(r['Correction']))
                            except: pass

            if 'Rush Hours' in xls.sheet_names:
                df_rush = pd.read_excel(xls, sheet_name='Rush Hours')
                df_rush.columns = [str(c).strip() for c in df_rush.columns]
                if set(['Date', 'Path', 'Start Time', 'End Time']).issubset(df_rush.columns):
                    for _, r in df_rush.iterrows():
                        if pd.notna(r['Date']) and pd.notna(r['Start Time']) and pd.notna(r['End Time']):
                            try:
                                d = pd.to_datetime(r['Date']).date()
                                p = normalize_path(r['Path'])
                                st_time = pd.to_datetime(f"{d} {str(r['Start Time']).strip()}").time()
                                et = pd.to_datetime(f"{d} {str(r['End Time']).strip()}").time()
                                if (d, p) not in rush_hours_dict: rush_hours_dict[(d, p)] = []
                                rush_hours_dict[(d, p)].append((st_time, et))
                            except: pass

            if 'VTS' in xls.sheet_names:
                df_vts = pd.read_excel(xls, sheet_name='VTS')
                df_vts.columns = [str(c).strip() for c in df_vts.columns]
                if set(['Date', 'Bus']).issubset(df_vts.columns):
                    for _, r in df_vts.iterrows():
                        if pd.notna(r['Date']):
                            try:
                                d = pd.to_datetime(r['Date']).date()
                                b = str(r['Bus']).strip().replace('EV', '20').replace('IB', '21')
                                b = re.sub(r'[^\d]', '', b)
                                if not b: b = str(r['Bus']).strip()
                                
                                vts_dict[(d, b)] = {}
                                for c in df_vts.columns:
                                    if c.lower() not in ['date', 'bus', 'rd1', 'rd2']:
                                        try:
                                            if pd.notna(r[c]): vts_dict[(d, b)][normalize_path(c)] = int(float(r[c]))
                                        except: pass
                            except: pass

        if master_dict: status_log.info(f"Loaded schedules for {len(master_dict)} routes.")
        return master_dict, master_times, intimations_dict, corrections_dict, rush_hours_dict, master_times_fast, vts_dict, path_distances
    except Exception as e:
        status_log.error(f"ERROR reading master file: {e}")
        return None, None, None, None, None, None, None, None

# ================= CORE PROCESSOR =================
def process_file(file_obj, params, authoritative_dict, authoritative_times, intimations_dict, corrections_dict, rush_hours_dict, authoritative_times_fast, path_distances):
    directional_path = extract_path_info(file_obj)
    target_time = authoritative_times.get(directional_path) if authoritative_times else None
    target_time_fast = authoritative_times_fast.get(directional_path) if authoritative_times_fast else None
    
    file_obj.seek(0)
    if file_obj.name.lower().endswith('.csv'):
        try: df = pd.read_csv(file_obj, header=1, encoding='utf-8')
        except UnicodeDecodeError: 
            file_obj.seek(0)
            df = pd.read_csv(file_obj, header=1, encoding='latin1')
    else:
        df = pd.read_excel(file_obj, header=1)

    df.columns = df.columns.astype(str).str.strip()
    
    arr_cols, cik_cols, plan_cols = {}, {}, {}
    for c in df.columns:
        cl = str(c).lower().strip()
        if 'arrival date' in cl:
            idx = 0 if cl == 'arrival date' else cl.split('arrival date.')[-1]
            try: arr_cols[int(idx)] = c
            except: pass
        elif 'çıkış date' in cl or 'cikis date' in cl:
            base = 'çıkış date' if 'çıkış date' in cl else 'cikis date'
            idx = 0 if cl == base else cl.split(base + '.')[-1]
            try: cik_cols[int(idx)] = c
            except: pass
        elif cl.startswith('plan'):
            idx = 0 if cl == 'plan' else cl.split('plan.')[-1]
            try: plan_cols[int(idx)] = c
            except: pass

    lower_cols = {c.lower(): c for c in df.columns}
    tot_stops_col = lower_cols.get("total stops", lower_cols.get("stop count total"))

    col_map = {
        "Schedule Time": lower_cols.get("schedule time", lower_cols.get("scheduled departure")),
        "Edge Code": lower_cols.get("edge code", lower_cols.get("bus", lower_cols.get("bus id"))),
        "REALIZED_KM": lower_cols.get("realized_km", lower_cols.get("realized km")),
        "Start Date": lower_cols.get("start date time", lower_cols.get("start date", lower_cols.get("date"))),
        "Actual Time": lower_cols.get("trip start time", lower_cols.get("actual departure time", lower_cols.get("çıkış date"))),
        "End Time": lower_cols.get("end date time", lower_cols.get("duty end time", lower_cols.get("end date"))),
        "Stop Count": lower_cols.get("trip stop count", lower_cols.get("trip_stop_count", lower_cols.get("stop count", lower_cols.get("visited stop count"))))
    }

    missing_critical = [k for k, v in col_map.items() if v is None and k not in ["Actual Time", "End Time", "Stop Count"]]
    if missing_critical: raise ValueError(f"Could not map essential columns: {', '.join(missing_critical)}")

    df = df.rename(columns={
        col_map["Schedule Time"]: "Schedule_Time", col_map["Edge Code"]: "Bus_ID",
        col_map["REALIZED_KM"]: "Realized_KM", col_map["Start Date"]: "Start_Date_Raw"
    })

    base_cols = list(df.columns)

    def make_raw(row, classification, link):
        d = {c: row[c] for c in base_cols if c in row.index}
        d['Path'] = directional_path
        d['Date'] = row['Date']
        d['Classification / Reason'] = classification
        d['Tracking Link'] = link
        return d

    cikis_cols_list = [c for c in df.columns if 'çıkış date' in str(c).lower() or 'cikis date' in str(c).lower()]
    if cikis_cols_list:
        temp_cikis = df[cikis_cols_list].replace(r'^\s*$', np.nan, regex=True)
        df['Actual_Registered_Stops'] = temp_cikis.notna().sum(axis=1)
    else:
        df['Actual_Registered_Stops'] = 0 

    df['Date'] = parse_datetime(df['Start_Date_Raw']).dt.date
    df = df.dropna(subset=['Date', 'Schedule_Time'])
    df['Sched_Time_Str'] = df['Schedule_Time'].astype(str).str.split().str[-1]
    df['Sched_DT'] = pd.to_datetime(df['Date'].astype(str) + ' ' + df['Sched_Time_Str'], errors='coerce')

    invalid_stops_log, path_statistics_log = [], []
    path_tot_stops = df[tot_stops_col].max() if tot_stops_col else (max(arr_cols.keys()) + 1 if arr_cols else "Unknown")
    path_exp_time = "Unknown"
    
    def parse_plan_to_dt(val):
        if pd.isna(val) or str(val).strip().lower() == 'nan': return pd.NaT
        if isinstance(val, (int, float)): return pd.to_datetime('1899-12-30') + pd.Timedelta(days=float(val))
        return pd.to_datetime(str(val), errors='coerce')

    if plan_cols:
        try:
            max_plan_col = plan_cols[max(plan_cols.keys())]
            min_plan_col = plan_cols[min(plan_cols.keys())]
            max_p = df[max_plan_col].apply(parse_plan_to_dt)
            min_p = df[min_plan_col].apply(parse_plan_to_dt)
            diffs = (max_p - min_p).dt.total_seconds() / 60.0
            diffs = diffs.apply(lambda x: x + 24*60 if pd.notna(x) and x < 0 else x)
            path_exp_time = round(diffs.max(), 2) if not diffs.isna().all() else "Unknown"
        except: pass
    
    path_statistics_log.append({
        "Path": directional_path, "Total Expected Stops": path_tot_stops, "Expected Travel Time (Mins)": path_exp_time
    })

    def force_obu_date(val, base_date, sched_dt):
        if pd.isna(val): return pd.NaT
        if isinstance(val, pd.Timestamp) or hasattr(val, 'strftime'): time_str = val.strftime('%H:%M:%S')
        else:
            val_str = str(val).replace('nan','').strip()
            if not val_str: return pd.NaT
            time_str = val_str.split(' ')[-1] if ' ' in val_str else val_str
            
        dt = pd.to_datetime(f"{base_date} {time_str}", errors='coerce')
        if pd.notna(dt) and pd.notna(sched_dt):
            if (sched_dt - dt).total_seconds() > 12 * 3600: dt += pd.Timedelta(days=1)
            elif (dt - sched_dt).total_seconds() > 12 * 3600: dt -= pd.Timedelta(days=1)
        return dt

    df['Actual_DT'] = pd.NaT
    df['Last_End_DT'] = pd.NaT
    df['Is_Imputed'] = False
    df['Incomplete_Route'] = False

    for idx, r in df.iterrows():
        stops_data = {}
        max_idx = max(list(arr_cols.keys()) + list(cik_cols.keys()) + list(plan_cols.keys()) + [0])
        for i in range(max_idx + 1):
            arr_raw = r[arr_cols[i]] if i in arr_cols else np.nan
            cik_raw = r[cik_cols[i]] if i in cik_cols else np.nan
            arr = force_obu_date(arr_raw, r['Date'], r['Sched_DT'])
            cik = force_obu_date(cik_raw, r['Date'], r['Sched_DT'])
            plan_raw = r[plan_cols[i]] if i in plan_cols else np.nan
            plan = parse_plan_to_dt(plan_raw)
            if pd.notna(arr) or pd.notna(cik) or pd.notna(plan):
                stops_data[i] = {'arr': arr, 'cik': cik, 'plan': plan}

        prev_cik = pd.NaT
        for k in sorted(stops_data.keys()):
            arr = stops_data[k]['arr']
            cik = stops_data[k]['cik']
            link = make_tracking_url(r['Bus_ID'], r['Sched_DT'], r['Sched_DT'] + pd.Timedelta(hours=1), display_text=r['Schedule_Time'])
            
            if pd.notna(arr) and pd.notna(cik) and cik < arr:
                invalid_stops_log.append({"Date": r['Date'], "Path": directional_path, "Bus ID": r['Bus_ID'], "Schedule Time": r['Schedule_Time'], "Stop Index": k, "Issue": "Çıkış < Arrival", "Tracking Link": link})
            if pd.notna(prev_cik) and pd.notna(arr) and arr < prev_cik:
                invalid_stops_log.append({"Date": r['Date'], "Path": directional_path, "Bus ID": r['Bus_ID'], "Schedule Time": r['Schedule_Time'], "Stop Index": k, "Issue": f"Arrival < Previous Çıkış (Stop {k-1})", "Tracking Link": link})
            if pd.notna(cik): prev_cik = cik
        
        first_cik = stops_data.get(0, {}).get('cik', pd.NaT)
        first_plan = stops_data.get(0, {}).get('plan', pd.NaT)
        is_invalid_first = pd.isna(first_cik)
        if not is_invalid_first and 1 in stops_data:
            if pd.notna(stops_data[1]['arr']) and first_cik > stops_data[1]['arr']: is_invalid_first = True
            if pd.notna(stops_data[1]['cik']) and first_cik > stops_data[1]['cik']: is_invalid_first = True

        imputed_val = pd.NaT
        if is_invalid_first and pd.notna(first_plan):
            for k in sorted(stops_data.keys()):
                if k > 0:
                    k_cik = stops_data[k]['cik']
                    k_plan = stops_data[k]['plan']
                    if pd.notna(k_cik) and pd.notna(k_plan) and (pd.isna(stops_data[k]['arr']) or k_cik >= stops_data[k]['arr']):
                        plan_diff = (k_plan - first_plan).total_seconds() / 60.0
                        if plan_diff < 0: plan_diff += 24 * 60
                        imputed_val = k_cik - pd.Timedelta(minutes=plan_diff)
                        df.at[idx, 'Actual_DT'] = imputed_val
                        df.at[idx, 'Is_Imputed'] = True
                        break
        
        if pd.isna(df.at[idx, 'Actual_DT']): df.at[idx, 'Actual_DT'] = first_cik
            
        for k in reversed(sorted(stops_data.keys())):
            arr = stops_data[k]['arr']
            if pd.notna(arr):
                df.at[idx, 'Last_End_DT'] = arr
                break

        if tot_stops_col and pd.notna(r[tot_stops_col]):
            tot_s = pd.to_numeric(r[tot_stops_col], errors='coerce')
            if pd.notna(tot_s) and r['Actual_Registered_Stops'] < (tot_s - 1):
                df.at[idx, 'Incomplete_Route'] = True

    df['Actual_Travel_Time_Mins'] = (df['Last_End_DT'] - df['Actual_DT']).dt.total_seconds() / 60.0
    df.loc[df['Actual_Travel_Time_Mins'] < 0, 'Actual_Travel_Time_Mins'] += 24 * 60
        
    if col_map["End Time"]: 
        df['End_DT'] = df.apply(lambda x: force_obu_date(x[col_map["End Time"]], x['Date'], x['Sched_DT']), axis=1)
    else: 
        df['End_DT'] = df['Actual_DT'] + pd.Timedelta(hours=1)

    # ---- OPTIONAL SANITY CHECKS ----
    sanity_log, raw_sanity_log = [], []
    if params["run_sanity_checks"]:
        elapsed_cols = [c for c in df.columns if str(c).lower().strip() in ['elapsed time', 'duration', 'travel time', 'trip time', 'geçen süre']]
        elapsed_col = elapsed_cols[0] if elapsed_cols else None
        
        for idx, r in df.iterrows():
            comments = []
            st_raw = r.get(col_map["Actual Time"])
            en_raw = r.get(col_map["End Time"])
            v_str = r[elapsed_col] if elapsed_col and pd.notna(r.get(elapsed_col)) else np.nan
            rep_stops = r[col_map["Stop Count"]] if col_map.get("Stop Count") else np.nan
            calc_stops = r.get('Actual_Registered_Stops', 0)
            
            if pd.notna(rep_stops):
                try:
                    if float(rep_stops) != float(calc_stops):
                        comments.append(f"Stops Mismatch: STIO={int(float(rep_stops))}, Computed={int(calc_stops)}")
                except: pass
                
            try:
                if pd.notna(st_raw) and pd.notna(en_raw):
                    def get_exact_dt(val, fallback_date):
                        if pd.isna(val): return pd.NaT
                        if isinstance(val, (pd.Timestamp, datetime)): return pd.Timestamp(val)
                        s = str(val).strip()
                        dt = pd.to_datetime(s, dayfirst=True, errors='coerce')
                        if pd.notna(dt):
                            if '-' not in s and '.' not in s and '/' not in s:
                                dt = pd.to_datetime(f"{fallback_date} {s}", errors='coerce')
                            return dt
                        return pd.NaT
                        
                    st_dt = get_exact_dt(st_raw, r['Date'])
                    en_dt = get_exact_dt(en_raw, r['Date'])
                    
                    if pd.notna(st_dt) and pd.notna(en_dt):
                        diff_mins = (en_dt - st_dt).total_seconds() / 60.0
                        if en_dt < st_dt:
                            comments.append("End Time is before Start Time")
                            diff_mins = abs(diff_mins) 
                        
                        if pd.notna(v_str):
                            el_mins = None
                            v_clean = str(v_str).strip()
                            if ':' in v_clean:
                                parts = v_clean.split(':')
                                if len(parts) == 3: el_mins = int(parts[0])*60 + int(parts[1]) + float(parts[2])/60.0
                                elif len(parts) == 2: el_mins = int(parts[0])*60 + float(parts[1])
                            else:
                                try: el_mins = float(v_clean)
                                except: pass
                            if el_mins is not None and abs(el_mins - diff_mins) > 2.0:
                                comments.append(f"Time Mismatch: STIO Col={v_clean}, Actual Diff={round(diff_mins, 1)}m")
            except: pass
            
            if comments:
                reason = " | ".join(comments)
                link = make_tracking_url(r['Bus_ID'], r['Sched_DT'], r.get('End_DT', r['Sched_DT']), display_text=r['Schedule_Time'])
                sanity_log.append({
                    "Date": r['Date'], "Path": directional_path, "Bus ID": r['Bus_ID'], "Schedule Time": r['Schedule_Time'], 
                    "STIO Start Time": st_raw if pd.notna(st_raw) else "N/A", "STIO End Time": en_raw if pd.notna(en_raw) else "N/A",
                    "STIO Elapsed Time": v_str if pd.notna(v_str) else "N/A", "STIO Trip Count": rep_stops if pd.notna(rep_stops) else "N/A",
                    "Actual Registered Stops": calc_stops, "Sanity Check Comment": reason, "Tracking Link": link
                })
                raw_sanity_log.append(make_raw(r, reason, link))
        
        return {
            "summary": [], "eff_details": [], "punct_details": [], "reg_details": [],
            "assigned": [], "trips": [], "reg_trips": [], "obu_errors": [], "cleaning": [], 
            "missed_trips": [], "accepted_trips": [], "bunched_trips": [], "strict_late_log": [], 
            "invalid_stops_log": [], "path_statistics_log": [], "auto_rush": [], "sanity_checks": sanity_log, 
            "raw_sanity_checks": raw_sanity_log, "raw_strict_late_log": [], "raw_cleaning": [], 
            "raw_obu": [], "raw_bunched": [], "raw_accepted": [], "raw_on_time": [], 
            "raw_off_time": [], "raw_reg_regular": [], "raw_reg_irregular": [], "ignored_trips": []
        }
    # ---------------------------------------------------
        
    has_master = authoritative_dict is not None and directional_path in authoritative_dict
    daily_master_schedules = {}
    
    for d, g in df.groupby('Date'):
        if has_master:
            master_times_for_date = []
            for sched_block in authoritative_dict[directional_path]:
                if sched_block['start'] <= d <= sched_block['end']:
                    master_times_for_date = sched_block['times']
                    break
            
            if master_times_for_date:
                valid_scheds = [pd.to_datetime(f"{d} {t}", errors='coerce') for t in master_times_for_date if pd.notna(pd.to_datetime(f"{d} {t}", errors='coerce'))]
            else:
                valid_scheds = g['Sched_DT'].dropna().unique()
        else:
            valid_scheds = g['Sched_DT'].dropna().unique()
            
        daily_master_schedules[d] = sorted([pd.Timestamp(ts) for ts in valid_scheds])

    headway_dict = {}
    for d, scheds in daily_master_schedules.items():
        prev_st = None
        for st in scheds:
            if prev_st is not None: headway_dict[(d, st)] = (st - prev_st).total_seconds() / 60.0
            else: headway_dict[(d, st)] = None
            prev_st = st

    cleaning_log, assigned_log, eff_details, punct_details, reg_details, kpi_results = [], [], [], [], [], []
    trip_log, trip_reg_log, obu_errors_log, missed_log, accepted_log, bunched_log = [], [], [], [], [], []
    raw_cleaning, raw_obu, raw_bunched, raw_accepted = [], [], [], []
    raw_on_time, raw_off_time, raw_reg_regular, raw_reg_irregular = [], [], [], []
    strict_late_log, raw_strict_late_log, auto_rush_periods_log = [], [], []

    min_km, min_stops = params["min_km"], params["min_stops"]
    df["Realized_KM"] = pd.to_numeric(df["Realized_KM"], errors="coerce")

    def log_cleaning(r, reason, link, acc_status):
        return {
            "Date": r['Date'], "Path": directional_path, "Bus": r['Bus_ID'], "Schedule Time": r['Schedule_Time'],
            "Stop Count (OBU)": r[col_map["Stop Count"]] if col_map.get("Stop Count") and col_map["Stop Count"] in r else "N/A",
            "Actual Registered Stops": r.get('Actual_Registered_Stops', 0), "Realized KM": r.get('Realized_KM', 0),
            "Time Taken (Mins)": round(r['Actual_Travel_Time_Mins'], 2) if pd.notna(r.get('Actual_Travel_Time_Mins')) else "N/A",
            "Acceptance Status": acc_status, "Reason": reason, "Tracking Link": link
        }

    if col_map.get("Stop Count"):
        df[col_map["Stop Count"]] = pd.to_numeric(df[col_map["Stop Count"]], errors="coerce")
        stops_filled = df[col_map["Stop Count"]].fillna(0)
        has_enough_stops = (df['Actual_Registered_Stops'] >= min_stops) | (stops_filled >= min_stops)
    else:
        has_enough_stops = df['Actual_Registered_Stops'] >= min_stops

    km_mask = df["Realized_KM"] < min_km
    low_km_drop = km_mask & ~has_enough_stops
    low_km_keep = km_mask & has_enough_stops
    stop_mask = (df["Realized_KM"] >= min_km) & ~has_enough_stops
    valid_mask = ~(low_km_drop | stop_mask)

    km_mode_series = df.loc[valid_mask, "Realized_KM"].mode()
    km_mode = km_mode_series.iloc[0] if not km_mode_series.empty else 0
    df['Trip_Weight'] = 1
    double_trip_mask = pd.Series(False, index=df.index)
    if km_mode > 0:
        double_trip_mask = (df["Realized_KM"] > km_mode * 2.5) & valid_mask
        df.loc[double_trip_mask, 'Trip_Weight'] = 2

    df.loc[~valid_mask, 'Trip_Weight'] = 0
    no_start_mask = valid_mask & df['Actual_DT'].isna()
    rmean = df.loc[valid_mask, "Realized_KM"].mean()
    rstd = df["Realized_KM"].std() if pd.notna(df["Realized_KM"].std()) else 0
    outlier_log_mask = (df["Realized_KM"] > (rmean + 3 * rstd)) & valid_mask & ~double_trip_mask
    df['Is_Outlier'] = (df["Realized_KM"] > (rmean + 3 * rstd)) | double_trip_mask

    broken_trip_mask = (~has_enough_stops) & (df['Actual_Travel_Time_Mins'] > 2.0)
    odo_start_col = lower_cols.get("odometer start km")
    odo_end_col = lower_cols.get("odometer end km")
    odometer_zero_mask = pd.Series(False, index=df.index)
    if odo_start_col in df.columns: odometer_zero_mask = odometer_zero_mask | (pd.to_numeric(df[odo_start_col], errors='coerce') == 0)
    if odo_end_col in df.columns: odometer_zero_mask = odometer_zero_mask | (pd.to_numeric(df[odo_end_col], errors='coerce') == 0)

    for idx, r in df.iterrows():
        reasons = []
        if low_km_drop.loc[idx]: reasons.append(f"Filtered: REALIZED_KM ({r.get('Realized_KM',0)}) below min")
        if stop_mask.loc[idx]: reasons.append(f"Filtered: Ghost Trip (Stops below min)")
        if low_km_keep.loc[idx]: reasons.append(f"Flagged: Low KM but Actual Stops are significant")
        if double_trip_mask.loc[idx]: reasons.append(f"Flagged: Distance > 2.5x Mode. Counted as 2 trips")
        if no_start_mask.loc[idx]: reasons.append(f"Flagged: Missing First Start Time")
        if outlier_log_mask.loc[idx]: reasons.append(f"Flagged: Distance Outlier")
        if broken_trip_mask.loc[idx]: reasons.append(f"Flagged: Possible broken trip split in two")
        if odometer_zero_mask.loc[idx]: reasons.append(f"Flagged: Odometer zero")
        
        if reasons:
            reason_str = " | ".join(reasons)
            link = make_tracking_url(r['Bus_ID'], r['Sched_DT'], r['End_DT'], display_text=r['Schedule_Time'])
            acc_status = f"Accepted as {int(r['Trip_Weight'])} trip{'s' if int(r['Trip_Weight']) != 1 else ''}"
            cleaning_log.append(log_cleaning(r, reason_str, link, acc_status))
            raw_entry = make_raw(r, reason_str, link)
            raw_entry['Acceptance Status'] = acc_status
            raw_cleaning.append(raw_entry)

    df_clean = df[valid_mask].copy()
    df_clean = df_clean.sort_values(by=['Date', 'Bus_ID', 'Sched_DT', 'Is_Outlier', 'Actual_DT'], ascending=[True, True, True, True, True])
    dup_mask_clean = df_clean.duplicated(subset=['Date', 'Bus_ID', 'Sched_DT'], keep='first')
    
    for _, r in df_clean[dup_mask_clean].iterrows():
        kept_row = df_clean[(~dup_mask_clean) & (df_clean['Date'] == r['Date']) & (df_clean['Bus_ID'] == r['Bus_ID']) & (df_clean['Sched_DT'] == r['Sched_DT'])]
        kept_time_val = kept_row['Actual_DT'].iloc[0] if not kept_row.empty else "Unknown"
        reason = f"Filtered: Duplicate Schedule Selection. Discarded Time: {r['Actual_DT']} | Kept Earlier: {kept_time_val}."
        link = make_tracking_url(r['Bus_ID'], r['Sched_DT'], r['End_DT'], display_text=r['Schedule_Time'])
        acc_status = "Accepted as 0 trips"
        cleaning_log.append(log_cleaning(r, reason, link, acc_status))
        raw_entry = make_raw(r, reason, link)
        raw_entry['Acceptance Status'] = acc_status
        raw_cleaning.append(raw_entry)
        
    df_clean = df_clean[~dup_mask_clean].copy()
    df_clean['Corrected_Sched_DT'] = df_clean['Sched_DT']
    df_clean['OBU_Error'] = False

    for idx, r in df_clean.iterrows():
        act = r['Actual_DT']
        obu = r['Sched_DT']
        d = r['Date']
        
        if r.get('Is_Imputed'):
            cand_scheds = []
            for cand_d in [d - pd.Timedelta(days=1), d, d + pd.Timedelta(days=1)]:
                cand_scheds.extend(daily_master_schedules.get(cand_d, []))
            master_scheds = cand_scheds if cand_scheds else daily_master_schedules.get(d, [])
        else:
            master_scheds = daily_master_schedules.get(d, [])

        if len(master_scheds) > 0 and pd.notna(act):
            diffs = np.abs(pd.Series(master_scheds) - act)
            closest_sched = master_scheds[diffs.idxmin()]
        elif len(master_scheds) > 0 and pd.notna(obu):
            diffs = np.abs(pd.Series(master_scheds) - obu)
            closest_sched = master_scheds[diffs.idxmin()]
        else:
            closest_sched = obu
            
        df_clean.at[idx, 'Corrected_Sched_DT'] = closest_sched
        
        if pd.notna(act) and closest_sched != obu:
            df_clean.at[idx, 'OBU_Error'] = True
            delay_obu = pd.Timedelta(act - obu).total_seconds() / 60.0 if pd.notna(obu) else 0
            delay_corr = pd.Timedelta(act - closest_sched).total_seconds() / 60.0
            link = make_tracking_url(r['Bus_ID'], obu, r['End_DT'], display_text=r['Schedule_Time'])
            obu_errors_log.append({
                "Date": d, "Path": directional_path, "Bus_ID": r['Bus_ID'],
                "OBU Selected Schedule": r['Schedule_Time'], "Corrected Schedule": closest_sched,
                "Actual Departure": act, "Delay vs OBU (Mins)": round(delay_obu, 2),
                "Delay vs Corrected (Mins)": round(delay_corr, 2), "Tracking Link": link
            })
            raw_obu.append(make_raw(r, f"OBU Error: Selected {r['Schedule_Time']} -> Corrected to {closest_sched}", link))

    df_clean['Delay_Mins'] = (df_clean['Actual_DT'] - df_clean['Corrected_Sched_DT']).dt.total_seconds() / 60.0

    def get_hw(row): return headway_dict.get((row['Date'], row['Corrected_Sched_DT']))
    df_clean['Headway_Mins'] = df_clean.apply(get_hw, axis=1)

    def get_delay_pct(row):
        hw = row['Headway_Mins']
        dm = row['Delay_Mins']
        if pd.notna(dm) and hw and hw > 0: return (dm / hw) * 100
        return np.nan
    df_clean['Delay_Pct'] = df_clean.apply(get_delay_pct, axis=1)
    df_clean = df_clean.sort_values(by=['Date', 'Corrected_Sched_DT', 'Actual_DT']).reset_index(drop=True)

    early_l_fixed = params["early_limit"]
    late_l_fixed = params["late_limit"]
    use_hw_pct = params["use_headway_pct"]
    early_pct = params["early_headway_pct"] / 100.0
    late_pct = params["late_headway_pct"] / 100.0

    def flag_punctuality(row):
        if row.get('Is_Imputed'): return "Imputed (Excluded)"
        delay = row['Delay_Mins']
        if pd.isna(delay): return "Unknown"
        d_early, d_late = early_l_fixed, late_l_fixed
        if use_hw_pct:
            hw = row['Headway_Mins']
            if hw is not None and hw > 0:
                d_early, d_late = hw * early_pct, hw * late_pct
        if delay < -d_early: return "Early"
        if delay > d_late: return "Late"
        return "On-Time"
        
    df_clean['Punctuality_Status'] = df_clean.apply(flag_punctuality, axis=1)

    fast_quotas = {}
    if intimations_dict:
        for (d, p), count in intimations_dict.items():
            if count < 0 and p == normalize_path(directional_path): fast_quotas[d] = abs(count)

    # Initial Regularity Status Assignment
    for idx, r in df_clean.iterrows():
        tt = r['Actual_Travel_Time_Mins']
        reg_status = "Unknown"
        act_time = r['Actual_DT'].time() if pd.notna(r['Actual_DT']) else None
        curr_rush_windows = rush_hours_dict.get((r['Date'], normalize_path(directional_path)), []) if rush_hours_dict else []
        in_rush_window = False
        if act_time:
            for (rst, ret) in curr_rush_windows:
                if rst <= act_time <= ret:
                    in_rush_window = True
                    break

        if r.get('Incomplete_Route'): reg_status = "Incomplete Route (Excluded)"
        elif r['Is_Outlier']: reg_status = "Distance Outlier (Excluded)"
        elif pd.notna(tt) and target_time is not None and target_time > 0:
            imp_fast = target_time * (1 - params["reg_imp_fast_pct"] / 100.0)
            imp_slow = target_time * (1 + params["reg_imp_slow_pct"] / 100.0)
            min_time = target_time * (1 - params["reg_fast_pct"] / 100.0)
            max_time = target_time * (1 + params["reg_slow_pct"] / 100.0)
            
            if target_time_fast and fast_quotas.get(r['Date'], 0) > 0:
                imp_fast_allowed = target_time_fast * (1 - params["reg_imp_fast_pct"] / 100.0)
                imp_fast = min(imp_fast, imp_fast_allowed)
            
            if tt < imp_fast or tt > imp_slow: reg_status = "Impossible Time (Excluded)"
            elif tt < min_time: 
                if target_time_fast and target_time_fast > 0 and fast_quotas.get(r['Date'], 0) > 0:
                    min_time_fast = target_time_fast * (1 - params["reg_fast_pct"] / 100.0)
                    if tt >= min_time_fast:
                        reg_status = "Regular (Approved Fast)"
                        fast_quotas[r['Date']] -= 1
                    else: reg_status = "Too Fast"
                else: reg_status = "Too Fast"
            elif tt > max_time: 
                if in_rush_window: reg_status = "Rush Hour Excluded (Slow)"
                else: reg_status = "Too Slow"
            else: reg_status = "Regular"
        else:
            if target_time is None: reg_status = "Missing Target Time"
            else: reg_status = "Missing Travel Time"

        df_clean.at[idx, 'Regularity_Status'] = reg_status

    # Apply Auto-Rush & Bunching Exclusions
    for date_val, group in df_clean.groupby('Date'):
        # 1. Punctuality Bunching Exemption
        sched_counts = group['Corrected_Sched_DT'].value_counts()
        bunched_times = sched_counts[sched_counts > 1].index
        for st_val in bunched_times:
            bunch_group = group[group['Corrected_Sched_DT'] == st_val]
            # Identify the single trip with the absolute minimum delay to keep evaluating
            best_idx = bunch_group['Delay_Mins'].abs().idxmin()
            for idx in bunch_group.index:
                if idx != best_idx:
                    df_clean.at[idx, 'Punctuality_Status'] = "Bunched (Excluded)"
        
        # 2. Auto Rush Hour Exemption
        if params["auto_rush_enable"]:
            group_sorted = df_clean[df_clean['Date'] == date_val].sort_values(by=['Corrected_Sched_DT', 'Actual_DT'])
            consec_indices = []
            for idx, r in group_sorted.iterrows():
                if r['Regularity_Status'] in ["Too Slow", "Rush Hour Excluded (Slow)"]:
                    consec_indices.append(idx)
                else:
                    if len(consec_indices) >= params["auto_rush_thresh"]:
                        for ci in consec_indices:
                            if df_clean.at[ci, 'Regularity_Status'] == "Too Slow":
                                df_clean.at[ci, 'Regularity_Status'] = "Auto Rush Excluded (Slow)"
                    consec_indices = []
            if len(consec_indices) >= params["auto_rush_thresh"]:
                for ci in consec_indices:
                    if df_clean.at[ci, 'Regularity_Status'] == "Too Slow":
                        df_clean.at[ci, 'Regularity_Status'] = "Auto Rush Excluded (Slow)"

    ignored_trips_log = []

    # Final Extraction & Detail Logging
    for _, r in df_clean.iterrows():
        link = make_tracking_url(r['Bus_ID'], r['Sched_DT'], r['End_DT'], display_text=r['Schedule_Time'])
        punct_status = r['Punctuality_Status']
        reg_status = r['Regularity_Status']
        
        # Determine if Ignored
        is_ignored = False
        reasons = []
        if "(Excluded)" in punct_status:
            is_ignored = True
            reasons.append(f"Punctuality: {punct_status}")
        if "(Excluded)" in reg_status:
            is_ignored = True
            reasons.append(f"Regularity: {reg_status}")
            
        if is_ignored:
            ignored_trips_log.append({
                "Date": r['Date'], "Path": directional_path, "Bus_ID": r['Bus_ID'],
                "Original Schedule": r['Schedule_Time'], "Corrected Schedule": r['Corrected_Sched_DT'],
                "Actual Departure": r['Actual_DT'], "Ignored Reason": " | ".join(reasons),
                "Tracking Link": link
            })
            
        trip_log.append({
            "Date": r['Date'], "Path": directional_path, "Bus_ID": r['Bus_ID'],
            "Original OBU Schedule": r['Schedule_Time'], "Corrected Schedule": r['Corrected_Sched_DT'],
            "Actual Departure": r['Actual_DT'], "Is Imputed Start": r.get('Is_Imputed', False),
            "Headway (Mins)": round(r['Headway_Mins'], 2) if pd.notna(r['Headway_Mins']) else None,
            "Delay (%)": round(r['Delay_Pct'], 2) if pd.notna(r['Delay_Pct']) else None,
            "Status": punct_status, "Tracking Link": link
        })
        
        if punct_status == "On-Time": raw_on_time.append(make_raw(r, "On-Time", link))
        elif punct_status in ["Early", "Late"]: raw_off_time.append(make_raw(r, punct_status, link))
        
        tt = r['Actual_Travel_Time_Mins']
        trip_reg_log.append({
            "Date": r['Date'], "Path": directional_path, "Bus_ID": r['Bus_ID'],
            "OBU Schedule": r['Schedule_Time'], "Corrected Schedule": r['Corrected_Sched_DT'],
            "First Start Time": r['Actual_DT'].strftime('%H:%M:%S') if pd.notna(r['Actual_DT']) else "N/A",
            "Last End Time": r['Last_End_DT'].strftime('%H:%M:%S') if pd.notna(r['Last_End_DT']) else "N/A",
            "Actual Travel Time (Mins)": round(tt, 2) if pd.notna(tt) else None,
            "Target Matrix Time (Mins)": target_time, "Status": reg_status, "Tracking Link": link
        })
        
        if reg_status in ["Regular", "Regular (Approved Fast)"]: raw_reg_regular.append(make_raw(r, f"{reg_status} Travel Time", link))
        elif reg_status in ["Too Fast", "Too Slow"]: raw_reg_irregular.append(make_raw(r, reg_status, link))


    e_up, e_tgt, e_bot = params["eff_upper"], params["eff_target"], params["eff_bottom"]
    pen_high, pen_low = params["eff_pen_high"], params["eff_pen_low"]
    p_up, p_bot, p_pen_const_val = params["punct_upper"], params["punct_bottom"], params["punct_pen"]
    r_up, r_bot, r_pen_const_val = params["reg_upper"], params["reg_bottom"], params["reg_pen"]

    for date_val, group in df_clean.groupby('Date'):
        assigned_schedules = daily_master_schedules.get(date_val, [])
        operated_schedules = set([pd.Timestamp(ts) for ts in group['Corrected_Sched_DT'] if pd.notna(ts)])
        prev_st = None
        
        sched_counts = group['Corrected_Sched_DT'].value_counts()
        bunched_times = sched_counts[sched_counts > 1].index
        
        consecutive_slow = []
        for _, r in group.iterrows():
            if r['Regularity_Status'] in ["Too Slow", "Auto Rush Excluded (Slow)"]:
                consecutive_slow.append(r)
            else:
                if len(consecutive_slow) >= params["auto_rush_thresh"]:
                    st_time = consecutive_slow[0]['Actual_DT'].strftime('%H:%M:%S')
                    en_time = consecutive_slow[-1]['Actual_DT'].strftime('%H:%M:%S')
                    trip_links = [make_tracking_url(x['Bus_ID'], x['Sched_DT'], x['End_DT'], display_text=x['Schedule_Time'], raw_url_only=True) for x in consecutive_slow]
                    auto_rush_periods_log.append({
                        "Date": date_val, "Path": directional_path, "Period Start": st_time, "Period End": en_time,
                        "Consecutive Slow Trips": len(consecutive_slow), "Tracking Links": " | ".join(trip_links)
                    })
                consecutive_slow = []
                
        if len(consecutive_slow) >= params["auto_rush_thresh"]:
            st_time = consecutive_slow[0]['Actual_DT'].strftime('%H:%M:%S')
            en_time = consecutive_slow[-1]['Actual_DT'].strftime('%H:%M:%S')
            trip_links = [make_tracking_url(x['Bus_ID'], x['Sched_DT'], x['End_DT'], display_text=x['Schedule_Time'], raw_url_only=True) for x in consecutive_slow]
            auto_rush_periods_log.append({
                "Date": date_val, "Path": directional_path, "Period Start": st_time, "Period End": en_time,
                "Consecutive Slow Trips": len(consecutive_slow), "Tracking Links": " | ".join(trip_links)
            })
        
        for st in bunched_times:
            bunch_group = group[group['Corrected_Sched_DT'] == st]
            for _, r in bunch_group.iterrows():
                link = make_tracking_url(r['Bus_ID'], r['Sched_DT'], r['End_DT'], display_text=r['Schedule_Time'])
                bunched_log.append({
                    "Date": date_val, "Path": directional_path, "Bunched Schedule Target": st,
                    "OBU Selected Schedule": r['Schedule_Time'], "Bus ID": r['Bus_ID'], "Actual Departure": r['Actual_DT'],
                    "Start Time Source": "Imputed" if r.get('Is_Imputed') else "OBU Recorded",
                    "Headway (Mins)": round(r['Headway_Mins'], 2) if pd.notna(r['Headway_Mins']) else None,
                    "Delay (%)": round(r['Delay_Pct'], 2) if pd.notna(r['Delay_Pct']) else None,
                    "Punctuality Status": r['Punctuality_Status'], "Tracking Link": link
                })
                raw_bunched.append(make_raw(r, f"Bunched at Target: {st}", link))

        filtered_st_for_date = set()
        for entry in cleaning_log:
            if entry["Date"] == date_val and entry["Path"] == directional_path:
                if "Filtered:" in entry["Reason"]:
                    try: filtered_st_for_date.add(pd.to_datetime(f"{date_val} {entry['Schedule Time']}"))
                    except: pass

        for st in assigned_schedules:
            headway = None
            if prev_st is not None: headway = round(pd.Timedelta(st - prev_st).total_seconds() / 60.0, 2)
                
            matching_trips = group[group['Corrected_Sched_DT'] == st]
            is_missed = len(matching_trips) == 0
            is_filtered = st in filtered_st_for_date
            is_bunched = len(matching_trips) > 1
            
            punct_list, reg_list, corr_list = [], [], []
            for _, r in matching_trips.iterrows():
                delay = r['Delay_Mins']
                status = r['Punctuality_Status']
                if r.get('Is_Imputed'): punct_list.append("IMPUTED EXCLUDED")
                elif pd.notna(delay):
                    d_int = int(round(delay))
                    sign = "+" if d_int > 0 else ""
                    if status == "On-Time": punct_list.append("TRUE" if d_int == 0 else f"{sign}{d_int}")
                    else: punct_list.append(f"{sign}{d_int} {status.upper()}")
                else: punct_list.append("Unknown")
                    
                reg = r['Regularity_Status']
                tt = r['Actual_Travel_Time_Mins']
                if reg in ["Regular", "Regular (Approved Fast)"]: reg_list.append("TRUE")
                else:
                    if pd.notna(tt): reg_list.append(f"{reg} - {int(round(tt))} MIN")
                    else: reg_list.append(reg)
                        
                if r['OBU_Error']: corr_list.append(str(r['Schedule_Time']))
                else: corr_list.append("FALSE")
                    
            corr_str = "FALSE" if all(c == "FALSE" for c in corr_list) and corr_list else (" | ".join(corr_list) if corr_list else "FALSE")
            if is_missed: corr_str = "N/A"
            
            assigned_log.append({
                "Date": date_val, "Path": directional_path, "Expected Schedule Time": st, "Headways in mins": headway,
                "Missed": is_missed, "Punctual": " | ".join(punct_list) if punct_list else "N/A",
                "Bunched": is_bunched, "Regular": " | ".join(reg_list) if reg_list else "N/A",
                "Corrected Departure": corr_str, "OBU Time in Filtered": is_filtered
            })
            
            if st not in operated_schedules:
                if has_master: missed_log.append({"Date": date_val, "Path": directional_path, "Missed Schedule Time": st})
            else:
                for _, r in matching_trips.iterrows():
                    link = make_tracking_url(r['Bus_ID'], r['Sched_DT'], r['End_DT'], display_text=r['Schedule_Time'])
                    accepted_log.append({
                        "Date": date_val, "Path": directional_path, "Accepted Schedule Time": st,
                        "OBU Selected Schedule": r['Schedule_Time'], "Bus ID": r['Bus_ID'], "Realized KM": r['Realized_KM'],
                        "Headway (Mins)": round(r['Headway_Mins'], 2) if pd.notna(r['Headway_Mins']) else None,
                        "Delay (%)": round(r['Delay_Pct'], 2) if pd.notna(r['Delay_Pct']) else None,
                        "Punctuality Status": r['Punctuality_Status'], "Tracking Link": link
                    })
                    raw_accepted.append(make_raw(r, f"Accepted for Target: {st}", link))
            prev_st = st
        
        assigned_count_raw = len(assigned_schedules)
        intimations = intimations_dict.get((date_val, normalize_path(directional_path)), 0) if intimations_dict else 0
        corrections = corrections_dict.get((date_val, normalize_path(directional_path)), 0) if corrections_dict else 0
        
        assigned_count_eff = max(0, assigned_count_raw - intimations)
        operated_count_gross = group['Trip_Weight'].sum()
        
        strict_late_deduction = 0
        if params["strict_eff"]:
            strict_mask = (group['Delay_Pct'] > 50) & (group['Punctuality_Status'] == 'Late')
            strict_late_trips = group[strict_mask]
            strict_late_deduction = strict_late_trips['Trip_Weight'].sum()
            
            for _, r in strict_late_trips.iterrows():
                link = make_tracking_url(r['Bus_ID'], r['Sched_DT'], r['End_DT'], display_text=r['Schedule_Time'])
                strict_late_log.append({
                    "Date": date_val, "Path": directional_path, "Expected Schedule": r['Corrected_Sched_DT'],
                    "OBU Selected Schedule": r['Schedule_Time'], "Bus ID": r['Bus_ID'], "Actual Departure": r['Actual_DT'],
                    "Headway (Mins)": round(r['Headway_Mins'], 2) if pd.notna(r['Headway_Mins']) else None,
                    "Delay (%)": round(r['Delay_Pct'], 2) if pd.notna(r['Delay_Pct']) else None, "Tracking Link": link
                })
                raw_strict_late_log.append(make_raw(r, "Strict Efficiency Deduction (>50% Late)", link))
        
        operated_count_eff = max(0, operated_count_gross - strict_late_deduction + corrections)
        missed_volume = max(0, assigned_count_eff - operated_count_eff)
        
        covered_slots_count = len(operated_schedules)
        uncovered_slots_count = assigned_count_raw - covered_slots_count
        bunched_trips_count = operated_count_gross - covered_slots_count
        
        extra_approved = abs(intimations) if intimations < 0 else 0
        penalized_bunched_trips = max(0, bunched_trips_count - extra_approved)
        
        eff_percent = (operated_count_eff / assigned_count_eff * 100) if assigned_count_eff > 0 else 0
        if eff_percent >= e_up: eff_pen_const = 0
        elif eff_percent >= e_tgt: eff_pen_const = pen_high
        elif eff_percent >= e_bot: eff_pen_const = pen_low
        else: eff_pen_const = pen_low
        eff_penalty = missed_volume * eff_pen_const

        operated_count_strict_bunching = max(0, operated_count_eff - penalized_bunched_trips)
        missed_volume_strict_bunching = max(0, assigned_count_eff - operated_count_strict_bunching)
        eff_percent_sb = (operated_count_strict_bunching / assigned_count_eff * 100) if assigned_count_eff > 0 else 0
        
        dist_factor = params["eff_pen_low"]
        if path_distances and normalize_path(directional_path) in path_distances:
            dist_factor = path_distances[normalize_path(directional_path)]
            
        eff_pen_const_sb = dist_factor
        eff_penalty_sb = missed_volume_strict_bunching * eff_pen_const_sb
        
        eff_details.append({
            "Date": date_val, "Path": directional_path, "Total Assigned Trips (Base)": assigned_count_raw, 
            "Intimations (Exempted)": intimations, "Net Target Assigned": assigned_count_eff,
            "Total Operated Trips (Gross Weighted)": operated_count_gross, 
            "Corrections (Wrong OBU)": corrections,
            "Strict Late Deductions": strict_late_deduction,
            "Net Operated Trips": operated_count_eff, "Volume Deficit": missed_volume, 
            "Uncovered Slots": uncovered_slots_count, "Bunched/Overlapped Trips": penalized_bunched_trips,
            "Efficiency Achieved (%)": round(eff_percent, 2), "Penalty Factor": eff_pen_const, "Calculated Efficiency Penalty": eff_penalty
        })
        
        on_time = group.loc[group['Punctuality_Status'] == "On-Time", 'Trip_Weight'].sum()
        early = group.loc[group['Punctuality_Status'] == "Early", 'Trip_Weight'].sum()
        late = group.loc[group['Punctuality_Status'] == "Late", 'Trip_Weight'].sum()
        off_time = early + late
        
        punct_evaluated_count = on_time + off_time
        punct_percent = (on_time / punct_evaluated_count * 100) if punct_evaluated_count > 0 else 0
        if punct_percent >= p_up: punct_pen_const = 0
        elif punct_percent >= p_bot: punct_pen_const = p_pen_const_val
        else: punct_pen_const = p_pen_const_val
        punct_penalty = off_time * punct_pen_const
        
        punct_details.append({
            "Date": date_val, "Path": directional_path, "Total Evaluated Trips (Weighted)": punct_evaluated_count,
            "On-Time Trips": on_time, "Early Trips": early, "Late Trips": late, "Total Off-Time Trips": off_time,
            "Punctuality Achieved (%)": round(punct_percent, 2), "Penalty Per Off-Time": punct_pen_const, "Calculated Punctuality Penalty": punct_penalty
        })
        
        reg_eval_mask = group['Regularity_Status'].isin(["Regular", "Regular (Approved Fast)", "Too Fast", "Too Slow"])
        reg_eval_count = group.loc[reg_eval_mask, 'Trip_Weight'].sum()
        reg_regular_count = group.loc[group['Regularity_Status'].isin(["Regular", "Regular (Approved Fast)"]), 'Trip_Weight'].sum()
        reg_fast = group.loc[group['Regularity_Status'] == "Too Fast", 'Trip_Weight'].sum()
        reg_slow = group.loc[group['Regularity_Status'] == "Too Slow", 'Trip_Weight'].sum()
        off_reg = reg_fast + reg_slow
        reg_rush_excluded = group.loc[group['Regularity_Status'] == "Rush Hour Excluded (Slow)", 'Trip_Weight'].sum()
        reg_approved_fast = group.loc[group['Regularity_Status'] == "Regular (Approved Fast)", 'Trip_Weight'].sum()
        
        reg_percent = (reg_regular_count / reg_eval_count * 100) if reg_eval_count > 0 else 0
        if reg_percent >= r_up: r_pen_const = 0
        elif reg_percent >= r_bot: r_pen_const = r_pen_const_val
        else: r_pen_const = r_pen_const_val
        reg_penalty = off_reg * r_pen_const

        valid_tt = group[reg_eval_mask].dropna(subset=['Actual_Travel_Time_Mins'])
        if not valid_tt.empty:
            max_idx = valid_tt['Actual_Travel_Time_Mins'].idxmax()
            min_idx = valid_tt['Actual_Travel_Time_Mins'].idxmin()
            daily_max = round(valid_tt.loc[max_idx, 'Actual_Travel_Time_Mins'], 2)
            daily_max_link = make_tracking_url(valid_tt.loc[max_idx, 'Bus_ID'], valid_tt.loc[max_idx, 'Sched_DT'], valid_tt.loc[max_idx, 'End_DT'])
            daily_min = round(valid_tt.loc[min_idx, 'Actual_Travel_Time_Mins'], 2)
            daily_min_link = make_tracking_url(valid_tt.loc[min_idx, 'Bus_ID'], valid_tt.loc[min_idx, 'Sched_DT'], valid_tt.loc[min_idx, 'End_DT'])
            daily_avg = round(valid_tt['Actual_Travel_Time_Mins'].mean(), 2)
        else:
            daily_max, daily_max_link, daily_min, daily_min_link, daily_avg = None, "N/A", None, "N/A", None
        
        reg_details.append({
            "Date": date_val, "Path": directional_path, "Total Evaluated Trips (Weighted)": reg_eval_count,
            "Regular Trips": reg_regular_count, "Too Fast Trips": reg_fast, "Too Slow Trips": reg_slow, "Total Irregular Trips": off_reg,
            "Rush Hour Excluded Trips": reg_rush_excluded, "Approved Fast Trips": reg_approved_fast,
            "Daily Max Time (Mins)": daily_max, "Daily Max Tracking Link": daily_max_link,
            "Daily Min Time (Mins)": daily_min, "Daily Min Tracking Link": daily_min_link,
            "Daily Average Time (Mins)": daily_avg, "Path Overall Avg Time (Mins)": None, "Path Weekday Avg Time (Mins)": None,  
            "Regularity Achieved (%)": round(reg_percent, 2), "Penalty Per Irregularity": r_pen_const, "Calculated Regularity Penalty": reg_penalty
        })
        
        kpi_results.append({
            "Date": date_val, "Path": directional_path, 
            "Base Assigned": assigned_count_raw, "Intimations": intimations, "Net Target Assigned": assigned_count_eff,
            "Gross Operated": operated_count_gross, "Corrections (Wrong OBU)": corrections, 
            "Strict Deductions": strict_late_deduction, "Net Operated": operated_count_eff,
            "Volume Deficit": missed_volume, "Uncovered Slots": uncovered_slots_count, 
            "Bunched Trips": penalized_bunched_trips, "Approved Extra (Neg. Intimations)": extra_approved,
            "Efficiency (%)": round(eff_percent, 2), "Eff Penalty Factor": eff_pen_const, "Eff Penalty": eff_penalty,
            "Strict Bunching Net Operated": operated_count_strict_bunching, 
            "Strict Bunching Eff (%)": round(eff_percent_sb, 2), "Strict Bunching Penalty": eff_penalty_sb,
            "On-Time": on_time + extra_approved, "Off-Time": off_time, 
            "Punctuality (%)": round((on_time + extra_approved) / (on_time + extra_approved + off_time) * 100 if (on_time + extra_approved + off_time) > 0 else 0, 2), 
            "Punct Penalty Factor": punct_pen_const, "Punct Penalty": punct_penalty,
            "Regularity (%)": round(reg_percent, 2), "Reg Penalty Factor": r_pen_const, "Reg Penalty": reg_penalty,
            "Total Daily Penalty": eff_penalty + punct_penalty + reg_penalty
        })

    return {
        "summary": kpi_results, "eff_details": eff_details, "punct_details": punct_details, "reg_details": reg_details,
        "assigned": assigned_log, "trips": trip_log, "reg_trips": trip_reg_log, "obu_errors": obu_errors_log,
        "cleaning": cleaning_log, "missed_trips": missed_log, "accepted_trips": accepted_log, "bunched_trips": bunched_log,
        "strict_late_log": strict_late_log, "invalid_stops_log": invalid_stops_log, "path_statistics_log": path_statistics_log,
        "auto_rush": auto_rush_periods_log, "sanity_checks": [], "raw_sanity_checks": [],
        "raw_strict_late_log": raw_strict_late_log, "raw_cleaning": raw_cleaning, 
        "raw_obu": raw_obu, "raw_bunched": raw_bunched, "raw_accepted": raw_accepted, "raw_on_time": raw_on_time, 
        "raw_off_time": raw_off_time, "raw_reg_regular": raw_reg_regular, "raw_reg_irregular": raw_reg_irregular,
        "ignored_trips": ignored_trips_log
    }

# ================= UI AND EXECUTION =================

with st.sidebar:
    st.header("🎛️ Operation Mode")
    app_mode = st.radio("Select Action:", ["🛠️ Process New Data", "📊 View Existing Dashboard", "ℹ️ About / Help"])
    st.markdown("---")

if app_mode == "ℹ️ About / Help":
    st.title("ℹ️ About & Documentation")
    st.markdown("This dashboard processes daily Kentkart schedule data (STIO) against your expected Master Matrix to compute Efficiency, Punctuality, and Regularity KPIs.")

    with st.expander("📝 How to Setup the Master File", expanded=True):
        st.markdown("""
        The Master File (Excel) configures the baseline for your operation. It supports the following sheets:
        * **Scheduled Times**: The core matrix. Columns represent routes (e.g., `Khanna Pul → Pims`).
            * Row 1: Route Name
            * Row 2: Valid Start Date (e.g., `2026-05-01`). Leave blank for permanent schedules.
            * Row 3: Valid End Date (e.g., `2026-05-31`).
            * Row 4+: Scheduled times (e.g., `06:00`, `06:15`).
        * **Total Time**: Target travel time for Regularity. Row 1: Normal target. Row 2 (Optional): Approved Fast Target.
        * **Distances**: Expected km for the route. Used as an efficiency penalty multiplier.
        * **Intimations**: Exemptions. Columns: `Date`, `Path`, `Intimation`, `Correction`. Negative intimations equal approved extra trips.
        * **Rush Hours**: Exempts slow trips. Columns: `Date`, `Path`, `Start Time`, `End Time`.
        * **VTS**: Hardware tracking overrides. Columns: `Date`, `Bus`, and specific route columns.
        """)

    with st.expander("📊 KPI Calculation Methodology"):
        st.markdown("""
        **1. Efficiency (Volume & Assignment)**
        * **Target:** `Total Assigned Schedules - Intimations`
        * **Operated:** Valid trips recorded. If `Strict Mode` is on, trips delayed >50% of headway are deducted.
        * **Bunching Deduction:** Trips occurring simultaneously on the same schedule slot are penalized as bunched, unless covered by negative intimations.
        * **Penalty:** `(Deficit Volume + Bunched Trips) * Penalty Factor`. The factor varies depending on whether you hit the Target/Bottom thresholds.

        **2. Punctuality (Timing)**
        * Evaluates the delay against the *Corrected Schedule* (the closest actual master schedule to the departure time, fixing driver OBU errors).
        * **Tolerances:** Can be fixed minutes (e.g., -2 to +5) or a dynamic percentage of the scheduled headway (e.g., 20%).
        * **Bunching Exemption:** If multiple trips are bunched on the same schedule, the most punctual one is evaluated normally. The extra bunched trips are excluded from punctuality penalties (since they are already penalized in Efficiency).
        * **Penalty:** Fixed penalty amount applied per Early or Late trip.

        **3. Regularity (Travel Time)**
        * Travel time is evaluated against the `Total Time` master sheet.
        * **Classifications:** Regular, Too Fast, Too Slow.
        * **Exemptions:** * Slow trips occurring during explicitly defined `Rush Hours` are exempted. 
            * **Auto-Rush Detection:** If enabled, any chain of consecutive slow trips (exceeding your threshold, default 3) is automatically exempted.
            * 'Approved Fast' quotas can exempt slightly fast trips based on intimations.
        * **Penalty:** Fixed penalty amount per Irregular trip.
        """)

    with st.expander("🛡️ Edge Cases & Data Cleaning"):
        st.markdown("""
        * **Ignored Trips Sheet:** Any trip that is exempted from KPI penalties (due to Imputation, Outlier status, Auto-Rush, or Bunching Exemption) is filtered out of your main impactful KPI detail sheets and aggregated in the `IGNORED_TRIPS` sheet for your review.
        * **Distance Outliers:** Trips with `REALIZED_KM` < minimum are dropped. Outliers (> Mean + 3*StdDev) are excluded from regularity calculations. Distances > 2.5x the mode are counted as double trips (Weight = 2).
        * **Ghost Trips:** Trips with fewer stops than the minimum threshold are ignored.
        * **Imputed Start Times:** If a trip lacks an explicit start time but has a planned schedule, the system retroactively imputes the start time based on the timestamp of the first recorded stop minus the expected duration to reach that stop.
        * **OBU Auto-Correction:** If a driver selects an incorrect schedule on the OBU (e.g., 06:00 instead of 06:15), the system automatically matches their actual departure time to the closest real schedule from the Master File to prevent false penalties.
        * **Date Rollovers:** Night shift times (e.g., past midnight) are automatically shifted by +1 day based on a 12-hour proximity check against the target schedule.
        """)

elif app_mode == "📊 View Existing Dashboard":
    st.title("📊 View Existing KPI Dashboard")
    st.markdown("Upload a previously generated `KPI_Dashboard.xlsx` to explore the interactive graphics and data tables instantly.")
    
    existing_file = st.file_uploader("Upload KPI Dashboard Excel File", type=["xlsx"])
    if existing_file:
        if load_existing_dashboard(existing_file):
            st.success("Dashboard successfully loaded!")
        else:
            st.warning("Please ensure you uploaded a valid KPI Dashboard generated by this tool.")

elif app_mode == "🛠️ Process New Data":
    with st.sidebar:
        st.header("⚙️ Dashboard Parameters")
        with st.expander("Trip Efficiency Parameters"):
            p_strict_eff = st.checkbox("Enable Strict Efficiency (Deduct >50% Late)", value=False)
            p_eff_upper = st.number_input("Upper Threshold (No Penalty >= %)", value=98.0)
            p_eff_target = st.number_input("Target Threshold (%)", value=91.0)
            p_eff_bottom = st.number_input("Bottom Limit (< %)", value=80.0)
            p_eff_pen_high = st.number_input("Missed Penalty (Target to Upper %)", value=40.0)
            p_eff_pen_low = st.number_input("Missed Penalty (< Target %)", value=45.0)

        with st.expander("Punctuality Parameters"):
            p_punct_upper = st.number_input("Upper Threshold (No Penalty >= %)", value=95.0, key="p1")
            p_punct_bottom = st.number_input("Bottom Limit (< %)", value=85.0, key="p2")
            p_punct_pen = st.number_input("Penalty Per Off-Time Trip", value=50.0, key="p3")
            p_use_headway = st.checkbox("Use Dynamic Headway Percentage for Tolerances", value=True)
            p_early_headway = st.number_input("Early Tolerance Headway (%)", value=20.0)
            p_late_headway = st.number_input("Late Tolerance Headway (%)", value=20.0)
            st.caption("Fixed Minute Limits (Fallback):")
            p_early_limit = st.number_input("Fixed Early Tolerance (mins)", value=2.0)
            p_late_limit = st.number_input("Fixed Late Tolerance (mins)", value=5.0)

        with st.expander("Regularity Parameters"):
            p_reg_upper = st.number_input("Upper Threshold (No Penalty >= %)", value=95.0, key="r1")
            p_reg_bottom = st.number_input("Bottom Limit (< %)", value=85.0, key="r2")
            p_reg_fast_pct = st.number_input("Too Fast Limit (+ %)", value=20.0)
            p_reg_slow_pct = st.number_input("Too Slow Limit (- %)", value=20.0)
            p_reg_imp_fast = st.number_input("Impossible Fast Limit (< %)", value=60.0)
            p_reg_imp_slow = st.number_input("Impossible Slow Limit (> %)", value=60.0)
            p_reg_pen = st.number_input("Penalty Per Irregular Trip", value=50.0, key="r3")
            st.caption("Auto-Rush Detection:")
            p_auto_rush = st.checkbox("Auto-Ignore Extended Rush Hours", value=True)
            p_auto_rush_thresh = st.number_input("Consecutive Slow Trips Threshold", value=3, min_value=1)

        with st.expander("Data Cleaning"):
            p_min_km = st.number_input("Minimum Valid REALIZED_KM", value=1.5)
            p_min_stops = st.number_input("Minimum Stop Count", value=5)

    params = {
        "strict_eff": p_strict_eff, "eff_upper": p_eff_upper, "eff_target": p_eff_target, "eff_bottom": p_eff_bottom,
        "eff_pen_high": p_eff_pen_high, "eff_pen_low": p_eff_pen_low, "punct_upper": p_punct_upper, "punct_bottom": p_punct_bottom,
        "use_headway_pct": p_use_headway, "early_headway_pct": p_early_headway, "late_headway_pct": p_late_headway,
        "early_limit": p_early_limit, "late_limit": p_late_limit, "punct_pen": p_punct_pen, "reg_upper": p_reg_upper,
        "reg_bottom": p_reg_bottom, "reg_fast_pct": p_reg_fast_pct, "reg_slow_pct": p_reg_slow_pct, 
        "reg_imp_fast_pct": p_reg_imp_fast, "reg_imp_slow_pct": p_reg_imp_slow, "reg_pen": p_reg_pen,
        "auto_rush_enable": p_auto_rush, "auto_rush_thresh": p_auto_rush_thresh,
        "min_km": p_min_km, "min_stops": p_min_stops, "run_sanity_checks": False
    }

    st.title("🛠️ Process Daily Kentkart Reports")
    col1, col2 = st.columns(2)
    with col1: master_file = st.file_uploader("1. Upload Master Matrix (Optional but Recommended)", type=["xlsx", "xls", "csv"])
    with col2: daily_files = st.file_uploader("2. Upload Daily Kentkart Reports (STIO)", type=["xlsx", "xls", "csv"], accept_multiple_files=True)
    
    st.markdown("---")
    opt_raw = st.checkbox("Generate Secondary Raw Data File (Full STIO Rows)", value=True)
    opt_audit = st.checkbox("Data Audit Mode (Sanity Checks ONLY - Skips all KPI logic)", value=False)
    params["run_sanity_checks"] = opt_audit

    if st.button("🚀 Run Analysis", type="primary", use_container_width=True):
        if not daily_files: st.error("Please upload at least one Daily Report.")
        else:
            with st.status("Initializing Analysis...", expanded=True) as status:
                if opt_audit:
                    master_dict, master_times, intimations_dict, corrections_dict = None, None, None, None
                    rush_hours_dict, master_times_fast, vts_dict, path_distances = None, None, None, None
                else:
                    master_dict, master_times, intimations_dict, corrections_dict, rush_hours_dict, master_times_fast, vts_dict, path_distances = load_master_data(master_file, status)
                
                master_data = {
                    "KPI_SUMMARY": [], "EFFICIENCY_DETAILS": [], "PUNCTUALITY_DETAILS": [], "REGULARITY_DETAILS": [],
                    "OBU_SELECTION_ERRORS": [], "ASSIGNED_TRIPS": [], "ACCEPTED_TRIPS": [], "BUNCHED_TRIPS": [], "MISSED_TRIPS": [], 
                    "STRICT_EFFICIENCY_DEDUCTIONS": [], "AUTO_DETECTED_RUSH": [], "STIO_SANITY_CHECKS": [], "ON_TIME_TRIPS": [], "OFF_TIME_TRIPS": [], 
                    "REGULAR_TIME_TRIPS": [], "IRREGULAR_TIME_TRIPS": [], "IGNORED_TRIPS": [], "DATA_CLEANING_LOG": [], "INVALID_STOP_TIMES": [], 
                    "PATH_STATISTICS": []
                }
                raw_master_data = {
                    "RAW_ON_TIME": [], "RAW_OFF_TIME": [], "RAW_REGULAR_TIME": [], "RAW_IRREGULAR_TIME": [], 
                    "RAW_ACCEPTED": [], "RAW_BUNCHED": [], "RAW_OBU_ERRORS": [], "RAW_STRICT_LATE": [], "RAW_CLEANING_LOG": [], "RAW_SANITY_CHECKS": []
                }
                
                for f in daily_files:
                    try:
                        status.write(f"Processing: **{f.name}**")
                        res = process_file(f, params, master_dict, master_times, intimations_dict, corrections_dict, rush_hours_dict, master_times_fast, path_distances)
                        
                        master_data["KPI_SUMMARY"].extend(res.get("summary", []))
                        master_data["EFFICIENCY_DETAILS"].extend(res.get("eff_details", []))
                        master_data["PUNCTUALITY_DETAILS"].extend(res.get("punct_details", []))
                        master_data["REGULARITY_DETAILS"].extend(res.get("reg_details", []))
                        master_data["OBU_SELECTION_ERRORS"].extend(res.get("obu_errors", []))
                        master_data["ASSIGNED_TRIPS"].extend(res.get("assigned", []))
                        master_data["ACCEPTED_TRIPS"].extend(res.get("accepted_trips", []))
                        master_data["BUNCHED_TRIPS"].extend(res.get("bunched_trips", []))
                        master_data["MISSED_TRIPS"].extend(res.get("missed_trips", []))
                        master_data["STRICT_EFFICIENCY_DEDUCTIONS"].extend(res.get("strict_late_log", []))
                        master_data["AUTO_DETECTED_RUSH"].extend(res.get("auto_rush", []))
                        master_data["STIO_SANITY_CHECKS"].extend(res.get("sanity_checks", []))
                        master_data["DATA_CLEANING_LOG"].extend(res.get("cleaning", []))
                        master_data["INVALID_STOP_TIMES"].extend(res.get("invalid_stops_log", []))
                        master_data["PATH_STATISTICS"].extend(res.get("path_statistics_log", []))
                        master_data["IGNORED_TRIPS"].extend(res.get("ignored_trips", []))
                        
                        # Only append impactful trips to the detail sheets
                        for trip in res.get("trips", []):
                            if trip["Status"] == "On-Time": master_data["ON_TIME_TRIPS"].append(trip)
                            elif trip["Status"] in ["Early", "Late"]: master_data["OFF_TIME_TRIPS"].append(trip)
                            
                        for trip in res.get("reg_trips", []):
                            if trip["Status"] in ["Regular", "Regular (Approved Fast)"]: master_data["REGULAR_TIME_TRIPS"].append(trip)
                            elif trip["Status"] in ["Too Fast", "Too Slow"]: master_data["IRREGULAR_TIME_TRIPS"].append(trip)
                            
                        raw_master_data["RAW_CLEANING_LOG"].extend(res.get("raw_cleaning", []))
                        raw_master_data["RAW_SANITY_CHECKS"].extend(res.get("raw_sanity_checks", []))
                        raw_master_data["RAW_OBU_ERRORS"].extend(res.get("raw_obu", []))
                        raw_master_data["RAW_STRICT_LATE"].extend(res.get("raw_strict_late_log", []))
                        raw_master_data["RAW_BUNCHED"].extend(res.get("raw_bunched", []))
                        raw_master_data["RAW_ACCEPTED"].extend(res.get("raw_accepted", []))
                        raw_master_data["RAW_ON_TIME"].extend(res.get("raw_on_time", []))
                        raw_master_data["RAW_OFF_TIME"].extend(res.get("raw_off_time", []))
                        raw_master_data["RAW_REGULAR_TIME"].extend(res.get("raw_reg_regular", []))
                        raw_master_data["RAW_IRREGULAR_TIME"].extend(res.get("raw_reg_irregular", []))

                    except Exception as e: st.error(f"Error on {f.name}: {str(e)}")

                if not master_data["KPI_SUMMARY"] and not opt_audit:
                    status.update(label="No valid data processed.", state="error")
                    st.stop()

                # --- Write to memory ---
                status.write("Compiling files into memory...")
                main_output = io.BytesIO()
                raw_output = io.BytesIO()
                
                excel_kpi_summary = [dict(row) for row in master_data["KPI_SUMMARY"]]
                for row in excel_kpi_summary:
                    row[" "] = "" 
                    row["-- QUICK NAVIGATION --"] = ""
                    row["Go To Accepted Trips"] = '=HYPERLINK("#\'ACCEPTED_TRIPS\'!A1", "↳ Open Accepted Trips")'
                    row["Go To Efficiency Details"] = '=HYPERLINK("#\'EFFICIENCY_DETAILS\'!A1", "↳ Open Efficiency Sheet")'
                    row["Go To Punctuality Details"] = '=HYPERLINK("#\'PUNCTUALITY_DETAILS\'!A1", "↳ Open Punctuality Sheet")'
                
                def format_excel_output(writer_obj):
                    from openpyxl.styles import Font
                    bold_font = Font(bold=True)
                    for sheetname in writer_obj.sheets:
                        ws = writer_obj.sheets[sheetname]
                        if sheetname == "KPI_SUMMARY":
                            for row in ws.iter_rows():
                                if row[0].value and ("(%)" in str(row[0].value) or "Total Daily Penalty" in str(row[0].value)):
                                    for cell in row: cell.font = bold_font
                        for col in ws.columns:
                            col_letter = col[0].column_letter
                            ws.column_dimensions[col_letter].width = 20

                with pd.ExcelWriter(main_output, engine="openpyxl") as writer:
                    if excel_kpi_summary:
                        pd.DataFrame(excel_kpi_summary).T.reset_index().to_excel(writer, sheet_name="KPI_SUMMARY", index=False, header=False)
                    for sheet_name, data in master_data.items():
                        if sheet_name == "KPI_SUMMARY": continue 
                        if data: pd.DataFrame(data).to_excel(writer, sheet_name=sheet_name, index=False)
                    format_excel_output(writer)
                    
                if opt_raw:
                    with pd.ExcelWriter(raw_output, engine="openpyxl") as writer:
                        for sheet_name, data in raw_master_data.items():
                            if data: pd.DataFrame(data).to_excel(writer, sheet_name=sheet_name, index=False)
                        format_excel_output(writer)
                
                # === SAVE TO SESSION STATE ===
                st.session_state.analysis_complete = True
                st.session_state.main_excel = main_output.getvalue()
                st.session_state.raw_excel = raw_output.getvalue() if opt_raw else None
                st.session_state.summary_data = master_data["KPI_SUMMARY"]
                st.session_state.all_accepted = master_data["ACCEPTED_TRIPS"]
                st.session_state.bunched_trips = master_data["BUNCHED_TRIPS"]
                st.session_state.auto_rush = master_data["AUTO_DETECTED_RUSH"]
                st.session_state.run_time = datetime.now().strftime("%Y%m%d_%H%M")
                st.session_state.loaded_from_file = False
                
                status.update(label="Analysis Complete! 🎉", state="complete", expanded=False)
            st.rerun()

# ================= RENDER UI IF COMPLETE =================
if st.session_state.analysis_complete and app_mode != "ℹ️ About / Help":
    
    if not st.session_state.loaded_from_file:
        st.success(f"✅ Data processed successfully at {st.session_state.run_time}")
        colA, colB = st.columns(2)
        with colA:
            st.download_button("📥 Download Generated Dashboard (Excel)", 
                               data=st.session_state.main_excel, 
                               file_name=f"KPI_Dashboard_{st.session_state.run_time}.xlsx", 
                               mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", 
                               use_container_width=True)
        if st.session_state.raw_excel:
            with colB:
                st.download_button("📥 Download Raw Data Extract (Excel)", 
                                   data=st.session_state.raw_excel, 
                                   file_name=f"Raw_STIO_Data_{st.session_state.run_time}.xlsx", 
                                   mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", 
                                   use_container_width=True)
        st.markdown("---")
    
    if st.session_state.summary_data:
        df_kpi = pd.DataFrame(st.session_state.summary_data)
        
        st.subheader("📈 Interactive Results Dashboard")
        
        avg_eff = df_kpi["Strict Bunching Eff (%)"].mean() if "Strict Bunching Eff (%)" in df_kpi.columns else 0
        avg_punct = df_kpi["Punctuality (%)"].mean() if "Punctuality (%)" in df_kpi.columns else 0
        avg_reg = df_kpi["Regularity (%)"].mean() if "Regularity (%)" in df_kpi.columns else 0
        total_penalty = df_kpi["Total Daily Penalty"].sum() if "Total Daily Penalty" in df_kpi.columns else 0
        
        m1, m2, m3, m4 = st.columns(4)
        m1.metric("Overall Efficiency (Strict)", f"{avg_eff:.2f}%")
        m2.metric("Overall Punctuality", f"{avg_punct:.2f}%")
        m3.metric("Overall Regularity", f"{avg_reg:.2f}%")
        m4.metric("Total Penalty Accrued", f"{total_penalty:,.0f}", delta_color="inverse")
        
        st.write("")
        tab1, tab2 = st.tabs(["📊 KPI Trend Charts", "🗄️ Deep-Dive Data Grid"])
        
        with tab1:
            if not df_kpi.empty:
                # Format Dates for cleaner X-Axis
                df_kpi['Date'] = pd.to_datetime(df_kpi['Date'], errors='coerce').dt.strftime('%Y-%m-%d')
                
                unique_paths = sorted(df_kpi['Path'].astype(str).unique())
                selected_path = st.selectbox("📍 Select Route to Analyze", unique_paths)
                
                df_filtered = df_kpi[df_kpi['Path'] == selected_path].sort_values("Date")
                
                if not df_filtered.empty:
                    c1, c2 = st.columns(2)
                    with c1:
                        if "Strict Bunching Eff (%)" in df_filtered.columns:
                            fig_eff = px.bar(df_filtered, x='Date', y='Strict Bunching Eff (%)', 
                                             text='Strict Bunching Eff (%)', title=f"Efficiency Trend: {selected_path}",
                                             color='Strict Bunching Eff (%)', color_continuous_scale="RdYlGn", range_color=[80, 100])
                            fig_eff.update_traces(texttemplate='%{text:.1f}%', textposition='outside')
                            fig_eff.update_layout(xaxis_type='category') # Forces discrete dates
                            st.plotly_chart(fig_eff, use_container_width=True)
                    with c2:
                        if "Punctuality (%)" in df_filtered.columns:
                            fig_punct = px.bar(df_filtered, x='Date', y='Punctuality (%)', 
                                               text='Punctuality (%)', title=f"Punctuality Trend: {selected_path}",
                                               color='Punctuality (%)', color_continuous_scale="RdYlGn", range_color=[80, 100])
                            fig_punct.update_traces(texttemplate='%{text:.1f}%', textposition='outside')
                            fig_punct.update_layout(xaxis_type='category')
                            st.plotly_chart(fig_punct, use_container_width=True)
                else:
                    st.info("No data available for the selected route.")
        
        with tab2:
            subtab1, subtab2, subtab3, subtab4 = st.tabs(["KPI Summary", "Accepted Trips Explorer", "Bunched Trips", "Auto Detected Rush"])
            
            with subtab1:
                st.dataframe(df_kpi, use_container_width=True)
                
            with subtab2:
                if st.session_state.all_accepted:
                    df_trips = pd.DataFrame(st.session_state.all_accepted)
                    if 'Tracking Link' in df_trips.columns:
                        df_trips['Clean URL'] = df_trips['Tracking Link'].apply(extract_raw_url)
                        st.dataframe(
                            df_trips.drop(columns=['Tracking Link'], errors='ignore'),
                            column_config={"Clean URL": st.column_config.LinkColumn("Kentkart VTS Link", display_text="Track Trip")},
                            use_container_width=True
                        )
                    else: st.dataframe(df_trips, use_container_width=True)
                else: st.info("No accepted trips data found.")
                
            with subtab3:
                if st.session_state.bunched_trips:
                    df_bunch = pd.DataFrame(st.session_state.bunched_trips)
                    if 'Tracking Link' in df_bunch.columns:
                        df_bunch['Clean URL'] = df_bunch['Tracking Link'].apply(extract_raw_url)
                        st.dataframe(
                            df_bunch.drop(columns=['Tracking Link'], errors='ignore'),
                            column_config={"Clean URL": st.column_config.LinkColumn("Kentkart VTS Link", display_text="Track Trip")},
                            use_container_width=True
                        )
                    else: st.dataframe(df_bunch, use_container_width=True)
                else: st.info("No bunched trips detected.")
                
            with subtab4:
                if st.session_state.auto_rush:
                    st.dataframe(pd.DataFrame(st.session_state.auto_rush), use_container_width=True)
                else: st.info("No auto-detected rush hour periods found.")
