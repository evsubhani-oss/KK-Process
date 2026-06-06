import streamlit as st
import pandas as pd
import plotly.express as px
import plotly.graph_objects as go
import io
from datetime import datetime

# ================= PAGE CONFIGURATION =================
st.set_page_config(page_title="Monthly KPI Aggregator", page_icon="📅", layout="wide")

st.title("📅 Monthly KPI & Penalty Aggregator")
st.markdown("""
Upload a month's worth of generated `KPI_Dashboard.xlsx` files. 
This tool aggregates all routes and days to calculate the **True Monthly Average**. 
Penalties are only applied if the company-wide average falls below the specified thresholds.
""")

# ================= SIDEBAR PARAMETERS =================
with st.sidebar:
    st.header("⚙️ Monthly Thresholds & Penalties")
    
    st.subheader("1. Efficiency (Volume)")
    thresh_eff = st.number_input("Efficiency Threshold (%)", value=98.0, step=0.1)
    rate_eff = st.number_input("Penalty per Missed Trip (Volume Deficit)", value=50.0, step=5.0)
    
    st.subheader("2. Bunching (Always Applied)")
    p_exempt_bunching = st.checkbox("Exempt Bunching Penalties caused by Imputed Trips", value=False)
    st.caption("Checked: Reverses bunching penalties mathematically caused by imputed/ghost trips, lowering your financial penalty pool.")
    
    st.subheader("3. Punctuality")
    thresh_punct = st.number_input("Punctuality Threshold (%)", value=95.0, step=0.1)
    rate_punct = st.number_input("Penalty per Off-Time Trip", value=50.0, step=5.0)
    
    st.subheader("4. Regularity")
    thresh_reg = st.number_input("Regularity Threshold (%)", value=95.0, step=0.1)
    rate_reg = st.number_input("Penalty per Irregular Trip", value=50.0, step=5.0)
    
    st.markdown("---")
    st.subheader("🧹 Data Cleaning")
    p_count_ignored = st.checkbox("Count Ignored Trips as Punctual/Regular", value=False)
    p_count_corrections = st.checkbox("Count Master File Corrections as Punctual/Regular", value=False)
    st.caption("Checked: Adds these trips back into the successful percentage calculation (forgives them).\n\nUnchecked: Completely removes them from the evaluated denominator without actively penalizing them.")

# ================= FILE UPLOAD =================
uploaded_files = st.file_uploader("Upload Daily KPI Dashboards (Excel)", type=["xlsx"], accept_multiple_files=True)

if uploaded_files:
    kpi_summaries = []
    reg_details = []
    
    with st.status("Aggregating Daily Files...", expanded=True) as status:
        for f in uploaded_files:
            try:
                xls = pd.ExcelFile(f)
                
                if "KPI_SUMMARY" in xls.sheet_names:
                    df_summ_raw = pd.read_excel(xls, sheet_name="KPI_SUMMARY", header=None)
                    df_summ = df_summ_raw.set_index(0).T
                    df_summ = df_summ.drop(columns=[c for c in df_summ.columns if str(c).startswith("Go To ") or c == " " or c == "-- QUICK NAVIGATION --"], errors='ignore')
                    for col in df_summ.columns:
                        if col != "Date" and col != "Path":
                            df_summ[col] = pd.to_numeric(df_summ[col], errors='coerce')
                    kpi_summaries.append(df_summ)
                
                if "REGULARITY_DETAILS" in xls.sheet_names:
                    df_reg = pd.read_excel(xls, sheet_name="REGULARITY_DETAILS")
                    reg_details.append(df_reg)
                    
            except Exception as e:
                st.error(f"Error reading {f.name}: {e}")
        
        status.update(label="Aggregation Complete!", state="complete")

    if kpi_summaries:
        # ================= DATA AGGREGATION =================
        df_summ_all = pd.concat(kpi_summaries, ignore_index=True)
        
        # 1. Efficiency & Bunching Parameters
        tot_target_assigned = df_summ_all['Net Target Assigned'].sum()
        tot_vol_deficit = df_summ_all['Volume Deficit'].sum()
        
        # Bunching Penalty Dynamic Re-calculation
        if 'Base Strict Bunching Penalty' in df_summ_all.columns and 'Potential Imputed Bunching Savings' in df_summ_all.columns:
            tot_base_bunching_penalty = df_summ_all['Base Strict Bunching Penalty'].sum()
            tot_potential_savings = df_summ_all['Potential Imputed Bunching Savings'].sum()
            
            if p_exempt_bunching:
                tot_bunched_penalty = max(0, tot_base_bunching_penalty - tot_potential_savings)
            else:
                tot_bunched_penalty = tot_base_bunching_penalty
        else:
            # Fallback for old files generated before this update
            tot_bunched_penalty = df_summ_all['Strict Bunching Penalty'].sum() if 'Strict Bunching Penalty' in df_summ_all.columns else 0
        
        # 2. Punctuality Parameters
        tot_on_time_raw = df_summ_all['On-Time'].sum()
        tot_off_time_raw = df_summ_all['Off-Time'].sum()
        
        tot_ignored_punct = df_summ_all['Ignored Counted Punctual'].sum() if 'Ignored Counted Punctual' in df_summ_all.columns else 0
        tot_corr_punct = df_summ_all['Corrections Counted Punctual'].sum() if 'Corrections Counted Punctual' in df_summ_all.columns else 0
        
        tot_on_time = tot_on_time_raw
        tot_off_time = tot_off_time_raw
        
        if not p_count_ignored:
            tot_on_time = max(0, tot_on_time - tot_ignored_punct)
        if not p_count_corrections:
            tot_on_time = max(0, tot_on_time - tot_corr_punct)
            
        tot_punct_evaluated = tot_on_time + tot_off_time

        # 3. Regularity Parameters
        if reg_details:
            df_reg_all = pd.concat(reg_details, ignore_index=True)
            tot_reg_regular_raw = df_reg_all['Regular Trips'].sum()
            tot_reg_irregular_raw = df_reg_all['Total Irregular Trips'].sum()
            
            tot_ignored_reg = df_reg_all['Ignored Counted Regular'].sum() if 'Ignored Counted Regular' in df_reg_all.columns else 0
            tot_corr_reg = df_reg_all['Corrections Counted Regular'].sum() if 'Corrections Counted Regular' in df_reg_all.columns else 0
            
            tot_reg_regular = tot_reg_regular_raw
            tot_reg_irregular = tot_reg_irregular_raw
            
            if not p_count_ignored:
                tot_reg_regular = max(0, tot_reg_regular - tot_ignored_reg)
            if not p_count_corrections:
                tot_reg_regular = max(0, tot_reg_regular - tot_corr_reg)
                
            tot_reg_evaluated = tot_reg_regular + tot_reg_irregular
        else:
            tot_reg_regular = 0
            tot_reg_irregular = 0
            tot_reg_evaluated = 0

        # ================= CALCULATE TRUE MONTHLY MEANS =================
        grand_eff = ((tot_target_assigned - tot_vol_deficit) / tot_target_assigned * 100) if tot_target_assigned > 0 else 100.0
        grand_punct = (tot_on_time / tot_punct_evaluated * 100) if tot_punct_evaluated > 0 else 100.0
        grand_reg = (tot_reg_regular / tot_reg_evaluated * 100) if tot_reg_evaluated > 0 else 100.0

        # ================= APPLY GATED PENALTIES =================
        pen_bunching = tot_bunched_penalty
        pen_eff = (tot_vol_deficit * rate_eff) if grand_eff < thresh_eff else 0
        pen_punct = (tot_off_time * rate_punct) if grand_punct < thresh_punct else 0
        pen_reg = (tot_reg_irregular * rate_reg) if grand_reg < thresh_reg else 0
        
        tot_penalties = pen_bunching + pen_eff + pen_punct + pen_reg

        # ================= VISUALIZATIONS =================
        st.markdown("---")
        st.subheader("🎯 True Monthly KPI Performance")
        
        def create_gauge(title, value, threshold):
            fig = go.Figure(go.Indicator(
                mode = "gauge+number",
                value = value,
                number = {'suffix': "%", 'valueformat': ".2f"},
                domain = {'x': [0, 1], 'y': [0, 1]},
                title = {'text': title, 'font': {'size': 18}},
                gauge = {
                    'axis': {'range': [None, 100], 'tickwidth': 1, 'tickcolor': "darkblue"},
                    'bar': {'color': "rgba(0,0,0,0)"},
                    'bgcolor': "white",
                    'borderwidth': 2,
                    'bordercolor': "gray",
                    'steps': [
                        {'range': [0, threshold], 'color': "#ffcccb"},   
                        {'range': [threshold, 100], 'color': "#d4edda"}  
                    ],
                    'threshold': {
                        'line': {'color': "black", 'width': 5},
                        'thickness': 0.75,
                        'value': value
                    }
                }
            ))
            fig.update_layout(height=250, margin=dict(l=10, r=10, t=40, b=10))
            return fig

        c1, c2, c3 = st.columns(3)
        with c1: st.plotly_chart(create_gauge(f"Efficiency (Target: {thresh_eff}%)", grand_eff, thresh_eff), use_container_width=True)
        with c2: st.plotly_chart(create_gauge(f"Punctuality (Target: {thresh_punct}%)", grand_punct, thresh_punct), use_container_width=True)
        with c3: st.plotly_chart(create_gauge(f"Regularity (Target: {thresh_reg}%)", grand_reg, thresh_reg), use_container_width=True)

        st.markdown("---")
        st.subheader("💸 Financial Penalty Waterfall")
        st.markdown("Penalties only appear below if the corresponding threshold gauge above is in the **red zone**.")
        
        fig_waterfall = go.Figure(go.Waterfall(
            name = "Penalties", orientation = "v",
            measure = ["relative", "relative", "relative", "relative", "total"],
            x = ["Bunching (Always)", "Efficiency Deficit", "Off-Time Trips", "Irregular Trips", "Total Penalties"],
            textposition = "outside",
            text = [f"-{pen_bunching:,.0f}", f"-{pen_eff:,.0f}", f"-{pen_punct:,.0f}", f"-{pen_reg:,.0f}", f"-{tot_penalties:,.0f}"],
            y = [-pen_bunching, -pen_eff, -pen_punct, -pen_reg, -tot_penalties],
            connector = {"line":{"color":"rgb(63, 63, 63)"}},
            decreasing = {"marker":{"color":"#ff4b4b"}},
            totals = {"marker":{"color":"#1f77b4"}}
        ))
        fig_waterfall.update_layout(height=400, margin=dict(l=10, r=10, t=30, b=10), showlegend=False)
        st.plotly_chart(fig_waterfall, use_container_width=True)

        # ================= ROUTE BY ROUTE ANALYSIS =================
        st.markdown("---")
        st.subheader("🔍 Route-by-Route Monthly Breakdown")
        
        cols_summ = ['Net Target Assigned', 'Volume Deficit', 'On-Time', 'Off-Time']
        if 'Strict Bunching Penalty' in df_summ_all.columns: cols_summ.append('Strict Bunching Penalty')
        if 'Base Strict Bunching Penalty' in df_summ_all.columns: cols_summ.append('Base Strict Bunching Penalty')
        if 'Potential Imputed Bunching Savings' in df_summ_all.columns: cols_summ.append('Potential Imputed Bunching Savings')
        if 'Ignored Counted Punctual' in df_summ_all.columns: cols_summ.append('Ignored Counted Punctual')
        if 'Corrections Counted Punctual' in df_summ_all.columns: cols_summ.append('Corrections Counted Punctual')
            
        route_summ = df_summ_all.groupby('Path')[cols_summ].sum().reset_index()
        if 'Ignored Counted Punctual' not in route_summ.columns: route_summ['Ignored Counted Punctual'] = 0
        if 'Corrections Counted Punctual' not in route_summ.columns: route_summ['Corrections Counted Punctual'] = 0
        
        if reg_details:
            cols_reg = ['Regular Trips', 'Total Irregular Trips']
            if 'Ignored Counted Regular' in df_reg_all.columns: cols_reg.append('Ignored Counted Regular')
            if 'Corrections Counted Regular' in df_reg_all.columns: cols_reg.append('Corrections Counted Regular')
                
            route_reg = df_reg_all.groupby('Path')[cols_reg].sum().reset_index()
            if 'Ignored Counted Regular' not in route_reg.columns: route_reg['Ignored Counted Regular'] = 0
            if 'Corrections Counted Regular' not in route_reg.columns: route_reg['Corrections Counted Regular'] = 0
            
            route_all = pd.merge(route_summ, route_reg, on='Path', how='left').fillna(0)
        else:
            route_all = route_summ.copy()
            route_all['Regular Trips'] = 0
            route_all['Total Irregular Trips'] = 0
            route_all['Ignored Counted Regular'] = 0
            route_all['Corrections Counted Regular'] = 0

        if not p_count_ignored:
            route_all['On-Time'] = (route_all['On-Time'] - route_all['Ignored Counted Punctual']).clip(lower=0)
            route_all['Regular Trips'] = (route_all['Regular Trips'] - route_all['Ignored Counted Regular']).clip(lower=0)
            
        if not p_count_corrections:
            route_all['On-Time'] = (route_all['On-Time'] - route_all['Corrections Counted Punctual']).clip(lower=0)
            route_all['Regular Trips'] = (route_all['Regular Trips'] - route_all['Corrections Counted Regular']).clip(lower=0)
            
        # Dynamically switch the Bunching penalty column based on the toggle for the route-by-route table
        if 'Base Strict Bunching Penalty' in route_all.columns and 'Potential Imputed Bunching Savings' in route_all.columns:
            if p_exempt_bunching:
                route_all['Strict Bunching Penalty'] = (route_all['Base Strict Bunching Penalty'] - route_all['Potential Imputed Bunching Savings']).clip(lower=0)
            else:
                route_all['Strict Bunching Penalty'] = route_all['Base Strict Bunching Penalty']

        route_all['Efficiency (%)'] = ((route_all['Net Target Assigned'] - route_all['Volume Deficit']) / route_all['Net Target Assigned'] * 100).fillna(0)
        route_all['Punctuality (%)'] = (route_all['On-Time'] / (route_all['On-Time'] + route_all['Off-Time']) * 100).fillna(0)
        route_all['Regularity (%)'] = (route_all['Regular Trips'] / (route_all['Regular Trips'] + route_all['Total Irregular Trips']) * 100).fillna(0)

        display_cols = ['Path', 'Efficiency (%)', 'Punctuality (%)', 'Regularity (%)', 'Volume Deficit', 'Strict Bunching Penalty', 'Off-Time', 'Total Irregular Trips']
        route_display = route_all[display_cols].copy()
        for col in ['Efficiency (%)', 'Punctuality (%)', 'Regularity (%)']:
            route_display[col] = route_display[col].round(2)
            
        st.dataframe(route_display, use_container_width=True)

        # ================= CLEAN DAILY PERFORMANCE TABLE =================
        st.markdown("---")
        st.subheader("📅 Cleaned Daily Performance Log")
        st.markdown("Daily records showing performance percentages. Daily monetary penalties have been scrubbed to prevent confusion.")
        
        cols_to_remove = [
            "Eff Penalty Factor", "Eff Penalty", "Strict Bunching Penalty", 
            "Punct Penalty Factor", "Punct Penalty", "Reg Penalty Factor", 
            "Reg Penalty", "Total Daily Penalty", "Strict Bunching Net Operated",
            "Strict Bunching Eff (%)", "Base Strict Bunching Penalty", 
            "Potential Imputed Bunching Savings", "Imputed Bunching Exempted"
        ]
        
        df_clean_log = df_summ_all.drop(columns=[c for c in cols_to_remove if c in df_summ_all.columns], errors='ignore')
        
        if 'Date' in df_clean_log.columns:
            df_clean_log['Date'] = pd.to_datetime(df_clean_log['Date'], errors='coerce').dt.strftime('%Y-%m-%d')
            
        st.dataframe(df_clean_log.astype(str), use_container_width=True)

        # ================= DOWNLOAD COMPRESSED MONTHLY SUMMARY SHEET =================
        st.markdown("---")
        st.subheader("📥 Export Aggregated Monthly Data")
        
        monthly_run_params = [
            {"Parameter": "Date of Generation", "Selected Value": datetime.now().strftime("%Y-%m-%d %H:%M:%S")},
            {"Parameter": "Efficiency Threshold (%)", "Selected Value": str(thresh_eff)},
            {"Parameter": "Efficiency Penalty per Deficit", "Selected Value": str(rate_eff)},
            {"Parameter": "Exempt Bunching caused by Imputed Trips", "Selected Value": str(p_exempt_bunching)},
            {"Parameter": "Punctuality Threshold (%)", "Selected Value": str(thresh_punct)},
            {"Parameter": "Punctuality Penalty per Deficit", "Selected Value": str(rate_punct)},
            {"Parameter": "Regularity Threshold (%)", "Selected Value": str(thresh_reg)},
            {"Parameter": "Regularity Penalty per Deficit", "Selected Value": str(rate_reg)},
            {"Parameter": "Count Ignored Trips enabled", "Selected Value": str(p_count_ignored)},
            {"Parameter": "Count Corrections enabled", "Selected Value": str(p_count_corrections)}
        ]
        
        monthly_totals_summary = [{
            "True Monthly Efficiency (%)": round(grand_eff, 2),
            "True Monthly Punctuality (%)": round(grand_punct, 2),
            "True Monthly Regularity (%)": round(grand_reg, 2),
            "Total Efficiency Penalty ($)": pen_eff,
            "Total Bunching Penalty ($)": pen_bunching,
            "Total Punctuality Penalty ($)": pen_punct,
            "Total Regularity Penalty ($)": pen_reg,
            "Grand Total Penalties ($)": tot_penalties
        }]

        output_buffer = io.BytesIO()
        with pd.ExcelWriter(output_buffer, engine="openpyxl") as writer:
            pd.DataFrame(monthly_totals_summary).T.reset_index().to_excel(writer, sheet_name="MONTHLY_KPI_OVERALL", index=False, header=["KPI Metric / Financial Total", "Value"])
            route_display.to_excel(writer, sheet_name="ROUTE_BREAKDOWN", index=False)
            df_clean_log.to_excel(writer, sheet_name="DAILY_HISTORICAL_LOG", index=False)
            pd.DataFrame(monthly_run_params).to_excel(writer, sheet_name="RUN_PARAMETERS", index=False)
            
            for sheetname in writer.sheets:
                ws = writer.sheets[sheetname]
                for col in ws.columns:
                    max_len = max(len(str(cell.value or '')) for cell in col)
                    col_letter = col[0].column_letter
                    ws.column_dimensions[col_letter].width = max(max_len + 3, 15)

        st.download_button(
            label="📥 Download Aggregated Monthly Report (Excel)",
            data=output_buffer.getvalue(),
            file_name=f"Aggregated_Monthly_KPI_{datetime.now().strftime('%Y%m%d_%H%M')}.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            use_container_width=True
        )
