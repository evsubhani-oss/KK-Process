import streamlit as st
import pandas as pd
import plotly.express as px
import plotly.graph_objects as go

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
    st.caption("Bunching penalties are applied regardless of the monthly average.")
    rate_bunching = st.number_input("Penalty per Bunched Trip", value=50.0, step=5.0)
    
    st.subheader("3. Punctuality")
    thresh_punct = st.number_input("Punctuality Threshold (%)", value=95.0, step=0.1)
    rate_punct = st.number_input("Penalty per Off-Time Trip", value=50.0, step=5.0)
    
    st.subheader("4. Regularity")
    thresh_reg = st.number_input("Regularity Threshold (%)", value=95.0, step=0.1)
    rate_reg = st.number_input("Penalty per Irregular Trip", value=50.0, step=5.0)

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
                    df_summ = pd.read_excel(xls, sheet_name="KPI_SUMMARY")
                    # Clean up navigation columns
                    df_summ = df_summ.drop(columns=[c for c in df_summ.columns if str(c).startswith("Go To ") or c == " " or c == "-- QUICK NAVIGATION --"], errors='ignore')
                    kpi_summaries.append(df_summ)
                
                if "REGULARITY_DETAILS" in xls.sheet_names:
                    df_reg = pd.read_excel(xls, sheet_name="REGULARITY_DETAILS")
                    reg_details.append(df_reg)
                    
            except Exception as e:
                st.error(f"Error reading {f.name}: {e}")
        
        status.update(label="Aggregation Complete!", state="complete")

    if kpi_summaries:
        # ================= DATA AGGREGATION =================
        df_all = pd.concat(kpi_summaries, ignore_index=True)
        
        # Grand Totals for Efficiency & Punctuality (from KPI_SUMMARY)
        tot_target_assigned = df_all['Net Target Assigned'].sum()
        tot_vol_deficit = df_all['Volume Deficit'].sum()
        tot_bunched = df_all['Bunched Trips'].sum()
        
        tot_on_time = df_all['On-Time'].sum()
        tot_off_time = df_all['Off-Time'].sum()
        tot_punct_evaluated = tot_on_time + tot_off_time

        # Grand Totals for Regularity (Preferably from REGULARITY_DETAILS for exact counts)
        if reg_details:
            df_reg_all = pd.concat(reg_details, ignore_index=True)
            tot_reg_evaluated = df_reg_all['Total Evaluated Trips (Weighted)'].sum()
            tot_reg_regular = df_reg_all['Regular Trips'].sum()
            tot_reg_irregular = df_reg_all['Total Irregular Trips'].sum()
        else:
            # Fallback if specific detail sheet is missing
            tot_reg_evaluated = 0
            tot_reg_regular = 0
            tot_reg_irregular = 0

        # ================= CALCULATE TRUE MONTHLY MEANS =================
        grand_eff = ((tot_target_assigned - tot_vol_deficit) / tot_target_assigned * 100) if tot_target_assigned > 0 else 0
        grand_punct = (tot_on_time / tot_punct_evaluated * 100) if tot_punct_evaluated > 0 else 0
        grand_reg = (tot_reg_regular / tot_reg_evaluated * 100) if tot_reg_evaluated > 0 else 0

        # ================= APPLY GATED PENALTIES =================
        pen_bunching = tot_bunched * rate_bunching
        
        pen_eff = (tot_vol_deficit * rate_eff) if grand_eff < thresh_eff else 0
        pen_punct = (tot_off_time * rate_punct) if grand_punct < thresh_punct else 0
        pen_reg = (tot_reg_irregular * rate_reg) if grand_reg < thresh_reg else 0
        
        tot_penalties = pen_bunching + pen_eff + pen_punct + pen_reg

        # ================= VISUALIZATIONS =================
        st.markdown("---")
        st.subheader("🎯 Monthly KPI Performance")
        
        def create_gauge(title, value, threshold):
            fig = go.Figure(go.Indicator(
                mode = "gauge+number",
                value = value,
                number = {'suffix': "%", 'valueformat': ".2f"},
                domain = {'x': [0, 1], 'y': [0, 1]},
                title = {'text': title, 'font': {'size': 18}},
                gauge = {
                    'axis': {'range': [None, 100], 'tickwidth': 1, 'tickcolor': "darkblue"},
                    'bar': {'color': "rgba(0,0,0,0)"}, # Hide default bar
                    'bgcolor': "white",
                    'borderwidth': 2,
                    'bordercolor': "gray",
                    'steps': [
                        {'range': [0, threshold], 'color': "#ffcccb"},   # Red zone
                        {'range': [threshold, 100], 'color': "#d4edda"}  # Green zone
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
        
        fig_waterfall = go.Figure(go.Waterfall(
            name = "Penalties", orientation = "v",
            measure = ["relative", "relative", "relative", "relative", "total"],
            x = ["Bunching Deductions", "Efficiency Deductions", "Punctuality Deductions", "Regularity Deductions", "Total Penalties"],
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
        st.subheader("🔍 Route-by-Route Breakdown (Who dragged the average down?)")
        
        # Group by Route for analysis
        route_grouped = df_all.groupby('Path').sum(numeric_only=True).reset_index()
        route_grouped['Route Efficiency (%)'] = ((route_grouped['Net Target Assigned'] - route_grouped['Volume Deficit']) / route_grouped['Net Target Assigned']) * 100
        route_grouped = route_grouped.sort_values('Route Efficiency (%)', ascending=True)

        fig_bar = px.bar(
            route_grouped, x='Route Efficiency (%)', y='Path', orientation='h',
            title="Efficiency by Route", text='Route Efficiency (%)',
            color='Route Efficiency (%)', color_continuous_scale="RdYlGn", range_color=[80, 100]
        )
        fig_bar.add_vline(x=thresh_eff, line_width=3, line_dash="dash", line_color="red", annotation_text="Threshold")
        fig_bar.update_traces(texttemplate='%{text:.1f}%', textposition='outside')
        fig_bar.update_layout(height=max(300, len(route_grouped) * 40))
        st.plotly_chart(fig_bar, use_container_width=True)

        # ================= CLEAN DAILY PERFORMANCE TABLE =================
        st.markdown("---")
        st.subheader("📅 Cleaned Daily Performance Log")
        st.markdown("Daily records showing performance percentages. Daily monetary penalties have been scrubbed.")
        
        cols_to_remove = [
            "Eff Penalty Factor", "Eff Penalty", "Strict Bunching Penalty", 
            "Punct Penalty Factor", "Punct Penalty", "Reg Penalty Factor", 
            "Reg Penalty", "Total Daily Penalty", "Strict Bunching Net Operated",
            "Strict Bunching Eff (%)"
        ]
        
        df_clean_log = df_all.drop(columns=[c for c in cols_to_remove if c in df_all.columns], errors='ignore')
        
        # Format dates nicely if present
        if 'Date' in df_clean_log.columns:
            df_clean_log['Date'] = pd.to_datetime(df_clean_log['Date'], errors='coerce').dt.strftime('%Y-%m-%d')
            
        st.dataframe(df_clean_log, use_container_width=True)
