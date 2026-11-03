import io
import streamlit as st
import pandas as pd
import numpy as np
import plotly.express as px
import plotly.graph_objects as go

import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter
from openpyxl.chart import BarChart, Reference

# ==============================================================================
# DEVELOPER CREDENTIALS & METADATA
# ==============================================================================
DEV_NAME = "Rana FaiSaL"
DEV_EMAIL = "mfaisal.ue@gmail.com"
DEV_TITLE = "Lead Analytics Architect & Developer"

# ==============================================================================
# CODE PASSWORD PROTECTION
# ==============================================================================
def check_password():
    def password_entered():
        if st.session_state["username"] == "admin" and st.session_state["password"] == "Rana@faisal":
            st.session_state["password_correct"] = True
            del st.session_state["password"]  # Security cleanup
        else:
            st.session_state["password_correct"] = False

    if "password_correct" not in st.session_state:
        st.markdown("### 🔒 Suthra Punjab Portal - Secure Login")
        st.text_input("Username", key="username")
        st.text_input("Password", type="password", key="password")
        st.button("Login", on_click=password_entered)
        st.caption(f"System Designed & Developed by **{DEV_NAME}** ({DEV_EMAIL})")
        return False
    elif not st.session_state["password_correct"]:
        st.markdown("### 🔒 Suthra Punjab Portal - Secure Login")
        st.text_input("Username", key="username")
        st.text_input("Password", type="password", key="password")
        st.button("Login", on_click=password_entered)
        st.error("❌ Galat Username ya Password!")
        return False
    else:
        return True

if not check_password():
    st.stop()

# ==============================================================================
# 1. PAGE SETUP & SUTHRA PUNJAB OFFICIAL BRANDING THEME
# ==============================================================================
st.set_page_config(
    page_title="Suthra Punjab - Muzaffargarh Fleet Command",
    page_icon="🌿",
    layout="wide",
    initial_sidebar_state="expanded"
)

# Custom Enterprise CSS with Suthra Punjab Emerald & Gold Palette
st.markdown(f"""
    <style>
    /* Global Base */
    .stApp {{
        background-color: #f7faf8;
    }}
    
    /* Official Top Banner */
    .sp-banner {{
        background: linear-gradient(135deg, #0d5c3a 0%, #157347 60%, #083c25 100%);
        padding: 24px 30px;
        border-radius: 12px;
        box-shadow: 0 4px 15px rgba(13, 92, 58, 0.2);
        margin-bottom: 25px;
        text-align: center;
        border-bottom: 4px solid #d4af37;
    }}
    .sp-gov-tag {{
        font-size: 13px;
        color: #d4af37;
        font-weight: 700;
        letter-spacing: 2px;
        text-transform: uppercase;
        margin-bottom: 4px;
    }}
    .sp-main-title {{
        font-size: 32px;
        color: #ffffff;
        font-weight: 800;
        letter-spacing: -0.5px;
        margin: 0;
    }}
    .sp-sub-title {{
        font-size: 15px;
        color: #e2f0d9;
        font-weight: 400;
        margin-top: 6px;
    }}
    
    /* Executive Metric Scorecards */
    .kpi-card {{
        background: #ffffff;
        padding: 16px;
        border-radius: 10px;
        box-shadow: 0 2px 8px rgba(0, 0, 0, 0.05);
        border-top: 4px solid #0d5c3a;
        text-align: center;
        min-height: 108px;
        transition: transform 0.2s ease, box-shadow 0.2s ease;
    }}
    .kpi-card:hover {{
        transform: translateY(-2px);
        box-shadow: 0 4px 12px rgba(13, 92, 58, 0.15);
    }}
    .kpi-title {{
        font-size: 11px;
        color: #5a6b7c;
        font-weight: 700;
        text-transform: uppercase;
        letter-spacing: 0.5px;
        margin-bottom: 5px;
    }}
    .kpi-value {{
        font-size: 26px;
        color: #0d5c3a;
        font-weight: 800;
        line-height: 1.1;
    }}
    .kpi-sub {{
        font-size: 12px;
        color: #27ae60;
        font-weight: 600;
        margin-top: 5px;
    }}
    .kpi-warn {{ border-top-color: #d4af37 !important; }}
    .kpi-alert {{ border-top-color: #c0392b !important; }}

    /* Analytical Note & Narrative Panels */
    .desc-box {{
        background-color: #ffffff;
        border-left: 4px solid #0d5c3a;
        padding: 14px 18px;
        border-radius: 6px;
        font-size: 13.5px;
        color: #2c3e50;
        margin-bottom: 20px;
        box-shadow: 0 1px 4px rgba(0,0,0,0.04);
        line-height: 1.6;
    }}
    .desc-box-warn {{
        background-color: #fffdf5;
        border-left: 4px solid #d4af37;
        padding: 14px 18px;
        border-radius: 6px;
        font-size: 13.5px;
        color: #7d5a00;
        margin-bottom: 20px;
        box-shadow: 0 1px 4px rgba(0,0,0,0.04);
        line-height: 1.6;
    }}

    /* Developer Footer Badge */
    .dev-footer {{
        background-color: #ffffff;
        border: 1px solid #e2e8f0;
        border-top: 3px solid #0d5c3a;
        padding: 16px;
        border-radius: 8px;
        text-align: center;
        margin-top: 35px;
        color: #4a5568;
        font-size: 13px;
    }}
    .dev-footer strong {{
        color: #0d5c3a;
    }}
    .dev-footer a {{
        color: #157347;
        text-decoration: none;
        font-weight: bold;
    }}
    .dev-footer a:hover {{
        text-decoration: underline;
    }}
    </style>
""", unsafe_allow_html=True)

# Official Title Banner
st.markdown("""
    <div class="sp-banner">
        <div class="sp-gov-tag">Government of the Punjab • Local Government & Community Development</div>
        <div class="sp-main-title">SUTHRA PUNJAB AGENCY — MUZAFFARGARH</div>
        <div class="sp-sub-title">Solid Waste Management Fleet Command Center, Transfer Point Verification & Beat Adherence Audit</div>
    </div>
""", unsafe_allow_html=True)

# ==============================================================================
# 2. SIDEBAR INGESTION & DATA PROCESSING PIPELINE
# ==============================================================================
st.sidebar.markdown("### 🏛️ Suthra Punjab Operations")
st.sidebar.markdown("**District Agency:** Muzaffargarh")
st.sidebar.write("---")

uploaded_files = st.sidebar.file_uploader(
    "Upload Vehicle Performance Sheets (.xlsx, .xls)", 
    accept_multiple_files=True, 
    type=['xlsx', 'xls']
)

if not uploaded_files:
    st.info("👋 Upload your Vehicle Performance Excel sheet from the sidebar to initialize the dashboard.")
    st.sidebar.write("---")
    st.sidebar.markdown(f"**Developed by:**<br>👨‍💻 {DEV_NAME}<br>✉️ [{DEV_EMAIL}](mailto:{DEV_EMAIL})", unsafe_allow_html=True)
    st.stop()

@st.cache_data
def load_and_standardize_data(files):
    dfs = []
    for file in files:
        try:
            temp_df = pd.read_excel(file, sheet_name='Vehicle Performance')
            dfs.append(temp_df)
        except Exception as e:
            st.sidebar.error(f"Error loading {file.name}: {e}")
            
    if not dfs:
        return pd.DataFrame()
        
    df = pd.concat(dfs, ignore_index=True)

    # 1. Clean numeric telemetry columns
    numeric_columns = [
        'Total km', 'Within beat km', 'Outside beat km', 
        'LR beat to WE trips (3 min stop)', 'LR beat to TCP trips (5 min stop)',
        'LR beat to Dumpsite trips (10 min stop)', 'WE to TCP trips (3 min WE stop, 5 min TCP stop)',
        'TCP to Dumpsite trips (VTCS trip loaded at a TCP)', 'WE to Dumpsite trips (VTCS trip loaded at a WE)',
        'Direct Dumpsite trips (VTCS trip with no TCP or WE before it)', 'VTCS trips (valid)',
        'VTCS trips without tracking (valid)'
    ]
    for col in numeric_columns:
        if col in df.columns:
            df[col] = pd.to_numeric(df[col].replace('-', np.nan), errors='coerce').fillna(0)

    # 2. Time conversions to decimal hours
    def convert_to_decimal_hours(series):
        clean_s = series.replace('-', '00:00:00').fillna('00:00:00')
        return pd.to_timedelta(clean_s, errors='coerce').dt.total_seconds().fillna(0) / 3600.0

    if 'Working hours' in df.columns:
        df['Working Hours (Hrs)'] = convert_to_decimal_hours(df['Working hours'])
    if 'Working hours within beat' in df.columns:
        df['Beat Working Hours (Hrs)'] = convert_to_decimal_hours(df['Working hours within beat'])
    if 'Working hours outside beat' in df.columns:
        df['Outside Working Hours (Hrs)'] = convert_to_decimal_hours(df['Working hours outside beat'])

    # 3. Categorize Waste Transfer and Disposal Trip Aggregates
    df['WE Drops'] = (
        df.get('LR beat to WE trips (3 min stop)', 0) + 
        df.get('WE to TCP trips (3 min WE stop, 5 min TCP stop)', 0) + 
        df.get('WE to Dumpsite trips (VTCS trip loaded at a WE)', 0)
    )
    df['TCP Drops'] = (
        df.get('LR beat to TCP trips (5 min stop)', 0) + 
        df.get('WE to TCP trips (3 min WE stop, 5 min TCP stop)', 0) + 
        df.get('TCP to Dumpsite trips (VTCS trip loaded at a TCP)', 0)
    )
    df['Dumpsite Trips'] = (
        df.get('LR beat to Dumpsite trips (10 min stop)', 0) + 
        df.get('TCP to Dumpsite trips (VTCS trip loaded at a TCP)', 0) + 
        df.get('WE to Dumpsite trips (VTCS trip loaded at a WE)', 0) + 
        df.get('Direct Dumpsite trips (VTCS trip with no TCP or WE before it)', 0) + 
        df.get('VTCS trips (valid)', 0)
    )
    df['Total Material Trips'] = df['WE Drops'] + df['TCP Drops'] + df['Dumpsite Trips']

    # 4. Standardize Tracking Status
    df['Live Status'] = np.where(df['Status'].str.contains('available', case=False, na=False), 'Online', 'Offline')

    # 5. Beat Adherence & Routing Compliance Logic
    def evaluate_beat_compliance(row):
        if row.get('Mapped with LR Beat') != 'Yes':
            return 'Not Beat-Mapped'
        w_km = row.get('Within beat km', 0)
        o_km = row.get('Outside beat km', 0)
        tot_km = row.get('Total km', 0)
        
        if tot_km == 0:
            return 'Zero Movement (Idle/Offline)'
        if w_km > 0 and o_km == 0:
            return '100% Inside Beat (Compliant)'
        if w_km > 0 and o_km > 0:
            return 'Mixed (Inside & Outside)'
        if w_km == 0 and o_km > 0:
            return '100% Outside Beat (Deviation)'
        return 'Inactive'

    df['Beat Compliance Status'] = df.apply(evaluate_beat_compliance, axis=1)

    if 'Assigned LR beats' in df.columns:
        df['Assigned LR beats'] = df['Assigned LR beats'].fillna('Unassigned')

    return df

raw_df = load_and_standardize_data(uploaded_files)
if raw_df.empty:
    st.warning("Koi valid records load nahi ho sakay. Sheet check karein.")
    st.stop()

# ==============================================================================
# 3. GLOBAL DRILL-DOWN FILTERS
# ==============================================================================
st.sidebar.header("🔍 Operational Filters")

all_districts = sorted(raw_df['District'].dropna().unique().tolist())
selected_districts = st.sidebar.multiselect("District", options=all_districts, default=all_districts)
f_df = raw_df[raw_df['District'].isin(selected_districts)]

all_tehsils = sorted(f_df['Tehsil'].dropna().unique().tolist())
selected_tehsils = st.sidebar.multiselect("Tehsil", options=all_tehsils, default=all_tehsils)
f_df = f_df[f_df['Tehsil'].isin(selected_tehsils)]

all_vtypes = sorted(f_df['Vehicle type'].dropna().unique().tolist())
selected_vtypes = st.sidebar.multiselect("Vehicle Type", options=all_vtypes, default=all_vtypes)
df = f_df[f_df['Vehicle type'].isin(selected_vtypes)]

# Sidebar Developer Credit Badge
st.sidebar.write("---")
st.sidebar.markdown(f"""
    <div style="background-color: #ffffff; padding: 10px; border-radius: 6px; border-left: 3px solid #0d5c3a; font-size: 12px;">
        <b>Architect & Developer:</b><br>
        👨‍💻 <b>{DEV_NAME}</b><br>
        ✉️ <a href="mailto:{DEV_EMAIL}">{DEV_EMAIL}</a>
    </div>
""", unsafe_allow_html=True)

if df.empty:
    st.warning("Muntakhib karda filter criteria par koi data dastiyab nahi hai.")
    st.stop()

# ==============================================================================
# 4. EXECUTIVE SCOREBOARD (PUNJAB GREEN THEME)
# ==============================================================================
total_fleet = len(df)
online_fleet = len(df[df['Live Status'] == 'Online'])
online_rate = round((online_fleet / total_fleet * 100), 1) if total_fleet > 0 else 0

total_km = round(df['Total km'].sum(), 1)
total_we = int(df['WE Drops'].sum())
total_tcp = int(df['TCP Drops'].sum())
total_dumps = int(df['Dumpsite Trips'].sum())

# Beat Adherence calculation for mapped fleet
mapped_lrs = df[df['Mapped with LR Beat'] == 'Yes']
compliant_count = len(mapped_lrs[mapped_lrs['Within beat km'] > 0])
adherence_rate = round((compliant_count / len(mapped_lrs) * 100), 1) if len(mapped_lrs) > 0 else 0

k1, k2, k3, k4, k5, k6 = st.columns(6)
k1.markdown(f'<div class="kpi-card"><div class="kpi-title">Total Fleet Assets</div><div class="kpi-value">{total_fleet}</div><div class="kpi-sub">{online_fleet} Tracking Live ({online_rate}%)</div></div>', unsafe_allow_html=True)
k2.markdown(f'<div class="kpi-card"><div class="kpi-title">Fleet Mileage</div><div class="kpi-value">{total_km:,.1f}</div><div class="kpi-sub">Total Kilometers</div></div>', unsafe_allow_html=True)
k3.markdown(f'<div class="kpi-card"><div class="kpi-title">WE Point Drops</div><div class="kpi-value">{total_we}</div><div class="kpi-sub">Waste Enclosures</div></div>', unsafe_allow_html=True)
k4.markdown(f'<div class="kpi-card"><div class="kpi-title">TCP Point Drops</div><div class="kpi-value">{total_tcp}</div><div class="kpi-sub">Transfer Points</div></div>', unsafe_allow_html=True)
k5.markdown(f'<div class="kpi-card"><div class="kpi-title">Dumpsite Drops</div><div class="kpi-value">{total_dumps}</div><div class="kpi-sub">Disposal Deliveries</div></div>', unsafe_allow_html=True)
k6.markdown(f'<div class="kpi-card kpi-warn"><div class="kpi-title">Beat Adherence</div><div class="kpi-value">{adherence_rate}%</div><div class="kpi-sub">{compliant_count} of {len(mapped_lrs)} Mapped LRs Active</div></div>', unsafe_allow_html=True)

st.write("---")

# ==============================================================================
# SECTION 1: TRANSFER STATIONS & DISPOSAL MONITORING (WE, TCP, DUMPSITE)
# ==============================================================================
st.subheader("📍 1. Waste Transfer Hubs & Dumpsite Drops (WE, TCP & Disposal Analytics)")

st.markdown("""
<div class="desc-box">
    <b>Suthra Punjab Operational Definitions & Telemetry Rules:</b><br>
    • <b>Waste Enclosure (WE):</b> Neighborhood collection points designed for primary localized collection. (Minimum recorded dwell time: 3 mins).<br>
    • <b>Temporary Collection Point (TCP):</b> Central secondary storage hub aggregated for bulk haulage. (Minimum recorded dwell time: 5 mins).<br>
    • <b>Disposal Dumpsite:</b> Final authorized disposal facility monitored under VTCS trip verification logic.
</div>
""", unsafe_allow_html=True)

tab_pts_charts, tab_pts_data = st.tabs(["📊 Transfer Point Analytics", "📋 Vehicle-Level Trip Manifest"])

with tab_pts_charts:
    col_p1, col_p2 = st.columns(2)
    
    with col_p1:
        drop_distribution = pd.DataFrame({
            'Facility Type': ['Waste Enclosure (WE)', 'Temporary Point (TCP)', 'Disposal Dumpsite'],
            'Total Drops': [total_we, total_tcp, total_dumps]
        })
        fig_facility_bar = px.bar(
            drop_distribution, 
            x='Facility Type', 
            y='Total Drops',
            text='Total Drops',
            color='Facility Type',
            color_discrete_sequence=['#0d5c3a', '#d4af37', '#157347'],
            title="Suthra Punjab: Drops Completed Across Facilities"
        )
        fig_facility_bar.update_traces(textposition='outside')
        fig_facility_bar.update_layout(showlegend=False)
        st.plotly_chart(fig_facility_bar, use_container_width=True)
        
    with col_p2:
        v_trips = df.groupby('Vehicle type')[['WE Drops', 'TCP Drops', 'Dumpsite Trips']].sum().reset_index()
        v_trips = v_trips[(v_trips['WE Drops'] > 0) | (v_trips['TCP Drops'] > 0) | (v_trips['Dumpsite Trips'] > 0)]
        if not v_trips.empty:
            fig_vtype_trips = px.bar(
                v_trips, 
                x='Vehicle type', 
                y=['WE Drops', 'TCP Drops', 'Dumpsite Trips'],
                barmode='group',
                title="Trips Distribution Segmented by Vehicle Classification",
                color_discrete_sequence=['#0d5c3a', '#d4af37', '#27ae60']
            )
            st.plotly_chart(fig_vtype_trips, use_container_width=True)
        else:
            st.info("Selected vehicle types ke leay koi transfer trips recorded nahi hain.")

with tab_pts_data:
    st.markdown("**Inspect Vehicles Visiting Specific Facility Points:**")
    filter_destination = st.radio(
        "Select Target Operational Facility:",
        ["All Operational Trips", "Vehicles Visiting Waste Enclosures (WE)", "Vehicles Visiting Transfer Points (TCP)", "Vehicles Visiting Final Dumpsite"],
        horizontal=True
    )
    
    if filter_destination == "Vehicles Visiting Waste Enclosures (WE)":
        manifest_df = df[df['WE Drops'] > 0]
    elif filter_destination == "Vehicles Visiting Transfer Points (TCP)":
        manifest_df = df[df['TCP Drops'] > 0]
    elif filter_destination == "Vehicles Visiting Final Dumpsite":
        manifest_df = df[df['Dumpsite Trips'] > 0]
    else:
        manifest_df = df[df['Total Material Trips'] > 0]

    cols_to_show = [
        'Tehsil', 'Vehicle', 'Vehicle type', 'Assigned LR beats', 
        'LR beat to WE trips (3 min stop)', 'LR beat to TCP trips (5 min stop)', 
        'WE to TCP trips (3 min WE stop, 5 min TCP stop)', 'VTCS trips (valid)',
        'Total km', 'Working hours', 'Live Status'
    ]
    st.dataframe(
        manifest_df[cols_to_show].sort_values(by='Total km', ascending=False),
        use_container_width=True
    )

st.write("---")

# ==============================================================================
# SECTION 2: LOADER RICKSHAW BEAT ROUTING & AUDIT
# ==============================================================================
st.subheader("🛺 2. Loader Rickshaw Beat Compliance & Exception Audit")

lr_df = df[df['Vehicle type'] == 'Loader Rickshaw']

if lr_df.empty:
    st.info("Muntakhib karda filter selection mein koi Loader Rickshaw shamil nahi hai.")
else:
    st.markdown("""
    <div class="desc-box">
        <b>Suthra Punjab Beat Compliance Framework:</b><br>
        • <b>100% Inside Beat (Compliant):</b> Rickshaws operating exclusively within designated beat boundaries.<br>
        • <b>Mixed Operation:</b> Valid operation inside beat plus transit to authorized WEs or TCPs.<br>
        • <b>100% Outside Beat (Deviation / Violation):</b> Vehicles moving on road but logging 0.0 KM inside assigned beat (Direct operational violation).<br>
        • <b>Underperforming Beats:</b> Beats assigned to vehicles that remained completely static (0 KM) or logged &lt; 5 KM.
    </div>
    """, unsafe_allow_html=True)

    tab_b1, tab_b2, tab_b3 = st.tabs([
        "📊 Beat Compliance Distribution", 
        "🚨 Non-Functional & Low-Activity Beats (< 5 KM)", 
        "📑 Granular Beat Directory Explorer"
    ])

    with tab_b1:
        c_b1, c_b2 = st.columns(2)
        with c_b1:
            comp_dist = lr_df['Beat Compliance Status'].value_counts().reset_index()
            comp_dist.columns = ['Status', 'Count']
            fig_pie_comp = px.pie(
                comp_dist, 
                names='Status', 
                values='Count', 
                hole=0.42,
                title="Suthra Punjab Fleet Beat Compliance Status",
                color='Status',
                color_discrete_map={
                    '100% Inside Beat (Compliant)': '#0d5c3a',
                    'Mixed (Inside & Outside)': '#d4af37',
                    '100% Outside Beat (Deviation)': '#c0392b',
                    'Zero Movement (Idle/Offline)': '#7f8c8d',
                    'Not Beat-Mapped': '#bdc3c7'
                }
            )
            st.plotly_chart(fig_pie_comp, use_container_width=True)
            
        with c_b2:
            km_summary = lr_df.groupby('Tehsil')[['Within beat km', 'Outside beat km']].sum().reset_index()
            fig_km_comp = px.bar(
                km_summary.melt(id_vars='Tehsil', var_name='Metric', value_name='Distance (KM)'),
                x='Tehsil', 
                y='Distance (KM)', 
                color='Metric', 
                barmode='group',
                title="Tehsil Breakdown: Inside Beat vs Outside Beat Mileage",
                color_discrete_map={'Within beat km': '#0d5c3a', 'Outside beat km': '#d4af37'}
            )
            st.plotly_chart(fig_km_comp, use_container_width=True)

    with tab_b2:
        zero_km_fleet = lr_df[(lr_df['Mapped with LR Beat'] == 'Yes') & (lr_df['Total km'] == 0)]
        low_km_fleet = lr_df[(lr_df['Mapped with LR Beat'] == 'Yes') & (lr_df['Total km'] > 0) & (lr_df['Total km'] < 5.0)]
        outside_fleet = lr_df[(lr_df['Mapped with LR Beat'] == 'Yes') & (lr_df['Within beat km'] == 0) & (lr_df['Outside beat km'] > 0)]

        st.markdown(f"""
        <div class="desc-box-warn">
            <b>Action Required Summary (Muzaffargarh District):</b><br>
            A total of <b>{len(zero_km_fleet)}</b> assigned beats registered <b>Zero Movement (0 KM)</b>, while <b>{len(low_km_fleet)}</b> beats logged critically sub-optimal mileage (&lt; 5 KM). Furthermore, <b>{len(outside_fleet)}</b> vehicles operated exclusively outside their assigned beats.
        </div>
        """, unsafe_allow_html=True)

        selected_audit_view = st.radio(
            "Select Exception Category for Direct Field Inspection:",
            ["Zero Movement Beats (0 KM)", "Low-Activity Beats (< 5 KM)", "100% Route Deviations (Zero Inside Beat KM)"],
            horizontal=True
        )

        if selected_audit_view == "Zero Movement Beats (0 KM)":
            audit_records = zero_km_fleet
        elif selected_audit_view == "Low-Activity Beats (< 5 KM)":
            audit_records = low_km_fleet
        else:
            audit_records = outside_fleet

        st.dataframe(
            audit_records[[
                'Tehsil', 'Vehicle', 'Assigned LR beats', 'Total km',
                'Within beat km', 'Outside beat km', 'Working hours', 
                'Beat Compliance Status', 'Live Status'
            ]].sort_values(by='Total km', ascending=True),
            use_container_width=True
        )

    with tab_b3:
        st.markdown("**Inspect Individual Assigned Beat Numbers:**")
        valid_beat_list = sorted([b for b in lr_df['Assigned LR beats'].unique() if b != 'Unassigned'])
        beat_filter_choice = st.selectbox("Select Target Beat Number to Audit:", options=["All Assigned Beats"] + valid_beat_list)

        if beat_filter_choice == "All Assigned Beats":
            beat_details = lr_df[lr_df['Mapped with LR Beat'] == 'Yes']
        else:
            beat_details = lr_df[lr_df['Assigned LR beats'] == beat_filter_choice]

        st.dataframe(
            beat_details[[
                'Tehsil', 'Vehicle', 'Assigned LR beats', 'Beat Compliance Status',
                'Within beat km', 'Outside beat km', 'Total km',
                'Working hours within beat', 'Working hours outside beat', 'Working hours', 'Live Status'
            ]],
            use_container_width=True
        )

st.write("---")

# ==============================================================================
# SECTION 3: ASSET UTILIZATION & EFFICIENCY CORRELATION
# ==============================================================================
st.subheader("📈 3. Fleet Utilization & Tracking Reliability")

st.markdown("""
<div class="desc-box">
    <b>Efficiency Analytics Note:</b><br>
    The scatter plot examines total engine working hours versus distance traveled. Assets logging high engine hours but low kilometers indicate excessive idling or static collection operations.
</div>
""", unsafe_allow_html=True)

col_corr1, col_corr2 = st.columns(2)

with col_corr1:
    fig_efficiency = px.scatter(
        df[df['Working Hours (Hrs)'] > 0],
        x='Working Hours (Hrs)', 
        y='Total km', 
        color='Vehicle type',
        hover_data=['Vehicle', 'Tehsil', 'Assigned LR beats'],
        title="Fleet Efficiency: Operating Hours vs Distance Covered (KM)",
        color_discrete_sequence=px.colors.qualitative.Dark24
    )
    st.plotly_chart(fig_efficiency, use_container_width=True)

with col_corr2:
    status_df = df.groupby(['Vehicle type', 'Live Status']).size().reset_index(name='Fleet Count')
    fig_tracking = px.bar(
        status_df, 
        x='Vehicle type', 
        y='Fleet Count', 
        color='Live Status',
        barmode='stack',
        title="Tracking Availability Profile (Online vs Offline)",
        color_discrete_map={'Online': '#0d5c3a', 'Offline': '#c0392b'}
    )
    st.plotly_chart(fig_tracking, use_container_width=True)

st.write("---")

# ==============================================================================
# SECTION 4: MASTER DATA DIRECTORY & FORMATTED EXCEL EXPORT (WITH EMBEDDED CHARTS)
# ==============================================================================
st.subheader("🗄️ 4. Master Data Directory & Formatted Export Center")

st.markdown("""
<div class="desc-box">
    <b>Executive Reporting Export:</b><br>
    Neeche diye gaye download buttons se aap filtered master audit data download kar sakte hain. <b>Excel (.xlsx)</b> option mukammal corporate formatting, developer attribution metadata aur <b>embedded bar charts</b> ke sath workbook generate karta hai.
</div>
""", unsafe_allow_html=True)

st.dataframe(df, use_container_width=True)

# Function to generate fully formatted Excel with Native Charts, Styles & Developer Credits
@st.cache_data
def generate_suthra_punjab_excel(data_frame):
    buffer = io.BytesIO()
    wb = openpyxl.Workbook()
    
    # -------------------------------------------------------------
    # TAB 1: EXECUTIVE SUMMARY SHEET WITH CHARTS
    # -------------------------------------------------------------
    ws_summary = wb.active
    ws_summary.title = "Executive_Summary"
    ws_summary.views.sheetView[0].showGridLines = True
    
    # Color & Font Definitions
    primary_green = "0D5C3A"
    accent_gold = "D4AF37"
    header_fill = PatternFill(start_color=primary_green, end_color=primary_green, fill_type="solid")
    gold_fill = PatternFill(start_color=accent_gold, end_color=accent_gold, fill_type="solid")
    light_fill = PatternFill(start_color="F2F7F4", end_color="F2F7F4", fill_type="solid")
    credit_fill = PatternFill(start_color="EAEAEA", end_color="EAEAEA", fill_type="solid")
    
    title_font = Font(name="Calibri", size=15, bold=True, color="FFFFFF")
    sub_font = Font(name="Calibri", size=11, bold=True, color=primary_green)
    tbl_header_font = Font(name="Calibri", size=11, bold=True, color="FFFFFF")
    bold_font = Font(name="Calibri", size=11, bold=True)
    normal_font = Font(name="Calibri", size=11)
    meta_font = Font(name="Calibri", size=10, italic=True, color="333333")
    meta_bold = Font(name="Calibri", size=10, bold=True, color="0D5C3A")
    
    thin_border = Border(
        left=Side(style='thin', color='DDDDDD'),
        right=Side(style='thin', color='DDDDDD'),
        top=Side(style='thin', color='DDDDDD'),
        bottom=Side(style='thin', color='DDDDDD')
    )

    # 1. Title Banner in Excel
    ws_summary.merge_cells("A1:G1")
    title_cell = ws_summary["A1"]
    title_cell.value = "SUTHRA PUNJAB AGENCY (MUZAFFARGARH) — EXECUTIVE FLEET AUDIT"
    title_cell.font = title_font
    title_cell.fill = header_fill
    title_cell.alignment = Alignment(horizontal="center", vertical="center")
    ws_summary.row_dimensions[1].height = 35

    # 2. Developer Attribution Block in Excel
    ws_summary.merge_cells("A2:G2")
    dev_cell = ws_summary["A2"]
    dev_cell.value = f"System Designed & Developed by: {DEV_NAME} | Contact: {DEV_EMAIL} | SWM Analytics Wing"
    dev_cell.font = meta_bold
    dev_cell.fill = credit_fill
    dev_cell.alignment = Alignment(horizontal="center", vertical="center")
    ws_summary.row_dimensions[2].height = 20

    # 3. Executive KPI Summary Table
    ws_summary["A4"].value = "Operational Key Performance Indicators (KPIs)"
    ws_summary["A4"].font = sub_font
    
    kpi_headers = ["Metric Parameter", "Audit Value", "Unit / Context"]
    for col_idx, text in enumerate(kpi_headers, start=1):
        cell = ws_summary.cell(row=5, column=col_idx, value=text)
        cell.font = tbl_header_font
        cell.fill = header_fill
        cell.alignment = Alignment(horizontal="center", vertical="center")
        cell.border = thin_border
    
    t_fleet = len(data_frame)
    on_fleet = len(data_frame[data_frame['Live Status'] == 'Online'])
    tot_km = round(data_frame['Total km'].sum(), 1)
    w_drops = int(data_frame['WE Drops'].sum())
    t_drops = int(data_frame['TCP Drops'].sum())
    d_trips = int(data_frame['Dumpsite Trips'].sum())
    
    kpi_rows = [
        ["Total Fleet Assets", t_fleet, "Active Registered Units"],
        ["Online / Tracking Fleet", on_fleet, f"Tracking Availability ({round(on_fleet/t_fleet*100,1)}%)"],
        ["Total Distance Traveled", tot_km, "Kilometers (KM)"],
        ["Waste Enclosures (WE) Drops", w_drops, "Neighborhood Collection Points"],
        ["Temporary Collection (TCP) Drops", t_drops, "Secondary Aggregation Points"],
        ["Final Dumpsite Deliveries", d_trips, "VTCS Valid Dumpsite Drops"]
    ]
    
    for r_idx, row_data in enumerate(kpi_rows, start=6):
        for c_idx, val in enumerate(row_data, start=1):
            cell = ws_summary.cell(row=r_idx, column=c_idx, value=val)
            cell.font = bold_font if c_idx == 1 else normal_font
            cell.border = thin_border
            if r_idx % 2 == 0:
                cell.fill = light_fill
            cell.alignment = Alignment(horizontal="left" if c_idx != 2 else "center", vertical="center")

    # 4. Drops Breakdown Table for Chart Source
    ws_summary["E4"].value = "Facility Drops Summary"
    ws_summary["E4"].font = sub_font
    
    ws_summary.cell(row=5, column=5, value="Facility Category").font = tbl_header_font
    ws_summary.cell(row=5, column=5).fill = gold_fill
    ws_summary.cell(row=5, column=5).alignment = Alignment(horizontal="center")
    
    ws_summary.cell(row=5, column=6, value="Recorded Drops").font = tbl_header_font
    ws_summary.cell(row=5, column=6).fill = gold_fill
    ws_summary.cell(row=5, column=6).alignment = Alignment(horizontal="center")

    drop_table_data = [
        ("Waste Enclosures (WE)", w_drops),
        ("Temporary Points (TCP)", t_drops),
        ("Final Dumpsite", d_trips)
    ]
    for idx, (cat, val) in enumerate(drop_table_data, start=6):
        c1 = ws_summary.cell(row=idx, column=5, value=cat)
        c2 = ws_summary.cell(row=idx, column=6, value=val)
        c1.font = normal_font; c2.font = bold_font
        c1.border = thin_border; c2.border = thin_border
        c2.alignment = Alignment(horizontal="center")

    # 5. Embed Native Openpyxl Bar Chart
    chart = BarChart()
    chart.type = "col"
    chart.style = 10
    chart.title = "Suthra Punjab: Facility Drops Overview"
    chart.y_axis.title = "Number of Drops / Trips"
    chart.x_axis.title = "Facility Classification"
    chart.height = 12
    chart.width = 18

    chart_data = Reference(ws_summary, min_col=6, min_row=5, max_row=8)
    chart_cats = Reference(ws_summary, min_col=5, min_row=6, max_row=8)
    chart.add_data(chart_data, titles_from_data=True)
    chart.set_categories(chart_cats)
    ws_summary.add_chart(chart, "A14")

    # Adjust Column Widths for Executive Summary
    for col in ws_summary.columns:
        max_len = max(len(str(cell.value or '')) for cell in col)
        col_letter = get_column_letter(col[0].column)
        ws_summary.column_dimensions[col_letter].width = max(max_len + 4, 15)

    # -------------------------------------------------------------
    # TAB 2: DETAILED MASTER DATA SHEET
    # -------------------------------------------------------------
    ws_data = wb.create_sheet(title="Master_Fleet_Data")
    ws_data.views.sheetView[0].showGridLines = True
    
    # Columns to include in export
    export_columns = [
        'Tehsil', 'Vehicle', 'Vehicle type', 'Assigned LR beats', 
        'Beat Compliance Status', 'Live Status', 'Total km', 
        'Within beat km', 'Outside beat km', 'Working hours',
        'WE Drops', 'TCP Drops', 'Dumpsite Trips', 'Total Material Trips'
    ]
    available_cols = [c for c in export_columns if c in data_frame.columns]
    data_to_write = data_frame[available_cols].copy()
    
    # Write Headers
    for c_idx, col_name in enumerate(available_cols, start=1):
        cell = ws_data.cell(row=1, column=c_idx, value=col_name)
        cell.font = tbl_header_font
        cell.fill = header_fill
        cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
        cell.border = thin_border
    ws_data.row_dimensions[1].height = 28

    # Write Data Rows
    for r_idx, row in enumerate(data_to_write.itertuples(index=False), start=2):
        for c_idx, value in enumerate(row, start=1):
            cell = ws_data.cell(row=r_idx, column=c_idx, value=value)
            cell.font = normal_font
            cell.border = thin_border
            if r_idx % 2 == 0:
                cell.fill = light_fill
            # Align numeric vs text
            if isinstance(value, (int, float)):
                cell.alignment = Alignment(horizontal="right", vertical="center")
            else:
                cell.alignment = Alignment(horizontal="left", vertical="center")
        ws_data.row_dimensions[r_idx].height = 20

    # Auto-adjust column widths for Data Sheet
    for col in ws_data.columns:
        max_len = max(len(str(cell.value or '')) for cell in col)
        col_letter = get_column_letter(col[0].column)
        ws_data.column_dimensions[col_letter].width = min(max(max_len + 3, 12), 35)

    # Save to memory buffer
    wb.save(buffer)
    return buffer.getvalue()

@st.cache_data
def generate_csv_binary(data_frame):
    return data_frame.to_csv(index=False).encode('utf-8')

col_dl1, col_dl2 = st.columns(2)

with col_dl1:
    csv_bytes = generate_csv_binary(df)
    st.download_button(
        label="📥 Download Data (Standard CSV)",
        data=csv_bytes,
        file_name='Suthra_Punjab_Muzaffargarh_Report.csv',
        mime='text/csv',
        use_container_width=True
    )

with col_dl2:
    excel_bytes = generate_suthra_punjab_excel(df)
    st.download_button(
        label="📊 Download Official Audit Report (Excel .xlsx with Charts)",
        data=excel_bytes,
        file_name='Suthra_Punjab_Muzaffargarh_Audit_Report.xlsx',
        mime='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
        use_container_width=True
    )

# ==============================================================================
# PERMANENT DEVELOPER SIGNATURE FOOTER
# ==============================================================================
st.write("---")
st.markdown(f"""
    <div class="dev-footer">
        <b>Suthra Punjab Fleet Command & Telemetry Engine</b> • District Muzaffargarh<br>
        Developed & Maintained by: <strong>{DEV_NAME}</strong> ({DEV_TITLE})<br>
        Inquiries & Operational Analytics Support: <a href="mailto:{DEV_EMAIL}">{DEV_EMAIL}</a>
    </div>
""", unsafe_allow_html=True)