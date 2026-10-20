"""EWB Operations Portal — unified Streamlit app.

Three tools under one roof:
  Tab 1 — EWB Endorsement Automation (Worklist vs. DataGrid → CMS uploads)
  Tab 2 — EWB Field Result Autostat Builder (daily field visits → autostat upload)
  Tab 3 — EWB BAU Daily Status Report (MC3 DRR → SP MADRID status report)
"""
import sys
from streamlit.runtime import exists

if __name__ == "__main__" and not exists():
    from streamlit.web import cli as stcli
    sys.argv = ["streamlit", "run", __file__]
    sys.exit(stcli.main())

import streamlit as st

st.set_page_config(
    page_title="EWB Operations Portal — SP MADRID",
    page_icon="🏦",
    layout="wide",
    initial_sidebar_state="expanded",
)

# ── Global CSS ────────────────────────────────────────────────────────────────
st.markdown("""
<style>
@import url('https://fonts.googleapis.com/css2?family=Inter:wght@300;400;500;600;700&display=swap');

html, body, [class*="css"] { font-family: 'Inter', sans-serif; }

/* Sidebar */
section[data-testid="stSidebar"] {
    background: linear-gradient(180deg, #0d1b2a 0%, #1b2838 100%);
    border-right: 1px solid #2a3a4a;
}
section[data-testid="stSidebar"] * { color: #e0e8f0 !important; }
section[data-testid="stSidebar"] .stSelectbox label,
section[data-testid="stSidebar"] .stTextInput label,
section[data-testid="stSidebar"] .stCheckbox label { color: #b0c4d8 !important; }

/* Header banner */
.portal-header {
    background: linear-gradient(135deg, #0f3460 0%, #16213e 50%, #0a1628 100%);
    border-radius: 12px;
    padding: 20px 28px 16px;
    margin-bottom: 20px;
    border: 1px solid #1a4070;
    box-shadow: 0 4px 20px rgba(0,0,0,0.4);
}
.portal-header h1 {
    color: #ffffff; font-size: 1.8rem; font-weight: 700; margin: 0 0 4px;
    letter-spacing: -0.5px;
}
.portal-header p {
    color: #7fb3d3; font-size: 0.85rem; margin: 0;
}
.portal-badge {
    display: inline-block;
    background: linear-gradient(90deg, #1f6fab, #2980b9);
    color: white; font-size: 0.7rem; font-weight: 600;
    padding: 3px 10px; border-radius: 20px; margin-right: 6px;
    letter-spacing: 0.5px; text-transform: uppercase;
}

/* Metric cards */
[data-testid="metric-container"] {
    background: linear-gradient(135deg, #1a2a3a, #243447);
    border: 1px solid #2a4060;
    border-radius: 10px;
    padding: 12px 16px;
    box-shadow: 0 2px 8px rgba(0,0,0,0.2);
}
[data-testid="metric-container"] label { color: #8cb4cc !important; font-size: 0.78rem !important; }
[data-testid="metric-container"] [data-testid="stMetricValue"] { color: #e8f4fd !important; font-weight: 700; }

/* Tab styling */
.stTabs [data-baseweb="tab-list"] {
    gap: 6px;
    background: transparent;
    border-bottom: 2px solid #1a3050;
}
.stTabs [data-baseweb="tab"] {
    background: #0d1b2a; border-radius: 8px 8px 0 0;
    color: #7090b0 !important; font-weight: 500;
    padding: 8px 20px; border: 1px solid #1a3050;
    border-bottom: none; transition: all 0.2s;
}
.stTabs [aria-selected="true"] {
    background: linear-gradient(135deg, #1f4e78, #2e75b6) !important;
    color: white !important; border-color: #2e75b6 !important;
}
.stTabs [data-baseweb="tab"]:hover { color: #a0c8e8 !important; }

/* Buttons */
.stButton > button[kind="primary"] {
    background: linear-gradient(135deg, #1f4e78, #2e75b6);
    border: none; border-radius: 8px; color: white; font-weight: 600;
    padding: 8px 24px; transition: all 0.2s;
    box-shadow: 0 2px 8px rgba(30,90,180,0.3);
}
.stButton > button[kind="primary"]:hover {
    background: linear-gradient(135deg, #2e75b6, #3e8fcf);
    box-shadow: 0 4px 16px rgba(30,90,180,0.5);
    transform: translateY(-1px);
}

/* Download buttons */
.stDownloadButton > button {
    background: linear-gradient(135deg, #145a32, #1e8449);
    border: none; border-radius: 8px; color: white; font-weight: 600;
    transition: all 0.2s;
}
.stDownloadButton > button:hover {
    background: linear-gradient(135deg, #1e8449, #27ae60);
    transform: translateY(-1px);
}

/* File uploader */
[data-testid="stFileUploader"] {
    background: #0d1b2a; border: 2px dashed #2a5080;
    border-radius: 10px; padding: 8px;
}
[data-testid="stFileUploader"]:hover { border-color: #3e7cb5; }

/* Step header */
.step-header {
    background: linear-gradient(90deg, #1F4E78 0%, #2E75B6 100%);
    color: white; border-radius: 8px; padding: 10px 16px;
    font-size: 1.05rem; font-weight: 700; margin-bottom: 8px;
}

/* Divider color */
hr { border-color: #1a3050 !important; }
</style>
""", unsafe_allow_html=True)

# ── Header ────────────────────────────────────────────────────────────────────
st.markdown("""
<div class="portal-header">
  <h1>🏦 EWB Operations Portal</h1>
  <p>
    <span class="portal-badge">SP MADRID</span>
    <span class="portal-badge">EastWest Bank</span>
    Endorsement Automation · Field Result Autostat · BAU Daily Status Report
  </p>
</div>
""", unsafe_allow_html=True)

# ── Sidebar nav note ──────────────────────────────────────────────────────────
with st.sidebar:
    st.markdown("""
    <div style='text-align:center; padding: 12px 0 4px;'>
      <span style='font-size:1.6rem'>🏦</span><br>
      <span style='font-weight:700; font-size:1rem; color:#e0e8f0;'>EWB Portal</span><br>
      <span style='font-size:0.75rem; color:#5a7a9a;'>SP MADRID Recovery</span>
    </div>
    <hr style='border-color:#2a3a4a; margin: 12px 0;'>
    """, unsafe_allow_html=True)
    st.markdown("Use the **tabs above** to switch between tools.")
    st.markdown("Settings for each tool appear here when that tab is active.")

# ── Import tab modules ────────────────────────────────────────────────────────
try:
    from modules.endo_tab import render_endo_tab
    _endo_ok = True
except ImportError as e:
    _endo_ok = False
    _endo_err = str(e)

from modules.autostat_tab import render_autostat_tab
from modules.drr_tab import render_drr_tab

# ── Tabs ──────────────────────────────────────────────────────────────────────
tab1, tab2, tab3 = st.tabs([
    "📑 Endorsement Automation",
    "📋 Field Result Autostat",
    "📊 BAU Daily Status Report",
])

with tab1:
    if _endo_ok:
        render_endo_tab()
    else:
        st.error(f"Could not load Endorsement tab: {_endo_err}")
        st.info("Make sure `endo_automation.py` is in the same folder as `app.py`.")

with tab2:
    render_autostat_tab()

with tab3:
    render_drr_tab()
