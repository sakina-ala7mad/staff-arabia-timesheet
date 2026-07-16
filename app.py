"""
app.py
=======
Top-level entry point. This file's only job is to lay out the two
sections as separate tabs and call each one's render() function. It
contains NO calculation logic of its own — that all lives in
timesheet_engine.py / section1_timesheet_ui.py (Section 1) and
ot_payroll_engine.py / section2_ot_payroll_ui.py (Section 2).

Run locally with:
    streamlit run app.py
"""

import streamlit as st
import section1_timesheet_ui
import section2_ot_payroll_ui

st.set_page_config(
    page_title="Staff Arabia HR Attendance Suite",
    page_icon="🗓️",
    layout="centered",
)

st.markdown(
    """
    <style>
    .stTabs [data-baseweb="tab-list"] { gap: 8px; }
    .stTabs [data-baseweb="tab"] {
        height: 46px;
        white-space: pre-wrap;
        border-radius: 8px 8px 0 0;
        font-weight: 600;
    }
    </style>
    """,
    unsafe_allow_html=True,
)

tab1, tab2 = st.tabs(["🗓️ Section 1 — Attendance Timesheet", "📊 Section 2 — OT & Payroll Summary"])

with tab1:
    section1_timesheet_ui.render()

with tab2:
    section2_ot_payroll_ui.render()
