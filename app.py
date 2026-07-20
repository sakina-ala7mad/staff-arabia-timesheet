"""
app.py
=======
Streamlit UI for Staff Arabia's HR Tools.

Two fully independent tools live here, side by side:
  1. Timesheet Generator          -> timesheet_engine.py
  2. OT & Attendance Summary      -> overtime_summary_engine.py

They do NOT share files, session state keys, or settings — each tab has its
own uploader, its own options, and its own download button, so using one
never affects or resets the other.

Run locally with:
    streamlit run app.py
"""

import io
import datetime
import pandas as pd
import streamlit as st

from timesheet_engine import (
    generate_timesheet,
    TimesheetError,
    DEFAULT_WEEKEND_DAYS as TS_DEFAULT_WEEKEND_DAYS,
)
from overtime_summary_engine import (
    generate_ot_summary,
    OTSummaryError,
    DEFAULT_WEEKEND_DAYS as OT_DEFAULT_WEEKEND_DAYS,
)
from daily_checkup_engine import (
    generate_daily_checkup,
    get_date_bounds,
    classify_files as dc_classify_files,
    load_dataframes as dc_load_dataframes,
    DailyCheckupError,
    STATUS_COMPLETE,
    STATUS_MISSING_OUT,
    STATUS_MISSING_IN,
    STATUS_NO_PUNCH,
)

st.set_page_config(
    page_title="Staff Arabia HR Tools",
    page_icon="🗓️",
    layout="centered",
)

DAY_NAMES = ["Monday", "Tuesday", "Wednesday", "Thursday", "Friday", "Saturday", "Sunday"]

# ── Global styling: clean, glossy, professional ─────────────────────────────
st.markdown(
    """
    <style>
    #MainMenu, footer {visibility: hidden;}

    .stApp {
        background: linear-gradient(180deg, #F7F9FC 0%, #EEF2F8 100%);
    }

    .block-container {
        padding-top: 2rem;
        max-width: 880px;
    }

    /* Hero header */
    .hero {
        background: linear-gradient(135deg, #1E3A5F 0%, #2C5282 55%, #0F766E 100%);
        border-radius: 18px;
        padding: 28px 32px;
        margin-bottom: 26px;
        box-shadow: 0 10px 30px rgba(30, 58, 95, 0.25);
    }
    .hero h1 {
        color: #FFFFFF;
        font-size: 1.65rem;
        font-weight: 700;
        margin: 0 0 6px 0;
    }
    .hero p {
        color: #DCEAF5;
        font-size: 0.95rem;
        margin: 0;
    }

    /* Section / card styling */
    .card {
        background: #FFFFFF;
        border-radius: 14px;
        padding: 22px 24px;
        box-shadow: 0 4px 18px rgba(20, 30, 60, 0.06);
        border: 1px solid rgba(20, 30, 60, 0.05);
        margin-bottom: 18px;
    }

    .section-badge {
        display: inline-block;
        padding: 4px 12px;
        border-radius: 999px;
        font-size: 0.75rem;
        font-weight: 600;
        letter-spacing: 0.02em;
        margin-bottom: 10px;
    }
    .badge-blue { background: #E0EAFB; color: #1E3A5F; }
    .badge-teal { background: #CCFBF1; color: #0F766E; }
    .badge-amber { background: #FEF3C7; color: #92400E; }

    /* Dashboard metric cards */
    [data-testid="stMetric"] {
        background: #FFFFFF;
        border-radius: 12px;
        padding: 14px 16px;
        box-shadow: 0 2px 10px rgba(20, 30, 60, 0.06);
        border: 1px solid rgba(20, 30, 60, 0.05);
    }

    /* Tabs */
    .stTabs [data-baseweb="tab-list"] {
        gap: 6px;
        background: #FFFFFF;
        padding: 6px;
        border-radius: 14px;
        box-shadow: 0 2px 10px rgba(20, 30, 60, 0.06);
    }
    .stTabs [data-baseweb="tab"] {
        border-radius: 10px;
        padding: 10px 18px;
        font-weight: 600;
        color: #4B5563;
    }
    .stTabs [aria-selected="true"] {
        background: linear-gradient(135deg, #2C5282, #0F766E);
        color: #FFFFFF !important;
    }

    /* Buttons */
    .stButton > button, .stDownloadButton > button {
        border-radius: 10px;
        font-weight: 600;
        border: none;
        padding: 0.6rem 1rem;
        box-shadow: 0 4px 14px rgba(15, 118, 110, 0.25);
    }

    /* File uploader */
    [data-testid="stFileUploaderDropzone"] {
        border-radius: 12px;
        background: #FAFBFF;
    }
    </style>
    """,
    unsafe_allow_html=True,
)

st.markdown(
    """
    <div class="hero">
        <h1>🗓️ Staff Arabia — HR Tools</h1>
        <p>Two independent tools in one place. Pick a tab below — each one has its own files, settings, and export, so they never interfere with each other.</p>
    </div>
    """,
    unsafe_allow_html=True,
)

tab_ts, tab_ot, tab_dc = st.tabs([
    "📋  Timesheet Generator",
    "🕐  OT & Attendance Summary",
    "✅  Daily Check-Up",
])


# ═════════════════════════════════════════════════════════════════════════
# TAB 1 — TIMESHEET GENERATOR  (unchanged logic, timesheet_engine.py)
# ═════════════════════════════════════════════════════════════════════════
with tab_ts:
    st.markdown('<span class="section-badge badge-blue">SECTION 1</span>', unsafe_allow_html=True)
    st.markdown("#### Per-Employee Daily Timesheet")
    st.write(
        "Upload the 3 required Excel files below. The app figures out which file "
        "is which automatically — file names don't matter."
    )

    with st.expander("What are the 3 files?", expanded=False):
        st.markdown(
            """
- **Attendance / System file** — must have a column called **`I/O`**
- **Employees Data file** — must have columns **`Employees Name`** and **`Title`**
- **Vacation Transaction file** — must have columns **`Vacation`** and **`From`**
            """
        )

    with st.expander("⚙️ Settings", expanded=False):
        st.caption("Defaults are fine for normal use — only change these if you know you need to.")
        ts_workday_hrs = st.number_input(
            "Standard workday length (hours)", min_value=1.0, max_value=24.0, value=8.0, step=0.5,
            key="ts_workday_hrs",
        )
        ts_work_start = st.text_input(
            "Official work start time (24h, e.g. 10:00)", value="10:00", key="ts_work_start"
        )
        st.caption("Weekend days:")
        ts_default_selected = [DAY_NAMES[i] for i in sorted(TS_DEFAULT_WEEKEND_DAYS)]
        ts_selected_weekend_days = st.multiselect(
            "Days treated as weekend", DAY_NAMES, default=ts_default_selected, key="ts_weekend_days"
        )
        ts_weekend_days = {DAY_NAMES.index(d) for d in ts_selected_weekend_days}

    ts_uploaded_files = st.file_uploader(
        "Upload all 3 Excel files here (you can select all 3 at once)",
        type=["xlsx", "xls"],
        accept_multiple_files=True,
        key="ts_uploader",
    )

    st.divider()

    if ts_uploaded_files:
        if len(ts_uploaded_files) != 3:
            st.warning(f"Please upload exactly 3 files. You've uploaded {len(ts_uploaded_files)}.")
        else:
            st.success(f"3 files ready: {', '.join(f.name for f in ts_uploaded_files)}")

            if st.button("🚀 Generate Timesheet", type="primary", use_container_width=True, key="ts_generate"):
                file_dict = {f.name: io.BytesIO(f.getvalue()) for f in ts_uploaded_files}
                progress_bar = st.progress(0, text="Starting...")

                def ts_update_progress(current, total):
                    progress_bar.progress(current / total, text=f"Processing employee {current}/{total}...")

                try:
                    with st.spinner("Reading and validating files..."):
                        wb, output_filename, stats = generate_timesheet(
                            file_dict,
                            workday_hrs=ts_workday_hrs,
                            work_start=ts_work_start,
                            weekend_days=ts_weekend_days,
                            progress_callback=ts_update_progress,
                        )

                    progress_bar.progress(1.0, text="Done!")
                    st.success(
                        f"✅ Timesheet generated: **{stats['num_employees']} employees** "
                        f"× **{stats['num_days']} days** ({stats['month_str']})"
                    )

                    buffer = io.BytesIO()
                    wb.save(buffer)
                    buffer.seek(0)

                    st.download_button(
                        label="⬇️ Download Timesheet (.xlsx)",
                        data=buffer,
                        file_name=output_filename,
                        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                        use_container_width=True,
                        key="ts_download",
                    )

                except TimesheetError as e:
                    st.error(f"⚠️ {e}")
                except Exception as e:
                    st.error(f"❌ Something unexpected went wrong: {e}")
    else:
        st.info("Waiting for you to upload the 3 files above.")


# ═════════════════════════════════════════════════════════════════════════
# TAB 2 — OT & ATTENDANCE SUMMARY  (new logic, overtime_summary_engine.py)
# ═════════════════════════════════════════════════════════════════════════
with tab_ot:
    st.markdown('<span class="section-badge badge-teal">SECTION 2</span>', unsafe_allow_html=True)
    st.markdown("#### Monthly OT & Attendance Summary")
    st.write(
        "Upload the 4 required Excel files below. One row per employee with "
        "Morning/Night overtime, cancelled day-offs, official holidays worked, "
        "unpaid leave days, and total working days."
    )

    with st.expander("What are the 4 files?", expanded=False):
        st.markdown(
            """
- **Attendance / Punches file** — must have a column called **`I/O`**
- **Vacation Transaction file** — must have columns **`Vacation`** and **`From`**
- **Employee master file** — must have columns **`Employees Name`** and **`Title`**
  (e.g. the "Data" sheet of the employee list workbook)
- **Official Holidays file** — must have an Arabic date column **`التاريخ`** and **`المناسبة`**
            """
        )

    with st.expander("How the numbers are calculated", expanded=False):
        st.markdown(
            """
- **Morning / Night OT Hours** — on regular working days only. Overtime is any
  time worked beyond the standard workday length, counted from *(time in + standard hours)*
  to *time out*. That overtime span is then split at the clock: the portion between
  **7:00–19:00** counts as Morning OT, the portion between **19:00–7:00** counts as Night OT.
  The workday's start time is flexible (9–11am) — this always measures from whenever the
  employee actually punched in, not a fixed clock time.
- **Cancel Day Offs / Days** — for weekends or official holidays where the employee worked
  *less than* the standard workday: hours ÷ standard hours, as a fractional day.
- **Official Holiday / Days** — for weekends or official holidays where the employee worked
  the *full* standard workday or more: counted as 1 full day.
- **Unpaid Leave** — regular working days with no punch at all **and** no vacation/mission/leave
  record of any kind (any status) covering that date.
- **Total Working Days** — every day the employee is credited for: days actually punched in,
  **plus** every Friday/Saturday and official holiday (counted automatically), **plus** any
  day covered by a leave/mission record. Only true Unpaid Leave days are excluded.
            """
        )

    with st.expander("⚙️ Settings", expanded=False):
        st.caption("Defaults are fine for normal use — only change these if you know you need to.")
        ot_workday_hrs = st.number_input(
            "Standard workday length (hours)", min_value=1.0, max_value=24.0, value=8.0, step=0.5,
            key="ot_workday_hrs",
        )
        st.caption("Weekend days:")
        ot_default_selected = [DAY_NAMES[i] for i in sorted(OT_DEFAULT_WEEKEND_DAYS)]
        ot_selected_weekend_days = st.multiselect(
            "Days treated as weekend", DAY_NAMES, default=ot_default_selected, key="ot_weekend_days"
        )
        ot_weekend_days = {DAY_NAMES.index(d) for d in ot_selected_weekend_days}

    ot_uploaded_files = st.file_uploader(
        "Upload all 4 Excel files here (you can select all 4 at once)",
        type=["xlsx", "xls"],
        accept_multiple_files=True,
        key="ot_uploader",
    )

    st.divider()

    if ot_uploaded_files:
        if len(ot_uploaded_files) != 4:
            st.warning(f"Please upload exactly 4 files. You've uploaded {len(ot_uploaded_files)}.")
        else:
            st.success(f"4 files ready: {', '.join(f.name for f in ot_uploaded_files)}")

            if st.button("🚀 Generate OT Summary", type="primary", use_container_width=True, key="ot_generate"):
                file_dict = {f.name: io.BytesIO(f.getvalue()) for f in ot_uploaded_files}
                progress_bar = st.progress(0, text="Starting...")

                def ot_update_progress(current, total):
                    progress_bar.progress(current / total, text=f"Processing employee {current}/{total}...")

                try:
                    with st.spinner("Reading and validating files..."):
                        wb, output_filename, stats = generate_ot_summary(
                            file_dict,
                            workday_hrs=ot_workday_hrs,
                            weekend_days=ot_weekend_days,
                            progress_callback=ot_update_progress,
                        )

                    progress_bar.progress(1.0, text="Done!")
                    st.success(
                        f"✅ Summary generated: **{stats['num_employees']} employees** "
                        f"× **{stats['num_days']} days** ({stats['month_str']}) — "
                        f"{stats['num_holidays_in_period']} official holiday(s) in range"
                    )

                    buffer = io.BytesIO()
                    wb.save(buffer)
                    buffer.seek(0)

                    st.download_button(
                        label="⬇️ Download OT & Attendance Summary (.xlsx)",
                        data=buffer,
                        file_name=output_filename,
                        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                        use_container_width=True,
                        key="ot_download",
                    )

                except OTSummaryError as e:
                    st.error(f"⚠️ {e}")
                except Exception as e:
                    st.error(f"❌ Something unexpected went wrong: {e}")
    else:
        st.info("Waiting for you to upload the 4 files above.")


# ═════════════════════════════════════════════════════════════════════════
# TAB 3 — DAILY CHECK-UP  (new logic, daily_checkup_engine.py)
# ═════════════════════════════════════════════════════════════════════════
with tab_dc:
    st.markdown('<span class="section-badge badge-amber">SECTION 3</span>', unsafe_allow_html=True)
    st.markdown("#### Daily Punch Check-Up")
    st.write(
        "Upload the 3 files below, pick a date, and instantly see who has an "
        "incomplete or missing punch record for that day — and whether they "
        "have a leave/mission on file to explain it."
    )

    with st.expander("What are the 3 files?", expanded=False):
        st.markdown(
            """
- **Attendance / Punches file** — must have a column called **`I/O`**
- **Vacation / Leave Transaction file** — must have columns **`Vacation`** and **`From`**
- **Employee master file** — must have columns **`Employees Name`** and **`Title`**
  (reads the *first* sheet — e.g. the "Data" sheet)
            """
        )

    with st.expander("What counts as an issue?", expanded=False):
        st.markdown(
            """
- **Punched In — Missing Out** — has a Punch In that day, no Punch Out.
- **Punched Out — Missing In** — has a Punch Out that day, no Punch In.
- **No Punch At All** — no punch record at all for that day.

For every flagged employee, the dashboard also checks the Vacation/Leave file
for **any** record (any type, any status) covering that date, so you can see
at a glance whether the gap is explained.
            """
        )

    dc_uploaded_files = st.file_uploader(
        "Upload all 3 Excel files here (you can select all 3 at once)",
        type=["xlsx", "xls"],
        accept_multiple_files=True,
        key="dc_uploader",
    )

    st.divider()

    if dc_uploaded_files:
        if len(dc_uploaded_files) != 3:
            st.warning(f"Please upload exactly 3 files. You've uploaded {len(dc_uploaded_files)}.")
        else:
            st.success(f"3 files ready: {', '.join(f.name for f in dc_uploaded_files)}")

            # Peek the attendance file's date range so the date picker is bounded correctly
            try:
                file_dict_peek = {f.name: io.BytesIO(f.getvalue()) for f in dc_uploaded_files}
                att_f, vac_f, emp_f = dc_classify_files(file_dict_peek)
                df_att_peek, _, _ = dc_load_dataframes(att_f, vac_f, emp_f)
                min_date, max_date = get_date_bounds(df_att_peek)

                col_a, col_b = st.columns([2, 1])
                with col_a:
                    check_date = st.date_input(
                        "Date to check",
                        value=max_date,
                        min_value=min_date,
                        max_value=max_date,
                        key="dc_check_date",
                    )
                with col_b:
                    include_complete = st.checkbox(
                        "Show everyone (incl. complete)", value=False, key="dc_include_complete"
                    )

                if st.button("🔍 Run Check-Up", type="primary", use_container_width=True, key="dc_run"):
                    file_dict = {f.name: io.BytesIO(f.getvalue()) for f in dc_uploaded_files}
                    with st.spinner("Checking attendance records..."):
                        records, stats, _ = generate_daily_checkup(
                            file_dict, check_date, include_complete=include_complete
                        )
                    st.session_state["dc_results"] = (records, stats)

            except DailyCheckupError as e:
                st.error(f"⚠️ {e}")
            except Exception as e:
                st.error(f"❌ Something unexpected went wrong: {e}")

            # ── Dashboard display ───────────────────────────────────────────
            if "dc_results" in st.session_state:
                records, stats = st.session_state["dc_results"]

                st.markdown(f"##### Results for **{stats['check_date'].strftime('%A, %d %B %Y')}**")

                m1, m2, m3, m4 = st.columns(4)
                m1.metric("✅ Complete", stats["complete"])
                m2.metric("🟠 Missing Out", stats["missing_out"])
                m3.metric("🔵 Missing In", stats["missing_in"])
                m4.metric("🔴 No Punch", stats["no_punch"])

                st.caption(
                    f"{stats['problem_total']} of {stats['total_employees']} employees "
                    f"need a look — table below."
                )

                if records:
                    df_view = pd.DataFrame(records)

                    status_options = df_view["Status"].unique().tolist()
                    filter_col1, filter_col2 = st.columns([2, 1])
                    with filter_col1:
                        status_filter = st.multiselect(
                            "Filter by status", status_options, default=status_options, key="dc_status_filter"
                        )
                    with filter_col2:
                        leave_filter = st.selectbox(
                            "Filter by leave", ["All", "Has Leave", "No Leave"], key="dc_leave_filter"
                        )

                    df_filtered = df_view[df_view["Status"].isin(status_filter)]
                    if leave_filter == "Has Leave":
                        df_filtered = df_filtered[df_filtered["Has Leave?"] == "Yes"]
                    elif leave_filter == "No Leave":
                        df_filtered = df_filtered[df_filtered["Has Leave?"] == "No"]

                    def _highlight_status(val):
                        colors = {
                            STATUS_COMPLETE: "background-color: #DCFCE7; color: #166534;",
                            STATUS_MISSING_OUT: "background-color: #FFEDD5; color: #9A3412;",
                            STATUS_MISSING_IN: "background-color: #DBEAFE; color: #1E40AF;",
                            STATUS_NO_PUNCH: "background-color: #FEE2E2; color: #991B1B;",
                        }
                        return colors.get(val, "")

                    def _highlight_leave(val):
                        if val == "Yes":
                            return "background-color: #DCFCE7; color: #166534; font-weight: 600;"
                        return "background-color: #FEE2E2; color: #991B1B; font-weight: 600;"

                    styled = (
                        df_filtered.style
                        .map(_highlight_status, subset=["Status"])
                        .map(_highlight_leave, subset=["Has Leave?"])
                    )

                    st.dataframe(styled, use_container_width=True, hide_index=True)

                    csv = df_filtered.to_csv(index=False).encode("utf-8-sig")
                    st.download_button(
                        "⬇️ Download this table (.csv)",
                        data=csv,
                        file_name=f"DailyCheckup_{stats['check_date'].strftime('%d%b%Y')}.csv",
                        mime="text/csv",
                        key="dc_download",
                    )
                else:
                    st.success("🎉 Everyone has a complete punch record for this date!")
    else:
        st.info("Waiting for you to upload the 3 files above.")
