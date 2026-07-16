"""
section2_ot_payroll_ui.py
============================
SECTION 2 — OT & Payroll Summary Sheet Generator (UI layer).

Fully independent from section1_timesheet_ui.py. Uses only
ot_payroll_engine.py for calculations. Nothing here touches Section 1's
files, session-state keys, or logic.
"""

import io
import datetime
import streamlit as st
import pandas as pd

from ot_payroll_engine import (
    generate_ot_payroll_report,
    export_report_to_excel,
    parse_holiday_file,
    OTPayrollError,
    DEFAULT_WEEKEND_DAYS,
    DEFAULT_WORKDAY_HRS,
)


def render():
    st.title("📊 OT & Payroll Summary Sheet Generator")
    st.write(
        "Upload an attendance punch-in/out file to get a per-employee summary of "
        "overtime, weekend/holiday work, unpaid leave, and total working days."
    )

    with st.expander("What are the columns in the output?", expanded=False):
        st.markdown(
            """
- **Morning OT Hours** — overtime worked between 07:00–19:00
- **Night OT Hours** — overtime worked between 19:00–07:00
- **Cancel Day Offs / Days** — weekend/holiday work under 8h, as a fraction of a day
- **Official Holiday / Days** — weekend/holiday work of 8h or more, 1 day each
- **Unpaid Leave** — regular workdays with no punch at all and no leave record on file
- **Total Working Days** — distinct days the employee punched in
            """
        )

    # ── Sidebar: Section 2 settings (independent keys from Section 1) ─────────
    with st.sidebar:
        st.header("⚙️ Section 2 Settings")

        workday_hrs = st.number_input(
            "Official work hours per day",
            min_value=1.0, max_value=24.0, value=DEFAULT_WORKDAY_HRS, step=0.5,
            key="s2_workday_hrs",
        )

        st.caption("Weekend days:")
        day_names = ["Monday", "Tuesday", "Wednesday", "Thursday", "Friday", "Saturday", "Sunday"]
        default_selected = [day_names[i] for i in sorted(DEFAULT_WEEKEND_DAYS)]
        selected_weekend_days = st.multiselect(
            "Days treated as weekend", day_names, default=default_selected, key="s2_weekend_days"
        )
        weekend_days = {day_names.index(d) for d in selected_weekend_days}

        st.divider()
        st.subheader("📅 Official Holidays")
        st.caption("Add dates manually, and/or upload a holidays file — both are combined.")

        if "s2_manual_holidays" not in st.session_state:
            st.session_state.s2_manual_holidays = []

        new_holiday = st.date_input(
            "Add a holiday date", value=None, key="s2_new_holiday_input", format="YYYY-MM-DD"
        )
        add_col, clear_col = st.columns(2)
        with add_col:
            if st.button("➕ Add date", use_container_width=True, key="s2_add_holiday"):
                if new_holiday and new_holiday not in st.session_state.s2_manual_holidays:
                    st.session_state.s2_manual_holidays.append(new_holiday)
        with clear_col:
            if st.button("🗑️ Clear all", use_container_width=True, key="s2_clear_holidays"):
                st.session_state.s2_manual_holidays = []

        if st.session_state.s2_manual_holidays:
            st.caption("Manually added holidays:")
            for d in sorted(st.session_state.s2_manual_holidays):
                col_a, col_b = st.columns([4, 1])
                col_a.write(f"• {d.strftime('%Y-%m-%d (%A)')}")
                if col_b.button("✕", key=f"s2_remove_{d}"):
                    st.session_state.s2_manual_holidays.remove(d)
                    st.rerun()

        holidays_upload = st.file_uploader(
            "Or upload a holidays list file (a 'Date' column)",
            type=["xlsx", "xls", "csv"],
            key="s2_holidays_file",
        )

    st.divider()

    # ── File uploads ────────────────────────────────────────────────────────
    st.subheader("1️⃣ Attendance punch file (required)")
    attendance_upload = st.file_uploader(
        "Upload the attendance punch-in/out Excel file",
        type=["xlsx", "xls"],
        key="s2_attendance_upload",
    )

    col1, col2 = st.columns(2)
    with col1:
        st.subheader("2️⃣ Employee data (optional)")
        st.caption("Adds Title / Department to the report. Without it, those show as \"—\".")
        employee_upload = st.file_uploader(
            "Upload employee data file",
            type=["xlsx", "xls"],
            key="s2_employee_upload",
        )
    with col2:
        st.subheader("3️⃣ Vacation records (optional)")
        st.caption("Used to tell real Unpaid Leave apart from any approved/pending absence.")
        vacation_upload = st.file_uploader(
            "Upload vacation transaction file",
            type=["xlsx", "xls"],
            key="s2_vacation_upload",
        )

    st.divider()

    if not attendance_upload:
        st.info("Waiting for the attendance file to generate the summary.")
        return

    if not employee_upload:
        st.caption("ℹ️ No employee file uploaded — Title and Department will show as \"—\".")
    if not vacation_upload:
        st.caption(
            "ℹ️ No vacation file uploaded — any day with no punch will count as Unpaid Leave."
        )

    if st.button("🚀 Generate OT & Payroll Summary", type="primary", use_container_width=True, key="s2_generate"):
        try:
            holiday_dates = set(st.session_state.s2_manual_holidays)
            if holidays_upload is not None:
                holidays_upload.seek(0)
                holiday_dates |= parse_holiday_file(holidays_upload)

            progress_bar = st.progress(0, text="Starting...")

            def update_progress(current, total):
                progress_bar.progress(current / total, text=f"Processing employee {current}/{total}...")

            attendance_upload.seek(0)
            emp_file = io.BytesIO(employee_upload.getvalue()) if employee_upload else None
            vac_file = io.BytesIO(vacation_upload.getvalue()) if vacation_upload else None

            with st.spinner("Reading and crunching the numbers..."):
                report_df, stats = generate_ot_payroll_report(
                    attendance_file=io.BytesIO(attendance_upload.getvalue()),
                    employee_file=emp_file,
                    vacation_file=vac_file,
                    holiday_dates=holiday_dates,
                    weekend_days=weekend_days,
                    workday_hrs=workday_hrs,
                    progress_callback=update_progress,
                )

            progress_bar.progress(1.0, text="Done!")

            st.success(
                f"✅ Summary generated for **{stats['num_employees']} employees** "
                f"over **{stats['num_days']} days** "
                f"({stats['date_min'].strftime('%d %b %Y')} → {stats['date_max'].strftime('%d %b %Y')})"
            )
            if holiday_dates:
                st.caption(f"📅 {len(holiday_dates)} holiday date(s) applied.")

            st.dataframe(report_df, use_container_width=True, hide_index=True)

            wb = export_report_to_excel(report_df)
            buffer = io.BytesIO()
            wb.save(buffer)
            buffer.seek(0)

            output_name = f"OT_Payroll_Summary_{stats['date_min'].strftime('%d%b')}_{stats['date_max'].strftime('%d%b_%Y')}.xlsx"
            st.download_button(
                label="⬇️ Export to Excel",
                data=buffer,
                file_name=output_name,
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                use_container_width=True,
                key="s2_download",
            )

        except OTPayrollError as e:
            st.error(f"⚠️ {e}")
        except Exception as e:
            st.error(f"❌ Something unexpected went wrong: {e}")
