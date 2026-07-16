"""
section1_timesheet_ui.py
=========================
SECTION 1 — Attendance Timesheet Generator (UI layer).

This is the exact same screen that used to be the whole of app.py — moved
into its own module, untouched, so Section 2 can be added as a separate tab
without risking any change to this section's behavior.

render() draws the whole section. app.py just calls render() inside its tab.
"""

import io
import streamlit as st

from timesheet_engine import generate_timesheet, TimesheetError, DEFAULT_WEEKEND_DAYS


def render():
    st.title("🗓️ Staff Arabia Timesheet Generator")
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

    # ── Sidebar: advanced settings (defaults match the original script) ───────
    with st.sidebar:
        st.header("⚙️ Section 1 Settings")
        st.caption("Defaults are fine for normal use — only change these if you know you need to.")

        workday_hrs = st.number_input(
            "Standard workday length (hours)", min_value=1.0, max_value=24.0, value=8.0, step=0.5,
            key="s1_workday_hrs",
        )
        work_start = st.text_input(
            "Official work start time (24h, e.g. 10:00)", value="10:00", key="s1_work_start"
        )

        st.caption("Weekend days:")
        day_names = ["Monday", "Tuesday", "Wednesday", "Thursday", "Friday", "Saturday", "Sunday"]
        default_selected = [day_names[i] for i in sorted(DEFAULT_WEEKEND_DAYS)]
        selected_weekend_days = st.multiselect(
            "Days treated as weekend", day_names, default=default_selected, key="s1_weekend_days"
        )
        weekend_days = {day_names.index(d) for d in selected_weekend_days}

    # ── File upload ──────────────────────────────────────────────────────────
    uploaded_files = st.file_uploader(
        "Upload all 3 Excel files here (you can select all 3 at once)",
        type=["xlsx", "xls"],
        accept_multiple_files=True,
        key="s1_uploader",
    )

    st.divider()

    if uploaded_files:
        if len(uploaded_files) != 3:
            st.warning(f"Please upload exactly 3 files. You've uploaded {len(uploaded_files)}.")
        else:
            st.success(f"3 files ready: {', '.join(f.name for f in uploaded_files)}")

            if st.button("🚀 Generate Timesheet", type="primary", use_container_width=True, key="s1_generate"):
                file_dict = {f.name: io.BytesIO(f.getvalue()) for f in uploaded_files}

                progress_bar = st.progress(0, text="Starting...")

                def update_progress(current, total):
                    progress_bar.progress(current / total, text=f"Processing employee {current}/{total}...")

                try:
                    with st.spinner("Reading and validating files..."):
                        wb, output_filename, stats = generate_timesheet(
                            file_dict,
                            workday_hrs=workday_hrs,
                            work_start=work_start,
                            weekend_days=weekend_days,
                            progress_callback=update_progress,
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
                        key="s1_download",
                    )

                except TimesheetError as e:
                    st.error(f"⚠️ {e}")
                except Exception as e:
                    st.error(f"❌ Something unexpected went wrong: {e}")
    else:
        st.info("Waiting for you to upload the 3 files above.")
