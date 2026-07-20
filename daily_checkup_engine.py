"""
daily_checkup_engine.py
=========================
Calculation logic for the "Daily Check-Up" dashboard.

Completely separate from timesheet_engine.py and overtime_summary_engine.py —
no shared imports, no shared state. This module answers one question for a
single chosen date: "who has an incomplete or missing punch record today,
and do they have a leave/mission on file to explain it?"

Required input files (any order, any file names — auto-detected by columns):
  1. Attendance / Punches file — needs a column named "I/O"
  2. Vacation / Leave Transaction file — needs columns "Vacation" + "From"
  3. Employee master file — needs columns "Employees Name" + "Title"
     (reads the FIRST sheet — e.g. the "Data" sheet)

Matching between files is done by employee Code (not Name).
"""

import pandas as pd


class DailyCheckupError(Exception):
    """Raised for any user-facing problem (missing columns, bad files, etc.)."""
    pass


# ── Status categories ───────────────────────────────────────────────────────
STATUS_COMPLETE = "Complete (In & Out)"
STATUS_MISSING_OUT = "Punched In - Missing Out"
STATUS_MISSING_IN = "Punched Out - Missing In"
STATUS_NO_PUNCH = "No Punch At All"

PROBLEM_STATUSES = {STATUS_MISSING_OUT, STATUS_MISSING_IN, STATUS_NO_PUNCH}


def norm_code(x):
    if pd.isna(x):
        return None
    s = str(x).strip()
    if s.endswith(".0"):
        s = s[:-2]
    return s if s else None


def _find_col(df, *candidates):
    lower_map = {c.lower().strip(): c for c in df.columns}
    for cand in candidates:
        if cand.lower() in lower_map:
            return lower_map[cand.lower()]
    return None


# ── Step 1: classify the 3 uploaded files ───────────────────────────────────
def classify_files(file_dict):
    """
    file_dict: {filename: file-like object}, exactly 3 files.
    Returns (att_file, vac_file, emp_file), each seeked to 0.
    """
    att_file = vac_file = emp_file = None
    problems = []

    for fname, fobj in file_dict.items():
        try:
            fobj.seek(0)
            df_peek = pd.read_excel(fobj, nrows=3)
            df_peek.columns = [str(c).strip() for c in df_peek.columns]
            cols_lower = [c.lower() for c in df_peek.columns]

            if "i/o" in cols_lower:
                att_file = fobj
            elif "vacation" in cols_lower and "from" in cols_lower:
                vac_file = fobj
            elif "employees name" in cols_lower and "title" in cols_lower:
                emp_file = fobj
        except Exception as e:
            problems.append(f"Could not read '{fname}': {e}")
        finally:
            fobj.seek(0)

    missing = []
    if not att_file:
        missing.append("Attendance/Punches file — needs a column named 'I/O'")
    if not vac_file:
        missing.append("Vacation/Leave Transaction file — needs columns 'Vacation' + 'From'")
    if not emp_file:
        missing.append("Employee master file — needs columns 'Employees Name' + 'Title'")

    if missing:
        msg = "Could not identify all 3 required files:\n- " + "\n- ".join(missing)
        if problems:
            msg += "\n\nAlso had trouble reading some files:\n- " + "\n- ".join(problems)
        raise DailyCheckupError(msg)

    return att_file, vac_file, emp_file


# ── Step 2: load & clean the 3 dataframes ───────────────────────────────────
def load_dataframes(att_file, vac_file, emp_file):
    df_att = pd.read_excel(att_file)
    df_vac = pd.read_excel(vac_file)
    df_emp = pd.read_excel(emp_file)  # first sheet only

    df_att.columns = [str(c).strip() for c in df_att.columns]
    df_vac.columns = [str(c).strip() for c in df_vac.columns]
    df_emp.columns = [str(c).strip() for c in df_emp.columns]

    # ---- Attendance ----
    a_code = _find_col(df_att, "Code")
    a_date = _find_col(df_att, "Date")
    a_time = _find_col(df_att, "Time")
    a_io = _find_col(df_att, "I/O")
    if not all([a_code, a_date, a_time, a_io]):
        raise DailyCheckupError("Attendance file is missing one of: Code, Date, Time, I/O columns.")

    df_att["Date"] = pd.to_datetime(df_att[a_date], errors="coerce").dt.date
    df_att["Time"] = pd.to_datetime(df_att[a_time], errors="coerce").dt.time
    df_att["CodeKey"] = df_att[a_code].map(norm_code)
    df_att["IO"] = df_att[a_io]
    df_att = df_att.dropna(subset=["Date", "CodeKey"])

    if df_att.empty:
        raise DailyCheckupError(
            "The Attendance file has no valid rows after cleaning. "
            "Check that 'Date' and 'Code' columns are filled in correctly."
        )

    # ---- Vacation / Leave (ANY status / ANY type counts as an excuse) ----
    v_code = _find_col(df_vac, "Code")
    v_from = _find_col(df_vac, "From")
    v_to = _find_col(df_vac, "To")
    v_vac = _find_col(df_vac, "Vacation")
    if not all([v_code, v_from, v_to]):
        raise DailyCheckupError("Vacation file is missing one of: Code, From, To columns.")

    df_vac["From"] = pd.to_datetime(df_vac[v_from], errors="coerce").dt.date
    df_vac["To"] = pd.to_datetime(df_vac[v_to], errors="coerce").dt.date
    df_vac["CodeKey"] = df_vac[v_code].map(norm_code)
    df_vac["VacationType"] = df_vac[v_vac].astype(str).str.strip() if v_vac else ""
    df_vac = df_vac.dropna(subset=["From", "To", "CodeKey"])

    # ---- Employee master ----
    e_code = _find_col(df_emp, "Code")
    e_name = _find_col(df_emp, "Employees Name")
    e_title = _find_col(df_emp, "Title")
    e_dept = _find_col(df_emp, "Department")
    if not all([e_code, e_name]):
        raise DailyCheckupError("Employee master file is missing 'Code' or 'Employees Name' columns.")

    df_emp["CodeKey"] = df_emp[e_code].map(norm_code)
    df_emp["Employees Name"] = df_emp[e_name].astype(str).str.strip()
    df_emp["Title_"] = df_emp[e_title].astype(str).str.strip() if e_title else ""
    df_emp["Dept_"] = df_emp[e_dept].astype(str).str.strip() if e_dept else ""

    df_emp_master = df_emp[
        df_emp["CodeKey"].notna()
        & df_emp["CodeKey"].str.match(r"^\d+$", na=False)
        & df_emp["Employees Name"].notna()
        & (df_emp["Employees Name"] != "")
        & (df_emp["Employees Name"].str.lower() != "nan")
    ].drop_duplicates(subset=["CodeKey"], keep="first").reset_index(drop=True)

    if df_emp_master.empty:
        raise DailyCheckupError(
            "The Employee master file has no valid numeric-coded rows after cleaning."
        )

    emp_by_code = df_emp_master.set_index("CodeKey").to_dict("index")

    return df_att, df_vac, emp_by_code


def get_date_bounds(df_att):
    """Returns (min_date, max_date) available in the attendance file."""
    return df_att["Date"].min(), df_att["Date"].max()


def _leave_info(df_vac, code_key, target_date):
    """Returns (has_leave: bool, leave_types: str) for this employee on this date."""
    rows = df_vac[df_vac["CodeKey"] == code_key]
    matches = []
    for _, row in rows.iterrows():
        if row["From"] <= target_date <= row["To"]:
            matches.append(row["VacationType"])
    if matches:
        return True, ", ".join(sorted(set(matches)))
    return False, ""


# ── Step 3: build the daily checkup table ───────────────────────────────────
def generate_daily_checkup(file_dict, check_date, include_complete=False):
    """
    file_dict: {filename: file-like object}, exactly 3 files.
    check_date: datetime.date to check.
    include_complete: if True, also include employees who punched in AND out
                       normally (status = Complete). Default False (only show
                       problem cases, which is the point of a daily checkup).

    Returns: (records: list[dict], stats: dict, df_att)
    """
    if len(file_dict) != 3:
        raise DailyCheckupError(f"Expected exactly 3 files, got {len(file_dict)}.")

    att_file, vac_file, emp_file = classify_files(file_dict)
    df_att, df_vac, emp_by_code = load_dataframes(att_file, vac_file, emp_file)

    day_df = df_att[df_att["Date"] == check_date]

    roster = sorted(emp_by_code.items(), key=lambda kv: kv[1]["Employees Name"])

    records = []
    counts = {STATUS_COMPLETE: 0, STATUS_MISSING_OUT: 0, STATUS_MISSING_IN: 0, STATUS_NO_PUNCH: 0}

    for code_key, info in roster:
        name = info["Employees Name"]
        title = info.get("Title_", "") or "—"
        dept = info.get("Dept_", "") or "—"
        title = title if title and str(title).lower() != "nan" else "—"
        dept = dept if dept and str(dept).lower() != "nan" else "—"

        person_day = day_df[day_df["CodeKey"] == code_key]
        ins = person_day[person_day["IO"] == "Punch In"]["Time"].dropna().tolist()
        outs = person_day[person_day["IO"] == "Punch Out"]["Time"].dropna().tolist()
        has_in = len(ins) > 0
        has_out = len(outs) > 0

        if has_in and has_out:
            status = STATUS_COMPLETE
        elif has_in and not has_out:
            status = STATUS_MISSING_OUT
        elif has_out and not has_in:
            status = STATUS_MISSING_IN
        else:
            status = STATUS_NO_PUNCH

        counts[status] += 1

        if status == STATUS_COMPLETE and not include_complete:
            continue

        has_leave, leave_types = _leave_info(df_vac, code_key, check_date)

        records.append({
            "Code": code_key,
            "Name": name,
            "Title / Position": title,
            "Department": dept,
            "Status": status,
            "Time In": min(ins).strftime("%H:%M") if ins else "",
            "Time Out": max(outs).strftime("%H:%M") if outs else "",
            "Has Leave?": "Yes" if has_leave else "No",
            "Leave Type": leave_types,
        })

    stats = {
        "check_date": check_date,
        "total_employees": len(roster),
        "complete": counts[STATUS_COMPLETE],
        "missing_out": counts[STATUS_MISSING_OUT],
        "missing_in": counts[STATUS_MISSING_IN],
        "no_punch": counts[STATUS_NO_PUNCH],
        "problem_total": counts[STATUS_MISSING_OUT] + counts[STATUS_MISSING_IN] + counts[STATUS_NO_PUNCH],
    }

    return records, stats, df_att
