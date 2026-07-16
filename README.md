# Staff Arabia HR Tools

A simple web app with two independent tools:

1. **Timesheet Generator** — per-employee daily attendance sheet (unchanged from before).
2. **OT & Attendance Summary** — one row per employee: Morning/Night overtime, cancelled
   day-offs, official holidays worked, unpaid leave, total working days.

Both tools live in the same app but are completely separated — different
tabs, different file uploads, different settings, different downloads.
Using one never resets or affects the other.

## Files

| File | Purpose |
|---|---|
| `timesheet_engine.py` | Calculation logic for **Section 1 — Timesheet Generator**. Untouched. |
| `overtime_summary_engine.py` | Calculation logic for **Section 2 — OT & Attendance Summary**. Brand new, fully independent — does not import from `timesheet_engine.py`. |
| `app.py` | The Streamlit screen: two tabs, one per tool, each with its own uploader, settings, and download button. |
| `requirements.txt` | The 3 packages needed: `streamlit`, `pandas`, `openpyxl`. |

Keeping the two engines separate means either one can be changed, replaced,
or extended later without any risk of breaking the other.

---

## 1. Run it locally in VS Code

**Requirements:** Python 3.9+ installed.

```bash
python -m venv venv
source venv/bin/activate        # on Windows: venv\Scripts\activate
pip install -r requirements.txt
streamlit run app.py
```

This opens a browser tab at `http://localhost:8501`. Pick a tab, upload the
right files for that tool, click **Generate**, then download the result.

---

## 2. Section 1 — Timesheet Generator

**3 files required** (any file names — auto-detected by columns):
- **Attendance / System file** — needs a column called **`I/O`**
- **Employees Data file** — needs columns **`Employees Name`** and **`Title`**
- **Vacation Transaction file** — needs columns **`Vacation`** and **`From`**

Settings (in the "⚙️ Settings" expander): standard workday length, official
work start time, weekend days.

---

## 3. Section 2 — OT & Attendance Summary

**4 files required** (any file names — auto-detected by columns):
- **Attendance / Punches file** — needs a column called **`I/O`**
- **Vacation Transaction file** — needs columns **`Vacation`** and **`From`**
- **Employee master file** — needs columns **`Employees Name`** and **`Title`**
  (reads the *first* sheet of the workbook — e.g. the "Data" sheet)
- **Official Holidays file** — needs an Arabic date column **`التاريخ`** and **`المناسبة`**

Output columns: `Code`, `Name`, `Title / Position`, `Department`,
`Morning OT Hours`, `Night OT Hours`, `Cancel Day Offs / Days`,
`Official Holiday / Days`, `Unpaid Leave`, `Total Working Days`.

### Calculation rules
- **Morning / Night OT Hours** — calculated only on regular working days (not
  weekends/holidays). Overtime = any time worked beyond the standard workday
  length, measured from *(time in + standard hours)* to *time out*. That
  overtime span is split at the clock: **7:00–19:00** → Morning OT,
  **19:00–7:00** → Night OT.
  _Example: in at 9:00, out at 20:00, standard = 8h → normal shift ends 17:00
  → 3h overtime (17:00–20:00) → split into 2h Morning OT (17:00–19:00) + 1h
  Night OT (19:00–20:00)._
- **Cancel Day Offs / Days** — for a weekend day or official holiday where the
  employee worked *less than* the standard workday: `hours worked ÷ standard
  hours`, added as a fractional day.
- **Official Holiday / Days** — for a weekend day or official holiday where
  the employee worked the *full* standard workday or more: counted as **1**
  full day for that date (capped, not fractional above 1).
- **Unpaid Leave** — a regular working day with **no punch at all** and
  **no vacation/mission/leave record of any kind** (any type, any status —
  approved, pending, or rejected all count as an excuse) covering that date.
- **Total Working Days** — count of days in the period the employee has a
  punch-in on record, regardless of day type.

---

## 4. Publish it for free (Streamlit Community Cloud)

1. Put this project on GitHub (public repo).
2. Upload `app.py`, `timesheet_engine.py`, `overtime_summary_engine.py`,
   `requirements.txt`.
3. Go to https://share.streamlit.io → **New app** → pick the repo, branch
   `main`, main file `app.py` → **Deploy**.

You'll get a permanent link like `https://staff-arabia-hr-tools.streamlit.app`.
It's a **public URL** by default — anyone with the link can open it and
upload files, so don't share it outside people you trust unless you add a
password gate.

### Alternatives
- **Render.com** free tier — supports private repos.
- **Hugging Face Spaces** — also free, also supports Streamlit directly.
