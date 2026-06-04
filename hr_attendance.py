"""
HR Attendance Calculator
========================
Phase 1: python hr_attendance.py phase1 fingerprint_file.xlsx
    -> produces: review_YYYY-MM.xlsx  (fill in hourly leave columns manually)

Phase 2: python hr_attendance.py phase2 review_YYYY-MM.xlsx
    -> produces: final_report_YYYY-MM.xlsx
"""

import sys
import pandas as pd
import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment
from openpyxl.utils import get_column_letter
from datetime import datetime

from helpers import detect_columns, parse_time, parse_bool_absent, style_header, style_data, truncate1, make_border


# ══════════════════════════════════════════════════════════════
#  PHASE 1 - produce review file
# ══════════════════════════════════════════════════════════════

def phase1(input_file, WORK_START_STR, WORK_END_STR, GRACE_MINUTES, MONTHLY_ALLOWANCE):
    df  = pd.read_excel(input_file)
    df.columns = df.columns.str.strip()
    col = detect_columns(df)

    ws_start = parse_time(WORK_START_STR)
    ws_end   = parse_time(WORK_END_STR)

    rows = []
    for _, r in df.iterrows():
        name     = str(r[col["name"]]).strip()
        raw_date = r[col["date"]]
        absent   = parse_bool_absent(r[col["absent"]])
        t_in     = parse_time(r[col["time_in"]])
        t_out    = parse_time(r[col["time_out"]])

        try:
            dt       = pd.to_datetime(raw_date)
            date_str = dt.strftime("%Y-%m-%d")
        except:
            date_str = str(raw_date)

        # Late minutes (after grace)
        late_min = 0.0
        if not absent and t_in and t_in > ws_start:
            diff = (t_in - ws_start).total_seconds() / 60
            late_min = diff if diff > GRACE_MINUTES else 0.0

        # Overtime minutes (after grace)
        extra_min = 0.0
        if not absent and t_out and t_out > ws_end:
            diff = (t_out - ws_end).total_seconds() / 60
            extra_min = diff if diff > GRACE_MINUTES else 0.0

        # Early departure (possible PM leave)
        early_out_min = 0.0
        if not absent and t_out and t_out < ws_end:
            early_out_min = (ws_end - t_out).total_seconds() / 60

        rows.append({
            "Name":                  name,
            "Date":                  date_str,
            "Check In":              str(r[col["time_in"]]).strip() if not absent else "",
            "Check Out":             str(r[col["time_out"]]).strip() if not absent else "",
            "Full Day Absent":       "True" if absent else "",
            "Late (min)":            round(late_min, 1),
            "Early Departure (min)": round(early_out_min, 1),
            "Overtime (min)":        round(extra_min, 1),
            # fill manually
            "Leave Type":            "",   # AM  /  PM  /  leave empty
            "Hourly Leave (hrs)":    "",   # e.g. 2.5
        })

    out_df = pd.DataFrame(rows)

    try:
        month_label = pd.to_datetime(out_df["Date"].iloc[0]).strftime("%Y-%m")
    except:
        month_label = "review"

    out_path = f"review_{month_label}.xlsx"

    TEAL  = "00695C"
    TEAL2 = "00897B"
    AMBER = "F57F17"
    WHITE = "FFFFFF"
    LGRAY = "F5F5F5"

    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "Review Data"

    ws.merge_cells("A1:J1")
    ws["A1"] = "Review File — Fill in the two orange columns for hourly leaves"
    ws["A1"].font      = Font(name="Arial", bold=True, size=13, color=WHITE)
    ws["A1"].fill      = PatternFill("solid", fgColor=TEAL)
    ws["A1"].alignment = Alignment(horizontal="center", vertical="center")
    ws.row_dimensions[1].height = 36

    ws.merge_cells("A2:J2")
    ws["A2"] = (
        "Leave Type: enter AM (morning leave) or PM (afternoon leave)   |   "
        "Hourly Leave (hrs): enter duration e.g. 2.5   |   Leave both empty if no hourly leave"
    )
    ws["A2"].font      = Font(name="Arial", size=10, color=WHITE)
    ws["A2"].fill      = PatternFill("solid", fgColor=AMBER)
    ws["A2"].alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
    ws.row_dimensions[2].height = 28

    headers    = list(out_df.columns)
    col_widths = [22, 14, 14, 14, 16, 14, 22, 16, 14, 20]

    for ci, (h, w) in enumerate(zip(headers, col_widths), start=1):
        cell = ws.cell(row=3, column=ci, value=h)
        bg   = AMBER if h in ("Leave Type", "Hourly Leave (hrs)") else TEAL2
        style_header(cell, bg=bg)
        ws.column_dimensions[get_column_letter(ci)].width = w
    ws.row_dimensions[3].height = 38

    for ri, (_, row) in enumerate(out_df.iterrows(), start=4):
        bg = LGRAY if ri % 2 == 0 else WHITE
        for ci, val in enumerate(row, start=1):
            cell     = ws.cell(row=ri, column=ci, value=val)
            col_name = headers[ci - 1]
            if col_name in ("Leave Type", "Hourly Leave (hrs)"):
                style_data(cell, bg="FFF8E1")
            else:
                style_data(cell, bg=bg)
        ws.row_dimensions[ri].height = 22

    # Save settings into a hidden sheet so phase2 can read them
    ws_cfg = wb.create_sheet("Settings")
    ws_cfg.sheet_state = "hidden"
    for row in [
        ("WORK_START_STR",    WORK_START_STR),
        ("WORK_END_STR",      WORK_END_STR),
        ("GRACE_MINUTES",     GRACE_MINUTES),
        ("MONTHLY_ALLOWANCE", MONTHLY_ALLOWANCE),
    ]:
        ws_cfg.append(row)

    wb.save(out_path)
    print(f"\n✅ Review file created: {out_path}")
    print(f"   Fill in 'Leave Type' (AM/PM) and 'Hourly Leave (hrs)', then run:")
    print(f"   python hr_attendance.py phase2 {out_path}\n")

    return out_path

# ══════════════════════════════════════════════════════════════
#  PHASE 2 - final report
# ══════════════════════════════════════════════════════════════

def phase2(review_file):
    # Load settings saved by phase1
    try:
        cfg = pd.read_excel(review_file, sheet_name="Settings", index_col=0, header=None)
        cfg.index = cfg.index.str.strip()
        WORK_START_STR    = str(cfg.loc["WORK_START_STR",    1]).strip()
        WORK_END_STR      = str(cfg.loc["WORK_END_STR",      1]).strip()
        GRACE_MINUTES     = int(cfg.loc["GRACE_MINUTES",     1])
        MONTHLY_ALLOWANCE = int(cfg.loc["MONTHLY_ALLOWANCE", 1])
        print(f"\n  Settings loaded from review file:")
        print(f"  Work hours     : {WORK_START_STR} - {WORK_END_STR}")
        print(f"  Grace period   : {GRACE_MINUTES} min")
        print(f"  Late allowance : {MONTHLY_ALLOWANCE} min/month\n")
    except Exception as ex:
        print(f"\n❌ Could not read settings from review file: {ex}")
        print("   Make sure you are using the review file produced by phase1.\n")
        sys.exit(1)

    df = pd.read_excel(review_file, sheet_name="Review Data", header=2)
    df.columns = df.columns.str.strip()

    employees = {}

    for _, r in df.iterrows():
        name = str(r.get("Name", "")).strip()
        if not name or name == "nan":
            continue

        if name not in employees:
            employees[name] = {
                "late_min_total":  0.0,
                "extra_min_total": 0.0,
                "absent_days":     0,
                "leave_hours":     0.0,
                "late_count":      0,
            }

        e = employees[name]

        # Full-day absence
        if str(r.get("Full Day Absent", "No")).strip().lower() in ("yes", "true", "1"):
            e["absent_days"] += 1
            continue

        # Hourly leave
        try:
            lh_val = r.get("Hourly Leave (hrs)", "")
            lh = float(lh_val) if str(lh_val).strip() not in ("", "nan") else 0.0
        except:
            lh = 0.0
        e["leave_hours"] += lh

        leave_type = str(r.get("Leave Type", "")).strip().upper()  # "AM", "PM", or ""

        # Late minutes
        # AM leave = morning leave = ignore Check In entirely for that day
        if leave_type != "AM":
            try:
                late = float(r.get("Late (min)", 0) or 0)
            except:
                late = 0.0
            e["late_min_total"] += late
            if late > 0:
                e["late_count"] += 1

        # Overtime minutes
        # PM leave = afternoon leave = ignore Check Out entirely for that day
        if leave_type != "PM":
            try:
                extra = float(r.get("Overtime (min)", 0) or 0)
            except:
                extra = 0.0
            e["extra_min_total"] += extra

    # Per-employee final calculation
    summary_rows = []
    for name, e in employees.items():
        late_after_allowance = max(0.0, e["late_min_total"] - MONTHLY_ALLOWANCE)
        net_min              = e["extra_min_total"] - late_after_allowance
        net_hours            = truncate1((net_min * 2) / 60)
        net_type             = "OT" if net_min >= 0 else "LATE"
        net_label            = f"{net_hours}h {net_type}"

        summary_rows.append({
            "Employee Name":          name,
            "Total Late (min)":       round(e["late_min_total"], 1),
            "Late Occurrences":       e["late_count"],
            "Total Overtime (min)":   round(e["extra_min_total"], 1),
            "Net (hrs) — OT or Late": net_label,
            "Leave Days":             e["absent_days"],
            "Hourly Leave (hrs)":     round(e["leave_hours"], 1),
            "_net_type":              net_type,
        })

    try:
        first_row   = pd.read_excel(review_file, sheet_name="Review Data", header=2).iloc[0]
        month_label = str(first_row.get("Date", ""))[:7]
    except:
        month_label = "report"

    out_path = f"final_report_{month_label}.xlsx"

    TEAL       = "00695C"
    TEAL2      = "00897B"
    WHITE      = "FFFFFF"
    LGRAY      = "F5F5F5"
    GREEN      = "E8F5E9"
    RED        = "FFEBEE"
    GREEN_TEXT = "1B5E20"
    RED_TEXT   = "B71C1C"

    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "Monthly Report"

    visible_keys = [k for k in summary_rows[0].keys() if not k.startswith("_")] if summary_rows else []
    n_cols       = len(visible_keys)

    ws.merge_cells(f"A1:{get_column_letter(n_cols)}1")
    ws["A1"] = "Monthly Attendance Report"
    ws["A1"].font      = Font(name="Arial", bold=True, size=15, color=WHITE)
    ws["A1"].fill      = PatternFill("solid", fgColor=TEAL)
    ws["A1"].alignment = Alignment(horizontal="center", vertical="center")
    ws.row_dimensions[1].height = 42

    ws.merge_cells(f"A2:{get_column_letter(n_cols)}2")
    ws["A2"] = (
        f"Work hours: {WORK_START_STR} - {WORK_END_STR}   |   "
        f"Grace period: {GRACE_MINUTES} min   |   "
        f"Monthly late allowance: {MONTHLY_ALLOWANCE} min   |   "
        f"Generated: {datetime.now().strftime('%Y-%m-%d %H:%M')}"
    )
    ws["A2"].font      = Font(name="Arial", size=10, color=WHITE)
    ws["A2"].fill      = PatternFill("solid", fgColor=TEAL2)
    ws["A2"].alignment = Alignment(horizontal="center", vertical="center")
    ws.row_dimensions[2].height = 22

    col_widths = [24, 20, 18, 22, 26, 14, 22]
    for ci, (h, w) in enumerate(zip(visible_keys, col_widths), start=1):
        cell = ws.cell(row=3, column=ci, value=h)
        style_header(cell, bg=TEAL2)
        ws.column_dimensions[get_column_letter(ci)].width = w
    ws.row_dimensions[3].height = 40

    for ri, row_data in enumerate(summary_rows, start=4):
        bg       = LGRAY if ri % 2 == 0 else WHITE
        net_type = row_data["_net_type"]
        for ci, key in enumerate(visible_keys, start=1):
            cell = ws.cell(row=ri, column=ci, value=row_data[key])
            if key == "Net (hrs) — OT or Late":
                if net_type == "OT":
                    style_data(cell, bg=GREEN, bold=True, color=GREEN_TEXT)
                else:
                    style_data(cell, bg=RED, bold=True, color=RED_TEXT)
            else:
                style_data(cell, bg=bg)
        ws.row_dimensions[ri].height = 26

    # Totals row
    total_row     = len(summary_rows) + 4
    summable_keys = {"Total Late (min)", "Late Occurrences", "Total Overtime (min)", "Leave Days", "Hourly Leave (hrs)"}

    for ci, key in enumerate(visible_keys, start=1):
        cell = ws.cell(row=total_row, column=ci)
        if key == "Employee Name":
            cell.value = "TOTAL"
        elif key in summable_keys:
            cell.value = round(sum(r[key] for r in summary_rows), 1)
        else:
            cell.value = ""
        style_header(cell, bg=TEAL)
    ws.row_dimensions[total_row].height = 28

    # Sheet 2: Calculation Logic
    ws2 = wb.create_sheet("Calculation Logic")
    ws2.column_dimensions["A"].width = 35
    ws2.column_dimensions["B"].width = 60

    explanation = [
        ("Day off",                "Friday (skipped automatically)"),
        ("Work start",             WORK_START_STR),
        ("Work end",               WORK_END_STR),
        ("Grace period",           f"{GRACE_MINUTES} min — under 8 min late/overtime is ignored"),
        ("Monthly late allowance", f"{MONTHLY_ALLOWANCE} min deducted from total late minutes"),
        ("Daily late calc",        "If Check In > 09:08 → minutes added to total"),
        ("Daily overtime calc",    "If Check Out > 17:08 → minutes added to total"),
        ("AM leave (morning)",     "Enter AM in Leave Type → Check In ignored, no late counted that day"),
        ("PM leave (afternoon)",   "Enter PM in Leave Type → Check Out ignored, no overtime counted that day"),
        ("Late after allowance",   "Total Late - 15 min  (minimum 0)"),
        ("Net minutes",            "Total Overtime - Late after allowance"),
        ("Net hours formula",      "( |Net min| x 2 ) / 60   — 1 decimal, no rounding"),
        ("Hourly leave",           "Entered manually in review file, shown separately in report"),
        ("Full-day absence",       "Absent = True/Yes rows, counted separately"),
    ]

    ws2.cell(row=1, column=1, value="Calculation Logic").font = Font(
        name="Arial", bold=True, size=13, color="00695C"
    )
    ws2.row_dimensions[1].height = 30

    for ri, (label, value) in enumerate(explanation, start=2):
        c1 = ws2.cell(row=ri, column=1, value=label)
        c2 = ws2.cell(row=ri, column=2, value=value)
        bg = "E0F2F1" if ri % 2 == 0 else WHITE
        for c in (c1, c2):
            c.font      = Font(name="Arial", size=11)
            c.fill      = PatternFill("solid", fgColor=bg)
            c.alignment = Alignment(horizontal="left", vertical="center")
            c.border    = make_border()
        ws2.row_dimensions[ri].height = 24

    wb.save(out_path)
    print(f"\n✅ Final report ready: {out_path}\n")

    return out_path
