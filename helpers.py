import math
from datetime import datetime

import pandas as pd
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side

def parse_time(val):
    if val is None or (isinstance(val, float) and math.isnan(val)):
        return None
    s = str(val).strip()
    if s in ("", "nan", "NaT", "None"):
        return None
    for fmt in ("%H:%M:%S", "%H:%M", "%I:%M %p", "%I:%M:%S %p"):
        try:
            return datetime.strptime(s, fmt)
        except ValueError:
            pass
    return None


def parse_bool_absent(value):
    if pd.isna(value) or str(value).strip() == "":
        return False

    if str(value).strip().lower() == "true":
        return True

    return False
def truncate1(x):
    """Truncate to 1 decimal WITHOUT rounding."""
    return math.floor(abs(x) * 10) / 10

def make_border(color="B0BEC5"):
    s = Side(style="thin", color=color)
    return Border(left=s, right=s, top=s, bottom=s)

def style_header(cell, bg="00695C", fg="FFFFFF", size=11, bold=True):
    cell.font      = Font(name="Arial", bold=bold, size=size, color=fg)
    cell.fill      = PatternFill("solid", fgColor=bg)
    cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
    cell.border    = make_border()

def style_data(cell, bg="FFFFFF", bold=False, color="000000"):
    cell.font      = Font(name="Arial", size=11, bold=bold, color=color)
    cell.fill      = PatternFill("solid", fgColor=bg)
    cell.alignment = Alignment(horizontal="center", vertical="center")
    cell.border    = make_border()

def detect_columns(df):
    """Map fingerprint file columns to logical names.
    Expected headers: Name, Date, Check in, Check out, Absent
    """
    col = {}
    for c in df.columns:
        n = str(c).strip().lower()
        if n == "name":                          col["name"]     = c
        elif n == "date":                        col["date"]     = c
        elif "clock" in n and "in" in n:         col["time_in"]  = c
        elif "clock" in n and "out" in n:        col["time_out"] = c
        elif "absent" in n:                      col["absent"]   = c
    missing = [k for k in ("name", "date", "time_in", "time_out", "absent") if k not in col]
    if missing:
        raise ValueError(
            f"Missing columns: {missing}\n"
            f"Available columns in file: {list(df.columns)}\n"
            f"Expected: Name, Date, Check in, Check out, Absent"
        )
    return col
