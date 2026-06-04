def get_settings():
    print("\n=== Attendance Calculator Settings ===")

    while True:
        val = input("Work start time        (default 09:00) : ").strip()
        val = val if val else "09:00"
        if parse_time_str(val):
            WORK_START_STR = val
            break
        print("  Invalid format. Please enter time as HH:MM (e.g. 09:00)")

    while True:
        val = input("Work end time          (default 17:00) : ").strip()
        val = val if val else "17:00"
        if parse_time_str(val):
            WORK_END_STR = val
            break
        print("  Invalid format. Please enter time as HH:MM (e.g. 17:00)")

    while True:
        val = input("Grace period (minutes) (default 8)     : ").strip()
        val = val if val else "8"
        if val.isdigit() and int(val) >= 0:
            GRACE_MINUTES = int(val)
            break
        print("  Please enter a whole number (e.g. 8)")

    while True:
        val = input("Monthly late allowance (default 15)    : ").strip()
        val = val if val else "15"
        if val.isdigit() and int(val) >= 0:
            MONTHLY_ALLOWANCE = int(val)
            break
        print("  Please enter a whole number (e.g. 15)")

    print(f"\n  Work hours      : {WORK_START_STR} - {WORK_END_STR}")
    print(f"  Grace period    : {GRACE_MINUTES} min")
    print(f"  Late allowance  : {MONTHLY_ALLOWANCE} min/month")
    print("======================================\n")

    return WORK_START_STR, WORK_END_STR, GRACE_MINUTES, MONTHLY_ALLOWANCE

def parse_time_str(s):
    """Validate HH:MM format."""
    from datetime import datetime
    try:
        datetime.strptime(s, "%H:%M")
        return True
    except:
        return False