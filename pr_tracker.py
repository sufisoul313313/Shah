"""
Canada PR Application Status Tracker
Reads PR_Tracker.xlsx and generates a formatted status report.
"""

import sys
import os
from datetime import datetime, timedelta
from pathlib import Path

try:
    import pandas as pd
    from openpyxl import load_workbook
except ImportError:
    print("Missing dependencies. Run: pip install pandas openpyxl")
    sys.exit(1)

EXCEL_PATH = Path.home() / "Downloads" / "PR_Tracker.xlsx"
REPORTS_DIR = Path(__file__).parent / "reports"
TODAY = datetime.today().date()
EXPIRY_ALERT_DAYS = 30

REQUIRED_COLUMNS = {
    "document name", "status", "date submitted", "expiry date", "notes"
}


def load_tracker(path: Path) -> pd.DataFrame:
    if not path.exists():
        print(f"ERROR: Excel file not found at: {path}")
        print("Creating a sample file there so you can fill it in.")
        create_sample_excel(path)
        sys.exit(0)

    df = pd.read_excel(path, engine="openpyxl")
    df.columns = df.columns.str.strip().str.lower()

    missing = REQUIRED_COLUMNS - set(df.columns)
    if missing:
        print(f"ERROR: Missing columns in Excel: {missing}")
        print(f"Found columns: {list(df.columns)}")
        sys.exit(1)

    # Normalize dates
    for col in ("date submitted", "expiry date"):
        df[col] = pd.to_datetime(df[col], errors="coerce").dt.date

    df["status"] = df["status"].str.strip().str.title()
    df["document name"] = df["document name"].str.strip()
    df["notes"] = df["notes"].fillna("").str.strip()

    return df


def create_sample_excel(path: Path):
    path.parent.mkdir(parents=True, exist_ok=True)
    import openpyxl
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "PR Tracker"

    headers = ["Document Name", "Status", "Date Submitted", "Expiry Date", "Notes"]
    ws.append(headers)

    sample_rows = [
        ["Medical Exam",          "Submitted", "2025-11-01", "2026-11-01", "Done at LifeLabs"],
        ["Police Certificate",    "Approved",  "2025-10-15", "2026-10-15", ""],
        ["Birth Certificate",     "Pending",   "",           "",           "Need certified copy"],
        ["Passport Copy",         "Approved",  "2025-09-01", "",           ""],
        ["Proof of Funds",        "Submitted", "2025-12-01", "",           "Bank statement"],
        ["Language Test (IELTS)", "Approved",  "2024-06-01", "2026-06-01", "Score: 8.0"],
        ["ECA Report",            "Expired",   "2023-01-10", "2025-01-10", "Needs renewal"],
        ["Photographs",           "Pending",   "",           "",           ""],
    ]
    for row in sample_rows:
        ws.append(row)

    # Auto-width columns
    for col in ws.columns:
        max_len = max(len(str(cell.value or "")) for cell in col) + 4
        ws.column_dimensions[col[0].column_letter].width = max_len

    wb.save(path)
    print(f"Sample file created: {path}")
    print("Fill it in with your real data and run this script again.")


def build_report(df: pd.DataFrame) -> str:
    total = len(df)
    approved = df[df["status"] == "Approved"]
    pending  = df[df["status"] == "Pending"]
    submitted = df[df["status"] == "Submitted"]
    expired  = df[df["status"] == "Expired"]

    pct_approved = round(len(approved) / total * 100) if total else 0

    # Expiring within 30 days (includes already-expired docs with a future expiry? no — upcoming only)
    def _expiring_soon(d):
        if d is None or d != d:  # catches None and NaT/NaN
            return False
        try:
            return TODAY <= d <= TODAY + timedelta(days=EXPIRY_ALERT_DAYS)
        except TypeError:
            return False

    expiring_soon = df[df["expiry date"].apply(_expiring_soon)]

    lines = []

    def separator(char="=", width=60):
        lines.append(char * width)

    def section(title):
        lines.append("")
        separator("-")
        lines.append(f"  {title}")
        separator("-")

    # Header
    separator()
    lines.append("       CANADA PR APPLICATION STATUS REPORT")
    lines.append(f"       Generated: {TODAY.strftime('%B %d, %Y')}")
    separator()

    # Summary
    section("OVERALL PROGRESS")
    bar_filled = int(pct_approved / 5)   # scale to 20-char bar
    bar = "█" * bar_filled + "░" * (20 - bar_filled)
    lines.append(f"  [{bar}] {pct_approved}%  ({len(approved)}/{total} documents approved)")
    lines.append("")
    lines.append(f"  Approved  : {len(approved)}")
    lines.append(f"  Submitted : {len(submitted)}")
    lines.append(f"  Pending   : {len(pending)}")
    lines.append(f"  Expired   : {len(expired)}")

    # Expiry alerts
    section("⚠  EXPIRING WITHIN 30 DAYS")
    if expiring_soon.empty:
        lines.append("  No documents expiring soon.")
    else:
        for _, row in expiring_soon.iterrows():
            days_left = (row["expiry date"] - TODAY).days
            lines.append(
                f"  • {row['document name']:<30}  expires {row['expiry date']}  "
                f"({days_left} day{'s' if days_left != 1 else ''} left)"
            )

    # Expired documents
    if not expired.empty:
        section("✗  EXPIRED DOCUMENTS")
        for _, row in expired.iterrows():
            exp = row["expiry date"] if pd.notna(row["expiry date"]) else "N/A"
            lines.append(f"  • {row['document name']:<30}  expired {exp}")
            if row["notes"]:
                lines.append(f"    Note: {row['notes']}")

    # Pending documents
    section("○  PENDING DOCUMENTS  (action required)")
    if pending.empty:
        lines.append("  No documents pending.")
    else:
        for _, row in pending.iterrows():
            lines.append(f"  • {row['document name']}")
            if row["notes"]:
                lines.append(f"    Note: {row['notes']}")

    # Submitted (in progress)
    section("→  SUBMITTED / IN PROGRESS")
    if submitted.empty:
        lines.append("  None.")
    else:
        for _, row in submitted.iterrows():
            sub_date = row["date submitted"] if pd.notna(row["date submitted"]) else "N/A"
            exp_date = f"  |  expires {row['expiry date']}" if pd.notna(row["expiry date"]) else ""
            lines.append(f"  • {row['document name']:<30}  submitted {sub_date}{exp_date}")

    # All documents table
    section("FULL DOCUMENT LIST")
    col_w = [30, 12, 15, 15]
    header_row = (
        f"  {'Document':<{col_w[0]}}  {'Status':<{col_w[1]}}"
        f"  {'Submitted':<{col_w[2]}}  {'Expiry':<{col_w[3]}}"
    )
    lines.append(header_row)
    lines.append("  " + "-" * (sum(col_w) + 6))
    for _, row in df.iterrows():
        sub  = str(row["date submitted"]) if pd.notna(row["date submitted"]) else "—"
        exp  = str(row["expiry date"])    if pd.notna(row["expiry date"])    else "—"
        status_marker = {
            "Approved": "✓", "Pending": "○", "Submitted": "→", "Expired": "✗"
        }.get(row["status"], " ")
        lines.append(
            f"  {row['document name']:<{col_w[0]}}  "
            f"{status_marker} {row['status']:<{col_w[1]-2}}  "
            f"{sub:<{col_w[2]}}  {exp:<{col_w[3]}}"
        )

    # Next steps
    section("NEXT STEPS")
    steps = []
    if not pending.empty:
        for _, row in pending.iterrows():
            steps.append(f"  [ ] Obtain and submit: {row['document name']}")
    if not expiring_soon.empty:
        for _, row in expiring_soon.iterrows():
            steps.append(f"  [ ] Renew before expiry: {row['document name']}")
    if not expired.empty:
        for _, row in expired.iterrows():
            steps.append(f"  [ ] Renew expired doc: {row['document name']}")
    if not steps:
        steps = ["  All documents are in order. Monitor for upcoming expiries."]
    lines.extend(steps)

    lines.append("")
    separator()
    lines.append("  Report saved. Good luck with your PR application!")
    separator()

    return "\n".join(lines)


def save_report(report_text: str) -> Path:
    REPORTS_DIR.mkdir(exist_ok=True)
    filename = REPORTS_DIR / f"PR_Status_Report_{TODAY.strftime('%Y-%m-%d')}.txt"
    filename.write_text(report_text, encoding="utf-8")
    return filename


def main():
    # Allow overriding the Excel path via command-line argument
    excel_path = Path(sys.argv[1]) if len(sys.argv) > 1 else EXCEL_PATH

    print(f"Loading: {excel_path}")
    df = load_tracker(excel_path)
    print(f"Loaded {len(df)} documents.")

    report = build_report(df)
    print(report)

    out_path = save_report(report)
    print(f"\nReport saved to: {out_path}")


if __name__ == "__main__":
    main()
