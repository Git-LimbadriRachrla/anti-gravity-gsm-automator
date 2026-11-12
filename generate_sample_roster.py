"""Initialize a styled sample `shift_roster.xlsx` workbook with realistic multi-timezone & overnight shifts."""

from __future__ import annotations

import argparse
from pathlib import Path
from typing import List, Sequence

from openpyxl import Workbook
from openpyxl.styles import Alignment, Font, PatternFill
from openpyxl.utils import get_column_letter
from openpyxl.worksheet.datavalidation import DataValidation

ROSTER_HEADERS: List[str] = [
    "Assignment_Group",
    "Engineer_Name",
    "Email_or_UPN",
    "Shift_Name",
    "Shift_Start",
    "Shift_End",
    "Timezone",
    "Active_Days",
    "On_Leave",
    "Is_Lead",
]

SAMPLE_ROWS: List[Sequence[object]] = [
    # Cloud-Ops-L2 (24x7 coverage across Morning, General, Night + Lead + Leave sample)
    (
        "Cloud-Ops-L2",
        "Jane Doe",
        "jane.doe@company.com",
        "Morning",
        "06:00",
        "14:30",
        "Asia/Kolkata",
        "Mon,Tue,Wed,Thu,Fri",
        "No",
        "True",
    ),
    (
        "Cloud-Ops-L2",
        "Arjun Mehta",
        "arjun.mehta@company.com",
        "General",
        "14:00",
        "22:30",
        "Asia/Kolkata",
        "Mon,Tue,Wed,Thu,Fri",
        "False",
        "False",
    ),
    (
        "Cloud-Ops-L2",
        "Priya Nair",
        "priya.nair@company.com",
        "General",
        "14:00",
        "22:30",
        "Asia/Kolkata",
        "Mon,Tue,Wed,Thu,Fri",
        "No",
        "False",
    ),
    (
        "Cloud-Ops-L2",
        "Rohan Kulkarni",
        "rohan.kulkarni@company.com",
        "Night",
        "22:00",
        "06:00",
        "Asia/Kolkata",
        "Mon,Tue,Wed,Thu,Fri,Sat,Sun",
        "False",
        "False",
    ),
    (
        "Cloud-Ops-L2",
        "Emily Carter",
        "emily.carter@company.com",
        "US-Day",
        "09:00",
        "17:30",
        "America/New_York",
        "Mon,Tue,Wed,Thu,Fri",
        "Yes",  # Currently on leave -> skipped automatically
        "False",
    ),
    # DB-Reliability-L3 (includes overnight US/IST coverage + Lead fallback)
    (
        "DB-Reliability-L3",
        "Carlos Rivera",
        "carlos.rivera@company.com",
        "Morning",
        "08:00",
        "16:30",
        "America/New_York",
        "Mon,Tue,Wed,Thu,Fri",
        "No",
        "True",
    ),
    (
        "DB-Reliability-L3",
        "Ananya Sharma",
        "ananya.sharma@company.com",
        "Night",
        "22:00",
        "06:30",
        "Asia/Kolkata",
        "Sun,Mon,Tue,Wed,Thu",
        "No",
        "False",
    ),
    # NetSec-Ops (EMEA + APAC shifts)
    (
        "NetSec-Ops",
        "Oliver Wright",
        "oliver.wright@company.com",
        "General",
        "08:30",
        "17:00",
        "Europe/London",
        "Mon,Tue,Wed,Thu,Fri",
        "No",
        "True",
    ),
    (
        "NetSec-Ops",
        "Vikram Singh",
        "vikram.singh@company.com",
        "Night",
        "21:30",
        "06:00",
        "Europe/London",
        "Mon,Tue,Wed,Thu,Fri,Sat",
        "No",
        "False",
    ),
]


def generate_roster_workbook(output_path: str | Path = "./data/shift_roster.xlsx") -> Path:
    """Create `shift_roster.xlsx` populated with realistic dummy shift data and validations."""
    target = Path(output_path)
    target.parent.mkdir(parents=True, exist_ok=True)

    wb = Workbook()
    ws = wb.active
    assert ws is not None
    ws.title = "Shift_Roster"

    # Header styling
    ws.append(ROSTER_HEADERS)
    header_fill = PatternFill(start_color="0F3D5E", end_color="0F3D5E", fill_type="solid")
    header_font = Font(color="FFFFFF", bold=True, size=11)

    widths = [22, 22, 30, 14, 14, 14, 22, 28, 12, 12]
    for col_idx, width in enumerate(widths, start=1):
        cell = ws.cell(row=1, column=col_idx)
        cell.fill = header_fill
        cell.font = header_font
        cell.alignment = Alignment(horizontal="center", vertical="center")
        ws.column_dimensions[get_column_letter(col_idx)].width = width

    # Populate sample rows
    for row_data in SAMPLE_ROWS:
        ws.append(list(row_data))

    # Add Excel dropdown validation for On_Leave (Col I) and Is_Lead (Col J)
    bool_dv = DataValidation(
        type="list", formula1='"True,False,Yes,No"', allow_blank=False
    )
    ws.add_data_validation(bool_dv)
    bool_dv.add("I2:J200")

    ws.freeze_panes = "A2"
    wb.save(target)
    wb.close()
    return target


def main() -> None:
    parser = argparse.ArgumentParser(
        description="Generate a sample shift_roster.xlsx workbook for Anti-Gravity GSM Automator."
    )
    parser.add_argument(
        "--output",
        default="./data/shift_roster.xlsx",
        help="Destination path for shift_roster.xlsx (default: ./data/shift_roster.xlsx)",
    )
    args = parser.parse_args()
    created_path = generate_roster_workbook(args.output)
    print(f"Sample shift roster successfully generated at: {created_path.resolve()}")


if __name__ == "__main__":
    main()
