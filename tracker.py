"""Excel-backed incident tracking ledger (`GSM_Incident_Tracker.xlsx`) with concurrency retry."""

from __future__ import annotations

from collections import defaultdict
from datetime import datetime, timezone
import logging
from pathlib import Path
import time
from typing import Callable, Dict, List, Optional, TypeVar

import openpyxl
from openpyxl.styles import Alignment, Font, PatternFill
from openpyxl.utils import get_column_letter
from openpyxl.workbook import Workbook
from openpyxl.worksheet.worksheet import Worksheet

from models import TicketStatus

logger = logging.getLogger("gsm_automator.tracker")

T = TypeVar("T")

TRACKER_COLUMNS: List[str] = [
    "Incident_ID",
    "Priority",
    "Assignment_Group",
    "Assigned_To",
    "Shift",
    "Assigned_At",
    "Teams_Alerted",
    "Status",
    "Remarks",
]

COLUMN_WIDTHS: Dict[str, int] = {
    "Incident_ID": 18,
    "Priority": 12,
    "Assignment_Group": 24,
    "Assigned_To": 34,
    "Shift": 20,
    "Assigned_At": 26,
    "Teams_Alerted": 16,
    "Status": 30,
    "Remarks": 50,
}


class IncidentTracker:
    """
    Maintains `GSM_Incident_Tracker.xlsx` with deduplication (`Incident_ID` upsert)
    and retry-based file-lock protection when operators open the workbook in Excel.
    """

    def __init__(
        self,
        tracker_path: str | Path = "./data/GSM_Incident_Tracker.xlsx",
        max_retries: int = 6,
        retry_delay_seconds: float = 2.5,
        backoff_multiplier: float = 1.5,
    ) -> None:
        self.tracker_path = Path(tracker_path)
        self.max_retries = max_retries
        self.retry_delay_seconds = retry_delay_seconds
        self.backoff_multiplier = backoff_multiplier
        self.ensure_initialized()

    def _run_with_file_lock_retry(
        self, operation_name: str, fn: Callable[[], T]
    ) -> T:
        """
        Execute an openpyxl read/write callable with retry + `time.sleep` backoff
        to survive transient Windows/Excel file locks (`PermissionError` / `OSError`).
        """
        delay = self.retry_delay_seconds
        last_exc: Optional[Exception] = None

        for attempt in range(1, self.max_retries + 1):
            try:
                return fn()
            except (PermissionError, OSError) as exc:
                last_exc = exc
                if attempt < self.max_retries:
                    logger.warning(
                        "Excel lock detected during '%s' on '%s' (attempt %d/%d): %s. "
                        "Waiting %.1fs for operator to close/release workbook...",
                        operation_name,
                        self.tracker_path,
                        attempt,
                        self.max_retries,
                        exc,
                        delay,
                    )
                    time.sleep(delay)
                    delay *= self.backoff_multiplier
                else:
                    logger.error(
                        "Failed '%s' on '%s' after %d attempts due to persistent file lock: %s",
                        operation_name,
                        self.tracker_path,
                        self.max_retries,
                        exc,
                    )

        raise PermissionError(
            f"Tracker file '{self.tracker_path}' remained locked after {self.max_retries} retries."
        ) from last_exc

    def ensure_initialized(self) -> None:
        """Create `GSM_Incident_Tracker.xlsx` with styled headers if it does not yet exist."""
        self.tracker_path.parent.mkdir(parents=True, exist_ok=True)
        if self.tracker_path.exists():
            return

        def _create() -> None:
            wb = Workbook()
            ws = wb.active
            assert ws is not None
            ws.title = "Incident_Ledger"
            self._apply_header_row(ws)
            wb.save(self.tracker_path)
            wb.close()
            logger.info("Initialized new tracking ledger at '%s'.", self.tracker_path)

        self._run_with_file_lock_retry("initialize_tracker", _create)

    @staticmethod
    def _apply_header_row(ws: Worksheet) -> None:
        ws.append(TRACKER_COLUMNS)
        header_fill = PatternFill(start_color="1F4E78", end_color="1F4E78", fill_type="solid")
        header_font = Font(color="FFFFFF", bold=True, size=11)

        for col_idx, col_name in enumerate(TRACKER_COLUMNS, start=1):
            cell = ws.cell(row=1, column=col_idx)
            cell.fill = header_fill
            cell.font = header_font
            cell.alignment = Alignment(horizontal="center", vertical="center")
            col_letter = get_column_letter(col_idx)
            ws.column_dimensions[col_letter].width = COLUMN_WIDTHS.get(col_name, 20)

        ws.freeze_panes = "A2"

    @staticmethod
    def _map_columns(ws: Worksheet) -> Dict[str, int]:
        """Map header column names to 1-based column indices."""
        col_map: Dict[str, int] = {}
        for col_idx in range(1, ws.max_column + 1):
            val = ws.cell(row=1, column=col_idx).value
            if val is not None:
                col_map[str(val).strip()] = col_idx
        return col_map

    def upsert_incident(
        self,
        incident_id: str,
        priority: str,
        assignment_group: str,
        assigned_to: str,
        shift: str,
        teams_alerted: bool,
        status: str | TicketStatus,
        remarks: str = "",
        assigned_at: Optional[str] = None,
    ) -> str:
        """
        Deduplicated upsert into `GSM_Incident_Tracker.xlsx`:
        - If `incident_id` does not exist -> appends a new row (`"INSERTED"`).
        - If `incident_id` already exists -> updates current status, timestamp, and metadata (`"UPDATED"`).
        """
        timestamp_str = assigned_at or datetime.now(timezone.utc).strftime(
            "%Y-%m-%d %H:%M:%S UTC"
        )
        status_str = status.value if isinstance(status, TicketStatus) else str(status)
        teams_str = "Yes" if teams_alerted else "No"
        target_id = incident_id.strip()

        def _upsert_tx() -> str:
            if not self.tracker_path.exists():
                self.ensure_initialized()

            wb = openpyxl.load_workbook(self.tracker_path)
            try:
                ws = wb.active
                if ws is None:
                    ws = wb.create_sheet("Incident_Ledger")
                    self._apply_header_row(ws)

                col_map = self._map_columns(ws)
                if "Incident_ID" not in col_map:
                    self._apply_header_row(ws)
                    col_map = self._map_columns(ws)

                id_col = col_map["Incident_ID"]
                existing_row_idx: Optional[int] = None

                for row_idx in range(2, ws.max_row + 1):
                    cell_val = ws.cell(row=row_idx, column=id_col).value
                    if cell_val is not None and str(cell_val).strip() == target_id:
                        existing_row_idx = row_idx
                        break

                row_values = {
                    "Incident_ID": target_id,
                    "Priority": priority,
                    "Assignment_Group": assignment_group,
                    "Assigned_To": assigned_to,
                    "Shift": shift,
                    "Assigned_At": timestamp_str,
                    "Teams_Alerted": teams_str,
                    "Status": status_str,
                    "Remarks": remarks,
                }

                if existing_row_idx is not None:
                    # Preserve prior 'Yes' on Teams_Alerted if already alerted earlier
                    prev_alerted = str(
                        ws.cell(row=existing_row_idx, column=col_map["Teams_Alerted"]).value
                        or ""
                    ).strip()
                    if prev_alerted.lower() == "yes" and not teams_alerted:
                        row_values["Teams_Alerted"] = "Yes"

                    for col_name, val in row_values.items():
                        if col_name in col_map:
                            ws.cell(row=existing_row_idx, column=col_map[col_name], value=val)
                    action = "UPDATED"
                else:
                    ordered_row = [row_values[col] for col in TRACKER_COLUMNS]
                    ws.append(ordered_row)
                    action = "INSERTED"

                wb.save(self.tracker_path)
                return action
            finally:
                wb.close()

        result = self._run_with_file_lock_retry(f"upsert_{target_id}", _upsert_tx)
        logger.info(
            "Tracker ledger %s row for %s | Status='%s' | Assigned_To='%s'.",
            result,
            target_id,
            status_str,
            assigned_to or "Unassigned",
        )
        return result

    def get_ticket_status(self, incident_id: str) -> Optional[str]:
        """Return recorded Status for `incident_id`, or None if unseen."""
        if not self.tracker_path.exists():
            return None

        target_id = incident_id.strip()

        def _read_status() -> Optional[str]:
            wb = openpyxl.load_workbook(self.tracker_path, read_only=True, data_only=True)
            try:
                ws = wb.active
                if ws is None:
                    return None
                rows = ws.iter_rows(values_only=True)
                header = next(rows, None)
                if not header:
                    return None
                headers = [str(h).strip() if h is not None else "" for h in header]
                if "Incident_ID" not in headers or "Status" not in headers:
                    return None
                id_idx = headers.index("Incident_ID")
                st_idx = headers.index("Status")

                for row in rows:
                    if row and row[id_idx] is not None and str(row[id_idx]).strip() == target_id:
                        return str(row[st_idx]) if row[st_idx] is not None else None
                return None
            finally:
                wb.close()

        return self._run_with_file_lock_retry(f"read_status_{target_id}", _read_status)

    def get_historical_assignment_counts(self) -> Dict[str, int]:
        """Count assignments per engineer UPN/name from ledger to seed load-balanced routing."""
        if not self.tracker_path.exists():
            return {}

        def _aggregate() -> Dict[str, int]:
            counts: Dict[str, int] = defaultdict(int)
            wb = openpyxl.load_workbook(self.tracker_path, read_only=True, data_only=True)
            try:
                ws = wb.active
                if ws is None:
                    return {}
                rows = ws.iter_rows(values_only=True)
                header = next(rows, None)
                if not header:
                    return {}
                headers = [str(h).strip() if h is not None else "" for h in header]
                if "Assigned_To" not in headers:
                    return {}
                asgn_idx = headers.index("Assigned_To")

                for row in rows:
                    if not row or row[asgn_idx] is None:
                        continue
                    raw_val = str(row[asgn_idx]).strip()
                    if not raw_val or raw_val.lower() in {"unassigned", "n/a", "-"}:
                        continue
                    # Extract UPN inside parentheses if formatted as 'Jane Doe (jane.doe@company.com)'
                    if "(" in raw_val and raw_val.endswith(")"):
                        upn = raw_val.rsplit("(", 1)[1][:-1].strip().lower()
                        if upn:
                            counts[upn] += 1
                            continue
                    counts[raw_val.lower()] += 1
                return dict(counts)
            finally:
                wb.close()

        return self._run_with_file_lock_retry("aggregate_assignments", _aggregate)
