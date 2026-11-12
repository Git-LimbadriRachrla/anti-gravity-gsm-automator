"""Dynamic static-Excel shift roster evaluation engine with overnight & load-balancing support."""

from __future__ import annotations

from collections import defaultdict
from datetime import datetime, time, timedelta, timezone
import io
import logging
from pathlib import Path
import time as time_mod
from typing import Dict, Iterable, List, Optional, Set, Tuple
from zoneinfo import ZoneInfo, ZoneInfoNotFoundError

import openpyxl

from models import Engineer, Incident, RosterResolution, RosterResolutionStatus

logger = logging.getLogger("gsm_automator.roster")

REQUIRED_COLUMNS: Set[str] = {
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
}

DAY_ALIASES: Dict[str, str] = {
    "mon": "Mon",
    "monday": "Mon",
    "tue": "Tue",
    "tues": "Tue",
    "tuesday": "Tue",
    "wed": "Wed",
    "wednesday": "Wed",
    "thu": "Thu",
    "thur": "Thu",
    "thurs": "Thu",
    "thursday": "Thu",
    "fri": "Fri",
    "friday": "Fri",
    "sat": "Sat",
    "saturday": "Sat",
    "sun": "Sun",
    "sunday": "Sun",
}

TRUTHY_STRINGS: Set[str] = {"true", "yes", "y", "1", "on", "leave"}
FALSY_STRINGS: Set[str] = {"false", "no", "n", "0", "off", ""}


def parse_bool(value: object, default: bool = False) -> bool:
    """Safely parse Excel boolean/string representations such as False, 'No', 'Yes', True."""
    if isinstance(value, bool):
        return value
    if value is None:
        return default
    if isinstance(value, (int, float)):
        return bool(value)
    normalized = str(value).strip().lower()
    if normalized in TRUTHY_STRINGS:
        return True
    if normalized in FALSY_STRINGS:
        return False
    return default


def parse_shift_time(value: object, field_name: str) -> time:
    """Parse HH:MM / HH:MM:SS strings or Excel datetime/time objects into datetime.time."""
    if isinstance(value, datetime):
        return value.time().replace(second=0, microsecond=0)
    if isinstance(value, time):
        return value.replace(second=0, microsecond=0)
    if value is None:
        raise ValueError(f"Missing value for {field_name}")

    raw = str(value).strip()
    for fmt in ("%H:%M", "%H:%M:%S", "%I:%M %p"):
        try:
            return datetime.strptime(raw, fmt).time()
        except ValueError:
            continue
    raise ValueError(f"Invalid time format '{raw}' in column '{field_name}' (expected HH:MM)")


def parse_active_days(value: object) -> frozenset[str]:
    """Parse comma-separated day strings (e.g. 'Mon,Tue,Wed,Thu,Fri') into canonical 3-letter names."""
    if value is None:
        return frozenset()
    tokens = [tok.strip().lower() for tok in str(value).split(",") if tok.strip()]
    parsed: Set[str] = set()
    for tok in tokens:
        canonical = DAY_ALIASES.get(tok)
        if not canonical:
            raise ValueError(f"Unrecognized weekday '{tok}' in Active_Days")
        parsed.add(canonical)
    return frozenset(parsed)


def is_engineer_in_shift_window(
    engineer: Engineer, reference_utc: Optional[datetime] = None
) -> bool:
    """
    Evaluate whether an engineer is actively on shift at `reference_utc`.

    Correctly handles midnight-crossing overnight shifts using `datetime.time`:
    - Standard shift (Shift_Start <= Shift_End, e.g. 06:00 -> 14:30):
        Active if today's localized weekday is in Active_Days AND Shift_Start <= now_time <= Shift_End.
    - Overnight shift (Shift_Start > Shift_End, e.g. 22:00 -> 06:00):
        - Pre-midnight segment (now_time >= Shift_Start, e.g. 23:15):
            Shift started *today* -> verify today's weekday is in Active_Days.
        - Post-midnight segment (now_time <= Shift_End, e.g. 02:30):
            Shift started *yesterday* -> verify yesterday's weekday is in Active_Days.
    """
    if engineer.on_leave or not engineer.active_days:
        return False

    utc_now = reference_utc or datetime.now(timezone.utc)
    if utc_now.tzinfo is None:
        utc_now = utc_now.replace(tzinfo=timezone.utc)

    try:
        tz = ZoneInfo(engineer.timezone)
    except ZoneInfoNotFoundError:
        logger.error(
            "Invalid IANA Timezone '%s' configured for engineer '%s' (%s).",
            engineer.timezone,
            engineer.engineer_name,
            engineer.email_or_upn,
        )
        return False

    local_now = utc_now.astimezone(tz)
    current_time: time = local_now.time().replace(second=0, microsecond=0)
    today_str = local_now.strftime("%a")
    yesterday_str = (local_now - timedelta(days=1)).strftime("%a")

    start = engineer.shift_start
    end = engineer.shift_end

    # Case 1: 24-hour shift (e.g., 00:00 to 00:00)
    if start == end:
        return today_str in engineer.active_days

    # Case 2: Standard intra-day shift (e.g., 06:00 to 14:30)
    if start < end:
        return (today_str in engineer.active_days) and (start <= current_time <= end)

    # Case 3: Overnight shift crossing midnight (e.g., 22:00 to 06:00)
    if current_time >= start:
        return today_str in engineer.active_days
    if current_time <= end:
        return yesterday_str in engineer.active_days

    return False


class RosterEngine:
    """
    Dynamically loads `shift_roster.xlsx` on each polling cycle and selects
    eligible shift engineers via Load-Balanced or Round-Robin scheduling.
    """

    def __init__(
        self,
        roster_path: str | Path,
        selection_strategy: str = "load_balanced",
        assign_to_lead_when_unmanned: bool = True,
        max_retries: int = 5,
        retry_delay_seconds: float = 2.0,
    ) -> None:
        self.roster_path = Path(roster_path)
        self.selection_strategy = selection_strategy.strip().lower()
        self.assign_to_lead_when_unmanned = assign_to_lead_when_unmanned
        self.max_retries = max_retries
        self.retry_delay_seconds = retry_delay_seconds

        self._engineers: List[Engineer] = []
        self._known_groups_canonical: Dict[str, str] = {}

        # Load-balancing & Round-Robin state persisted across polling cycles
        self._assignment_counts: Dict[str, int] = defaultdict(int)
        self._last_assigned_epoch: Dict[str, float] = defaultdict(float)
        self._rr_cursors: Dict[str, int] = defaultdict(int)

    def seed_assignment_counts(self, counts_by_assignee: Dict[str, int]) -> None:
        """Seed historical workload counters from GSM_Incident_Tracker.xlsx on startup."""
        for key, count in counts_by_assignee.items():
            self._assignment_counts[key.strip().lower()] += count
        logger.info(
            "Seeded workload counters for %d engineers from tracking ledger.",
            len(counts_by_assignee),
        )

    def _read_workbook_bytes_with_retry(self) -> bytes:
        """Read Excel file bytes into memory quickly to avoid holding SMB/Windows file locks."""
        if not self.roster_path.exists():
            raise FileNotFoundError(f"Shift roster file not found: {self.roster_path}")

        delay = self.retry_delay_seconds
        last_exc: Optional[Exception] = None
        for attempt in range(1, self.max_retries + 1):
            try:
                return self.roster_path.read_bytes()
            except (PermissionError, OSError) as exc:
                last_exc = exc
                logger.warning(
                    "Shift roster '%s' is temporarily locked (attempt %d/%d): %s. Retrying in %.1fs...",
                    self.roster_path,
                    attempt,
                    self.max_retries,
                    exc,
                    delay,
                )
                time_mod.sleep(delay)
                delay *= 1.5

        raise PermissionError(
            f"Unable to read locked roster file '{self.roster_path}' after {self.max_retries} attempts."
        ) from last_exc

    def reload_roster(self) -> int:
        """
        Re-read `shift_roster.xlsx` dynamically on each polling cycle so shift/leave edits
        by Team Leads take effect immediately without restarting the daemon.
        """
        raw_bytes = self._read_workbook_bytes_with_retry()
        wb = openpyxl.load_workbook(
            filename=io.BytesIO(raw_bytes), read_only=True, data_only=True
        )
        try:
            sheet = wb.active
            if sheet is None:
                raise ValueError(f"Workbook '{self.roster_path}' contains no active worksheet.")

            rows_iter = sheet.iter_rows(values_only=True)
            header_row = next(rows_iter, None)
            if not header_row:
                raise ValueError(f"Workbook '{self.roster_path}' is empty.")

            headers = [str(col).strip() if col is not None else "" for col in header_row]
            missing = REQUIRED_COLUMNS - set(headers)
            if missing:
                raise ValueError(
                    f"Roster schema mismatch in '{self.roster_path}'. Missing columns: {sorted(missing)}"
                )

            col_idx = {name: idx for idx, name in enumerate(headers) if name}
            loaded_engineers: List[Engineer] = []
            groups_map: Dict[str, str] = {}

            for row_num, row in enumerate(rows_iter, start=2):
                if not row or all(cell is None or str(cell).strip() == "" for cell in row):
                    continue

                try:
                    group = str(row[col_idx["Assignment_Group"]] or "").strip()
                    name = str(row[col_idx["Engineer_Name"]] or "").strip()
                    upn = str(row[col_idx["Email_or_UPN"]] or "").strip()
                    shift_name = str(row[col_idx["Shift_Name"]] or "").strip()
                    tz_name = str(row[col_idx["Timezone"]] or "").strip()

                    if not group or not name or not upn:
                        logger.warning(
                            "Skipping incomplete roster row %d (missing group, name, or UPN).",
                            row_num,
                        )
                        continue

                    shift_start = parse_shift_time(row[col_idx["Shift_Start"]], "Shift_Start")
                    shift_end = parse_shift_time(row[col_idx["Shift_End"]], "Shift_End")
                    active_days = parse_active_days(row[col_idx["Active_Days"]])
                    on_leave = parse_bool(row[col_idx["On_Leave"]], default=False)
                    is_lead = parse_bool(row[col_idx["Is_Lead"]], default=False)

                    # Validate IANA timezone eagerly during reload
                    ZoneInfo(tz_name)

                    eng = Engineer(
                        assignment_group=group,
                        engineer_name=name,
                        email_or_upn=upn,
                        shift_name=shift_name,
                        shift_start=shift_start,
                        shift_end=shift_end,
                        timezone=tz_name,
                        active_days=active_days,
                        on_leave=on_leave,
                        is_lead=is_lead,
                    )
                    loaded_engineers.append(eng)
                    groups_map[group.lower()] = group
                except Exception as exc:  # noqa: BLE001
                    logger.error(
                        "Skipping malformed roster row %d in '%s': %s",
                        row_num,
                        self.roster_path,
                        exc,
                    )

            self._engineers = loaded_engineers
            self._known_groups_canonical = groups_map
            logger.info(
                "Reloaded shift roster '%s': %d valid engineer entries across %d assignment groups.",
                self.roster_path.name,
                len(self._engineers),
                len(self._known_groups_canonical),
            )
            return len(self._engineers)
        finally:
            wb.close()

    def _select_candidate(self, group_key: str, candidates: Iterable[Engineer]) -> Engineer:
        """Select the next engineer using Load-Balanced or Round-Robin strategy."""
        ordered = sorted(candidates, key=lambda e: e.email_or_upn.lower())
        if len(ordered) == 1:
            chosen = ordered[0]
            self._record_selection(chosen)
            return chosen

        if self.selection_strategy == "round_robin":
            cursor = self._rr_cursors[group_key] % len(ordered)
            chosen = ordered[cursor]
            self._rr_cursors[group_key] = cursor + 1
            self._record_selection(chosen)
            return chosen

        # Default: "load_balanced" (least tickets assigned -> oldest assignment timestamp -> UPN)
        def sort_key(eng: Engineer) -> Tuple[int, float, str]:
            upn_key = eng.email_or_upn.lower()
            name_key = eng.engineer_name.lower()
            count = self._assignment_counts.get(upn_key, 0)
            if name_key != upn_key:
                count += self._assignment_counts.get(name_key, 0)
            last_ts = self._last_assigned_epoch.get(upn_key, 0.0)
            return (count, last_ts, upn_key)

        chosen = min(ordered, key=sort_key)
        self._record_selection(chosen)
        return chosen

    def _record_selection(self, engineer: Engineer) -> None:
        upn_key = engineer.email_or_upn.lower()
        self._assignment_counts[upn_key] += 1
        self._last_assigned_epoch[upn_key] = time_mod.time()

    def resolve_engineer(
        self, incident: Incident, reference_utc: Optional[datetime] = None
    ) -> RosterResolution:
        """
        Determine the optimal shift engineer for `incident`.

        Returns:
            - `ASSIGNED_SHIFT` if one or more active shift engineers match.
            - `ASSIGNED_LEAD_FALLBACK` if no shift engineer is active, `assign_to_lead_when_unmanned`
              is True, and an available Lead exists for the group.
            - `NO_PERSONNEL_AVAILABLE` if group exists in roster, but nobody (nor lead) is available.
            - `UNKNOWN_ASSIGNMENT_GROUP` if `incident.assignment_group` is absent from the roster.
        """
        group_key = incident.assignment_group.strip().lower()
        group_members = [
            eng for eng in self._engineers if eng.assignment_group.lower() == group_key
        ]

        if not group_members:
            msg = (
                f"Assignment_Group '{incident.assignment_group}' for incident "
                f"'{incident.incident_id}' does not exist in shift_roster.xlsx."
            )
            logger.error(msg)
            return RosterResolution(
                status=RosterResolutionStatus.UNKNOWN_ASSIGNMENT_GROUP,
                engineer=None,
                reason=msg,
            )

        eligible_on_shift = [
            eng
            for eng in group_members
            if is_engineer_in_shift_window(eng, reference_utc=reference_utc)
        ]

        if eligible_on_shift:
            chosen = self._select_candidate(group_key, eligible_on_shift)
            logger.info(
                "Incident %s (%s) -> Selected active shift engineer '%s' (%s, Shift=%s).",
                incident.incident_id,
                incident.assignment_group,
                chosen.engineer_name,
                chosen.email_or_upn,
                chosen.shift_name,
            )
            return RosterResolution(
                status=RosterResolutionStatus.ASSIGNED_SHIFT,
                engineer=chosen,
                reason=f"Assigned via {self.selection_strategy} among {len(eligible_on_shift)} active shift engineer(s).",
            )

        # Fallback: check designated Is_Lead engineer(s) for that Assignment_Group
        if self.assign_to_lead_when_unmanned:
            available_leads = [
                eng for eng in group_members if eng.is_lead and not eng.on_leave
            ]
            if available_leads:
                lead = self._select_candidate(f"{group_key}:lead", available_leads)
                logger.warning(
                    "No active shift engineer for '%s' (Incident %s). Falling back to group Lead '%s' (%s).",
                    incident.assignment_group,
                    incident.incident_id,
                    lead.engineer_name,
                    lead.email_or_upn,
                )
                return RosterResolution(
                    status=RosterResolutionStatus.ASSIGNED_LEAD_FALLBACK,
                    engineer=lead,
                    reason="No active shift engineer detected; escalated to designated Shift/Group Lead.",
                )

        logger.warning(
            "No active shift personnel or available Lead detected for group '%s' (Incident %s).",
            incident.assignment_group,
            incident.incident_id,
        )
        return RosterResolution(
            status=RosterResolutionStatus.NO_PERSONNEL_AVAILABLE,
            engineer=None,
            reason="No active shift personnel detected in roster.",
        )
