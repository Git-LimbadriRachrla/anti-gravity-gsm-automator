"""Unit tests covering overnight shifts, timezone handling, leave filtering, and load-balancing."""

from __future__ import annotations

from datetime import datetime, time, timezone
from pathlib import Path
from zoneinfo import ZoneInfo

from generate_sample_roster import generate_roster_workbook
from models import Engineer, Incident, RosterResolutionStatus
from roster_engine import RosterEngine, is_engineer_in_shift_window
from tracker import IncidentTracker


def test_overnight_shift_pre_and_post_midnight() -> None:
    """Verify 22:00 -> 06:00 Mon-Fri shift covers Fri 23:30 and Sat 03:00, but NOT Mon 03:00."""
    ist = ZoneInfo("Asia/Kolkata")
    night_eng = Engineer(
        assignment_group="Cloud-Ops-L2",
        engineer_name="Night Engineer",
        email_or_upn="night@company.com",
        shift_name="Night",
        shift_start=time(22, 0),
        shift_end=time(6, 0),
        timezone="Asia/Kolkata",
        active_days=frozenset({"Mon", "Tue", "Wed", "Thu", "Fri"}),
        on_leave=False,
        is_lead=False,
    )

    # Friday 23:30 IST -> Active (started Friday night)
    fri_late = datetime(2026, 11, 13, 23, 30, tzinfo=ist).astimezone(timezone.utc)
    assert is_engineer_in_shift_window(night_eng, fri_late) is True

    # Saturday 03:15 IST -> Active (spillover from Friday 22:00 shift!)
    sat_early = datetime(2026, 11, 14, 3, 15, tzinfo=ist).astimezone(timezone.utc)
    assert is_engineer_in_shift_window(night_eng, sat_early) is True

    # Monday 03:15 IST -> Inactive (would belong to Sunday 22:00 shift, and Sun not in Active_Days)
    mon_early = datetime(2026, 11, 16, 3, 15, tzinfo=ist).astimezone(timezone.utc)
    assert is_engineer_in_shift_window(night_eng, mon_early) is False


def test_load_balancing_and_unknown_group(tmp_path: Path) -> None:
    roster_file = generate_roster_workbook(tmp_path / "shift_roster.xlsx")
    engine = RosterEngine(roster_file, selection_strategy="load_balanced")
    engine.reload_roster()

    # Thursday 15:30 IST -> Both Arjun Mehta & Priya Nair are active on General shift (14:00-22:30)
    ist = ZoneInfo("Asia/Kolkata")
    thu_afternoon = datetime(2026, 11, 12, 15, 30, tzinfo=ist).astimezone(timezone.utc)

    inc1 = Incident("INC-1", "P1", "Test 1", "Cloud-Ops-L2", "2026-11-12T10:00:00Z")
    inc2 = Incident("INC-2", "P2", "Test 2", "Cloud-Ops-L2", "2026-11-12T10:01:00Z")

    res1 = engine.resolve_engineer(inc1, reference_utc=thu_afternoon)
    res2 = engine.resolve_engineer(inc2, reference_utc=thu_afternoon)

    assert res1.status == RosterResolutionStatus.ASSIGNED_SHIFT
    assert res2.status == RosterResolutionStatus.ASSIGNED_SHIFT
    assert res1.engineer is not None and res2.engineer is not None
    assert res1.engineer.email_or_upn != res2.engineer.email_or_upn

    # Unknown group test
    inc_unknown = Incident("INC-99", "P1", "Unknown", "NonExistent-Group", "2026-11-12T10:02:00Z")
    res_unk = engine.resolve_engineer(inc_unknown, reference_utc=thu_afternoon)
    assert res_unk.status == RosterResolutionStatus.UNKNOWN_ASSIGNMENT_GROUP


def test_tracker_deduplication(tmp_path: Path) -> None:
    tracker = IncidentTracker(tmp_path / "GSM_Incident_Tracker.xlsx")
    a1 = tracker.upsert_incident(
        incident_id="INC-500",
        priority="P1",
        assignment_group="Cloud-Ops-L2",
        assigned_to="Unassigned",
        shift="N/A",
        teams_alerted=False,
        status="Unassigned - No Shift Personnel",
        remarks="Initial check",
    )
    a2 = tracker.upsert_incident(
        incident_id="INC-500",
        priority="P1",
        assignment_group="Cloud-Ops-L2",
        assigned_to="Jane Doe (jane.doe@company.com)",
        shift="Morning",
        teams_alerted=True,
        status="Assigned",
        remarks="Assigned on retry",
    )
    assert a1 == "INSERTED"
    assert a2 == "UPDATED"
    assert tracker.get_ticket_status("INC-500") == "Assigned"
