"""Domain data models and enumerations for the Anti-Gravity GSM Automator."""

from __future__ import annotations

from dataclasses import dataclass
from datetime import time
from enum import Enum
from typing import FrozenSet, Optional


class RosterResolutionStatus(str, Enum):
    """Outcome categories when evaluating the shift roster for an incident."""

    ASSIGNED_SHIFT = "ASSIGNED_SHIFT"
    ASSIGNED_LEAD_FALLBACK = "ASSIGNED_LEAD_FALLBACK"
    NO_PERSONNEL_AVAILABLE = "NO_PERSONNEL_AVAILABLE"
    UNKNOWN_ASSIGNMENT_GROUP = "UNKNOWN_ASSIGNMENT_GROUP"


class TicketStatus(str, Enum):
    """Lifecycle states stored in GSM_Incident_Tracker.xlsx."""

    ASSIGNED = "Assigned"
    ASSIGNED_TO_LEAD = "Assigned (Lead Fallback)"
    UNASSIGNED_NO_SHIFT = "Unassigned - No Shift Personnel"
    ERROR_UNKNOWN_GROUP = "Error - Unknown Assignment Group"
    ERROR_PORTAL_FAILURE = "Error - GSM Portal Update Failed"


@dataclass(frozen=True)
class Engineer:
    """Represents a single row from shift_roster.xlsx."""

    assignment_group: str
    engineer_name: str
    email_or_upn: str
    shift_name: str
    shift_start: time
    shift_end: time
    timezone: str
    active_days: FrozenSet[str]
    on_leave: bool
    is_lead: bool


@dataclass(frozen=True)
class Incident:
    """Represents an unassigned incident scraped from the GSM Portal."""

    incident_id: str
    priority: str
    summary: str
    assignment_group: str
    timestamp: str

    @property
    def normalized_priority(self) -> str:
        """Normalize priority values like '1 - Critical' or 'p1' into 'P1'–'P4'."""
        raw = self.priority.strip().upper()
        for level in ("P1", "P2", "P3", "P4"):
            if level in raw:
                return level
        if raw and raw[0] in ("1", "2", "3", "4"):
            return f"P{raw[0]}"
        return raw


@dataclass(frozen=True)
class RosterResolution:
    """Result returned by RosterEngine when resolving an assignee for an Incident."""

    status: RosterResolutionStatus
    engineer: Optional[Engineer]
    reason: str
