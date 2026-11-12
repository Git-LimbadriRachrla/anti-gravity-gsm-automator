"""Main daemon & polling orchestrator for the Anti-Gravity GSM Automator."""

from __future__ import annotations

import argparse
import logging
import signal
import sys
import time
from types import FrameType
from typing import Any, Dict, Optional

from gsm_worker import GSMAuthenticationError, GSMWorker
from logger_setup import load_config, setup_logging
from models import Incident, RosterResolutionStatus, TicketStatus
from roster_engine import RosterEngine
from teams_notifier import TeamsNotifier
from tracker import IncidentTracker

PRIORITY_SORT_ORDER: Dict[str, int] = {"P1": 1, "P2": 2, "P3": 3, "P4": 4}


class GSMAutomationOrchestrator:
    """Coordinates dynamic roster reloading, GSM scraping/assignment, Teams alerts, and Excel tracking."""

    def __init__(self, config: Dict[str, Any]) -> None:
        self.config = config
        self._shutdown_requested = False

        paths_cfg = config.get("paths", {})
        concurrency_cfg = config.get("excel_concurrency", {})
        roster_cfg = config.get("roster", {})
        polling_cfg = config.get("polling", {})

        self.poll_interval_sec: int = int(polling_cfg.get("interval_seconds", 60))
        self.run_once: bool = bool(polling_cfg.get("run_once", False))
        self.unmanned_work_note: str = str(
            roster_cfg.get(
                "unmanned_work_note", "No active shift personnel detected in roster."
            )
        )

        self.logger = logging.getLogger("gsm_automator.main")

        self.tracker = IncidentTracker(
            tracker_path=paths_cfg.get("tracker_xlsx", "./data/GSM_Incident_Tracker.xlsx"),
            max_retries=int(concurrency_cfg.get("max_retries", 6)),
            retry_delay_seconds=float(concurrency_cfg.get("retry_delay_seconds", 2.5)),
            backoff_multiplier=float(concurrency_cfg.get("backoff_multiplier", 1.5)),
        )

        self.roster_engine = RosterEngine(
            roster_path=paths_cfg.get("roster_xlsx", "./data/shift_roster.xlsx"),
            selection_strategy=str(roster_cfg.get("selection_strategy", "load_balanced")),
            assign_to_lead_when_unmanned=bool(
                roster_cfg.get("assign_to_lead_when_unmanned", True)
            ),
            max_retries=int(concurrency_cfg.get("max_retries", 6)),
            retry_delay_seconds=float(concurrency_cfg.get("retry_delay_seconds", 2.5)),
        )

        # Seed historical assignment counts so load-balancing continues smoothly across restarts
        historical_counts = self.tracker.get_historical_assignment_counts()
        if historical_counts:
            self.roster_engine.seed_assignment_counts(historical_counts)

        self.notifier = TeamsNotifier(config.get("notifications", {}))
        self.worker = GSMWorker(
            portal_config=config.get("gsm_portal", {}),
            screenshots_dir=paths_cfg.get("screenshots_dir", "./logs/screenshots"),
        )

    def request_shutdown(self, signum: int, _frame: Optional[FrameType]) -> None:
        sig_name = signal.Signals(signum).name
        self.logger.info(
            "Received %s — completing active operation before graceful shutdown...", sig_name
        )
        self._shutdown_requested = True

    def _sleep_interruptible(self, seconds: int) -> None:
        end_time = time.monotonic() + seconds
        while time.monotonic() < end_time and not self._shutdown_requested:
            time.sleep(min(1.0, end_time - time.monotonic()))

    def process_incident(self, incident: Incident) -> None:
        """Evaluate roster, update GSM ticket, notify Teams (P1/P2), and persist to Excel ledger."""
        prior_status = self.tracker.get_ticket_status(incident.incident_id)
        if prior_status in {
            TicketStatus.ASSIGNED.value,
            TicketStatus.ASSIGNED_TO_LEAD.value,
        }:
            self.logger.debug(
                "Incident %s already recorded as '%s' in ledger; skipping duplicate assignment.",
                incident.incident_id,
                prior_status,
            )
            return

        ticket_url = self.worker.build_ticket_url(incident.incident_id)
        resolution = self.roster_engine.resolve_engineer(incident)

        # 1. Eligible shift engineer or Lead Fallback found
        if resolution.status in {
            RosterResolutionStatus.ASSIGNED_SHIFT,
            RosterResolutionStatus.ASSIGNED_LEAD_FALLBACK,
        }:
            engineer = resolution.engineer
            assert engineer is not None
            is_lead_fallback = (
                resolution.status == RosterResolutionStatus.ASSIGNED_LEAD_FALLBACK
            )
            portal_note = (
                "Assigned to designated Shift/Group Lead by Anti-Gravity GSM Automator "
                "(no active shift personnel detected in window)."
                if is_lead_fallback
                else None
            )

            assigned_ok = self.worker.assign_incident(
                incident_id=incident.incident_id,
                engineer=engineer,
                work_note=portal_note,
            )
            if not assigned_ok:
                self.tracker.upsert_incident(
                    incident_id=incident.incident_id,
                    priority=incident.normalized_priority,
                    assignment_group=incident.assignment_group,
                    assigned_to="Unassigned",
                    shift=engineer.shift_name,
                    teams_alerted=False,
                    status=TicketStatus.ERROR_PORTAL_FAILURE,
                    remarks=f"Failed to save assignment to {engineer.email_or_upn} on GSM portal.",
                )
                return

            teams_alerted = self.notifier.send_incident_alert(
                incident=incident,
                engineer=engineer,
                ticket_url=ticket_url,
                is_lead_fallback=is_lead_fallback,
            )
            ledger_status = (
                TicketStatus.ASSIGNED_TO_LEAD
                if is_lead_fallback
                else TicketStatus.ASSIGNED
            )
            self.tracker.upsert_incident(
                incident_id=incident.incident_id,
                priority=incident.normalized_priority,
                assignment_group=incident.assignment_group,
                assigned_to=f"{engineer.engineer_name} ({engineer.email_or_upn})",
                shift=engineer.shift_name,
                teams_alerted=teams_alerted,
                status=ledger_status,
                remarks=resolution.reason,
            )
            return

        # 2. Group exists in roster, but zero active shift personnel (and no available Lead)
        if resolution.status == RosterResolutionStatus.NO_PERSONNEL_AVAILABLE:
            already_flagged = prior_status == TicketStatus.UNASSIGNED_NO_SHIFT.value
            if not already_flagged:
                self.worker.post_work_note(incident.incident_id, self.unmanned_work_note)
                alerted = False
                if self.notifier.should_notify_for_incident(incident):
                    alerted = self.notifier.send_supervisor_alert(
                        incident=incident,
                        ticket_url=ticket_url,
                        alert_reason=(
                            f"No active shift personnel or available Lead detected in roster "
                            f"for high-priority group '{incident.assignment_group}'."
                        ),
                    )
            else:
                alerted = False

            self.tracker.upsert_incident(
                incident_id=incident.incident_id,
                priority=incident.normalized_priority,
                assignment_group=incident.assignment_group,
                assigned_to="Unassigned",
                shift="N/A",
                teams_alerted=alerted,
                status=TicketStatus.UNASSIGNED_NO_SHIFT,
                remarks=self.unmanned_work_note,
            )
            return

        # 3. Unknown Assignment_Group (Roster Mismatch)
        if resolution.status == RosterResolutionStatus.UNKNOWN_ASSIGNMENT_GROUP:
            already_errored = prior_status == TicketStatus.ERROR_UNKNOWN_GROUP.value
            alerted = False
            if not already_errored:
                alerted = self.notifier.send_supervisor_alert(
                    incident=incident,
                    ticket_url=ticket_url,
                    alert_reason=resolution.reason,
                )

            self.tracker.upsert_incident(
                incident_id=incident.incident_id,
                priority=incident.normalized_priority,
                assignment_group=incident.assignment_group,
                assigned_to="Unassigned",
                shift="N/A",
                teams_alerted=alerted,
                status=TicketStatus.ERROR_UNKNOWN_GROUP,
                remarks=resolution.reason,
            )

    def run_polling_cycle(self) -> None:
        """Execute a single end-to-end roster reload + GSM incident processing cycle."""
        self.roster_engine.reload_roster()
        incidents = self.worker.fetch_unassigned_incidents()
        if not incidents:
            self.logger.info("No unassigned incidents detected on GSM dashboard.")
            return

        # Prioritize P1 -> P2 -> P3 -> P4
        incidents_sorted = sorted(
            incidents,
            key=lambda inc: PRIORITY_SORT_ORDER.get(inc.normalized_priority, 99),
        )

        for incident in incidents_sorted:
            if self._shutdown_requested:
                break
            try:
                self.process_incident(incident)
            except Exception as exc:  # noqa: BLE001
                self.logger.exception(
                    "Unhandled error processing incident %s: %s",
                    incident.incident_id,
                    exc,
                )

    def run_forever(self) -> int:
        """Main daemon entrypoint with browser lifecycle and graceful signal shutdown."""
        self.logger.info("Starting Anti-Gravity GSM Automator...")
        with self.worker:
            cycle_num = 0
            while not self._shutdown_requested:
                cycle_num += 1
                self.logger.info("--- Starting Polling Cycle #%d ---", cycle_num)
                try:
                    self.run_polling_cycle()
                except GSMAuthenticationError as auth_err:
                    self.logger.critical("Fatal authentication error: %s", auth_err)
                    return 2
                except FileNotFoundError as fnf_err:
                    self.logger.error(
                        "%s — Run `python generate_sample_roster.py` to initialize a sample roster.",
                        fnf_err,
                    )
                except Exception as exc:  # noqa: BLE001
                    self.logger.exception(
                        "Error during polling cycle #%d: %s", cycle_num, exc
                    )

                if self.run_once or self._shutdown_requested:
                    break

                self.logger.info(
                    "Polling cycle #%d complete. Sleeping %ds until next cycle...",
                    cycle_num,
                    self.poll_interval_sec,
                )
                self._sleep_interruptible(self.poll_interval_sec)

        self.logger.info("Anti-Gravity GSM Automator stopped cleanly.")
        return 0


def parse_args() -> argparse.Namespace:
    parser = argparse.ArgumentParser(
        description="Anti-Gravity GSM Automator — Shift-aware incident routing & Teams alerting."
    )
    parser.add_argument(
        "--config",
        default="config.yaml",
        help="Path to YAML configuration file (default: config.yaml)",
    )
    parser.add_argument(
        "--once",
        action="store_true",
        help="Run a single polling cycle and exit.",
    )
    parser.add_argument(
        "--mock",
        action="store_true",
        help="Enable GSM mock mode (simulates GSM portal scraping & assignment without live browser).",
    )
    return parser.parse_args()


def main() -> int:
    args = parse_args()
    config = load_config(args.config)

    if args.once:
        config.setdefault("polling", {})["run_once"] = True
    if args.mock:
        config.setdefault("gsm_portal", {})["mock_mode"] = True

    log_cfg = config.get("logging", {})
    paths_cfg = config.get("paths", {})
    setup_logging(
        log_file=paths_cfg.get("log_file", "./logs/gsm_automation.log"),
        level=str(log_cfg.get("level", "INFO")),
        max_bytes=int(log_cfg.get("max_bytes", 10 * 1024 * 1024)),
        backup_count=int(log_cfg.get("backup_count", 5)),
    )

    orchestrator = GSMAutomationOrchestrator(config)
    signal.signal(signal.SIGINT, orchestrator.request_shutdown)
    signal.signal(signal.SIGTERM, orchestrator.request_shutdown)

    return orchestrator.run_forever()


if __name__ == "__main__":
    sys.exit(main())
