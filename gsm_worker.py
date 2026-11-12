"""Playwright-based GSM Portal scraper and ticket assignment automation engine."""

from __future__ import annotations

from datetime import datetime, timezone
import logging
from pathlib import Path
from typing import Any, Dict, List, Optional, Set

from models import Engineer, Incident

logger = logging.getLogger("gsm_automator.worker")


class GSMAuthenticationError(RuntimeError):
    """Raised when SSO/MFA session has expired while running in headless mode."""


class GSMWorker:
    """
    Automates the GSM Incident Portal via Playwright.

    Supports:
    - Persistent Chromium profile (`user_data_dir`) to preserve Enterprise SSO / MFA tokens.
    - Cookie/storage-state JSON (`storage_state_path`) import & automatic refresh.
    - Extraction of unassigned incidents (`Incident ID`, `Priority`, `Summary`,
      `Assignment Group`, `Timestamp`).
    - Ticket assignment and internal work-note updates with failure screenshots.
    - Safe `mock_mode` for local integration testing and dry runs.
    """

    def __init__(
        self,
        portal_config: Dict[str, Any],
        screenshots_dir: str | Path = "./logs/screenshots",
    ) -> None:
        self.cfg = portal_config
        self.base_url: str = self.cfg.get("base_url", "https://gsm.company.internal").rstrip("/")
        self.dashboard_url: str = self.cfg.get(
            "dashboard_url", f"{self.base_url}/incidents/unassigned"
        )
        self.ticket_url_template: str = self.cfg.get(
            "ticket_url_template", f"{self.base_url}/incidents/view/{{incident_id}}"
        )
        self.user_data_dir: str = str(self.cfg.get("user_data_dir") or "").strip()
        self.storage_state_path: str = str(self.cfg.get("storage_state_path") or "").strip()
        self.headless: bool = bool(self.cfg.get("headless", False))
        self.mock_mode: bool = bool(self.cfg.get("mock_mode", False))
        self.nav_timeout_ms: int = int(self.cfg.get("navigation_timeout_ms", 30000))
        self.action_timeout_ms: int = int(self.cfg.get("action_timeout_ms", 15000))
        self.selectors: Dict[str, str] = self.cfg.get("selectors", {})
        self.screenshots_dir = Path(screenshots_dir)

        self._pw: Any = None
        self._browser: Any = None
        self._context: Any = None
        self._page: Any = None

        # Used in mock_mode so assigned mock tickets are not infinitely re-returned
        self._mock_assigned_ids: Set[str] = set()

    def __enter__(self) -> GSMWorker:
        self.start()
        return self

    def __exit__(self, exc_type: Any, exc_val: Any, exc_tb: Any) -> None:
        self.close()

    def build_ticket_url(self, incident_id: str) -> str:
        """Return direct hyperlink to a GSM ticket."""
        return self.ticket_url_template.format(incident_id=incident_id)

    def start(self) -> None:
        """Launch Playwright browser context (persistent profile or cookie storage state)."""
        if self.mock_mode:
            logger.info("GSMWorker initialized in MOCK MODE (no live browser launched).")
            return

        from playwright.sync_api import sync_playwright

        self._pw = sync_playwright().start()
        if self.user_data_dir:
            profile_path = Path(self.user_data_dir)
            profile_path.mkdir(parents=True, exist_ok=True)
            logger.info(
                "Launching persistent Chromium context at '%s' (headless=%s)...",
                profile_path,
                self.headless,
            )
            self._context = self._pw.chromium.launch_persistent_context(
                user_data_dir=str(profile_path),
                headless=self.headless,
                viewport={"width": 1440, "height": 900},
            )
            self._page = (
                self._context.pages[0]
                if self._context.pages
                else self._context.new_page()
            )
        else:
            logger.info("Launching standard Chromium browser (headless=%s)...", self.headless)
            self._browser = self._pw.chromium.launch(headless=self.headless)
            ctx_kwargs: Dict[str, Any] = {"viewport": {"width": 1440, "height": 900}}
            if self.storage_state_path and Path(self.storage_state_path).exists():
                ctx_kwargs["storage_state"] = self.storage_state_path
                logger.info("Loaded session storage state from '%s'.", self.storage_state_path)
            self._context = self._browser.new_context(**ctx_kwargs)
            self._page = self._context.new_page()

        self._page.set_default_navigation_timeout(self.nav_timeout_ms)
        self._page.set_default_timeout(self.action_timeout_ms)

    def close(self) -> None:
        """Persist session state if configured and cleanly shut down Playwright resources."""
        if self.mock_mode:
            return

        try:
            self._persist_storage_state()
        except Exception as exc:  # noqa: BLE001
            logger.debug("Could not save storage state during close: %s", exc)

        for resource_name, resource in (
            ("context", self._context),
            ("browser", self._browser),
            ("playwright", self._pw),
        ):
            if resource is not None:
                try:
                    method = getattr(resource, "close", None) or getattr(resource, "stop", None)
                    if method:
                        method()
                except Exception as exc:  # noqa: BLE001
                    logger.debug("Error closing %s: %s", resource_name, exc)

        self._page = None
        self._context = None
        self._browser = None
        self._pw = None

    def _persist_storage_state(self) -> None:
        """Save updated SSO cookies/localStorage when storage_state_path is configured."""
        if self._context is not None and self.storage_state_path:
            state_file = Path(self.storage_state_path)
            state_file.parent.mkdir(parents=True, exist_ok=True)
            self._context.storage_state(path=str(state_file))

    def _capture_failure_screenshot(self, context_label: str) -> Optional[Path]:
        """Save diagnostic screenshot on scraping or DOM interaction failure."""
        if self._page is None:
            return None
        try:
            self.screenshots_dir.mkdir(parents=True, exist_ok=True)
            ts = datetime.now(timezone.utc).strftime("%Y%m%dT%H%M%SZ")
            safe_label = "".join(c if c.isalnum() or c in "-_" else "_" for c in context_label)
            out_path = self.screenshots_dir / f"{ts}_{safe_label}.png"
            self._page.screenshot(path=str(out_path), full_page=True)
            logger.info("Saved diagnostic screenshot: %s", out_path)
            return out_path
        except Exception as exc:  # noqa: BLE001
            logger.warning("Failed to capture diagnostic screenshot: %s", exc)
            return None

    def _verify_authenticated(self) -> None:
        """Detect SSO/MFA login prompts and allow interactive completion when non-headless."""
        sso_selector = self.selectors.get(
            "sso_login_indicator", "input[type='email'], #i0116, .sso-login-prompt"
        )
        if self._page.locator(sso_selector).count() == 0:
            return

        if self.headless:
            self._capture_failure_screenshot("sso_expired_headless")
            raise GSMAuthenticationError(
                "SSO/MFA login screen detected while running in headless mode. "
                "Re-run once with GSM_HEADLESS=false to refresh persistent browser session."
            )

        logger.warning(
            "SSO/MFA authentication prompt detected! Waiting up to 120s for interactive sign-in..."
        )
        ready_sel = self.selectors.get(
            "dashboard_ready", "table#incident-grid, [data-testid='incident-table']"
        )
        self._page.wait_for_selector(ready_sel, timeout=120_000)
        self._persist_storage_state()
        logger.info("Interactive SSO/MFA sign-in succeeded; session persisted.")

    def fetch_unassigned_incidents(self) -> List[Incident]:
        """Navigate to the GSM Incident Dashboard and extract unassigned incidents."""
        if self.mock_mode:
            return self._generate_mock_incidents()

        if self._page is None:
            self.start()

        try:
            logger.info("Navigating to GSM Dashboard: %s", self.dashboard_url)
            self._page.goto(self.dashboard_url, wait_until="domcontentloaded")
            self._verify_authenticated()

            ready_sel = self.selectors.get(
                "dashboard_ready", "table#incident-grid, [data-testid='incident-table']"
            )
            self._page.wait_for_selector(ready_sel, timeout=self.nav_timeout_ms)

            row_sel = self.selectors.get(
                "unassigned_rows", "table#incident-grid tbody tr.unassigned-row"
            )
            col_id = self.selectors.get("col_incident_id", "td.col-incident-id")
            col_prio = self.selectors.get("col_priority", "td.col-priority")
            col_sum = self.selectors.get("col_summary", "td.col-summary")
            col_grp = self.selectors.get("col_assignment_group", "td.col-assignment-group")
            col_ts = self.selectors.get("col_timestamp", "td.col-created-at")

            rows = self._page.locator(row_sel)
            row_count = rows.count()
            incidents: List[Incident] = []

            for idx in range(row_count):
                row = rows.nth(idx)
                inc_id = row.locator(col_id).inner_text().strip()
                prio = row.locator(col_prio).inner_text().strip()
                summary = row.locator(col_sum).inner_text().strip()
                group = row.locator(col_grp).inner_text().strip()
                ts_val = row.locator(col_ts).inner_text().strip()

                if not inc_id:
                    continue

                incidents.append(
                    Incident(
                        incident_id=inc_id,
                        priority=prio,
                        summary=summary,
                        assignment_group=group,
                        timestamp=ts_val,
                    )
                )

            self._persist_storage_state()
            logger.info("Extracted %d unassigned incident(s) from GSM dashboard.", len(incidents))
            return incidents

        except Exception as exc:
            self._capture_failure_screenshot("fetch_unassigned_failed")
            logger.exception("Failed to scrape unassigned incidents from GSM portal: %s", exc)
            raise

    def assign_incident(
        self,
        incident_id: str,
        engineer: Engineer,
        work_note: Optional[str] = None,
    ) -> bool:
        """Open ticket in GSM portal, populate Assigned To (+ optional work note), and save."""
        if self.mock_mode:
            self._mock_assigned_ids.add(incident_id)
            logger.info(
                "[MOCK GSM] Assigned ticket %s -> %s (%s). Note=%r",
                incident_id,
                engineer.engineer_name,
                engineer.email_or_upn,
                work_note,
            )
            return True

        ticket_url = self.build_ticket_url(incident_id)
        try:
            logger.info(
                "Assigning GSM ticket %s to '%s' (%s)...",
                incident_id,
                engineer.engineer_name,
                engineer.email_or_upn,
            )
            self._page.goto(ticket_url, wait_until="domcontentloaded")
            self._verify_authenticated()

            assignee_sel = self.selectors.get(
                "assignee_input", "input#assignee-field, [name='assigned_to']"
            )
            note_sel = self.selectors.get(
                "work_note_textarea", "textarea#work-notes, [name='work_notes']"
            )
            save_sel = self.selectors.get(
                "save_button", "button#save-incident, button[type='submit']"
            )

            assignee_field = self._page.locator(assignee_sel).first
            assignee_field.wait_for(state="visible", timeout=self.action_timeout_ms)
            assignee_field.fill(engineer.email_or_upn)
            assignee_field.press("Enter")

            if work_note:
                note_field = self._page.locator(note_sel).first
                if note_field.count() > 0:
                    note_field.fill(work_note)

            save_btn = self._page.locator(save_sel).first
            save_btn.click()
            self._page.wait_for_load_state("networkidle", timeout=self.nav_timeout_ms)
            self._persist_storage_state()

            logger.info("Successfully saved GSM assignment for %s.", incident_id)
            return True

        except Exception as exc:  # noqa: BLE001
            self._capture_failure_screenshot(f"assign_failed_{incident_id}")
            logger.error("Failed to assign GSM ticket %s: %s", incident_id, exc)
            return False

    def post_work_note(self, incident_id: str, note: str) -> bool:
        """Post an internal work note on an unassigned ticket and save state."""
        if self.mock_mode:
            logger.info("[MOCK GSM] Posted work note on %s: %r", incident_id, note)
            return True

        ticket_url = self.build_ticket_url(incident_id)
        try:
            logger.info("Posting work note on GSM ticket %s: '%s'", incident_id, note)
            self._page.goto(ticket_url, wait_until="domcontentloaded")
            self._verify_authenticated()

            note_sel = self.selectors.get(
                "work_note_textarea", "textarea#work-notes, [name='work_notes']"
            )
            save_sel = self.selectors.get(
                "save_button", "button#save-incident, button[type='submit']"
            )

            note_field = self._page.locator(note_sel).first
            note_field.wait_for(state="visible", timeout=self.action_timeout_ms)
            note_field.fill(note)

            save_btn = self._page.locator(save_sel).first
            save_btn.click()
            self._page.wait_for_load_state("networkidle", timeout=self.nav_timeout_ms)
            self._persist_storage_state()
            return True

        except Exception as exc:  # noqa: BLE001
            self._capture_failure_screenshot(f"worknote_failed_{incident_id}")
            logger.error("Failed to post work note on GSM ticket %s: %s", incident_id, exc)
            return False

    def _generate_mock_incidents(self) -> List[Incident]:
        """Return realistic sample incidents for dry-run / testing workflows."""
        now_iso = datetime.now(timezone.utc).strftime("%Y-%m-%d %H:%M:%S UTC")
        samples = [
            Incident(
                incident_id="INC-900101",
                priority="P1",
                summary="Production Kubernetes API Server High Latency in ap-south-1",
                assignment_group="Cloud-Ops-L2",
                timestamp=now_iso,
            ),
            Incident(
                incident_id="INC-900102",
                priority="P2",
                summary="Primary PostgreSQL Read-Replica Replication Lag > 120s",
                assignment_group="DB-Reliability-L3",
                timestamp=now_iso,
            ),
            Incident(
                incident_id="INC-900103",
                priority="P3",
                summary="VPN Split-Tunnel Route Advertisement Warning on Edge Gateway",
                assignment_group="NetSec-Ops",
                timestamp=now_iso,
            ),
            Incident(
                incident_id="INC-900104",
                priority="P2",
                summary="Unrecognized Mainframe Batch Queue Stalled",
                assignment_group="Legacy-Mainframe-Unknown",
                timestamp=now_iso,
            ),
        ]
        return [inc for inc in samples if inc.incident_id not in self._mock_assigned_ids]
