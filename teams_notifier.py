"""Microsoft Teams Incoming Webhook notifier using Adaptive Cards v1.4 with UPN @mentions."""

from __future__ import annotations

import logging
import time
from typing import Any, Dict, Iterable, Optional

import requests

from models import Engineer, Incident

logger = logging.getLogger("gsm_automator.teams")


class TeamsNotifier:
    """
    Dispatches rich Adaptive Card notifications to Microsoft Teams Incoming Webhooks
    for high-severity incidents (P1/P2) and Shift Supervisor roster mismatch escalations.
    """

    def __init__(self, notif_config: Dict[str, Any]) -> None:
        self.incident_webhook_url: str = str(
            notif_config.get("incident_webhook_url") or ""
        ).strip()
        self.supervisor_webhook_url: str = str(
            notif_config.get("supervisor_webhook_url") or self.incident_webhook_url
        ).strip()

        levels: Iterable[str] = notif_config.get("high_priority_levels", ["P1", "P2"])
        self.high_priority_levels = {str(lvl).strip().upper() for lvl in levels}

        self.http_timeout: int = int(notif_config.get("http_timeout_seconds", 10))
        self.max_retries: int = int(notif_config.get("max_retries", 3))

        sup_cfg = notif_config.get("supervisor", {}) or {}
        self.supervisor_name: str = str(sup_cfg.get("name", "Shift Supervisor")).strip()
        self.supervisor_upn: str = str(sup_cfg.get("email_or_upn", "")).strip()

    def should_notify_for_incident(self, incident: Incident) -> bool:
        """Return True if the incident's normalized priority requires an immediate Teams alert."""
        return incident.normalized_priority in self.high_priority_levels

    @staticmethod
    def _is_configured_webhook(url: str) -> bool:
        return bool(url) and "YOUR-" not in url and url.startswith("https://")

    def build_incident_adaptive_card(
        self,
        incident: Incident,
        engineer: Engineer,
        ticket_url: str,
        is_lead_fallback: bool = False,
    ) -> Dict[str, Any]:
        """Construct an MS Teams Adaptive Card v1.4 payload with native UPN @mention."""
        prio = incident.normalized_priority
        banner_style = "attention" if prio == "P1" else "warning"
        mention_tag = f"<at>{engineer.engineer_name}</at>"
        assignment_note = (
            f"{mention_tag} *(Lead Escalation)*" if is_lead_fallback else mention_tag
        )

        card_content: Dict[str, Any] = {
            "$schema": "http://adaptivecards.io/schemas/adaptive-card.json",
            "type": "AdaptiveCard",
            "version": "1.4",
            "body": [
                {
                    "type": "Container",
                    "style": banner_style,
                    "bleed": True,
                    "items": [
                        {
                            "type": "TextBlock",
                            "text": f"🚨 [{prio} SEVERITY] GSM Incident Assigned — {incident.incident_id}",
                            "weight": "Bolder",
                            "size": "Medium",
                            "wrap": True,
                        }
                    ],
                },
                {
                    "type": "TextBlock",
                    "text": incident.summary,
                    "weight": "Bolder",
                    "size": "Large",
                    "wrap": True,
                    "spacing": "Medium",
                },
                {
                    "type": "FactSet",
                    "spacing": "Medium",
                    "facts": [
                        {"title": "Incident ID:", "value": incident.incident_id},
                        {"title": "Severity:", "value": f"🚨 {prio}"},
                        {"title": "Assignment Group:", "value": incident.assignment_group},
                        {"title": "Assigned Engineer:", "value": engineer.engineer_name},
                        {
                            "title": "Shift Window:",
                            "value": (
                                f"{engineer.shift_name} "
                                f"({engineer.shift_start.strftime('%H:%M')}–"
                                f"{engineer.shift_end.strftime('%H:%M')} {engineer.timezone})"
                            ),
                        },
                        {"title": "Reported Timestamp:", "value": incident.timestamp},
                    ],
                },
                {
                    "type": "TextBlock",
                    "text": f"👋 Action required: {assignment_note} — please acknowledge this incident immediately in GSM.",
                    "wrap": True,
                    "spacing": "Medium",
                },
            ],
            "actions": [
                {
                    "type": "Action.OpenUrl",
                    "title": f"🔗 Open {incident.incident_id} in GSM Portal",
                    "url": ticket_url,
                }
            ],
            "msteams": {
                "width": "Full",
                "entities": [
                    {
                        "type": "mention",
                        "text": mention_tag,
                        "mentioned": {
                            "id": engineer.email_or_upn,
                            "name": engineer.engineer_name,
                        },
                    }
                ],
            },
        }

        return {
            "type": "message",
            "attachments": [
                {
                    "contentType": "application/vnd.microsoft.card.adaptive",
                    "contentUrl": None,
                    "content": card_content,
                }
            ],
        }

    def send_incident_alert(
        self,
        incident: Incident,
        engineer: Engineer,
        ticket_url: str,
        is_lead_fallback: bool = False,
    ) -> bool:
        """Send Adaptive Card alert if incident priority is in high_priority_levels (P1/P2)."""
        if not self.should_notify_for_incident(incident):
            logger.debug(
                "Skipping Teams notification for %s (priority %s not in %s).",
                incident.incident_id,
                incident.normalized_priority,
                sorted(self.high_priority_levels),
            )
            return False

        payload = self.build_incident_adaptive_card(
            incident=incident,
            engineer=engineer,
            ticket_url=ticket_url,
            is_lead_fallback=is_lead_fallback,
        )
        return self._post_webhook(
            url=self.incident_webhook_url,
            payload=payload,
            context=f"Incident {incident.incident_id} ({incident.normalized_priority})",
        )

    def send_supervisor_alert(
        self,
        incident: Incident,
        ticket_url: str,
        alert_reason: str,
    ) -> bool:
        """Alert the Shift Supervisor when an unknown Assignment_Group or unmanned queue is encountered."""
        entities = []
        mention_line = f"⚠️ Escalated to **{self.supervisor_name}**."
        if self.supervisor_upn:
            mention_tag = f"<at>{self.supervisor_name}</at>"
            mention_line = f"⚠️ Escalated to {mention_tag} for immediate manual triage."
            entities.append(
                {
                    "type": "mention",
                    "text": mention_tag,
                    "mentioned": {
                        "id": self.supervisor_upn,
                        "name": self.supervisor_name,
                    },
                }
            )

        payload: Dict[str, Any] = {
            "type": "message",
            "attachments": [
                {
                    "contentType": "application/vnd.microsoft.card.adaptive",
                    "contentUrl": None,
                    "content": {
                        "$schema": "http://adaptivecards.io/schemas/adaptive-card.json",
                        "type": "AdaptiveCard",
                        "version": "1.4",
                        "body": [
                            {
                                "type": "Container",
                                "style": "attention",
                                "bleed": True,
                                "items": [
                                    {
                                        "type": "TextBlock",
                                        "text": f"🚨 Supervisor Escalation — {incident.incident_id} ({incident.normalized_priority})",
                                        "weight": "Bolder",
                                        "size": "Medium",
                                        "wrap": True,
                                    }
                                ],
                            },
                            {
                                "type": "TextBlock",
                                "text": alert_reason,
                                "wrap": True,
                                "color": "Attention",
                                "weight": "Bolder",
                            },
                            {
                                "type": "FactSet",
                                "facts": [
                                    {"title": "Incident ID:", "value": incident.incident_id},
                                    {"title": "Priority:", "value": incident.normalized_priority},
                                    {"title": "Summary:", "value": incident.summary},
                                    {
                                        "title": "Assignment Group:",
                                        "value": incident.assignment_group or "(Blank)",
                                    },
                                ],
                            },
                            {"type": "TextBlock", "text": mention_line, "wrap": True},
                        ],
                        "actions": [
                            {
                                "type": "Action.OpenUrl",
                                "title": f"🔗 Triage {incident.incident_id} in GSM Portal",
                                "url": ticket_url,
                            }
                        ],
                        "msteams": {"width": "Full", "entities": entities},
                    },
                }
            ],
        }
        return self._post_webhook(
            url=self.supervisor_webhook_url,
            payload=payload,
            context=f"Supervisor escalation for {incident.incident_id}",
        )

    def _post_webhook(
        self, url: str, payload: Dict[str, Any], context: str
    ) -> bool:
        """POST JSON payload to MS Teams Incoming Webhook with transient error backoff."""
        if not self._is_configured_webhook(url):
            logger.warning(
                "Teams Webhook URL is unconfigured or placeholder; skipping live HTTP POST for %s.",
                context,
            )
            return False

        delay = 1.5
        for attempt in range(1, self.max_retries + 1):
            try:
                resp = requests.post(
                    url,
                    json=payload,
                    headers={"Content-Type": "application/json"},
                    timeout=self.http_timeout,
                )
                if 200 <= resp.status_code < 300:
                    logger.info("Sent Teams Adaptive Card notification for %s.", context)
                    return True

                if resp.status_code in {429, 500, 502, 503, 504} and attempt < self.max_retries:
                    logger.warning(
                        "Transient HTTP %d from Teams Webhook for %s (attempt %d/%d). Retrying in %.1fs...",
                        resp.status_code,
                        context,
                        attempt,
                        self.max_retries,
                        delay,
                    )
                    time.sleep(delay)
                    delay *= 2.0
                    continue

                logger.error(
                    "Teams Webhook failed for %s: HTTP %d — %s",
                    context,
                    resp.status_code,
                    resp.text[:300],
                )
                return False

            except requests.RequestException as exc:
                if attempt < self.max_retries:
                    logger.warning(
                        "Network error posting Teams Webhook for %s (attempt %d/%d): %s. Retrying in %.1fs...",
                        context,
                        attempt,
                        self.max_retries,
                        exc,
                        delay,
                    )
                    time.sleep(delay)
                    delay *= 2.0
                else:
                    logger.error(
                        "Exhausted retries posting Teams Webhook for %s: %s", context, exc
                    )
        return False
