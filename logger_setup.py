"""Structured console and rotating file logging setup + configuration loader."""

from __future__ import annotations

import logging
from logging.handlers import RotatingFileHandler
import os
from pathlib import Path
import sys
from typing import Any, Dict

from dotenv import load_dotenv
import yaml

LOG_FORMAT = (
    "%(asctime)s | %(levelname)-8s | %(name)-20s | %(funcName)-22s | %(message)s"
)
DATE_FORMAT = "%Y-%m-%d %H:%M:%S%z"


def setup_logging(
    log_file: str | Path = "./logs/gsm_automation.log",
    level: str = "INFO",
    max_bytes: int = 10 * 1024 * 1024,
    backup_count: int = 5,
) -> logging.Logger:
    """Configure root logger with structured console output and RotatingFileHandler."""
    log_path = Path(log_file)
    log_path.parent.mkdir(parents=True, exist_ok=True)

    numeric_level = getattr(logging, level.upper(), logging.INFO)
    root_logger = logging.getLogger("gsm_automator")
    root_logger.setLevel(numeric_level)
    root_logger.propagate = False

    # Clear existing handlers if re-initialized
    if root_logger.handlers:
        root_logger.handlers.clear()

    formatter = logging.Formatter(fmt=LOG_FORMAT, datefmt=DATE_FORMAT)

    # 1. Console Handler (UTF-8 safe)
    console_stream = sys.stdout
    if hasattr(sys.stdout, "reconfigure"):
        try:
            sys.stdout.reconfigure(encoding="utf-8")
        except Exception:  # noqa: BLE001
            pass
    console_handler = logging.StreamHandler(console_stream)
    console_handler.setLevel(numeric_level)
    console_handler.setFormatter(formatter)

    # 2. Rotating File Handler
    file_handler = RotatingFileHandler(
        filename=log_path,
        maxBytes=max_bytes,
        backupCount=backup_count,
        encoding="utf-8",
    )
    file_handler.setLevel(numeric_level)
    file_handler.setFormatter(formatter)

    root_logger.addHandler(console_handler)
    root_logger.addHandler(file_handler)

    return root_logger


def _str_to_bool(val: str) -> bool:
    return val.strip().lower() in {"1", "true", "yes", "y", "on"}


def load_config(config_path: str | Path = "config.yaml") -> Dict[str, Any]:
    """Load YAML configuration and merge environment variable overrides from .env."""
    load_dotenv(override=False)

    cfg_file = Path(config_path)
    if cfg_file.exists():
        with cfg_file.open("r", encoding="utf-8") as fh:
            config: Dict[str, Any] = yaml.safe_load(fh) or {}
    else:
        config = {}

    config.setdefault("paths", {})
    config.setdefault("logging", {})
    config.setdefault("polling", {})
    config.setdefault("excel_concurrency", {})
    config.setdefault("roster", {})
    config.setdefault("gsm_portal", {})
    config.setdefault("notifications", {})
    config["notifications"].setdefault("supervisor", {})

    # Apply environment variable overrides if present
    if os.getenv("ROSTER_FILE_PATH"):
        config["paths"]["roster_xlsx"] = os.environ["ROSTER_FILE_PATH"]
    if os.getenv("TRACKER_FILE_PATH"):
        config["paths"]["tracker_xlsx"] = os.environ["TRACKER_FILE_PATH"]
    if os.getenv("LOG_FILE_PATH"):
        config["paths"]["log_file"] = os.environ["LOG_FILE_PATH"]

    if os.getenv("GSM_PORTAL_BASE_URL"):
        config["gsm_portal"]["base_url"] = os.environ["GSM_PORTAL_BASE_URL"]
    if os.getenv("GSM_DASHBOARD_URL"):
        config["gsm_portal"]["dashboard_url"] = os.environ["GSM_DASHBOARD_URL"]
    if os.getenv("GSM_BROWSER_USER_DATA_DIR"):
        config["gsm_portal"]["user_data_dir"] = os.environ["GSM_BROWSER_USER_DATA_DIR"]
    if os.getenv("GSM_HEADLESS") is not None:
        config["gsm_portal"]["headless"] = _str_to_bool(os.environ["GSM_HEADLESS"])
    if os.getenv("GSM_MOCK_MODE") is not None:
        config["gsm_portal"]["mock_mode"] = _str_to_bool(os.environ["GSM_MOCK_MODE"])

    if os.getenv("TEAMS_INCIDENT_WEBHOOK_URL"):
        config["notifications"]["incident_webhook_url"] = os.environ[
            "TEAMS_INCIDENT_WEBHOOK_URL"
        ]
    if os.getenv("TEAMS_SUPERVISOR_WEBHOOK_URL"):
        config["notifications"]["supervisor_webhook_url"] = os.environ[
            "TEAMS_SUPERVISOR_WEBHOOK_URL"
        ]
    if os.getenv("SUPERVISOR_NAME"):
        config["notifications"]["supervisor"]["name"] = os.environ["SUPERVISOR_NAME"]
    if os.getenv("SUPERVISOR_UPN"):
        config["notifications"]["supervisor"]["email_or_upn"] = os.environ[
            "SUPERVISOR_UPN"
        ]

    return config
