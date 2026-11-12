# Anti-Gravity GSM Automator

Production-ready, modular Python automation system to scrape unassigned incidents from the **GSM Portal**, dynamically evaluate a live-editable **Excel Shift Roster (`shift_roster.xlsx`)** with multi-timezone and overnight shift support, assign incidents via **Playwright**, dispatch **Microsoft Teams Adaptive Cards** (`P1`/`P2` + Supervisor escalations), and maintain an Excel-concurrency-safe ledger (`GSM_Incident_Tracker.xlsx`).

---

## Directory Structure

```text
anti_gravity_gsm_automator/
├── .env.example                 # Environment secrets template (URLs, Webhooks, Supervisor UPN)
├── config.yaml                  # Central runtime, selector, concurrency, and routing settings
├── requirements.txt             # Locked production dependencies
├── README.md                    # Setup & operational guide
├── models.py                    # Domain dataclasses & enums (Engineer, Incident, TicketStatus)
├── logger_setup.py              # Structured console + RotatingFileHandler + .env/YAML loader
├── roster_engine.py             # Dynamic Excel shift evaluator (overnight shifts + load balancing)
├── gsm_worker.py                # Playwright GSM scraper & ticket updater (persistent SSO profile)
├── teams_notifier.py            # MS Teams Incoming Webhook Adaptive Card v1.4 sender (@mentions)
├── tracker.py                   # Deduplicated GSM_Incident_Tracker.xlsx manager with lock retries
├── main.py                      # Polling daemon orchestrator with graceful SIGINT/SIGTERM handling
├── generate_sample_roster.py    # Initializes sample data/shift_roster.xlsx with dropdown validations
├── data/                        # Runtime Excel files (shift_roster.xlsx, GSM_Incident_Tracker.xlsx)
├── logs/                        # Rotating gsm_automation.log & diagnostic failure screenshots
└── tests/
    └── test_roster_engine.py    # Pytest suite for overnight windows, load-balancing, and upserts
```

---

## Quick Start

### 1. Create Virtual Environment & Install Dependencies

```powershell
python -m venv .venv
.\.venv\Scripts\Activate.ps1
pip install -r requirements.txt
playwright install chromium
```

### 2. Configure Environment & Generate Sample Shift Roster

```powershell
Copy-Item .env.example .env
# Edit .env and config.yaml with your GSM URLs, CSS selectors, and Teams Webhook URLs

python generate_sample_roster.py
```

### 3. Dry-Run Verification (Mock GSM Mode)

Test the end-to-end roster evaluation, Teams payload building, and `GSM_Incident_Tracker.xlsx` creation safely without touching the live GSM portal:

```powershell
python main.py --mock --once
```

### 4. First-Time SSO / MFA Session Bootstrap & Production Daemon

1. Set `GSM_HEADLESS=false` in `.env` on your first run so you can complete SSO/MFA interactively; Playwright saves cookies/tokens into `./browser_profile`.
2. Start the continuous polling daemon:

```powershell
python main.py
```

### 5. Run Unit Tests

```powershell
pytest -v
```
