# Dental Production Dashboard

A single-file Streamlit dashboard for a dental center's insurance-claims / production export
(`.xlsx` or `.xls`). It shows gross and net production, chair yield and utilization, provider
performance, service mix by procedure category, and revenue-cycle figures (payers, denials,
pending insurance).

## Privacy

Patient names, chart numbers, member/policy IDs, claim and authorization IDs are dropped or
one-way hashed (with a random per-session salt) right after the file is loaded. They are never
displayed, and error messages are scrubbed of patient identifiers.

When hosted, uploaded files are still processed on the hosting server. Make the deployed app
**private** and share it only with people allowed to see clinic data.

## Run locally (Windows)

Double-click `run_dashboard.bat`, or:

```bat
set DENTAL_DASHBOARD_LOCAL=1
python -m pip install -r requirements.txt
python -m streamlit run app.py
```

Local mode adds a "file path" box (defaulting to `Downloads\Production.xlsx`) and lets you save
the chair map to `chair_map.json` next to `app.py`.

## Deploy (Streamlit Community Cloud)

1. New app → this repo, branch `main`, main file `app.py`.
2. Advanced settings → Python **3.12** (3.11 is the minimum for pandas 3).
3. After deploying, set the app to private and invite viewers.

Do **not** set `DENTAL_DASHBOARD_LOCAL` on a hosted deployment: it would let visitors read files
from the server's disk. Hosted, the app accepts uploads only, and the chair map is kept for the
browser session; use **Download chair map** and **Load saved chair map** to reuse it.

## Expected columns

The header row is detected automatically. Required (any one name per field):

| Field | Accepted column names |
|---|---|
| Date | Date, Appointment Date, Transaction Date, Service Date |
| Provider | Completed Doctor, Claim Clinician, Provider, Clinician, Doctor |
| Procedure code | Activity Code, Procedure Code, USCLS Code, ADA Code, Code |
| Gross | Gross, Gross Fees, Gross Production |

Optional: Start Time / End Time (enables chair hours and the occupancy heatmap), Operatory /
Chair, Activity Price, Patient Share, Payment Amount, Rejected Amount, Date Settlement,
Insurance Company, Status, Denial Reason, Chart #, Encounter ID, Patient Age.
