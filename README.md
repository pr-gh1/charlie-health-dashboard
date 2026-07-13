# Charlie Health Attendance & Billing Dashboard

Upload an Attendance & Billing export (`.xlsx`, or a delimited `.txt`/`.csv`
of the same wide-grid layout) and this app:

1. Runs it through `transform_attendance.py` -- the same parser used to build
   the Excel deliverable -- to unpivot the raw grid into Sessions / Patients
   / Rates tables.
2. Lays the KPIs out around a usage narrative rather than a flat chart grid:
   **Snapshot** (L7D/L1M/L3M trailing-window tiles -- is right now normal?),
   **Trend** (two grouped chart pairs, single axis throughout -- how did we
   get here?), **Composition** (LOS distribution, Commercial vs Medicaid --
   who are we treating?), **Narrative** (LLM synthesis + chat + change
   tracking).
3. Lets you download the same Sessions/Patients/Rates workbook the CLI
   pipeline produces.
4. Optionally asks Claude for a narrative summary of the KPIs, lets you chat
   with the real Sessions/Patients data (not just the on-screen aggregates),
   and can track what changed between uploads.

## Files

- `app.py` -- Streamlit UI: upload, filters, KPI tiles, charts, narrative tabs.
- `transform_attendance.py` -- the parsing pipeline (canonical copy; same
  file also runs standalone from the command line -- see its own docstring).
- `kpis.py` -- weekly rollups, trailing-window snapshots, and payor
  composition, all payor-filterable off the same Sessions/Patients tables.
- `chat_tools.py` -- the "Ask a question" tab's tool-use loop: Claude gets a
  fixed set of scoped query functions (`list_patients`, `get_patient_detail`,
  `weekly_trend`, `aggregate_stats`) over the real data, not free-form code
  execution.
- `snapshot_store.py` -- persists KPI snapshots to a GitHub Gist so the
  "What changed" tab can diff this upload against the last one, even though
  the app itself is otherwise stateless.
- `sample_data/sample_attendance_export.xlsx` -- the case study dataset, for
  the "Use sample Charlie Health data" button.

## Run locally

```bash
pip install -r requirements.txt
cp .streamlit/secrets.toml.example .streamlit/secrets.toml
# edit .streamlit/secrets.toml -- see "Secrets" below for what each one enables
streamlit run app.py
```

Every secret is optional and the app degrades gracefully without it -- see
below for exactly what each one turns on.

## Secrets

| Secret | Enables | Without it |
|---|---|---|
| `ANTHROPIC_API_KEY` | KPI narrative, chat, change-summary narrative | Those sections show a message explaining they're disabled; everything else still works |
| `ANTHROPIC_MODEL` | Optional override, defaults to `claude-sonnet-5` | Uses the default |
| `GITHUB_TOKEN` + `GIST_ID` | "What changed" tab's cross-upload diffing | Tab shows a setup message; rest of the app unaffected |

### Setting up the Gist (for "What changed")

1. Create a private Gist at [gist.github.com](https://gist.github.com) with
   one file, e.g. `charlie_health_snapshot.json`, containing `{}`. Copy the
   Gist ID from its URL (`https://gist.github.com/<user>/<GIST_ID>`).
2. Create a
   [fine-grained personal access token](https://github.com/settings/tokens?type=beta)
   scoped only to **Gists: read and write**.
3. Add both as secrets (`GITHUB_TOKEN`, `GIST_ID`).

The app never auto-overwrites this snapshot -- the "What changed" tab always
shows you the diff first, and only saves a new baseline when you click
**Save this as the new snapshot**.

## Deploy to a public URL (Streamlit Community Cloud, free)

1. Push this folder to a GitHub repo (public or private both work).
2. Go to [share.streamlit.io](https://share.streamlit.io), sign in with
   GitHub.
3. Click **New app**, pick the repo/branch, and set the main file to
   `app.py`.
4. Open **Advanced settings -> Secrets** and paste the contents of your
   local `.streamlit/secrets.toml`.
5. Click **Deploy**. You'll get a `https://<something>.streamlit.app` URL --
   that's the public link for the case study's Part 4 submission.

Redeploys happen automatically on every push to the connected branch.

## Design notes

- `transform_attendance.py` finds the rate block and patient rows by
  scanning for labels ("Rates", "Patient info"), not fixed row numbers, so
  it keeps working if a future export gains/loses patients or shifts a row.
  It accepts `.xlsx` or delimited `.txt`/`.csv` of the same layout --
  verified to produce byte-identical output from both formats.
- Every chart is single-axis. Two series that share a unit (e.g. IOP vs OPT
  attendance rate) can share a chart; two different units (e.g. patient
  count and revenue) get separate charts instead of a dual-axis overlay,
  which is easy to misread as correlation that isn't there.
- The OPT attendance-rate line is styled distinctly (dashed, amber, larger
  markers) rather than just a different color, because OPT runs far fewer
  weekly sessions than IOP -- its rate is a noisier signal from a much
  smaller sample, and the styling is meant to say "read this one more
  loosely," not just "this is a different series."
- The chat's tools are a fixed, scoped set of pandas queries, not arbitrary
  code execution -- a question (or anything hidden in uploaded data) can
  only ever trigger one of the four defined functions.
- The narrative/chat/change-summary LLM calls are all on-demand
  (button-triggered), not run automatically on every page load, to avoid
  unnecessary API cost on each rerun.
- The "What changed" diff itself is free and automatic (pure Python
  comparison); only turning that diff into a written explanation costs an
  API call, and that step is separately gated behind its own button.
- The exact JSON sent to the model is always shown in an expander, so any
  narrative's numbers can be checked against what was actually sent.
