"""
Charlie Health -- Attendance & Billing Dashboard

Upload the wide-format Attendance & Billing export (.xlsx, or a delimited
.txt/.csv of the same layout) and this app runs it through the same
transform_attendance.py pipeline used to build the Excel deliverable,
computes the Part 1 KPIs, and lays them out around a specific usage
narrative rather than a flat grid of every possible chart:

  1. Snapshot  -- is right now normal? (L7D/L1M/L3M KPI tiles)
  2. Trend     -- how did we get here? (census/attendance, revenue/billing,
     and LOS-by-discharge-cohort, single axis throughout)
  3. Composition -- who are we treating? (LOS distribution, payor mix)
  4. Narrative -- LLM synthesis of the above, plus a chat that can query
     the real data (not just the aggregates already on screen), plus a
     diff-based "what changed since last snapshot" section.

Nothing here recomputes the parsing logic -- transform_attendance.py is the
single source of truth for "how do we turn the raw grid into Sessions /
Patients / Rates," whether it's run from the command line or from this app.
"""

import io
import json
import os
import tempfile
from datetime import datetime, timezone
from pathlib import Path

import plotly.graph_objects as go
import streamlit as st

from transform_attendance import load_grid, parse_grid, write_workbook, RAW_SHEET_DEFAULT
from kpis import (build_frames, weekly_rollup, summary_stats, trailing_window_stats,
                   payor_composition, los_by_discharge_month, active_patient_stats,
                   WINDOWS, DEFAULT_INACTIVITY_DAYS)
from chat_tools import answer_question
from snapshot_store import load_last_snapshot, save_snapshot, compute_diff, llm_change_summary

st.set_page_config(page_title="Charlie Health Attendance Dashboard", layout="wide")


# ---------------------------------------------------------------------------
# Secrets / config
# ---------------------------------------------------------------------------

def _secret(name, default=None):
    try:
        return st.secrets.get(name, os.environ.get(name, default))
    except Exception:
        return os.environ.get(name, default)


def get_api_key():
    return _secret("ANTHROPIC_API_KEY")


def get_model_name():
    return _secret("ANTHROPIC_MODEL", "claude-sonnet-5")


def get_github_secrets():
    return _secret("GITHUB_TOKEN"), _secret("GIST_ID")


# ---------------------------------------------------------------------------
# Data loading (unchanged pipeline, reused as-is)
# ---------------------------------------------------------------------------

@st.cache_data(show_spinner="Parsing attendance export...")
def run_pipeline(file_bytes: bytes, filename: str):
    suffix = Path(filename).suffix or ".xlsx"
    with tempfile.NamedTemporaryFile(suffix=suffix, delete=False) as tmp:
        tmp.write(file_bytes)
        tmp_path = tmp.name
    try:
        grid, source_wb = load_grid(tmp_path, RAW_SHEET_DEFAULT)
        rates, sessions, patients = parse_grid(grid)
    finally:
        os.unlink(tmp_path)
    return rates, sessions, patients, (source_wb is not None)


def build_output_workbook_bytes(rates, sessions, patients, file_bytes, filename, had_raw_sheet):
    source_wb = None
    if had_raw_sheet:
        suffix = Path(filename).suffix or ".xlsx"
        with tempfile.NamedTemporaryFile(suffix=suffix, delete=False) as tmp:
            tmp.write(file_bytes)
            tmp_path = tmp.name
        import openpyxl
        source_wb = openpyxl.load_workbook(tmp_path, data_only=True)
        os.unlink(tmp_path)

    buf = io.BytesIO()
    write_workbook(buf, rates, sessions, patients, source_wb, RAW_SHEET_DEFAULT)
    buf.seek(0)
    return buf


@st.cache_data(show_spinner="Checking for a previous snapshot...")
def cached_previous_snapshot(file_bytes: bytes, filename: str, gist_id: str):
    # Cache key includes file identity so this only hits the network once
    # per distinct upload, not on every widget interaction rerun.
    token, gid = get_github_secrets()
    return load_last_snapshot(token, gid)


# ---------------------------------------------------------------------------
# Chart helper -- single y-axis always, per design system guidance
# ---------------------------------------------------------------------------

def line_chart(df, x_col, series, y_title, pct=False):
    """series: list of (col, name, color, dash, mode)"""
    fig = go.Figure()
    for col, name, color, dash, mode in series:
        fig.add_trace(go.Scatter(
            x=df[x_col], y=df[col], mode=mode, name=name,
            line=dict(color=color, dash=dash, width=2),
            marker=dict(size=6) if "markers" in mode else None,
        ))
    fig.update_layout(
        yaxis=dict(title=y_title, tickformat=".0%" if pct else None),
        legend=dict(orientation="h") if len(series) > 1 else dict(visible=False),
        margin=dict(t=10, l=10, r=10, b=10),
        height=260,
    )
    return fig


def bar_chart(df, x_col, y_col, y_title, color="#2a78d6"):
    fig = go.Figure()
    fig.add_trace(go.Bar(x=df[x_col], y=df[y_col], marker_color=color))
    fig.update_layout(yaxis=dict(title=y_title), margin=dict(t=10, l=10, r=10, b=10), height=260)
    return fig


# ---------------------------------------------------------------------------
# Sidebar: data source + global filter
# ---------------------------------------------------------------------------

st.sidebar.title("Data source")
uploaded = st.sidebar.file_uploader(
    "Upload Attendance & Billing export", type=["xlsx", "xlsm", "txt", "csv", "tsv"]
)
use_sample = st.sidebar.button("Use sample Charlie Health data")

if "active_file" not in st.session_state:
    st.session_state.active_file = None

if uploaded is not None:
    st.session_state.active_file = (uploaded.getvalue(), uploaded.name)
elif use_sample:
    sample_path = Path(__file__).parent / "sample_data" / "sample_attendance_export.xlsx"
    st.session_state.active_file = (sample_path.read_bytes(), sample_path.name)

if st.session_state.active_file is None:
    st.title("Charlie Health -- Attendance & Billing Dashboard")
    st.info(
        "Upload an Attendance & Billing export in the sidebar (.xlsx, or a "
        "delimited .txt/.csv of the same layout), or click **Use sample "
        "Charlie Health data** to see it with the case study dataset."
    )
    st.stop()

file_bytes, filename = st.session_state.active_file

try:
    rates, sessions, patients, had_raw_sheet = run_pipeline(file_bytes, filename)
except Exception as e:
    st.error(f"Couldn't parse {filename}: {e}")
    st.stop()

rates_df, sessions_df, patients_df = build_frames(rates, sessions, patients)
weekly_df_all = weekly_rollup(sessions_df, patients_df)
stats = summary_stats(sessions_df, patients_df, weekly_df_all)
stats["captured_at"] = datetime.now(timezone.utc).isoformat()

st.sidebar.success(f"Loaded {filename}")
st.sidebar.caption(
    f"{stats['totals']['patients']} patients · "
    f"{stats['totals']['scheduled_sessions']} scheduled sessions · "
    f"{stats['date_range']['start']} to {stats['date_range']['end']}"
)

xlsx_buf = build_output_workbook_bytes(rates, sessions, patients, file_bytes, filename, had_raw_sheet)
st.sidebar.download_button(
    "Download Sessions/Patients/Rates workbook",
    data=xlsx_buf,
    file_name="attendance_billing_transformed.xlsx",
    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
)

st.sidebar.divider()
payor_filter = st.sidebar.radio("Payor", ["All", "Commercial", "Medicaid"], horizontal=True)
weekly_df = weekly_rollup(sessions_df, patients_df, payor=None if payor_filter == "All" else payor_filter)


# ---------------------------------------------------------------------------
# Header + usage narrative
# ---------------------------------------------------------------------------

st.title("Charlie Health -- Attendance & Billing Dashboard")
with st.expander("How to use this dashboard", expanded=False):
    st.markdown(
        "**Snapshot** (below) answers *is right now normal* -- pick a "
        "trailing window (L7D/L1M/L3M) and check the headline numbers "
        "before digging further. **Trend** answers *how did we get here* "
        "-- two chart pairs, census/attendance and revenue/billing, each "
        "on a single axis so nothing is visually distorted by a second "
        "scale. **Composition** answers *who are we treating* -- length "
        "of stay and payor mix, since payor is close to a 1.5x swing in "
        "billed rate. **Narrative** (bottom) synthesizes the above and "
        "lets you ask follow-up questions the charts don't answer directly. "
        "Use the **Payor** filter in the sidebar to cut the Trend and "
        "Snapshot sections to Commercial or Medicaid only."
    )

st.divider()


# ---------------------------------------------------------------------------
# 1. Snapshot
# ---------------------------------------------------------------------------

st.subheader("Snapshot")
window_label = st.radio("Window", list(WINDOWS.keys()), index=0, horizontal=True, label_visibility="collapsed")
win = trailing_window_stats(
    sessions_df, patients_df, WINDOWS[window_label],
    payor=None if payor_filter == "All" else payor_filter,
)

c1, c2, c3, c4, c5, c6 = st.columns(6)
c1.metric("Patients in treatment", win["patients_in_treatment"])
c2.metric("Patients attending IOP", win["iop_patients"])
c3.metric("New admissions", win["new_admissions"])
c4.metric("IOP attendance rate", f"{win['iop_attendance_rate']:.1%}" if win["iop_attendance_rate"] is not None else "—")
c5.metric("Billable IOP sessions", win["billable_iop_sessions"])
c6.metric("Avg daily revenue", f"${win['avg_daily_revenue']:,.0f}")
st.caption(f"{window_label} · {win['start']} to {win['end']} · {payor_filter} payor")

st.divider()


# ---------------------------------------------------------------------------
# 2. Trend
# ---------------------------------------------------------------------------

st.subheader("Trend")

st.markdown("**Census & attendance** -- is the clinical engine healthy?")
col_a, col_b = st.columns(2)
with col_a:
    fig = line_chart(weekly_df, "week", [
        ("patients_in_treatment", "In treatment (any service)", "#2a78d6", None, "lines"),
        ("iop_patients", "Attending IOP", "#5dcaa5", "dash", "lines"),
    ], "Patients")
    st.plotly_chart(fig, width="stretch")
    st.caption(
        "Any gap between the two lines is patients active on OPT only that "
        "week -- e.g. IOP alumni stepping down, or someone between IOP "
        "blocks."
    )
with col_b:
    fig = line_chart(weekly_df, "week", [
        ("iop_attendance_rate", "IOP", "#2a78d6", None, "lines"),
        ("opt_attendance_rate", "OPT", "#eda100", "dash", "lines+markers"),
    ], "Attendance rate", pct=True)
    st.plotly_chart(fig, width="stretch")
    st.caption(
        "OPT runs far fewer sessions/week than IOP, so its rate (amber, "
        "dashed, markers) is noisier -- read it as a rougher signal, not "
        "a precise trend the way the IOP line is."
    )

st.markdown("**Revenue & billing** -- is that engine converting to dollars?")
col_c, col_d = st.columns(2)
with col_c:
    fig = bar_chart(weekly_df, "week", "iop_attended", "Billable IOP sessions")
    st.plotly_chart(fig, width="stretch")
with col_d:
    fig = line_chart(weekly_df, "week", [("revenue", "Revenue", "#1baf7a", None, "lines")], "Revenue ($)")
    st.plotly_chart(fig, width="stretch")

st.markdown("**Length of stay** -- are completed episodes getting longer or shorter?")
los_payor = None if payor_filter == "All" else payor_filter
los_trend = los_by_discharge_month(sessions_df, patients_df, payor=los_payor)
active_stats = active_patient_stats(sessions_df, patients_df, payor=los_payor)
if los_trend.empty:
    st.caption("Not enough completed episodes yet to trend LOS.")
else:
    fig = bar_chart(los_trend, "discharge_month", "avg_los_weeks", "Avg LOS at discharge (weeks)", color="#7f77dd")
    st.plotly_chart(fig, width="stretch")
    st.caption(
        f"Completed episodes only (last session {DEFAULT_INACTIVITY_DAYS}+ days before the data's most "
        f"recent date) -- a patient still actively attending doesn't have a "
        f"finished LOS yet, so including them would bias recent months down. "
        f"Separately: {active_stats['count']} patients are currently active"
        + (f", averaging {active_stats['avg_tenure_weeks']} weeks so far (not "
           f"included in the chart above)." if active_stats["count"] else ".")
    )

st.divider()


# ---------------------------------------------------------------------------
# 3. Composition
# ---------------------------------------------------------------------------

st.subheader("Composition")
col_e, col_f = st.columns(2)
with col_e:
    st.markdown("**Length of stay distribution**")
    fig = go.Figure()
    fig.add_trace(go.Histogram(x=patients_df["los_weeks"], nbinsx=20, marker_color="#2a78d6"))
    fig.add_vline(x=10, line_dash="dash", line_color="gray", annotation_text="~10wk typical program")
    fig.update_layout(xaxis_title="LOS (weeks)", yaxis_title="Patients", height=260, margin=dict(t=10, l=10, r=10, b=10))
    st.plotly_chart(fig, width="stretch")
with col_f:
    st.markdown("**Payor mix** -- close to a 1.5x spread in billed rate")
    comp_df = payor_composition(sessions_df, patients_df)
    st.dataframe(
        comp_df.rename(columns={
            "payor": "Payor", "patients": "Patients", "avg_los_weeks": "Avg LOS (wks)",
            "attendance_rate": "Attendance rate", "total_revenue": "Total revenue",
            "avg_billed_iop_rate": "Avg billed IOP rate",
        }),
        width="stretch", hide_index=True,
    )

with st.expander("Weekly KPI table"):
    show_cols = ["week", "patients_in_treatment", "new_admissions",
                 "iop_scheduled", "iop_attended", "iop_attendance_rate",
                 "opt_scheduled", "opt_attended", "opt_attendance_rate",
                 "avg_billed_iop_rate", "revenue", "avg_daily_revenue"]
    st.dataframe(weekly_df[show_cols], width="stretch")

with st.expander("Patient-level table"):
    st.dataframe(patients_df, width="stretch")

st.divider()


# ---------------------------------------------------------------------------
# 4. Narrative: KPI summary, what-changed diff, and chat
# ---------------------------------------------------------------------------

st.subheader("Narrative")

api_key = get_api_key()
client = None
if api_key:
    from anthropic import Anthropic
    client = Anthropic(api_key=api_key)
else:
    st.warning(
        "No ANTHROPIC_API_KEY configured, so narrative generation and chat "
        "are disabled. Add it as a Streamlit secret to enable this section "
        "-- see README.md."
    )

tab_summary, tab_changed, tab_chat = st.tabs(["KPI summary", "What changed", "Ask a question"])

with tab_summary:
    if client:
        if st.button("Generate insights"):
            with st.spinner("Asking Claude for a summary..."):
                try:
                    prompt = (
                        "You are a clinical operations analyst summarizing weekly "
                        "KPI trends for Charlie Health leadership. Using ONLY the "
                        "JSON data below -- do not invent numbers that aren't in "
                        "it -- write 4-6 concise bullet points covering attendance "
                        "trends, revenue trends, payor mix, and anything that looks "
                        "like a risk or an opportunity. Cite the specific figures "
                        "you reference.\n\n" + json.dumps(stats, indent=2, default=str)
                    )
                    resp = client.messages.create(model=get_model_name(), max_tokens=600,
                                                    messages=[{"role": "user", "content": prompt}])
                    st.session_state["insights_text"] = resp.content[0].text
                except Exception as e:
                    st.error(f"Insight generation failed: {e}")
        if "insights_text" in st.session_state:
            st.markdown(st.session_state["insights_text"])
        else:
            st.caption("Click **Generate insights** to summarize the KPIs above.")
        with st.expander("Exact data sent to the model"):
            st.json(stats)

with tab_changed:
    gh_token, gist_id = get_github_secrets()
    if not gh_token or not gist_id:
        st.info(
            "Not configured -- add GITHUB_TOKEN and GIST_ID as Streamlit "
            "secrets to enable change tracking across uploads. See README.md."
        )
    else:
        previous_stats = cached_previous_snapshot(file_bytes, filename, gist_id)
        diff = compute_diff(previous_stats, stats)
        if diff is None:
            st.info("No previous snapshot found yet -- this will be the baseline.")
        else:
            st.caption(f"Comparing against snapshot captured {diff['previous_captured_at']}")
            totals_changes = diff["changes"]["totals"]
            cols = st.columns(len(totals_changes))
            for col, (k, v) in zip(cols, totals_changes.items()):
                col.metric(k.replace("_", " "), v["current"], delta=v["delta"])
            with st.expander("Full diff"):
                st.json(diff)
            if client and st.button("Explain what changed"):
                with st.spinner("Asking Claude to interpret the changes..."):
                    try:
                        st.session_state["change_summary"] = llm_change_summary(client, get_model_name(), diff, stats)
                    except Exception as e:
                        st.error(f"Change summary failed: {e}")
            if "change_summary" in st.session_state:
                st.markdown(st.session_state["change_summary"])

        if st.button("Save this as the new snapshot"):
            try:
                save_snapshot(gh_token, gist_id, stats)
                st.success("Snapshot saved -- future uploads will diff against this one.")
                cached_previous_snapshot.clear()
            except Exception as e:
                st.error(f"Couldn't save snapshot: {e}")

with tab_chat:
    if not client:
        st.caption("Configure ANTHROPIC_API_KEY to enable chat.")
    else:
        if "chat_history" not in st.session_state:
            st.session_state.chat_history = []
        if "chat_display" not in st.session_state:
            st.session_state.chat_display = []

        for role, text in st.session_state.chat_display:
            with st.chat_message(role):
                st.markdown(text)

        question = st.chat_input("Ask about the data, e.g. 'who are the 10 worst-attendance Commercial patients?'")
        if question:
            st.session_state.chat_display.append(("user", question))
            with st.chat_message("user"):
                st.markdown(question)
            with st.chat_message("assistant"):
                with st.spinner("Querying the data..."):
                    answer, new_history = answer_question(
                        client, get_model_name(), question, sessions_df, patients_df,
                        history=st.session_state.chat_history,
                    )
                    st.markdown(answer)
            st.session_state.chat_history = new_history
            st.session_state.chat_display.append(("assistant", answer))
