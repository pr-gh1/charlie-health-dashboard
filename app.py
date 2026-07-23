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

import pandas as pd
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


@st.cache_data(show_spinner=False)
def load_forecast_baseline():
    """Frozen N24M forecast (base case + go-big case), generated once from
    N24M_Revenue_Projection.xlsx and checked into the repo -- see
    extract_n24m_baseline.py. In production this file would be regenerated
    by Finance on a regular cadence, not on every dashboard load; the
    dashboard's job is to compare actuals against whatever baseline is
    currently checked in, not to re-run the forecast itself.
    """
    path = Path(__file__).parent / "forecast" / "n24m_baseline.json"
    if not path.exists():
        return None
    return json.loads(path.read_text())


def forecast_window_estimate(forecast, scenario_key, metric_key, start, end):
    """Day-weighted proration of a monthly forecast baseline onto an
    arbitrary [start, end] window, so a trailing L7D/L1M/L3M snapshot window
    (which rarely lines up with calendar month boundaries) can still be
    compared against the monthly N24M model. `census` is an intensity
    metric so overlapping days are averaged (weighted by days contributed);
    `admits`/`total_revenue` are flow metrics so each month's total is
    prorated by the fraction of that month falling inside the window, then
    summed across every month the window touches. Returns None if the
    window falls entirely outside the forecast horizon (e.g. the base
    sample dataset, which ends the day before the forecast starts).
    """
    scenario = forecast[scenario_key]
    months = forecast["months"]
    is_intensity = metric_key == "census"
    weighted_sum, weight_total, total, any_data = 0.0, 0, 0.0, False

    for i, m in enumerate(months):
        m_start = pd.Timestamp(m + "-01")
        m_end = m_start + pd.offsets.MonthEnd(1)
        overlap_start, overlap_end = max(start, m_start), min(end, m_end)
        if overlap_start > overlap_end:
            continue
        days = (overlap_end - overlap_start).days + 1
        val = scenario.get(metric_key, [None] * len(months))[i]
        if val is None:
            continue
        any_data = True
        if is_intensity:
            weighted_sum += val * days
            weight_total += days
        else:
            days_in_month = (m_end - m_start).days + 1
            total += val * (days / days_in_month)

    if not any_data:
        return None
    return weighted_sum / weight_total if is_intensity else total


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

DATA_DIR = Path(__file__).parent
PRESETS = {
    "Base data (Mar 2021, no forecast overlap)": DATA_DIR / "sample_data" / "sample_attendance_export.xlsx",
    "September (+6 months vs. forecast)": DATA_DIR / "test_data" / "sample_attendance_export_v3_september.xlsx",
}

st.sidebar.title("Data source")
st.sidebar.caption(
    "In production these would auto-refresh from a blob store as new "
    "exports land, rather than being picked here -- these two are "
    "checked-in snapshots standing in for that: today's data, and six "
    "months out, so you can see how actuals track against the N24M "
    "forecast partway through the projection window."
)
preset_choice = st.sidebar.radio("Preset dataset", list(PRESETS.keys()), index=0)
uploaded = st.sidebar.file_uploader(
    "...or upload your own Attendance & Billing export",
    type=["xlsx", "xlsm", "txt", "csv", "tsv"],
)

# Precedence: an explicit upload always wins (and keeps winning across
# reruns, since Streamlit's file_uploader holds the file until it's
# cleared); otherwise whichever preset is selected in the radio above
# drives the dashboard. This is also what makes the app usable with zero
# clicks -- the radio defaults to Base data, so first load just works.
if uploaded is not None:
    st.session_state.active_file = (uploaded.getvalue(), uploaded.name)
else:
    preset_path = PRESETS[preset_choice]
    st.session_state.active_file = (preset_path.read_bytes(), preset_path.name)

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
snap_payor = None if payor_filter == "All" else payor_filter
window_days = WINDOWS[window_label]
win = trailing_window_stats(sessions_df, patients_df, window_days, payor=snap_payor)

# Prior period = the immediately preceding window of the same length
# (non-overlapping), so L7D compares to the L7D right before it, L1M to
# the L1M right before that, etc. -- a like-for-like "is this normal"
# reference that doesn't require a second dataset to be loaded.
prev_end = pd.Timestamp(win["start"]) - pd.Timedelta(days=1)
prev = trailing_window_stats(sessions_df, patients_df, window_days, payor=snap_payor, as_of=prev_end)

forecast = load_forecast_baseline()
win_start_ts, win_end_ts = pd.Timestamp(win["start"]), pd.Timestamp(win["end"])
fc_census = fc_admits = fc_revenue = None
if forecast is not None:
    fc_census = forecast_window_estimate(forecast, "base_case", "census", win_start_ts, win_end_ts)
    fc_admits = forecast_window_estimate(forecast, "base_case", "admits", win_start_ts, win_end_ts)
    fc_revenue = forecast_window_estimate(forecast, "base_case", "total_revenue", win_start_ts, win_end_ts)
fc_avg_daily_revenue = fc_revenue / window_days if fc_revenue is not None else None

def _delta(cur, prior, pct=False, pp=False):
    if cur is None or prior is None:
        return None
    d = cur - prior
    if pp:
        return f"{d * 100:+.1f}pp"
    if pct and prior:
        return f"{d:+,.0f} ({d / prior:+.0%})"
    return f"{d:+,.0f}"

st.caption("**vs. prior period** -- same-length window immediately before this one")
c1, c2, c3, c4, c5, c6 = st.columns(6)
c1.metric("Patients in treatment", win["patients_in_treatment"],
          delta=_delta(win["patients_in_treatment"], prev["patients_in_treatment"]))
c2.metric("Patients attending IOP", win["iop_patients"],
          delta=_delta(win["iop_patients"], prev["iop_patients"]))
c3.metric("New admissions", win["new_admissions"],
          delta=_delta(win["new_admissions"], prev["new_admissions"]))
c4.metric("IOP attendance rate",
          f"{win['iop_attendance_rate']:.1%}" if win["iop_attendance_rate"] is not None else "—",
          delta=_delta(win["iop_attendance_rate"], prev["iop_attendance_rate"], pp=True))
c5.metric("Billable IOP sessions", win["billable_iop_sessions"],
          delta=_delta(win["billable_iop_sessions"], prev["billable_iop_sessions"]))
c6.metric("Avg daily revenue", f"${win['avg_daily_revenue']:,.0f}",
          delta=_delta(win["avg_daily_revenue"], prev["avg_daily_revenue"]))
st.caption(f"{window_label} · {win['start']} to {win['end']} · {payor_filter} payor")

st.caption(
    "**vs. base-case forecast** -- N24M model expectation for this same window "
    "(revenue-only scope, so it only covers the metrics the model actually "
    "projects)"
)
if forecast is None:
    st.caption("No forecast baseline found.")
elif fc_census is None:
    st.caption(
        "This window falls before the forecast horizon starts "
        f"({forecast['months'][0]}) -- switch to the September preset "
        "(sidebar) to see a window that overlaps it."
    )
else:
    d1, d2, d3 = st.columns(3)
    d1.metric("Patients in treatment", win["patients_in_treatment"],
              delta=_delta(win["patients_in_treatment"], fc_census))
    d2.metric("New admissions", win["new_admissions"],
              delta=_delta(win["new_admissions"], fc_admits))
    d3.metric("Avg daily revenue", f"${win['avg_daily_revenue']:,.0f}",
              delta=_delta(win["avg_daily_revenue"], fc_avg_daily_revenue))
    st.caption(
        "Patients attending IOP, IOP attendance rate, and billable IOP "
        "sessions have no counterpart in the N24M model (it projects "
        "census and revenue, not attendance behavior) -- the prior-period "
        "row above is the trend read for those."
    )

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
        "Individual rate: each patient's own attended/scheduled ratio that "
        "week, averaged across patients -- not a pooled ratio of total "
        "attended over total scheduled sessions. OPT runs far fewer "
        "sessions/week than IOP, so its rate (amber, dashed, markers) is "
        "noisier -- read it as a rougher signal, not a precise trend the "
        "way the IOP line is."
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
    fig = bar_chart(los_trend, "discharge_month", "avg_los_appointments", "Avg LOS at discharge (appointments attended)", color="#7f77dd")
    st.plotly_chart(fig, width="stretch")
    st.caption(
        f"LOS = number of appointments ATTENDED per patient stay -- the "
        f"billing-relevant measure of treatment received, not a calendar-time "
        f"span (which conflates episode length with scheduling gaps). "
        f"Completed episodes only (last attended session "
        f"{DEFAULT_INACTIVITY_DAYS}+ days before the data's most "
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
    fig.add_trace(go.Histogram(x=patients_df["los_appointments"], nbinsx=20, marker_color="#2a78d6"))
    median_appts = patients_df["los_appointments"].median()
    fig.add_vline(x=median_appts, line_dash="dash", line_color="gray",
                  annotation_text=f"median: {median_appts:.0f} appts")
    fig.update_layout(xaxis_title="LOS (appointments attended)", yaxis_title="Patients", height=260, margin=dict(t=10, l=10, r=10, b=10))
    st.plotly_chart(fig, width="stretch")
with col_f:
    st.markdown("**Payor mix** -- close to a 1.5x spread in billed rate")
    comp_df = payor_composition(sessions_df, patients_df)
    st.dataframe(
        comp_df.rename(columns={
            "payor": "Payor", "patients": "Patients", "avg_los_appointments": "Avg LOS (appts)",
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
# 4. Forecast: actuals vs. the N24M revenue projection
# ---------------------------------------------------------------------------

st.subheader("Forecast")

# `forecast` already loaded above, in Snapshot -- reused here as-is.
if forecast is None:
    st.info(
        "No forecast baseline found (forecast/n24m_baseline.json). This "
        "section compares actuals against the N24M revenue projection "
        "model once that file is present."
    )
else:
    st.caption(
        f"Baseline: N24M_Revenue_Projection.xlsx, frozen as of "
        f"{forecast['baseline_cutoff']} and checked into the repo -- not "
        f"re-forecast on every load. In production this file would be "
        f"regenerated by Finance on a regular cadence (e.g. monthly) as "
        f"actuals land in blob storage, the same way the underlying "
        f"attendance export would auto-refresh rather than requiring a "
        f"manual upload. This view is always blended across payors, "
        f"regardless of the Payor filter above, since the forecast itself "
        f"doesn't split census/revenue by payor at the output level."
    )

    # Weekly grain throughout -- matches the Trend section's cadence and
    # gives a much more legible variance read than a 24-point monthly bar
    # chart would (a single bad/good week doesn't get smeared across a
    # whole month before it's visible).
    fc_months = forecast["months"]
    fc_start = pd.Timestamp(fc_months[0] + "-01")
    fc_end = pd.Timestamp(fc_months[-1] + "-01") + pd.offsets.MonthEnd(1)
    first_monday = fc_start - pd.Timedelta(days=fc_start.weekday())
    fc_weeks = pd.date_range(first_monday, fc_end, freq="7D")

    fc_rows = []
    for wk in fc_weeks:
        wk_end = wk + pd.Timedelta(days=6)
        fc_rows.append({
            "week": wk,
            "base_revenue": forecast_window_estimate(forecast, "base_case", "total_revenue", wk, wk_end),
            "growth_revenue": forecast_window_estimate(forecast, "growth_case", "total_revenue", wk, wk_end),
            "base_census": forecast_window_estimate(forecast, "base_case", "census", wk, wk_end),
            "growth_census": forecast_window_estimate(forecast, "growth_case", "census", wk, wk_end),
        })
    plot_df = pd.DataFrame(fc_rows)
    plot_df = plot_df.merge(
        weekly_df_all[["week", "revenue", "patients_in_treatment"]].rename(
            columns={"revenue": "actual_revenue", "patients_in_treatment": "actual_census"}
        ),
        on="week", how="left",
    )
    plot_df["variance_revenue"] = plot_df["actual_revenue"] - plot_df["base_revenue"]

    n_actual = plot_df["actual_revenue"].notna().sum()
    if n_actual == 0:
        st.caption(
            "No actuals fall inside the forecast window (Apr 2021 onward) "
            "yet -- the base sample dataset runs Aug 2020-Mar 2021, right "
            "up to the forecast's start. Switch to the September preset "
            "(sidebar) to see actuals plotted against the forecast six "
            "months into the projection period."
        )

    col_fc1, col_fc2 = st.columns(2)
    with col_fc1:
        st.markdown("**Revenue -- actual vs. forecast, weekly**")
        fig = line_chart(plot_df, "week", [
            ("actual_revenue", "Actual", "#1baf7a", None, "lines+markers"),
            ("base_revenue", "Base case", "#2a78d6", "dash", "lines"),
            ("growth_revenue", "Go-big case (staged hiring)", "#7f77dd", "dot", "lines"),
        ], "Revenue ($)")
        st.plotly_chart(fig, width="stretch")
    with col_fc2:
        st.markdown("**Census -- actual vs. forecast, weekly**")
        fig = line_chart(plot_df, "week", [
            ("actual_census", "Actual", "#1baf7a", None, "lines+markers"),
            ("base_census", "Base case", "#2a78d6", "dash", "lines"),
            ("growth_census", "Go-big case (staged hiring)", "#7f77dd", "dot", "lines"),
        ], "Patients in treatment")
        st.plotly_chart(fig, width="stretch")

    if n_actual > 0:
        st.markdown("**Revenue variance vs. base case, weekly** -- actual minus base-case forecast")
        variance_df = plot_df[plot_df["actual_revenue"].notna()]
        fig = bar_chart(
            variance_df, "week", "variance_revenue", "Variance ($)",
            color="#1baf7a",
        )
        st.plotly_chart(fig, width="stretch")
        st.caption(
            "Positive = ahead of the base-case forecast; negative = behind "
            "it. Weekly base-case revenue here is the monthly N24M figure "
            "prorated by day-count, not a separately modeled weekly number "
            "-- so a short first/last week of a month shows a smaller "
            "forecast bar, by design. Compared against base case (not "
            "go-big case) since the rep-headcount increase behind the "
            "go-big case hasn't actually been executed against yet -- base "
            "case is the fairer bar for 'are we tracking as expected' "
            "until that changes."
        )

st.divider()


# ---------------------------------------------------------------------------
# 5. Narrative: KPI summary, what-changed diff, and chat
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
                    # Filter to text blocks -- some models can return non-text
                    # blocks (e.g. thinking) ahead of the actual answer, and
                    # content[0] isn't reliably the text block.
                    st.session_state["insights_text"] = "".join(
                        b.text for b in resp.content if getattr(b, "type", None) == "text"
                    )
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
