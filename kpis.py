"""
kpis.py

Turns the Sessions/Patients/Rates records produced by transform_attendance.py
into the KPI tables the case study's Part 1 asks for, plus a couple of
extras. Kept separate from transform_attendance.py on purpose: that module's
job is "reshape the raw export faithfully," this module's job is
"compute business metrics from the reshaped data" -- so a future change to
how a KPI is defined never touches the (already-verified) parsing logic.

All functions take/return pandas DataFrames built from the plain list-of-dict
records that transform_attendance.parse_grid() returns.
"""

import pandas as pd


def build_frames(rates, sessions, patients):
    """Convert the raw records from parse_grid() into pandas DataFrames,
    with rate/revenue attached to each session."""
    rates_df = pd.DataFrame(rates)
    sessions_df = pd.DataFrame(sessions)
    sessions_df["date"] = pd.to_datetime(sessions_df["date"])
    sessions_df["billable"] = sessions_df["attended"]  # see Sessions tab note

    sessions_df = sessions_df.merge(
        rates_df, on=["session_type", "payor"], how="left"
    )
    sessions_df["revenue"] = sessions_df["billable"] * sessions_df["rate"]

    patients_df = pd.DataFrame(patients)
    agg = sessions_df.groupby("patient_id").agg(
        first_session_date=("date", "min"),
        last_session_date=("date", "max"),
        scheduled_sessions=("date", "count"),
        attended_sessions=("attended", "sum"),
        iop_attended=("attended", lambda s: s[sessions_df.loc[s.index, "session_type"] == "IOP"].sum()),
        opt_attended=("attended", lambda s: s[sessions_df.loc[s.index, "session_type"] == "OPT"].sum()),
        total_revenue=("revenue", "sum"),
    ).reset_index()
    agg["los_days"] = (agg["last_session_date"] - agg["first_session_date"]).dt.days
    agg["los_weeks"] = agg["los_days"] / 7
    agg["attendance_rate"] = agg["attended_sessions"] / agg["scheduled_sessions"]
    patients_df = patients_df.merge(agg, on="patient_id", how="left")

    return rates_df, sessions_df, patients_df


def _week_start(dates):
    """Monday of the week each date falls in."""
    return dates - pd.to_timedelta(dates.dt.weekday, unit="D")


def filter_sessions(sessions_df, payor=None):
    if payor and payor != "All":
        return sessions_df[sessions_df["payor"] == payor]
    return sessions_df


def filter_patients(patients_df, payor=None):
    if payor and payor != "All":
        return patients_df[patients_df["payor"] == payor]
    return patients_df


def weekly_rollup(sessions_df, patients_df, payor=None):
    """One row per (week), covering every metric Part 1 asks to track over
    time: patients in treatment, IOP attendance, billable sessions,
    attendance rates, avg billed IOP rate/session, avg daily revenue --
    plus new admissions, a natural read on b2b clinical-outreach volume
    per the case study's hint that all admissions come from outreach reps.

    `payor` optionally restricts everything to "Commercial" or "Medicaid"
    (or leave as None / "All" for the blended view) -- same shape either
    way, so the dashboard can reuse one code path for both.
    """
    sessions_df = filter_sessions(sessions_df, payor)
    patients_df = filter_patients(patients_df, payor)

    s = sessions_df.copy()
    s["week"] = _week_start(s["date"])

    weekly = s.groupby("week").agg(
        scheduled_sessions=("date", "count"),
        attended_sessions=("attended", "sum"),
        revenue=("revenue", "sum"),
    ).reset_index()

    iop = s[s.session_type == "IOP"].groupby("week").agg(
        iop_scheduled=("date", "count"),
        iop_attended=("attended", "sum"),
        iop_revenue=("revenue", "sum"),
    ).reset_index()
    opt = s[s.session_type == "OPT"].groupby("week").agg(
        opt_scheduled=("date", "count"),
        opt_attended=("attended", "sum"),
    ).reset_index()

    # Distinct headcount, not a session count or a rate -- "patients attending
    # IOP" answers a different question than iop_attended (session volume) or
    # iop_attendance_rate (a ratio). 10 patients attending once each and 2
    # patients attending five times each look identical in session counts;
    # this is what tells them apart.
    iop_patients = (
        s[(s.session_type == "IOP") & (s.attended == 1)]
        .groupby("week")["patient_id"].nunique()
        .reset_index(name="iop_patients")
    )

    weekly = weekly.merge(iop, on="week", how="left").merge(opt, on="week", how="left")
    weekly = weekly.merge(iop_patients, on="week", how="left")
    weekly["iop_patients"] = weekly["iop_patients"].fillna(0).astype(int)
    weekly["attendance_rate"] = weekly["attended_sessions"] / weekly["scheduled_sessions"]
    weekly["iop_attendance_rate"] = weekly["iop_attended"] / weekly["iop_scheduled"]
    weekly["opt_attendance_rate"] = weekly["opt_attended"] / weekly["opt_scheduled"]
    weekly["avg_billed_iop_rate"] = weekly["iop_revenue"] / weekly["iop_attended"]
    weekly["avg_daily_revenue"] = weekly["revenue"] / 7

    p = patients_df.copy()
    p["admit_week"] = _week_start(p["first_session_date"])
    admits = p.groupby("admit_week").size().reset_index(name="new_admissions")
    weekly = weekly.merge(admits, left_on="week", right_on="admit_week", how="left").drop(columns=["admit_week"])
    weekly["new_admissions"] = weekly["new_admissions"].fillna(0).astype(int)

    # Patients "in treatment" that week = active anywhere between their
    # first and last observed session, inclusive of that week's Mon-Sun span.
    weeks = weekly["week"].tolist()
    in_treatment = []
    for wk in weeks:
        wk_end = wk + pd.Timedelta(days=6)
        active = ((p["first_session_date"] <= wk_end) & (p["last_session_date"] >= wk)).sum()
        in_treatment.append(active)
    weekly["patients_in_treatment"] = in_treatment

    return weekly.sort_values("week").reset_index(drop=True)


WINDOWS = {"L7D": 7, "L1M": 30, "L3M": 90}


def trailing_window_stats(sessions_df, patients_df, window_days, payor=None, as_of=None):
    """KPI snapshot for a trailing N-day window ending at the most recent
    date in the data (or `as_of`). Computed directly off daily session-level
    data rather than weekly_rollup(), because 7/30/90-day windows don't line
    up with week boundaries -- this is "is right now normal," the weekly
    charts are "how did we get here."
    """
    s = filter_sessions(sessions_df, payor)
    p = filter_patients(patients_df, payor)
    end = as_of or s["date"].max()
    start = end - pd.Timedelta(days=window_days - 1)
    win = s[(s["date"] >= start) & (s["date"] <= end)]

    iop = win[win.session_type == "IOP"]
    opt = win[win.session_type == "OPT"]
    iop_scheduled, iop_attended = len(iop), int(iop["attended"].sum())
    opt_scheduled, opt_attended = len(opt), int(opt["attended"].sum())
    iop_patient_count = int(iop[iop.attended == 1]["patient_id"].nunique())
    revenue = float(win["revenue"].sum())
    iop_revenue = float(iop["revenue"].sum())

    patients_in_treatment = int(((p["first_session_date"] <= end) & (p["last_session_date"] >= start)).sum())
    new_admissions = int(((p["first_session_date"] >= start) & (p["first_session_date"] <= end)).sum())

    return {
        "window_days": window_days,
        "start": start.strftime("%Y-%m-%d"),
        "end": end.strftime("%Y-%m-%d"),
        "payor": payor or "All",
        "patients_in_treatment": patients_in_treatment,
        "new_admissions": new_admissions,
        "iop_patients": iop_patient_count,
        "iop_attendance_rate": round(iop_attended / iop_scheduled, 3) if iop_scheduled else None,
        "opt_attendance_rate": round(opt_attended / opt_scheduled, 3) if opt_scheduled else None,
        "billable_iop_sessions": iop_attended,
        "avg_billed_iop_rate": round(iop_revenue / iop_attended, 2) if iop_attended else None,
        "avg_daily_revenue": round(revenue / window_days, 2),
        "total_revenue": round(revenue, 2),
    }


DEFAULT_INACTIVITY_DAYS = 14


def classify_episodes(patients_df, sessions_df, inactivity_days=DEFAULT_INACTIVITY_DAYS):
    """Splits patients into 'completed' vs 'active' episodes based on how
    recently their last session was, relative to the newest date in the
    data. This matters for LOS specifically: a patient who's still actively
    attending doesn't have a finished LOS yet, just a still-running clock,
    and folding their current tenure into an LOS average biases recent
    numbers down for no clinical reason -- they haven't left early, the
    data just hasn't caught up to them yet.

    `inactivity_days` is a heuristic, not a real discharge flag (the source
    data doesn't have one) -- a typical IOP patient attends several times a
    week, so ~2 weeks with no session is a reasonable "probably done" signal.
    """
    max_date = sessions_df["date"].max()
    cutoff = max_date - pd.Timedelta(days=inactivity_days)
    p = patients_df.copy()
    p["episode_status"] = "active"
    p.loc[p["last_session_date"] < cutoff, "episode_status"] = "completed"
    return p


def los_by_discharge_month(sessions_df, patients_df, payor=None, inactivity_days=DEFAULT_INACTIVITY_DAYS):
    """Avg LOS trended by discharge month, completed episodes only -- see
    classify_episodes() for why active (still-running) episodes are
    excluded rather than averaged in."""
    p = filter_patients(classify_episodes(patients_df, sessions_df, inactivity_days), payor)
    completed = p[p["episode_status"] == "completed"].copy()
    if completed.empty:
        return pd.DataFrame(columns=["discharge_month", "avg_los_weeks", "patients"])
    completed["discharge_month"] = completed["last_session_date"].dt.to_period("M").dt.to_timestamp()
    monthly = completed.groupby("discharge_month").agg(
        avg_los_weeks=("los_weeks", "mean"),
        patients=("patient_id", "count"),
    ).reset_index()
    monthly["avg_los_weeks"] = monthly["avg_los_weeks"].round(1)
    return monthly.sort_values("discharge_month")


def active_patient_stats(sessions_df, patients_df, payor=None, inactivity_days=DEFAULT_INACTIVITY_DAYS):
    """Count and avg tenure-so-far of patients still in an active episode --
    context alongside the LOS trend, deliberately not part of the LOS
    average itself."""
    p = filter_patients(classify_episodes(patients_df, sessions_df, inactivity_days), payor)
    active = p[p["episode_status"] == "active"]
    if active.empty:
        return {"count": 0, "avg_tenure_weeks": None}
    max_date = sessions_df["date"].max()
    tenure_weeks = (max_date - active["first_session_date"]).dt.days / 7
    return {"count": int(len(active)), "avg_tenure_weeks": round(float(tenure_weeks.mean()), 1)}


def payor_composition(sessions_df, patients_df):
    """Per-payor breakdown -- patients, avg LOS, attendance rate, revenue,
    avg billed IOP rate -- for the Composition section. The case study's own
    rate table (IOP Commercial $340 vs Medicaid $220, OPT $120 vs $90) is
    close to a 1.5x spread, so payor mix is a first-order revenue lever,
    not a footnote."""
    rows = []
    for payor in sorted(patients_df["payor"].unique()):
        p = patients_df[patients_df["payor"] == payor]
        s = sessions_df[sessions_df["payor"] == payor]
        iop_attended = s[(s.session_type == "IOP") & (s.attended == 1)]
        rows.append({
            "payor": payor,
            "patients": int(len(p)),
            "avg_los_weeks": round(float(p["los_weeks"].mean()), 1),
            "attendance_rate": round(float(s["attended"].sum() / len(s)), 3) if len(s) else None,
            "total_revenue": round(float(s["revenue"].sum()), 2),
            "avg_billed_iop_rate": round(float(iop_attended["revenue"].sum() / len(iop_attended)), 2)
            if len(iop_attended) else None,
        })
    return pd.DataFrame(rows)


def summary_stats(sessions_df, patients_df, weekly_df):
    """Compact, LLM-ready summary: totals, most-recent-4-weeks vs the 4
    weeks before that, trailing-window snapshots, and payor breakdown.
    Every number here is exactly what gets shown in the dashboard --
    nothing is computed specially for the LLM, so the narrative can't
    drift from the charts.
    """
    recent = weekly_df.tail(4)
    prior = weekly_df.iloc[-8:-4] if len(weekly_df) >= 8 else weekly_df.head(0)

    def block(df):
        if df.empty:
            return None
        return {
            "weeks": [d.strftime("%Y-%m-%d") for d in df["week"]],
            "avg_patients_in_treatment": round(df["patients_in_treatment"].mean(), 1),
            "avg_iop_attendance_rate": round(df["iop_attendance_rate"].mean(), 3),
            "avg_opt_attendance_rate": round(df["opt_attendance_rate"].mean(), 3),
            "total_billable_iop_sessions": int(df["iop_attended"].sum()),
            "avg_billed_iop_rate": round(df["avg_billed_iop_rate"].mean(), 2),
            "avg_daily_revenue": round(df["avg_daily_revenue"].mean(), 2),
            "total_new_admissions": int(df["new_admissions"].sum()),
        }

    return {
        "date_range": {
            "start": sessions_df["date"].min().strftime("%Y-%m-%d"),
            "end": sessions_df["date"].max().strftime("%Y-%m-%d"),
        },
        "totals": {
            "patients": int(patients_df["patient_id"].nunique()),
            "scheduled_sessions": int(len(sessions_df)),
            "attended_sessions": int(sessions_df["attended"].sum()),
            "total_revenue": round(float(sessions_df["revenue"].sum()), 2),
            "avg_los_weeks": round(float(patients_df["los_weeks"].mean()), 1),
        },
        "payor_mix": {
            payor: int(count)
            for payor, count in patients_df["payor"].value_counts().items()
        },
        "most_recent_4_weeks": block(recent),
        "prior_4_weeks": block(prior),
        "trailing_windows": {
            label: trailing_window_stats(sessions_df, patients_df, days)
            for label, days in WINDOWS.items()
        },
        "by_payor": payor_composition(sessions_df, patients_df).to_dict(orient="records"),
    }
