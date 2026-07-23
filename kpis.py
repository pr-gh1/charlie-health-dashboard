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

    # LOS: duration between a patient's FIRST and LAST ATTENDED appointment,
    # not first/last scheduled entry. Using all scheduled sessions (the
    # original approach) lets a booked-but-never-attended or no-show
    # appointment stretch a patient's measured stay past when they actually
    # last showed up -- LOS should reflect actual attended duration, per the
    # appointment record itself, not the calendar span of the booking log.
    attended_only = sessions_df[sessions_df["attended"] == 1]
    los_dates = attended_only.groupby("patient_id").agg(
        first_attended_date=("date", "min"),
        last_attended_date=("date", "max"),
    ).reset_index()
    agg = agg.merge(los_dates, on="patient_id", how="left")
    agg["los_days"] = (agg["last_attended_date"] - agg["first_attended_date"]).dt.days
    agg["los_weeks"] = agg["los_days"] / 7

    # Attendance rate: each patient's OWN attended/scheduled ratio. This is
    # the individual (patient-level) rate -- kept distinct from a pooled/
    # blended rate (total attended sessions / total scheduled sessions
    # across all patients), which the downstream weekly/trailing/payor
    # views used to compute instead. Reporting should average these
    # individual rates, not pool the underlying counts first.
    agg["attendance_rate"] = agg["attended_sessions"] / agg["scheduled_sessions"]

    # LOS, the headline KPI: number of appointments ATTENDED per patient
    # stay -- not a calendar-time span. Appointment count is the stable,
    # billing-relevant measure of how much treatment someone actually
    # received; calendar weeks conflates that with scheduling gaps (holidays,
    # attendance cadence, etc.) that don't reflect episode length. los_days/
    # los_weeks are still computed below and kept on the table, but only as
    # an internal timing figure for the N24M and Part 2 models (which need a
    # calendar-time discharge lag) -- they are not the reported LOS KPI.
    agg["los_appointments"] = agg["attended_sessions"]

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

    # Individual attendance rate: each patient's own attended/scheduled
    # ratio for the week, averaged across patients -- not a pooled ratio of
    # total attended / total scheduled sessions across everyone. A pooled
    # rate is silently weighted toward whichever patients had the most
    # sessions that week; the individual average treats every patient's
    # attendance equally regardless of their volume.
    def _individual_rate(frame, type_filter=None):
        f = frame if type_filter is None else frame[frame.session_type == type_filter]
        per_patient = f.groupby(["week", "patient_id"]).agg(
            scheduled=("date", "count"), attended=("attended", "sum")
        ).reset_index()
        per_patient["rate"] = per_patient["attended"] / per_patient["scheduled"]
        return per_patient.groupby("week")["rate"].mean().reset_index()

    weekly = weekly.merge(
        _individual_rate(s).rename(columns={"rate": "attendance_rate"}), on="week", how="left"
    )
    weekly = weekly.merge(
        _individual_rate(s, "IOP").rename(columns={"rate": "iop_attendance_rate"}), on="week", how="left"
    )
    weekly = weekly.merge(
        _individual_rate(s, "OPT").rename(columns={"rate": "opt_attendance_rate"}), on="week", how="left"
    )
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


def monthly_rollup(sessions_df, patients_df, payor=None):
    """One row per calendar month -- census (avg patients in treatment) and
    revenue (IOP/OPT/total), same shape/grain as the N24M forecast baseline
    so actuals can be plotted directly against it. Census is an average of
    the weekly in-treatment headcount within the month (the forecast's
    census figure is likewise a within-month average, not a point-in-time
    snapshot), so the two are comparable rather than apples-to-oranges.
    """
    weekly = weekly_rollup(sessions_df, patients_df, payor=payor)
    if weekly.empty:
        return weekly.assign(month=[])

    s = filter_sessions(sessions_df, payor).copy()
    s["month"] = s["date"].dt.to_period("M").dt.to_timestamp()

    monthly = s.groupby("month").agg(revenue=("revenue", "sum")).reset_index()
    iop = s[s.session_type == "IOP"].groupby("month").agg(iop_revenue=("revenue", "sum")).reset_index()
    opt = s[s.session_type == "OPT"].groupby("month").agg(opt_revenue=("revenue", "sum")).reset_index()
    monthly = monthly.merge(iop, on="month", how="left").merge(opt, on="month", how="left")
    monthly[["iop_revenue", "opt_revenue"]] = monthly[["iop_revenue", "opt_revenue"]].fillna(0)

    w = weekly.copy()
    w["month"] = w["week"].dt.to_period("M").dt.to_timestamp()
    census = w.groupby("month")["patients_in_treatment"].mean().reset_index(name="census")

    monthly = monthly.merge(census, on="month", how="left")
    monthly["month_label"] = monthly["month"].dt.strftime("%Y-%m")
    return monthly.sort_values("month").reset_index(drop=True)


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

    # Individual (per-patient) attendance rate, averaged -- not pooled
    # attended/scheduled counts. See weekly_rollup()'s _individual_rate for
    # the same logic applied per-week instead of per-window.
    def _individual_rate(frame):
        if frame.empty:
            return None
        per_patient = frame.groupby("patient_id").agg(
            scheduled=("date", "count"), attended=("attended", "sum")
        )
        rate = (per_patient["attended"] / per_patient["scheduled"]).mean()
        return round(float(rate), 3)

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
        "iop_attendance_rate": _individual_rate(iop),
        "opt_attendance_rate": _individual_rate(opt),
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
    # Keyed off last ATTENDED date, matching the LOS fix -- a stale future
    # scheduled entry that was never attended shouldn't make a patient look
    # "still active," and vice versa.
    p.loc[p["last_attended_date"] < cutoff, "episode_status"] = "completed"
    return p


def los_by_discharge_month(sessions_df, patients_df, payor=None, inactivity_days=DEFAULT_INACTIVITY_DAYS):
    """Avg LOS (appointments attended) trended by discharge month, completed
    episodes only -- see classify_episodes() for why active (still-running)
    episodes are excluded rather than averaged in. Also carries avg_los_weeks
    alongside as a secondary/reference figure -- not the headline KPI, but
    needed by the N24M and Part 2 models for their calendar-time discharge
    lag."""
    p = filter_patients(classify_episodes(patients_df, sessions_df, inactivity_days), payor)
    completed = p[p["episode_status"] == "completed"].copy()
    if completed.empty:
        return pd.DataFrame(columns=["discharge_month", "avg_los_appointments", "avg_los_weeks", "patients"])
    completed["discharge_month"] = completed["last_attended_date"].dt.to_period("M").dt.to_timestamp()
    monthly = completed.groupby("discharge_month").agg(
        avg_los_appointments=("los_appointments", "mean"),
        avg_los_weeks=("los_weeks", "mean"),
        patients=("patient_id", "count"),
    ).reset_index()
    monthly["avg_los_appointments"] = monthly["avg_los_appointments"].round(1)
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
    tenure_weeks = (max_date - active["first_attended_date"]).dt.days / 7
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
        # Individual attendance rate: mean of each patient's own rate
        # (already computed per-patient in build_frames), not pooled
        # attended/scheduled counts across the payor group.
        rows.append({
            "payor": payor,
            "patients": int(len(p)),
            "avg_los_appointments": round(float(p["los_appointments"].mean()), 1),
            "attendance_rate": round(float(p["attendance_rate"].mean()), 3) if len(p) else None,
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
            "avg_los_appointments": round(float(patients_df["los_appointments"].mean()), 1),
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
