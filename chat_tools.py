"""
chat_tools.py

Lets the dashboard's chat answer specific, drill-down questions ("who are
the 10 worst-attendance patients?") instead of only reasoning over the same
aggregate numbers already on screen. Claude gets a small, fixed set of
scoped query functions -- not arbitrary code execution -- so a question
(or a prompt-injection attempt buried in uploaded data) can only ever
trigger one of these four safe, bounded pandas queries, never free-form
eval.

Each tool function takes plain JSON-safe kwargs and returns plain
JSON-safe data (rounded floats, string dates) built from the same
sessions_df/patients_df the rest of the dashboard uses -- so an answer
in the chat can't drift from what the charts show.
"""

import json

from kpis import weekly_rollup, filter_sessions, filter_patients

MAX_TOOL_ROUNDS = 5
MAX_LIST_LIMIT = 50

TOOLS = [
    {
        "name": "list_patients",
        "description": (
            "List patients matching filters, sorted and limited. Use this "
            "for questions like 'who has the worst attendance' or 'which "
            "patients have been in treatment longest'."
        ),
        "input_schema": {
            "type": "object",
            "properties": {
                "payor": {"type": "string", "enum": ["All", "Commercial", "Medicaid"]},
                "min_attendance_rate": {"type": "number"},
                "max_attendance_rate": {"type": "number"},
                "min_los_appointments": {"type": "number"},
                "max_los_appointments": {"type": "number"},
                "sort_by": {
                    "type": "string",
                    "enum": ["attendance_rate", "los_appointments", "total_revenue",
                             "scheduled_sessions", "first_session_date"],
                    "default": "attendance_rate",
                },
                "ascending": {"type": "boolean", "default": True},
                "limit": {"type": "integer", "default": 10, "maximum": MAX_LIST_LIMIT},
            },
        },
    },
    {
        "name": "get_patient_detail",
        "description": "Full record and session history for one patient by ID.",
        "input_schema": {
            "type": "object",
            "properties": {"patient_id": {"type": "string"}},
            "required": ["patient_id"],
        },
    },
    {
        "name": "weekly_trend",
        "description": (
            "Weekly time series for one metric, optionally filtered by "
            "payor and/or date range. Use for 'how has X trended' questions."
        ),
        "input_schema": {
            "type": "object",
            "properties": {
                "metric": {
                    "type": "string",
                    "enum": ["patients_in_treatment", "iop_attendance_rate",
                              "opt_attendance_rate", "revenue", "avg_daily_revenue",
                              "new_admissions", "iop_attended"],
                },
                "payor": {"type": "string", "enum": ["All", "Commercial", "Medicaid"]},
                "start_date": {"type": "string", "description": "YYYY-MM-DD, optional"},
                "end_date": {"type": "string", "description": "YYYY-MM-DD, optional"},
            },
            "required": ["metric"],
        },
    },
    {
        "name": "aggregate_stats",
        "description": (
            "Totals (scheduled/attended sessions, revenue, avg LOS, "
            "attendance rate) over a filtered/dated slice of the data. Use "
            "for 'what was total revenue in X' type questions."
        ),
        "input_schema": {
            "type": "object",
            "properties": {
                "payor": {"type": "string", "enum": ["All", "Commercial", "Medicaid"]},
                "session_type": {"type": "string", "enum": ["All", "IOP", "OPT"]},
                "start_date": {"type": "string", "description": "YYYY-MM-DD, optional"},
                "end_date": {"type": "string", "description": "YYYY-MM-DD, optional"},
            },
        },
    },
]


def _date_filter(df, start_date, end_date):
    if start_date:
        df = df[df["date"] >= start_date]
    if end_date:
        df = df[df["date"] <= end_date]
    return df


def _tool_list_patients(sessions_df, patients_df, **kw):
    p = filter_patients(patients_df, kw.get("payor"))
    if kw.get("min_attendance_rate") is not None:
        p = p[p["attendance_rate"] >= kw["min_attendance_rate"]]
    if kw.get("max_attendance_rate") is not None:
        p = p[p["attendance_rate"] <= kw["max_attendance_rate"]]
    if kw.get("min_los_appointments") is not None:
        p = p[p["los_appointments"] >= kw["min_los_appointments"]]
    if kw.get("max_los_appointments") is not None:
        p = p[p["los_appointments"] <= kw["max_los_appointments"]]

    sort_by = kw.get("sort_by", "attendance_rate")
    ascending = kw.get("ascending", True)
    limit = min(kw.get("limit", 10) or 10, MAX_LIST_LIMIT)
    p = p.sort_values(sort_by, ascending=ascending).head(limit)

    cols = ["patient_id", "payor", "attendance_rate", "los_appointments",
            "scheduled_sessions", "attended_sessions", "total_revenue",
            "first_session_date", "last_session_date"]
    out = p[cols].copy()
    for c in ("first_session_date", "last_session_date"):
        out[c] = out[c].dt.strftime("%Y-%m-%d")
    out["attendance_rate"] = out["attendance_rate"].round(3)
    out["total_revenue"] = out["total_revenue"].round(2)
    return out.to_dict(orient="records")


def _tool_get_patient_detail(sessions_df, patients_df, **kw):
    pid = kw["patient_id"]
    p = patients_df[patients_df["patient_id"] == pid]
    if p.empty:
        return {"error": f"No patient found with id {pid!r}"}
    record = p.iloc[0].to_dict()
    for k, v in list(record.items()):
        if hasattr(v, "strftime"):
            record[k] = v.strftime("%Y-%m-%d")
        elif hasattr(v, "item"):
            record[k] = v.item()

    s = sessions_df[sessions_df["patient_id"] == pid].sort_values("date")
    sessions_list = [
        {
            "date": row["date"].strftime("%Y-%m-%d"),
            "session_type": row["session_type"],
            "attended": int(row["attended"]),
            "revenue": round(float(row["revenue"]), 2),
        }
        for _, row in s.iterrows()
    ]
    record["sessions"] = sessions_list
    return record


def _tool_weekly_trend(sessions_df, patients_df, **kw):
    metric = kw["metric"]
    payor = kw.get("payor")
    weekly = weekly_rollup(sessions_df, patients_df, payor=payor)
    weekly = _date_filter(weekly.rename(columns={"week": "date"}), kw.get("start_date"), kw.get("end_date"))
    if metric not in weekly.columns:
        return {"error": f"Unknown metric {metric!r}"}
    return [
        {"week": d.strftime("%Y-%m-%d"), "value": round(float(v), 3) if pd_notna(v) else None}
        for d, v in zip(weekly["date"], weekly[metric])
    ]


def pd_notna(v):
    import pandas as pd
    return pd.notna(v)


def _tool_aggregate_stats(sessions_df, patients_df, **kw):
    s = filter_sessions(sessions_df, kw.get("payor"))
    session_type = kw.get("session_type")
    if session_type and session_type != "All":
        s = s[s["session_type"] == session_type]
    s = _date_filter(s, kw.get("start_date"), kw.get("end_date"))
    p = filter_patients(patients_df, kw.get("payor"))
    if kw.get("start_date") or kw.get("end_date"):
        # Scope "patients"/"avg_los_appointments" to those with activity in the
        # window, not the whole payor cohort -- otherwise a Q1-only query
        # would silently report every Commercial patient ever seen.
        p = p[p["patient_id"].isin(s["patient_id"].unique())]

    scheduled = len(s)
    attended = int(s["attended"].sum())
    return {
        "scheduled_sessions": scheduled,
        "attended_sessions": attended,
        "attendance_rate": round(attended / scheduled, 3) if scheduled else None,
        "total_revenue": round(float(s["revenue"].sum()), 2),
        "avg_los_appointments": round(float(p["los_appointments"].mean()), 1) if len(p) else None,
        "patients": int(p["patient_id"].nunique()),
    }


_DISPATCH = {
    "list_patients": _tool_list_patients,
    "get_patient_detail": _tool_get_patient_detail,
    "weekly_trend": _tool_weekly_trend,
    "aggregate_stats": _tool_aggregate_stats,
}


def dispatch_tool(name, tool_input, sessions_df, patients_df):
    fn = _DISPATCH.get(name)
    if fn is None:
        return {"error": f"Unknown tool {name!r}"}
    try:
        return fn(sessions_df, patients_df, **tool_input)
    except Exception as e:
        return {"error": str(e)}


def answer_question(client, model, question, sessions_df, patients_df, history=None):
    """Runs Claude's standard tool-use loop: ask -> (maybe) call a tool ->
    feed the result back -> repeat, capped at MAX_TOOL_ROUNDS so a stuck
    loop can't run away on cost. Returns (answer_text, new_history) where
    new_history is the full message list, ready to pass back in on the
    next question for multi-turn context.
    """
    system = (
        "You are a data analyst for Charlie Health's Attendance & Billing "
        "dashboard. Answer questions using the provided tools to query the "
        "real Sessions/Patients data -- never guess or invent numbers. If "
        "a question is ambiguous (e.g. 'worst attendance' could mean "
        "lowest rate or fewest sessions), pick a reasonable interpretation "
        "and say which you used. Keep answers concise and cite the actual "
        "figures returned by the tools."
    )
    messages = list(history or []) + [{"role": "user", "content": question}]

    for _ in range(MAX_TOOL_ROUNDS):
        resp = client.messages.create(
            model=model,
            max_tokens=800,
            system=system,
            tools=TOOLS,
            messages=messages,
        )
        messages.append({"role": "assistant", "content": resp.content})

        if resp.stop_reason != "tool_use":
            text = "".join(b.text for b in resp.content if b.type == "text")
            return text, messages

        tool_results = []
        for block in resp.content:
            if block.type == "tool_use":
                result = dispatch_tool(block.name, block.input, sessions_df, patients_df)
                tool_results.append({
                    "type": "tool_result",
                    "tool_use_id": block.id,
                    "content": json.dumps(result, default=str),
                })
        messages.append({"role": "user", "content": tool_results})

    return "Hit the tool-call limit for this question -- try narrowing it down.", messages
