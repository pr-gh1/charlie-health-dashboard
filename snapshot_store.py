"""
snapshot_store.py

Persists the last computed KPI snapshot (summary_stats() output) to a
private GitHub Gist, so the app can tell you what changed between this
upload and the last one -- even though the app itself is otherwise
stateless (Streamlit Community Cloud's filesystem doesn't survive a
restart, and each upload is parsed fresh in memory).

Two layers, deliberately separate:
  - compute_diff() is free, deterministic, no LLM -- a structured
    before/after comparison of the numbers. This alone is useful and
    costs nothing to show.
  - llm_change_summary() is the on-demand, paid step that turns that
    diff into a business-context narrative ("this looks like normal
    post-holiday attrition" vs. just reciting that a number went down).

Network failures degrade gracefully -- if the Gist is unreachable or not
configured, load_last_snapshot() returns None (treated as "first run")
rather than crashing the dashboard.
"""

import json

import requests

GITHUB_API = "https://api.github.com"
SNAPSHOT_FILENAME = "charlie_health_snapshot.json"


def load_last_snapshot(token, gist_id):
    """Returns the last saved stats dict, or None if there isn't one yet
    (first run) or the store isn't configured/reachable."""
    if not token or not gist_id:
        return None
    try:
        resp = requests.get(
            f"{GITHUB_API}/gists/{gist_id}",
            headers={"Authorization": f"token {token}", "Accept": "application/vnd.github+json"},
            timeout=10,
        )
        resp.raise_for_status()
        files = resp.json().get("files", {})
        f = files.get(SNAPSHOT_FILENAME)
        if not f:
            return None
        content = f.get("content")
        if f.get("truncated") and f.get("raw_url"):
            content = requests.get(f["raw_url"], timeout=10).text
        return json.loads(content) if content else None
    except (requests.RequestException, json.JSONDecodeError, KeyError):
        return None


def save_snapshot(token, gist_id, stats):
    """Overwrites the stored snapshot with the current stats. Raises on
    failure (unlike load_last_snapshot) so a misconfigured secret is loud
    when you're trying to save, not silently swallowed."""
    if not token or not gist_id:
        raise ValueError("GITHUB_TOKEN and GIST_ID must both be configured to save a snapshot.")
    payload = {"files": {SNAPSHOT_FILENAME: {"content": json.dumps(stats, indent=2, default=str)}}}
    resp = requests.patch(
        f"{GITHUB_API}/gists/{gist_id}",
        headers={"Authorization": f"token {token}", "Accept": "application/vnd.github+json"},
        json=payload,
        timeout=10,
    )
    resp.raise_for_status()
    return resp.json()


def _cmp_block(prev_block, curr_block, keys):
    out = {}
    for k in keys:
        pv, cv = prev_block.get(k), curr_block.get(k)
        if pv is None or cv is None or isinstance(pv, str) or isinstance(cv, str):
            continue
        delta = cv - pv
        pct = round(delta / pv * 100, 1) if pv else None
        out[k] = {"previous": pv, "current": cv, "delta": round(delta, 3), "pct_change": pct}
    return out


def compute_diff(previous_stats, current_stats):
    """Free, deterministic before/after comparison -- no LLM. Returns None
    if there's no previous snapshot to compare against."""
    if not previous_stats:
        return None

    changes = {
        "totals": _cmp_block(
            previous_stats.get("totals", {}), current_stats.get("totals", {}),
            ["patients", "scheduled_sessions", "attended_sessions", "total_revenue", "avg_los_appointments"],
        ),
    }
    for window in ("L7D", "L1M", "L3M"):
        pw = previous_stats.get("trailing_windows", {}).get(window, {})
        cw = current_stats.get("trailing_windows", {}).get(window, {})
        changes[window] = _cmp_block(
            pw, cw,
            ["patients_in_treatment", "iop_attendance_rate", "opt_attendance_rate",
             "billable_iop_sessions", "avg_billed_iop_rate", "avg_daily_revenue",
             "total_revenue", "new_admissions"],
        )

    prev_payor = {r["payor"]: r for r in previous_stats.get("by_payor", [])}
    curr_payor = {r["payor"]: r for r in current_stats.get("by_payor", [])}
    by_payor = {}
    for payor in sorted(set(prev_payor) | set(curr_payor)):
        by_payor[payor] = _cmp_block(
            prev_payor.get(payor, {}), curr_payor.get(payor, {}),
            ["patients", "avg_los_appointments", "attendance_rate", "total_revenue", "avg_billed_iop_rate"],
        )
    changes["by_payor"] = by_payor

    return {
        "previous_captured_at": previous_stats.get("captured_at"),
        "current_captured_at": current_stats.get("captured_at"),
        "changes": changes,
    }


def llm_change_summary(client, model, diff, current_stats):
    """On-demand narrative: what changed and why it might matter, grounded
    strictly in the diff + current stats (not a free recitation)."""
    if diff is None:
        return "No previous snapshot to compare against yet -- this is the first captured run."

    prompt = (
        "You are a clinical operations analyst for Charlie Health. Below is "
        "a structured diff between the last captured KPI snapshot and the "
        "current one, plus the full current stats for context. Using ONLY "
        "these numbers, write a short business-context summary of what "
        "changed and why it might matter -- not a recitation of the diff, "
        "but what a COO should take away (e.g. does this look like a "
        "seasonal pattern, a payor mix shift, a capacity or admissions "
        "problem, something worth a follow-up). If nothing changed "
        "meaningfully, say so plainly rather than inventing a story.\n\n"
        "DIFF:\n" + json.dumps(diff, indent=2, default=str) + "\n\n"
        "CURRENT STATS:\n" + json.dumps(current_stats, indent=2, default=str)
    )
    resp = client.messages.create(model=model, max_tokens=500, messages=[{"role": "user", "content": prompt}])
    # Some models may return non-text content blocks (e.g. thinking) ahead of
    # the actual text -- filter to text blocks specifically rather than
    # assuming content[0] is text.
    return "".join(b.text for b in resp.content if getattr(b, "type", None) == "text")
