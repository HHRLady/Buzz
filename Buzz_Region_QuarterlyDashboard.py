"""
Buzz_Region_QuarterlyDashboard.py
===================================
Generates a single-page HTML quarterly dashboard for HQ founder meetings
and Emma's regional review.

Usage:
    python Buzz_Region_QuarterlyDashboard.py             # current quarter
    python Buzz_Region_QuarterlyDashboard.py --year 2026 --quarter 1

Output:
    Buzz_Region_Curated/Region_Quarterly_Dashboard_YYYY_QN.html
"""

import argparse
import json
import re as _re
from datetime import datetime
from math import floor
from pathlib import Path

import pandas as pd

# ── Paths ──────────────────────────────────────────────────────────────────────
BASE    = Path(__file__).resolve().parent
CURATED = BASE / "Buzz_Region_Curated"
REF     = BASE / "Buzz_Region_Ref"

# ── Brand colours ──────────────────────────────────────────────────────────────
TEAL   = "#00A19A"
ORANGE = "#F39200"
PINK   = "#D60B52"
LIME   = "#B6BD00"
DARK   = "#1A1A2E"
GREY   = "#6B7280"
LGREY  = "#F3F4F6"

# ── Town config ────────────────────────────────────────────────────────────────
TOWNS_ORDER  = ["MarketHarborough", "Leicester", "Loughborough", "Lutterworth", "Hinckley"]
TOWN_LABELS  = {
    "MarketHarborough": "Market Harborough",
    "Leicester":        "Leicester",
    "Loughborough":     "Loughborough",
    "Lutterworth":      "Lutterworth",
    "Hinckley":         "Hinckley",
}
TOWN_SHORT = {
    "MarketHarborough": "MH",
    "Leicester":        "LE",
    "Loughborough":     "LB",
    "Lutterworth":      "LW",
    "Hinckley":         "HI",
}
TOWN_COLOURS = {
    "MarketHarborough": TEAL,
    "Leicester":        ORANGE,
    "Loughborough":     LIME,
    "Lutterworth":      PINK,
    "Hinckley":         "#8B5CF6",
}

# ── EE thresholds ──────────────────────────────────────────────────────────────
EE_PAYING_TARGET  = 25
EE_AMBASSADOR_MIN = 2
EE_COMBINED_MIN   = 4  # combined Buzz Plus + sponsors for 5-star
EE_PLUS_FULL      = 5
EE_SPONSOR_FULL   = 4
SPONSOR_SLOTS     = 4

# ── Cross-town exclusions ──────────────────────────────────────────────────────
# Fallback hardcoded set — the live set is loaded from exclude_region.csv at runtime
CROSS_TOWN_EXCLUDE = {
    "warwickshire@business-buzz.org",  # James Brodie — HQ Regional Lead (Buddha Connect)
    "hello@andkarenhall.co.uk",        # Karen Hall — moved away from the area
}


def load_cross_town_exclusions():
    """Return set of emails to exclude from cross-town table.
    Merges exclude_region.csv (all entries) with the hardcoded fallback set."""
    excluded = set(CROSS_TOWN_EXCLUDE)
    csv_path = BASE / "exclude_region.csv"
    if csv_path.exists():
        try:
            df = pd.read_csv(csv_path)
            df.columns = [c.strip().lower() for c in df.columns]
            if "email" in df.columns:
                emails = (df["email"].astype(str).str.strip().str.lower())
                excluded |= set(emails[(emails != "") & (emails != "nan")])
        except Exception:
            pass
    return excluded


def event_excellence_star(ee_paying, ee_ambs, active_plus=0, sponsors=0, is_new=False):
    """Return (star_str, description) based on HQ star criteria.

    New events are evaluated against the months they have run within the
    Nov–Oct conference period — is_new adds a note but does not bypass scoring.
    """
    new_sfx = " \u00b7 new event, partial period data" if is_new else ""
    if not ee_paying:
        return "\u2014", f"Working toward 3\u2605 (avg below {EE_PAYING_TARGET}){new_sfx}"
    if ee_ambs and (active_plus + sponsors) >= EE_COMBINED_MIN:
        if active_plus >= EE_PLUS_FULL and sponsors >= EE_SPONSOR_FULL:
            return "\u2728 5\u2605", f"Special prize eligible{new_sfx}"
        return "5\u2605", f"5-star Event Excellence{new_sfx}"
    if ee_ambs:
        return "4\u2605", f"4-star Event Excellence{new_sfx}"
    return "3\u2605", f"3-star Event Excellence{new_sfx}"


# ══════════════════════════════════════════════════════════════════════════════
# QUARTER UTILITIES
# ══════════════════════════════════════════════════════════════════════════════

def quarter_months(year: int, q: int):
    """Return the three YYYY-MM strings for a given quarter."""
    start_month = (q - 1) * 3 + 1
    return [f"{year}-{m:02d}" for m in range(start_month, start_month + 3)]


def quarter_label(year: int, q: int):
    return f"Q{q} {year}"


def prior_quarter(year: int, q: int):
    if q == 1:
        return year - 1, 4
    return year, q - 1


def current_quarter():
    now = datetime.now()
    return now.year, (now.month - 1) // 3 + 1


# ══════════════════════════════════════════════════════════════════════════════
# DATA LOADING
# ══════════════════════════════════════════════════════════════════════════════

def _safe_read(path, sheet=None):
    try:
        if sheet:
            return pd.read_excel(path, sheet_name=sheet)
        return pd.read_excel(path)
    except Exception:
        return pd.DataFrame()


def _nc(df):
    if not df.empty:
        df.columns = [str(c).strip().lower() for c in df.columns]
    return df


def _clean_company(name):
    if not name or str(name).strip().lower() in ("", "nan"):
        return ""
    name = str(name).strip()
    parts = name.split(",")
    if len(parts) > 2:
        return parts[0].strip()
    if len(parts) == 2:
        second = parts[1].strip()
        if _re.match(r'^\d+\s', second) or len(second) > 30:
            return parts[0].strip()
    return name


def load_trend():
    ov = _nc(_safe_read(CURATED / "region_dashboard.xlsx", sheet="Town_Overview"))
    tr = _nc(_safe_read(CURATED / "region_dashboard.xlsx", sheet="Attendance_Trend_12m"))
    # Normalise event_month to YYYY-MM string regardless of how Excel stored the date
    if not tr.empty and "event_month" in tr.columns:
        tr["event_month"] = tr["event_month"].astype(str).str[:7]
    return ov, tr


def load_master():
    p = CURATED / "region_master_people.csv"
    if not p.exists():
        return pd.DataFrame()
    df = pd.read_csv(p)
    return _nc(df)


def load_ambassadors():
    df = _nc(_safe_read(REF / "buzz_ambassadors.xlsx"))
    by_town = {}
    if not df.empty and "town_code" in df.columns:
        for tc, grp in df.groupby("town_code"):
            ambs = []
            for _, r in grp.iterrows():
                name = str(r.get("name", "") or "").strip()
                if not name or name.lower() == "nan":
                    continue
                trained = str(r.get("ambassador_training_complete", "")).strip().lower() in ("yes", "true", "1")
                has_left = str(r.get("has_left", "")).strip().lower() in ("yes", "true", "1")
                if not has_left:
                    ambs.append({"name": name, "trained": trained})
            by_town[tc] = ambs
    return by_town


def load_hosts():
    df = _nc(_safe_read(REF / "buzz_hosts.xlsx"))
    hosts = {}
    if not df.empty and "town_code" in df.columns:
        for _, r in df.iterrows():
            tc = str(r.get("town_code", "")).strip()
            is_ct  = str(r.get("is_caretaker", "")).strip().lower() in ("yes", "true", "1")
            is_new = str(r.get("is_new_event", "")).strip().lower() in ("yes", "true", "1")
            name = str(r.get("caretaker_name" if is_ct else "host_name", "") or "").strip()
            if is_ct:
                name_raw = str(r.get("caretaker_name", "") or "").strip()
                hosts[tc] = {"name": name_raw, "caretaker": True, "new_event": is_new}
            else:
                name_raw = str(r.get("host_name", "") or "").strip()
                if name_raw and name_raw.lower() != "nan":
                    hosts[tc] = {"name": name_raw, "caretaker": False, "new_event": is_new}
    return hosts


def load_plus_by_town():
    """Return (active_counts, strong_counts) — both from live source files."""
    active_counts = {}
    try:
        df = pd.read_excel(REF / "buzz_plus_members.xlsx")
        df.columns = [str(c).strip().lower() for c in df.columns]
        if "status" in df.columns:
            df = df[df["status"].astype(str).str.strip().str.lower() == "active"]
        if "town_code" in df.columns:
            active_counts = df.groupby("town_code").size().to_dict()
    except Exception:
        pass

    strong_counts   = {}
    possible_counts = {}
    try:
        bp = pd.read_excel(CURATED / "buzzplus_intelligence.xlsx",
                           sheet_name="Town_Plus_Summary")
        bp.columns = [str(c).strip().lower() for c in bp.columns]
        for _, r in bp.iterrows():
            tc = str(r.get("town_code", "")).strip()
            strong_counts[tc]   = int(pd.to_numeric(r.get("buzzplus_strong_prospects", 0), errors="coerce") or 0)
            possible_counts[tc] = int(pd.to_numeric(r.get("buzzplus_possible_prospects", 0), errors="coerce") or 0)
    except Exception:
        pass

    return active_counts, strong_counts, possible_counts


def load_regional_sponsors():
    sp = _nc(_safe_read(REF / "buzz_regional_sponsors.xlsx"))
    ch = _nc(_safe_read(REF / "buzz_regional_charity.xlsx"))
    sp_names = []
    if not sp.empty and "sponsor_name" in sp.columns:
        sp_names = [str(n).strip() for n in sp["sponsor_name"] if str(n).strip() not in ("", "nan")]
    ch_name = ""
    if not ch.empty and "charity_name" in ch.columns:
        rows = [str(n).strip() for n in ch["charity_name"] if str(n).strip() not in ("", "nan")]
        ch_name = rows[0] if rows else ""
    return {"sp_names": sp_names, "sp_count": len(sp_names), "ch_name": ch_name}


# ══════════════════════════════════════════════════════════════════════════════
# DATA ASSEMBLY
# ══════════════════════════════════════════════════════════════════════════════

def quarterly_attendance(trend_df, months):
    """Return {town_code: [val_m1, val_m2, val_m3]} — None where no event."""
    result = {}
    for code in TOWNS_ORDER:
        t = trend_df[trend_df["town_code"] == code] if not trend_df.empty else pd.DataFrame()
        vals = []
        for m in months:
            row = t[t["event_month"] == m] if not t.empty else pd.DataFrame()
            if not row.empty:
                v = pd.to_numeric(row.iloc[0].get("unique_attendees"), errors="coerce")
                vals.append(int(v) if pd.notna(v) else None)
            else:
                vals.append(None)  # No event this month
        result[code] = vals
    return result


def quarter_avg(vals):
    """Average of non-None values."""
    active = [v for v in vals if v is not None]
    return round(sum(active) / len(active), 1) if active else None


def build_town_quarters(trend_df, overview, year, q, prev_year, prev_q,
                        ambassadors, hosts, plus_counts, plus_strong=None, plus_possible=None):
    q_months   = quarter_months(year, q)
    pq_months  = quarter_months(prev_year, prev_q)

    towns = []
    for code in TOWNS_ORDER:
        ov = overview[overview["town_code"] == code] if not overview.empty else pd.DataFrame()
        cur_vals  = quarterly_attendance(trend_df, q_months)[code]
        prev_vals = quarterly_attendance(trend_df, pq_months)[code]

        cur_avg  = quarter_avg(cur_vals)
        prev_avg = quarter_avg(prev_vals)
        num_events = sum(1 for v in cur_vals if v is not None)

        delta = None
        if cur_avg is not None and prev_avg is not None:
            delta = round(cur_avg - prev_avg, 1)
        arrow = ("▲" if delta and delta > 3 else
                 ("▼" if delta and delta < -3 else "→"))
        arrow_col = (LIME if arrow == "▲" else (PINK if arrow == "▼" else ORANGE))

        active_plus  = plus_counts.get(code, 0)
        sponsors     = int(ov["current_sponsors"].iloc[0]) if not ov.empty and "current_sponsors" in ov.columns else 0
        # Load strong/possible from buzzplus_intelligence (bypasses stale rollup)
        ov_strong    = int(ov["buzzplus_strong_prospects"].iloc[0]) if not ov.empty and "buzzplus_strong_prospects" in ov.columns else 0
        ov_possible  = int(ov["buzzplus_possible_prospects"].iloc[0]) if not ov.empty and "buzzplus_possible_prospects" in ov.columns else 0
        strong_plus  = (plus_strong  or {}).get(code, ov_strong)
        possible_plus = (plus_possible or {}).get(code, ov_possible)
        lockouts    = str(ov["industry_lockouts"].iloc[0]).strip() if not ov.empty and "industry_lockouts" in ov.columns else ""
        if lockouts.lower() in ("nan", ""):
            lockouts = "None"

        ambs = ambassadors.get(code, [])
        trained_count = sum(1 for a in ambs if a["trained"])
        host_info = hosts.get(code, {})

        is_new = host_info.get("new_event", False)

        ee_paying   = cur_avg is not None and cur_avg >= EE_PAYING_TARGET
        ee_ambs     = trained_count >= EE_AMBASSADOR_MIN
        ee_combined = (active_plus + sponsors) >= EE_COMBINED_MIN
        ee_score    = sum([ee_paying, ee_ambs, ee_combined])
        ee_star, ee_star_desc = event_excellence_star(
            ee_paying, ee_ambs,
            active_plus=active_plus, sponsors=sponsors, is_new=is_new,
        )

        rag = (
            "green" if ee_paying and sponsors >= 1 else
            ("red" if (cur_avg is not None and cur_avg < 18) else "amber")
        )

        towns.append(dict(
            code=code, label=TOWN_LABELS[code], short=TOWN_SHORT[code],
            colour=TOWN_COLOURS[code],
            cur_vals=cur_vals, prev_vals=prev_vals,
            cur_avg=cur_avg, prev_avg=prev_avg, delta=delta,
            num_events=num_events,
            arrow=arrow, arrow_col=arrow_col,
            active_plus=active_plus, strong_plus=strong_plus, possible_plus=possible_plus,
            sponsors=sponsors, lockouts=lockouts,
            ambs=ambs, trained_count=trained_count, host_info=host_info,
            is_new=is_new,
            ee_paying=ee_paying, ee_ambs=ee_ambs,
            ee_combined=ee_combined, ee_score=ee_score,
            ee_star=ee_star, ee_star_desc=ee_star_desc,
            rag=rag,
        ))
    return towns


def build_region_summary(towns, year, q, trend_df, master):
    q_months = quarter_months(year, q)
    # Total events
    total_events = sum(t["num_events"] for t in towns)
    # Total visit count for the quarter
    total_visits = sum(
        v for t in towns for v in t["cur_vals"] if v is not None
    )
    # Best and lowest performing towns by avg
    with_data = [t for t in towns if t["cur_avg"] is not None]
    best  = max(with_data, key=lambda t: t["cur_avg"]) if with_data else None
    worst = min(with_data, key=lambda t: t["cur_avg"]) if with_data else None
    region_avg = (
        round(sum(t["cur_avg"] for t in with_data) / len(with_data), 1)
        if with_data else None
    )
    # Unduplicated region headcount
    total_people = 0
    if not master.empty and "role_region" in master.columns:
        team_roles = {"host", "ambassador", "regional lead"}
        is_team = master["role_region"].astype(str).str.strip().str.lower().isin(team_roles)
        total_people = len(master[~is_team])
    return dict(
        total_events=total_events,
        total_visits=total_visits,
        region_avg=region_avg,
        best=best,
        worst=worst,
        total_people=total_people,
    )


def build_cross_town(master, top_n=8):
    if master.empty:
        return []
    excluded_emails = load_cross_town_exclusions()
    team_roles = {"host", "ambassador", "regional lead"}
    excl = master["email"].astype(str).str.strip().str.lower().isin(excluded_emails)
    df = master[~excl & (master["towns_visited_count"] >= 2)].copy()
    df["visits_n"] = pd.to_numeric(df["visits_region"], errors="coerce").fillna(0)
    df = df.sort_values(["visits_n", "towns_visited_count"], ascending=False).head(top_n)
    result = []
    for _, row in df.iterrows():
        name     = str(row.get("name", "") or "").strip().title()
        company  = _clean_company(str(row.get("company", "") or ""))
        role_raw = str(row.get("role_region", "") or "").strip().lower()
        role     = role_raw if role_raw in team_roles else ""
        result.append(dict(
            name=name, company=company,
            visits=int(row["visits_n"]),
            towns_count=int(row.get("towns_visited_count", 0)),
            towns=str(row.get("towns_visited", "") or ""),
            role=role,
        ))
    return result


# ══════════════════════════════════════════════════════════════════════════════
# HTML
# ══════════════════════════════════════════════════════════════════════════════

def _dot(met):
    col = LIME if met else PINK
    sym = "●" if met else "○"
    return f'<span style="color:{col};font-size:13px;">{sym}</span>'


def _rag(rag):
    col = {"green": LIME, "amber": ORANGE, "red": PINK}.get(rag, GREY)
    lbl = {"green": "Strong", "amber": "Steady", "red": "Watch"}.get(rag, "—")
    return col, lbl


def snap_card(colour, value, label, sub=""):
    sub_html = f'<div style="font-size:10px;color:{GREY};margin-top:2px;">{sub}</div>' if sub else ""
    return f"""
    <div style="flex:1;min-width:120px;background:#fff;border-top:4px solid {colour};
         border-radius:8px;padding:14px 12px;text-align:center;
         box-shadow:0 1px 4px rgba(0,0,0,0.07);">
      <div style="font-size:26px;font-weight:800;color:{colour};">{value}</div>
      <div style="font-size:11px;color:{GREY};margin-top:4px;font-weight:600;
           text-transform:uppercase;letter-spacing:0.5px;">{label}</div>
      {sub_html}
    </div>"""


def four_quarter_chart_html(trend_df, year, q):
    """Grouped column chart: towns on x-axis, 4 quarter datasets (oldest → newest).

    Shows avg attendance per event for each town across 4 consecutive quarters.
    """
    # Build the 4 quarters in chronological order (oldest first)
    quarters = []
    cy, cq = year, q
    for _ in range(4):
        quarters.insert(0, (cy, cq))
        cy, cq = prior_quarter(cy, cq)

    town_labels = [TOWN_LABELS[c] for c in TOWNS_ORDER]

    # Alpha shading: oldest = most faded, newest = full colour
    alphas = ["44", "77", "AA", ""]  # for 4 quarters oldest→newest

    datasets = []
    for i, (qy, qq) in enumerate(quarters):
        months = quarter_months(qy, qq)
        avgs = []
        for code in TOWNS_ORDER:
            town_df = trend_df[trend_df["town_code"] == code] if not trend_df.empty else pd.DataFrame()
            vals = []
            for m in months:
                row = town_df[town_df["event_month"] == m] if not town_df.empty else pd.DataFrame()
                if not row.empty:
                    v = pd.to_numeric(row.iloc[0].get("unique_attendees"), errors="coerce")
                    if pd.notna(v):
                        vals.append(float(v))
            avgs.append(round(sum(vals) / len(vals), 1) if vals else 0)

        ql = quarter_label(qy, qq)
        bg_colours = [TOWN_COLOURS[c] + alphas[i] for c in TOWNS_ORDER]
        datasets.append({
            "label": ql,
            "data": avgs,
            "backgroundColor": bg_colours,
            "borderRadius": 4,
            "borderWidth": 0,
        })

    labels_js   = json.dumps(town_labels)
    datasets_js = json.dumps(datasets)

    return f"""
    <canvas id="qBarChart" style="width:100%;max-height:300px;"></canvas>
    <script src="https://cdnjs.cloudflare.com/ajax/libs/Chart.js/4.4.0/chart.umd.min.js"></script>
    <script>
    new Chart(document.getElementById('qBarChart'), {{
      type: 'bar',
      data: {{ labels: {labels_js}, datasets: {datasets_js} }},
      options: {{
        responsive: true,
        plugins: {{
          legend: {{ position: 'bottom', labels: {{ font: {{ size: 11 }}, boxWidth: 12 }} }},
          tooltip: {{ mode: 'index', intersect: false }}
        }},
        scales: {{
          y: {{
            beginAtZero: true,
            grid: {{ color: '#f0f0f0' }},
            ticks: {{ font: {{ size: 11 }} }},
            title: {{ display: true, text: 'Avg attendees per event', font: {{ size: 11 }} }}
          }},
          x: {{ grid: {{ display: false }}, ticks: {{ font: {{ size: 11 }} }} }}
        }}
      }}
    }});
    </script>
    <div style="margin-top:8px;font-size:10px;color:{GREY};text-align:center;">
      4-quarter comparison — lighter bars = earlier quarters &nbsp;|&nbsp;
      Event Excellence target: {EE_PAYING_TARGET} avg guests
    </div>"""


def month_sparkline(vals, colour, months):
    """Mini bar sparkline for 3 monthly values."""
    mx = max((v for v in vals if v is not None), default=1) or 1
    bars = ""
    for i, v in enumerate(vals):
        if v is None:
            bars += f'<div title="No event {months[i]}" style="width:14px;height:4px;background:#e5e7eb;border-radius:2px;"></div>'
        else:
            h = max(4, int((v / mx) * 36))
            m_label = datetime.strptime(months[i] + "-01", "%Y-%m-%d").strftime("%b")
            bars += (
                f'<div style="display:flex;flex-direction:column;align-items:center;gap:1px;">'
                f'<span style="font-size:8px;color:{GREY};">{v}</span>'
                f'<div title="{m_label}: {v}" style="width:14px;height:{h}px;background:{colour};border-radius:2px 2px 0 0;"></div>'
                f'<span style="font-size:8px;color:{GREY};">{m_label}</span>'
                f'</div>'
            )
    return f'<div style="display:flex;align-items:flex-end;gap:4px;margin-top:4px;">{bars}</div>'


def town_row_html(t, q_months):
    rag_col, rag_lbl = _rag(t["rag"])
    delta_str = ""
    if t["delta"] is not None:
        sign = "+" if t["delta"] > 0 else ""
        col  = LIME if t["delta"] > 1 else (PINK if t["delta"] < -1 else ORANGE)
        delta_str = f'<span style="color:{col};font-weight:700;">{sign}{t["delta"]}</span>'

    cur_avg_str  = str(t["cur_avg"])  if t["cur_avg"]  is not None else "—"
    prev_avg_str = str(t["prev_avg"]) if t["prev_avg"] is not None else "—"

    ambs_str = ""
    for a in t["ambs"]:
        tick = "✓" if a["trained"] else "·"
        col  = LIME if a["trained"] else GREY
        ambs_str += f'<span style="color:{col};">{tick} {a["name"]}</span><br>'

    host_name = t["host_info"].get("name", "") if t["host_info"] else ""
    caretaker = t["host_info"].get("caretaker", False) if t["host_info"] else False
    host_str  = host_name
    if caretaker and host_name:
        host_str += ' <span style="font-size:9px;background:#FFF3CD;color:#856404;padding:1px 4px;border-radius:3px;">caretaker</span>'

    sparkline = month_sparkline(t["cur_vals"], t["colour"], q_months)

    # Star rating — compute colour outside f-string to avoid backslash in expression
    star_str = t.get("ee_star", "—")
    star_col = LIME if ("\u2605" in star_str or "\u2728" in star_str) else GREY

    no_amb_html = f'<span style="color:{GREY};">—</span>'

    return f"""
    <tr style="border-bottom:1px solid #f0f0f0;">
      <td style="padding:10px 6px;min-width:130px;">
        <div style="display:flex;align-items:center;gap:6px;">
          <span style="width:10px;height:10px;border-radius:50%;background:{t['colour']};display:inline-block;flex-shrink:0;"></span>
          <strong style="font-size:12px;color:{DARK};">{t['label']}</strong>
        </div>
        <div style="font-size:10px;color:{GREY};margin-top:2px;">{host_str}</div>
      </td>
      <td style="padding:10px 6px;text-align:center;">
        {sparkline}
      </td>
      <td style="padding:10px 6px;text-align:center;font-size:14px;font-weight:800;color:{t['colour']};">{cur_avg_str}</td>
      <td style="padding:10px 6px;text-align:center;font-size:12px;color:{GREY};">{prev_avg_str}</td>
      <td style="padding:10px 6px;text-align:center;font-size:12px;">{delta_str}</td>
      <td style="padding:10px 6px;text-align:center;">
        <span style="background:{rag_col};color:#fff;padding:2px 8px;border-radius:12px;font-size:10px;font-weight:700;">{rag_lbl}</span>
      </td>
      <td style="padding:10px 6px;font-size:10px;color:{DARK};line-height:1.6;">{ambs_str or no_amb_html}</td>
      <td style="padding:10px 6px;text-align:center;font-size:12px;">
        {_dot(t['ee_paying'])} {_dot(t['ee_ambs'])} {_dot(t['ee_combined'])}
        <div style="font-size:12px;font-weight:800;margin-top:3px;color:{star_col};">{star_str}</div>
      </td>
    </tr>"""


def _role_badge(role):
    """Return a small HTML pill badge for a team role, or empty string."""
    if not role:
        return ""
    label  = {"regional lead": "RL", "host": "Host", "ambassador": "Amb"}.get(role, role.title())
    colour = {"regional lead": TEAL, "host": ORANGE, "ambassador": LIME}.get(role, GREY)
    return (
        f'<span style="background:{colour}22;color:{colour};border:1px solid {colour}44;'
        f'padding:1px 5px;border-radius:3px;font-size:9px;font-weight:700;'
        f'margin-left:5px;vertical-align:middle;">{label}</span>'
    )


def cross_town_rows(cross_town):
    rows = ""
    for p in cross_town:
        town_badges = ""
        for tc_name in p["towns"].split(", "):
            code = next((c for c, l in TOWN_LABELS.items() if l == tc_name.strip()), None)
            col  = TOWN_COLOURS.get(code, GREY) if code else GREY
            short = TOWN_SHORT.get(code, tc_name[:2]) if code else tc_name[:2]
            town_badges += (
                f'<span style="background:{col}22;color:{col};padding:1px 5px;'
                f'border-radius:3px;font-size:10px;font-weight:700;margin-right:3px;">{short}</span>'
            )
        name_cell = p['name'] + _role_badge(p.get('role', ''))
        rows += f"""
        <tr style="border-bottom:1px solid #f0f0f0;">
          <td style="padding:8px 6px;font-weight:600;font-size:12px;">{name_cell}</td>
          <td style="padding:8px 6px;font-size:11px;color:{GREY};">{p['company']}</td>
          <td style="padding:8px 6px;text-align:center;font-weight:700;">{p['visits']}</td>
          <td style="padding:8px 6px;">{town_badges}</td>
        </tr>"""
    return rows


def regional_sponsors_html(rsp):
    sp_max   = 4
    sp_names = rsp.get("sp_names", [])
    ch_name  = rsp.get("ch_name", "")
    sp_slots = ""
    for i in range(sp_max):
        if i < len(sp_names):
            sp_slots += (
                f'<span style="display:inline-block;padding:3px 10px;margin:2px;'
                f'background:{TEAL}22;color:{TEAL};border:1px solid {TEAL}44;'
                f'border-radius:5px;font-size:11px;font-weight:600;">{sp_names[i]}</span>'
            )
        else:
            sp_slots += (
                f'<span style="display:inline-block;padding:3px 10px;margin:2px;'
                f'background:#f9f9f9;color:{GREY};border:1px dashed #ddd;'
                f'border-radius:5px;font-size:11px;">Slot {i+1} available</span>'
            )
    ch_html = (
        f'<span style="padding:3px 10px;background:{LIME}22;color:{LIME};'
        f'border:1px solid {LIME}44;border-radius:5px;font-size:11px;font-weight:600;">{ch_name}</span>'
    ) if ch_name else (
        f'<span style="padding:3px 10px;background:#f9f9f9;color:{GREY};'
        f'border:1px dashed #ddd;border-radius:5px;font-size:11px;">'
        f'Not yet confirmed</span>'
    )
    return f"""
  <div class="section">
    <div class="section-title">Regional sponsors &amp; charity partner</div>
    <div style="display:flex;gap:20px;flex-wrap:wrap;">
      <div style="flex:2;min-width:220px;">
        <div style="font-size:11px;color:{GREY};margin-bottom:6px;font-weight:600;">
          Sponsors ({len(sp_names)}/{sp_max} slots)</div>
        <div>{sp_slots}</div>
      </div>
      <div style="flex:1;min-width:160px;">
        <div style="font-size:11px;color:{GREY};margin-bottom:6px;font-weight:600;">Charity partner</div>
        <div>{ch_html}</div>
      </div>
    </div>
  </div>"""


def auto_observations(towns, summary, year, q):
    obs = []

    # New events — positive note first
    new_towns = [t for t in towns if t.get("is_new")]
    if new_towns:
        names = ", ".join(t["label"] for t in new_towns)
        obs.append((LIME, f"New events — {names}",
            f"Still in early build phase. Evaluation covers their months within the Nov\u2013Oct "
            f"conference period. Focus on consistency and first sponsor conversations."))

    # Strong towns
    strong = [t for t in towns if t["rag"] == "green" and not t.get("is_new")]
    if strong:
        names = ", ".join(t["label"] for t in strong)
        obs.append((LIME, "Performing above target",
            f"{names} averaging above {EE_PAYING_TARGET} this quarter. "
            f"On track for Event Excellence progression \u2014 focus now on sponsor and Buzz Plus to move up the star rating."))

    # Growth (exclude new events — not enough prior history)
    gainers = [t for t in towns if t["delta"] is not None and t["delta"] >= 1 and not t.get("is_new")]
    if gainers:
        names = ", ".join(f"{t['label']} (+{t['delta']})" for t in gainers)
        obs.append((TEAL, "Quarter-on-quarter growth",
            f"{names} \u2014 momentum is building. Explore what\u2019s working and share across other towns."))

    # Declining (exclude new events)
    decliners = [t for t in towns if t["delta"] and t["delta"] <= -3 and not t.get("is_new")]
    if decliners:
        parts = []
        for t in decliners:
            is_ct = t.get("host_info", {}).get("caretaker", False)
            if is_ct:
                parts.append(f"{t['label']} ({t['delta']}) \u2014 review attendee re-engagement and new face pipeline")
            else:
                host_name = t.get("host_info", {}).get("name", "the host")
                parts.append(f"{t['label']} ({t['delta']}) \u2014 connect with {host_name} on re-engagement and new face conversion")
        obs.append((PINK, "Quarter-on-quarter decline", "; ".join(parts) + "."))

    # Buzz Plus — show all 5 towns
    plus_parts = []
    for t in towns:
        if t.get("is_new") and t["strong_plus"] == 0:
            plus_parts.append(f"{TOWN_SHORT[t['code']]}: building (new event)")
        elif t["active_plus"] > 0:
            plus_parts.append(f"{TOWN_SHORT[t['code']]}: {t['active_plus']} active, {t['strong_plus']} strong prospects")
        else:
            plus_parts.append(f"{TOWN_SHORT[t['code']]}: {t['strong_plus']} strong, {t.get('possible_plus',0)} possible, 0 active")
    obs.append((TEAL, "Buzz Plus position \u2014 all events",
        " \u00b7 ".join(plus_parts) + ". "
        "Brief hosts on named strong prospects before events and make warm introductions."))

    # Sponsor position — show all 5 towns
    sp_parts = []
    for t in towns:
        new_flag = " (new)" if t.get("is_new") else ""
        sp_parts.append(f"{TOWN_SHORT[t['code']]}: {t['sponsors']}/{SPONSOR_SLOTS}{new_flag}")
    obs.append((ORANGE, "Sponsor position \u2014 all events",
        " \u00b7 ".join(sp_parts) + f". "
        f"Event Excellence needs Buzz Plus + sponsors combined ≥{EE_COMBINED_MIN}. "
        f"Use existing regional sponsors as social proof when approaching new ones."))

    # Ambassador training
    under_trained = [t for t in towns if t["trained_count"] < EE_AMBASSADOR_MIN]
    if under_trained:
        parts = [f"{t['label']} ({t['trained_count']}/{EE_AMBASSADOR_MIN} trained)" for t in under_trained]
        obs.append((ORANGE, "Ambassador training",
            f"{', '.join(parts)} \u2014 below the threshold for 4-star Event Excellence. "
            f"Training completion needed before the awards window (Oct) closes."))

    return obs[:7]


def render_html(towns, summary, cross_town, regional_sp,
                year, q, prev_year, prev_q, trend_df, generated,
                is_in_progress=False):
    ql  = quarter_label(year, q)
    pql = quarter_label(prev_year, prev_q)
    q_months  = quarter_months(year, q)
    pq_months = quarter_months(prev_year, prev_q)

    chart = four_quarter_chart_html(trend_df, year, q)

    # Snap cards
    region_avg_str = f"{summary['region_avg']}" if summary["region_avg"] is not None else "—"
    best_str  = f"{summary['best']['label']} ({summary['best']['cur_avg']})"  if summary["best"]  else "—"
    worst_str = f"{summary['worst']['label']} ({summary['worst']['cur_avg']})" if summary["worst"] else "—"
    snap_row = "".join([
        snap_card(TEAL,   summary["total_events"],  "Events in quarter"),
        snap_card(ORANGE, summary["total_visits"],  "Total event visits"),
        snap_card(PINK,   summary["total_people"],  "People in region", sub="unduplicated"),
        snap_card(LIME,   region_avg_str,            "Region avg / event"),
        snap_card(TEAL,   best_str,                  "Best performing town"),
    ])

    # Town performance table
    town_rows_html = "".join(town_row_html(t, q_months) for t in towns)

    # Observations
    observations = auto_observations(towns, summary, year, q)
    obs_html = ""
    for i, (col, title, body) in enumerate(observations, 1):
        obs_html += f"""
        <div style="display:flex;gap:12px;align-items:flex-start;padding:12px 0;
             {'border-bottom:1px solid #f0f0f0;' if i < len(observations) else ''}">
          <div style="width:28px;height:28px;border-radius:50%;background:{col};color:#fff;
               display:flex;align-items:center;justify-content:center;font-weight:800;
               font-size:12px;flex-shrink:0;">{i}</div>
          <div>
            <div style="font-weight:700;font-size:13px;color:{DARK};margin-bottom:3px;">{title}</div>
            <div style="font-size:12px;color:{GREY};line-height:1.5;">{body}</div>
          </div>
        </div>"""

    # Cross-town
    ct_rows = cross_town_rows(cross_town)
    rsp_html = regional_sponsors_html(regional_sp)

    return f"""<!DOCTYPE html>
<html lang="en">
<head>
<meta charset="UTF-8">
<meta name="viewport" content="width=device-width, initial-scale=1.0">
<title>Business Buzz — {ql} Regional Review</title>
<style>
  * {{ box-sizing:border-box; margin:0; padding:0; }}
  body {{ font-family:'Century Gothic','Gill Sans',Calibri,sans-serif; background:{LGREY}; color:{DARK}; }}
  .container {{ max-width:1100px; margin:0 auto; padding:24px 20px; }}
  .section {{ background:#fff; border-radius:12px; padding:20px 22px;
              box-shadow:0 2px 8px rgba(0,0,0,0.06); margin-bottom:20px; }}
  .section-title {{ font-size:13px; font-weight:800; color:{TEAL};
                    text-transform:uppercase; letter-spacing:1px; margin-bottom:14px; }}
  table {{ width:100%; border-collapse:collapse; }}
  th {{ text-align:left; font-size:11px; color:{GREY}; font-weight:700;
        text-transform:uppercase; padding:6px 6px 10px; letter-spacing:0.5px; border-bottom:2px solid #f0f0f0; }}
  @media print {{
    body {{ background:#fff; }}
    .section {{ box-shadow:none; border:1px solid #e5e7eb; page-break-inside:avoid; }}
  }}
</style>
</head>
<body>
<div class="container">

  <!-- HEADER -->
  <div style="background:{ORANGE};color:#fff;border-radius:12px;padding:20px 24px;
       margin-bottom:{'8px' if is_in_progress else '20px'};display:flex;justify-content:space-between;align-items:center;">
    <div>
      <div style="font-size:10px;font-weight:700;letter-spacing:2px;text-transform:uppercase;
           opacity:0.7;margin-bottom:4px;">Business Buzz · Leicestershire &amp; Rutland</div>
      <div style="font-size:24px;font-weight:800;">{ql} Regional Review</div>
      <div style="font-size:13px;opacity:0.8;margin-top:4px;">
        4-quarter trend view &nbsp;·&nbsp; 5 active towns
      </div>
    </div>
    <div style="text-align:right;font-size:12px;opacity:0.8;">
      <div style="font-size:14px;font-weight:700;opacity:1;color:#fff;margin-bottom:4px;">
        Awards window: Nov–Oct
      </div>
      <div>Event Excellence criteria: ≥{EE_PAYING_TARGET} guests · {EE_AMBASSADOR_MIN} trained ambs</div>
      <div>Buzz Plus + sponsors combined ≥{EE_COMBINED_MIN}</div>
      <div style="margin-top:8px;font-size:11px;">Generated {generated}</div>
    </div>
  </div>

  {'<!-- IN PROGRESS BANNER --><div style="background:#FFF8EC;border:1px solid ' + ORANGE + '44;border-radius:8px;padding:10px 18px;margin-bottom:20px;display:flex;align-items:center;gap:10px;"><div style="width:10px;height:10px;border-radius:50%;background:' + ORANGE + ';flex-shrink:0;animation:pulse 1.5s infinite;"></div><div style="font-size:12px;color:#9A5800;"><strong>Quarter in progress</strong> — data shown is live and will update as events run. Final figures available once ' + ql + ' closes at end of ' + ("March" if q==1 else "June" if q==2 else "September" if q==3 else "December") + ' ' + str(year) + '.</div></div><style>@keyframes pulse{0%,100%{opacity:1;}50%{opacity:0.4;}}</style>' if is_in_progress else ''}

  <!-- REGIONAL SNAPSHOT -->
  <div class="section">
    <div class="section-title">Quarter at a glance — {ql}</div>
    <div style="display:flex;gap:10px;flex-wrap:wrap;">{snap_row}</div>
  </div>

  <!-- 4-QUARTER TREND CHART -->
  <div class="section">
    <div class="section-title">4-quarter attendance trend — average attendees per event</div>
    {chart}
  </div>

  <!-- TOWN PERFORMANCE TABLE -->
  <div class="section">
    <div class="section-title">Town performance detail</div>
    <div style="font-size:11px;color:{GREY};margin-bottom:10px;">
      ✓ = ambassador training complete &nbsp;|&nbsp;
      Event Excellence dots: guests · ambassadors · Buzz Plus · sponsors
    </div>
    <table>
      <tr>
        <th>Town / Host</th>
        <th style="text-align:center;">{ql} monthly</th>
        <th style="text-align:center;">{ql} avg</th>
        <th style="text-align:center;">{pql} avg</th>
        <th style="text-align:center;">Change</th>
        <th style="text-align:center;">Status</th>
        <th>Ambassadors</th>
        <th style="text-align:center;">Rating</th>
      </tr>
      {town_rows_html}
    </table>
  </div>

  <!-- OBSERVATIONS & ACTIONS -->
  <div class="section">
    <div class="section-title">Key observations &amp; actions — {ql}</div>
    {obs_html}
  </div>

  {rsp_html}

  <!-- CROSS-TOWN CHAMPIONS -->
  <div class="section">
    <div class="section-title">Cross-town engagement — most active across the region</div>
    <div style="font-size:11px;color:{GREY};margin-bottom:12px;">
      People attending 2+ towns — strongest ambassador and Buzz Plus prospects.
    </div>
    <table>
      <tr>
        <th>Name</th><th>Company</th>
        <th style="text-align:center;">Region visits</th>
        <th>Towns attended</th>
      </tr>
      {ct_rows}
    </table>
  </div>

  <!-- FOOTER -->
  <div style="text-align:center;font-size:11px;color:{GREY};padding:16px 0;">
    Business Buzz · Leicestershire &amp; Rutland · {ql} Review · Generated {generated}
  </div>

</div>
</body>
</html>"""


# ══════════════════════════════════════════════════════════════════════════════
# MAIN
# ══════════════════════════════════════════════════════════════════════════════

def main():
    parser = argparse.ArgumentParser()
    parser.add_argument("--year",    type=int, default=None)
    parser.add_argument("--quarter", type=int, default=None, choices=[1, 2, 3, 4])
    args = parser.parse_args()

    if args.year and args.quarter:
        year, q = args.year, args.quarter
    else:
        # Default to the current quarter (in progress if run mid-quarter)
        year, q = current_quarter()

    prev_year, prev_q = prior_quarter(year, q)
    generated = datetime.now().strftime("%d %B %Y").lstrip("0")
    ql = quarter_label(year, q)

    # A quarter is in progress if we're currently inside it
    is_in_progress = (year, q) == current_quarter()

    print(f"[Quarterly Dashboard] Building {ql} {'(in progress)' if is_in_progress else '(completed)'} (vs {quarter_label(prev_year, prev_q)})")

    overview, trend_df = load_trend()
    master      = load_master()
    ambassadors = load_ambassadors()
    hosts       = load_hosts()
    plus_counts, plus_strong, plus_possible = load_plus_by_town()
    regional_sp = load_regional_sponsors()

    towns      = build_town_quarters(trend_df, overview, year, q, prev_year, prev_q,
                                     ambassadors, hosts, plus_counts,
                                     plus_strong=plus_strong, plus_possible=plus_possible)
    summary    = build_region_summary(towns, year, q, trend_df, master)
    cross_town = build_cross_town(master)

    html = render_html(towns, summary, cross_town, regional_sp,
                       year, q, prev_year, prev_q, trend_df, generated,
                       is_in_progress=is_in_progress)

    out = CURATED / f"Region_Quarterly_Dashboard_{year}_Q{q}.html"
    out.parent.mkdir(parents=True, exist_ok=True)
    out.write_text(html, encoding="utf-8")
    print(f"[Quarterly Dashboard] Written to {out}")


if __name__ == "__main__":
    main()
