"""
Buzz_Region_MonthlyDashboard.py
================================
Generates a single-page HTML regional dashboard for:
  - Monthly RL review (all 5 towns at a glance without attending each one)
  - Founder/HQ meetings (clean, data-led, boardroom-ready)

Output:
  Buzz_Region_Curated/Region_Monthly_Dashboard.html

Run as part of Buzz_Region_AllBuild.py (after Dashboard and 321 intelligence scripts).
"""

import json
import os
from datetime import datetime
from pathlib import Path

import pandas as pd
from dateutil.relativedelta import relativedelta

# ── Paths ──────────────────────────────────────────────────────────────────────
BASE            = Path(__file__).resolve().parent
CURATED         = BASE / "Buzz_Region_Curated"
OUT_HTML        = CURATED / "Region_Monthly_Dashboard.html"

# ── Brand colours ──────────────────────────────────────────────────────────────
TEAL   = "#00A19A"
ORANGE = "#F39200"
PINK   = "#D60B52"
LIME   = "#B6BD00"
DARK   = "#1A1A2E"
GREY   = "#6B7280"
LGREY  = "#F3F4F6"

# ── Town display order ─────────────────────────────────────────────────────────
TOWNS_ORDER  = ["MarketHarborough", "Leicester", "Loughborough", "Lutterworth", "Hinckley"]
TOWN_LABELS  = {
    "MarketHarborough": "Market Harborough",
    "Leicester":        "Leicester",
    "Loughborough":     "Loughborough",
    "Lutterworth":      "Lutterworth",
    "Hinckley":         "Hinckley",
}
TOWN_SHORT   = {
    "MarketHarborough": "MH",
    "Leicester":        "LE",
    "Loughborough":     "LB",
    "Lutterworth":      "LW",
    "Hinckley":         "HI",
}
# Chart colours per town
TOWN_COLOURS = {
    "MarketHarborough": TEAL,
    "Leicester":        ORANGE,
    "Loughborough":     LIME,
    "Lutterworth":      PINK,
    "Hinckley":         "#8B5CF6",
}

# ── Event Excellence thresholds (HQ criteria) ─────────────────────────────────
# 3★: avg ≥25 paying guests (best 10 months Nov–Oct)
# 4★: 3★ + 2 trained ambassadors active ≥3 months
# 5★: 4★ + combined Buzz Plus + sponsors ≥4 (any mix)
# Special ✨: 5★ + 4 sponsors AND 5 Buzz Plus
EE_PAYING_TARGET    = 25   # avg paying guests for 3-star+
EE_AMBASSADOR_MIN   = 2    # trained ambassadors for 4-star+
EE_COMBINED_MIN     = 4    # combined Buzz Plus + sponsors for 5-star
EE_PLUS_FULL        = 5    # full Buzz Plus complement for special prize
EE_SPONSOR_FULL     = 4    # full sponsor complement for special prize


def event_excellence_star(ee_paying, ee_ambassadors, active_plus=0, sponsors=0, is_new=False):
    """Return (star_str, description) based on HQ star criteria.

    New events are still evaluated against HQ criteria — the conference period
    (Nov–Oct) covers whatever months they have run so far. is_new adds a note
    to the description but does NOT bypass scoring.
    """
    new_sfx = " · new event, partial period data" if is_new else ""
    if not ee_paying:
        return "—", f"Working toward 3\u2605 (avg below {EE_PAYING_TARGET}){new_sfx}"
    if ee_ambassadors and (active_plus + sponsors) >= EE_COMBINED_MIN:
        if active_plus >= EE_PLUS_FULL and sponsors >= EE_SPONSOR_FULL:
            return "\u2728 5\u2605", f"Special prize eligible{new_sfx}"
        return "5\u2605", f"5-star Event Excellence{new_sfx}"
    if ee_ambassadors:
        return "4\u2605", f"4-star Event Excellence{new_sfx}"
    return "3\u2605", f"3-star Event Excellence{new_sfx}"


# ══════════════════════════════════════════════════════════════════════════════
# DATA LOADING
# ══════════════════════════════════════════════════════════════════════════════

def _safe_read(path, sheet=None, **kwargs):
    try:
        if sheet:
            return pd.read_excel(path, sheet_name=sheet, **kwargs)
        return pd.read_excel(path, **kwargs)
    except Exception:
        return pd.DataFrame()


def _nc(df):
    """Normalise column names to lowercase stripped."""
    if not df.empty:
        df.columns = [str(c).strip().lower() for c in df.columns]
    return df


def load_town_overview():
    return _nc(_safe_read(CURATED / "region_dashboard.xlsx", sheet="Town_Overview"))


def load_attendance_trend():
    tr = _nc(_safe_read(CURATED / "region_dashboard.xlsx", sheet="Attendance_Trend_12m"))
    # Normalise event_month to YYYY-MM string regardless of how Excel stored the date
    if not tr.empty and "event_month" in tr.columns:
        tr["event_month"] = tr["event_month"].astype(str).str[:7]
    return tr


def load_sponsor_capacity():
    return _nc(_safe_read(CURATED / "sponsor_intelligence.xlsx", sheet="Sponsor_Capacity"))


def load_sponsor_attendance():
    return _nc(_safe_read(CURATED / "sponsor_intelligence.xlsx", sheet="Sponsor_Attendance"))


def load_master_people():
    path = CURATED / "region_master_people.csv"
    if not path.exists():
        return pd.DataFrame()
    df = pd.read_csv(path)
    return _nc(df)


def load_321_town_summary():
    return _nc(_safe_read(CURATED / "buzz_321_intelligence.xlsx", sheet="Town_Month_Summary"))


def load_active_plus_by_town():
    """Read active Buzz Plus members and strong prospects directly from source files.

    Active members: buzz_plus_members.xlsx (bypasses stale rollup cache).
    Strong prospects: buzzplus_intelligence.xlsx / Town_Plus_Summary (live data).
    """
    # Active members
    counts = {}
    members = {}
    try:
        df = pd.read_excel(BASE / "Buzz_Region_Ref" / "buzz_plus_members.xlsx")
        df.columns = [str(c).strip().lower() for c in df.columns]
        if "status" in df.columns:
            df = df[df["status"].astype(str).str.strip().str.lower() == "active"]
        if "town_code" in df.columns:
            counts = df.groupby("town_code").size().to_dict()
            for tc, grp in df.groupby("town_code"):
                names = [str(r.get("member_name","") or "").strip()
                         for _, r in grp.iterrows()
                         if str(r.get("member_name","")).strip().lower() not in ("", "nan")]
                members[tc] = names
    except Exception:
        pass

    # Strong prospects — load from buzzplus_intelligence.xlsx Town_Plus_Summary
    strong = {}
    possible = {}
    try:
        bp = pd.read_excel(CURATED / "buzzplus_intelligence.xlsx",
                           sheet_name="Town_Plus_Summary")
        bp.columns = [str(c).strip().lower() for c in bp.columns]
        if "town_code" in bp.columns:
            for _, r in bp.iterrows():
                tc = str(r.get("town_code", "")).strip()
                strong[tc]   = int(pd.to_numeric(r.get("buzzplus_strong_prospects", 0), errors="coerce") or 0)
                possible[tc] = int(pd.to_numeric(r.get("buzzplus_possible_prospects", 0), errors="coerce") or 0)
    except Exception:
        pass

    return counts, members, strong, possible


def load_ambassadors():
    """Load ambassador reference file — name, town, training status."""
    path = BASE / "Buzz_Region_Ref" / "buzz_ambassadors.xlsx"
    df = _nc(_safe_read(path))
    by_town = {}
    if not df.empty and "town_code" in df.columns:
        for tc, grp in df.groupby("town_code"):
            amb_list = []
            for _, r in grp.iterrows():
                name = str(r.get("name","") or "").strip()
                if not name or name.lower() == "nan":
                    continue
                trained_raw = str(r.get("ambassador_training_complete","")).strip().lower()
                trained = trained_raw in ("yes", "true", "1")
                has_left = str(r.get("has_left","")).strip().lower() in ("yes","true","1")
                if not has_left:
                    amb_list.append({"name": name, "trained": trained})
            by_town[tc] = amb_list
    return by_town


def load_hosts():
    """Load host reference file — includes caretaker flag and new event flag."""
    path = BASE / "Buzz_Region_Ref" / "buzz_hosts.xlsx"
    df = _nc(_safe_read(path))
    hosts = {}
    if not df.empty and "town_code" in df.columns:
        for _, r in df.iterrows():
            tc = str(r.get("town_code","")).strip()
            is_ct  = str(r.get("is_caretaker","")).strip().lower() in ("yes","true","1")
            is_new = str(r.get("is_new_event","")).strip().lower() in ("yes","true","1")
            if is_ct:
                name = str(r.get("caretaker_name","") or "").strip()
                hosts[tc] = {"name": name, "caretaker": True, "new_event": is_new}
            else:
                name = str(r.get("host_name","") or "").strip()
                if name and name.lower() != "nan":
                    hosts[tc] = {"name": name, "caretaker": False, "new_event": is_new}
    return hosts


def load_regional_sponsors():
    """Load regional-level sponsor + charity files."""
    sp_path = BASE / "Buzz_Region_Ref" / "buzz_regional_sponsors.xlsx"
    ch_path = BASE / "Buzz_Region_Ref" / "buzz_regional_charity.xlsx"
    sponsors = _nc(_safe_read(sp_path))
    charity  = _nc(_safe_read(ch_path))
    sp_count  = len(sponsors) if not sponsors.empty else 0
    ch_active = not charity.empty and len(charity) > 0 and any(
        str(r.get("charity_name","")).strip() for _, r in charity.iterrows()
        if str(r.get("charity_name","")).strip() not in ("","nan")
    )
    sp_names = []
    if not sponsors.empty and "sponsor_name" in sponsors.columns:
        sp_names = [str(n).strip() for n in sponsors["sponsor_name"] if str(n).strip() not in ("","nan")]
    ch_name = ""
    if not charity.empty and "charity_name" in charity.columns:
        rows = [str(n).strip() for n in charity["charity_name"] if str(n).strip() not in ("","nan")]
        ch_name = rows[0] if rows else ""
    return {"sp_count": sp_count, "sp_max": 4, "sp_names": sp_names,
            "ch_active": ch_active, "ch_name": ch_name}


# ══════════════════════════════════════════════════════════════════════════════
# DATA ASSEMBLY
# ══════════════════════════════════════════════════════════════════════════════

def _val(df, col, default=0):
    if df.empty or col not in df.columns:
        return default
    v = pd.to_numeric(df.iloc[0][col], errors="coerce")
    return int(v) if pd.notna(v) else default


def _fval(df, col, default=0.0):
    if df.empty or col not in df.columns:
        return default
    v = pd.to_numeric(df.iloc[0][col], errors="coerce")
    return round(float(v), 1) if pd.notna(v) else default


def _sval(df, col, default=""):
    if df.empty or col not in df.columns:
        return default
    return str(df.iloc[0].get(col, default) or default).strip()


SPONSOR_TOTAL_SLOTS = 4  # Each town has 4 sponsor slots; EE needs ≥2

def build_town_data(overview, trend, sponsor_cap, sponsor_att, data_321,
                    plus_counts=None, plus_members=None, plus_strong=None,
                    plus_possible=None, ambassadors=None, hosts=None):
    towns = []
    if plus_counts is None:   plus_counts   = {}
    if plus_members is None:  plus_members  = {}
    if plus_strong is None:   plus_strong   = {}
    if plus_possible is None: plus_possible = {}
    if ambassadors is None:   ambassadors   = {}
    if hosts is None:         hosts         = {}

    for code in TOWNS_ORDER:
        label  = TOWN_LABELS[code]
        short  = TOWN_SHORT[code]
        colour = TOWN_COLOURS[code]

        ov  = overview[overview["town_code"] == code] if not overview.empty else pd.DataFrame()
        sp  = sponsor_cap[sponsor_cap["town_code"] == code] if not sponsor_cap.empty else pd.DataFrame()
        att = sponsor_att[sponsor_att["town_code"] == code] if not sponsor_att.empty else pd.DataFrame()
        tr  = trend[trend["town_code"] == code].sort_values("event_month") if not trend.empty else pd.DataFrame()
        d3  = data_321[data_321["town_code"] == code].sort_values("event_month_key") if not data_321.empty else pd.DataFrame()

        avg          = _fval(ov, "avg_unique_per_event_12m")
        num_events   = _val(ov, "num_events_12m")
        total_12m    = _val(ov, "total_unique_attendees_12m")
        # Load strong/possible prospects from buzzplus_intelligence directly
        strong_plus  = plus_strong.get(code, _val(ov, "buzzplus_strong_prospects"))
        poss_plus    = plus_possible.get(code, _val(ov, "buzzplus_possible_prospects"))
        # Active members: direct from buzz_plus_members (bypasses stale rollup cache)
        active_plus  = plus_counts.get(code, _val(ov, "active_plus_members"))
        plus_names   = plus_members.get(code, [])
        sponsors     = _val(ov, "current_sponsors")
        slots        = _val(ov, "capacity_left")
        att_focus    = _sval(ov, "attendance_focus")
        lockouts     = _sval(ov, "industry_lockouts", "None")

        # Host and ambassadors from reference files
        host_info = hosts.get(code, {})
        town_ambs = ambassadors.get(code, [])

        # Attendance trend (last 6 months)
        trend_vals   = []
        trend_months = []
        if not tr.empty and "unique_attendees" in tr.columns:
            t6 = tr.tail(6)
            trend_vals   = [int(v) for v in t6["unique_attendees"].tolist()]
            trend_months = list(t6["event_month"].tolist())

        # Last 3 values for arrow direction
        last3 = trend_vals[-3:] if len(trend_vals) >= 3 else trend_vals
        if len(last3) >= 2:
            delta = last3[-1] - last3[-2]
            arrow = "▲" if delta > 3 else ("▼" if delta < -3 else "→")
            arrow_col = LIME if delta > 3 else (PINK if delta < -3 else ORANGE)
        else:
            arrow, arrow_col = "→", GREY

        latest_count = trend_vals[-1] if trend_vals else 0
        latest_month = trend_months[-1] if trend_months else "—"

        # Sponsor compliance
        sponsor_flags = []
        if not att.empty:
            for _, row in att.iterrows():
                flag = str(row.get("compliance_flag", ""))
                company = str(row.get("sponsor_company", "")).strip()
                raw_pct = float(row.get("attendance_pct", 0) or 0)
                # If value is already stored as a percentage (>1) don't multiply again
                pct = round(raw_pct) if raw_pct > 1 else round(raw_pct * 100)
                if company and company.lower() != "nan":
                    sponsor_flags.append({"company": company, "flag": flag, "pct": pct})

        # 3-2-1 most recent data
        q_interaction = None
        q_met3        = None
        q_121s        = None
        q_brought     = None
        if not d3.empty:
            latest_321 = d3.iloc[-1]
            q_interaction = latest_321.get("avg_interaction_rate")
            q_met3        = latest_321.get("total_met_3_new")
            q_121s        = latest_321.get("total_one_2_ones")
            q_brought     = latest_321.get("total_brought_some1")

        # RAG
        if att_focus == "Strong" and latest_count >= EE_PAYING_TARGET:
            rag = "green"
        elif att_focus == "Low" or latest_count < 18:
            rag = "red"
        else:
            rag = "amber"

        # New event flag from host reference
        is_new = host_info.get("new_event", False)

        # Event Excellence criteria
        trained_count  = sum(1 for a in town_ambs if a["trained"])
        ee_paying      = avg >= EE_PAYING_TARGET
        ee_ambassadors = trained_count >= EE_AMBASSADOR_MIN
        ee_combined    = (active_plus + sponsors) >= EE_COMBINED_MIN
        ee_score       = sum([ee_paying, ee_ambassadors, ee_combined])
        ee_star, ee_star_desc = event_excellence_star(
            ee_paying, ee_ambassadors,
            active_plus=active_plus, sponsors=sponsors, is_new=is_new,
        )

        towns.append(dict(
            code=code, label=label, short=short, colour=colour,
            avg=avg, num_events=num_events, total_12m=total_12m,
            strong_plus=strong_plus, poss_plus=poss_plus, plus_possible=poss_plus,
            active_plus=active_plus, plus_names=plus_names,
            sponsors=sponsors, slots=slots, att_focus=att_focus, lockouts=lockouts,
            trend_vals=trend_vals, trend_months=trend_months,
            latest_count=latest_count, latest_month=latest_month,
            arrow=arrow, arrow_col=arrow_col,
            sponsor_flags=sponsor_flags,
            q_interaction=q_interaction, q_met3=q_met3,
            q_121s=q_121s, q_brought=q_brought,
            rag=rag,
            host_info=host_info, town_ambs=town_ambs, trained_count=trained_count,
            is_new=is_new,
            ee_paying=ee_paying, ee_ambassadors=ee_ambassadors,
            ee_combined=ee_combined, ee_score=ee_score,
            ee_star=ee_star, ee_star_desc=ee_star_desc,
        ))
    return towns


# Fallback hardcoded exclusions — live set is loaded from exclude_region.csv at runtime
CROSS_TOWN_EXCLUDE_EMAILS = {
    "warwickshire@business-buzz.org",  # James Brodie — HQ Regional Lead (Buddha Connect)
    "hello@andkarenhall.co.uk",        # Karen Hall — moved away from the area
}


def load_cross_town_exclusions():
    """Return set of emails to exclude from cross-town table.
    Merges exclude_region.csv with the hardcoded fallback set."""
    excluded = set(CROSS_TOWN_EXCLUDE_EMAILS)
    csv_path = BASE / "exclude_region.csv"
    if csv_path.exists():
        try:
            df = pd.read_csv(csv_path)
            df.columns = [c.strip().lower() for c in df.columns]
            if "email" in df.columns:
                emails = df["email"].astype(str).str.strip().str.lower()
                excluded |= set(emails[(emails != "") & (emails != "nan")])
        except Exception:
            pass
    return excluded


def build_cross_town(master, top_n=10):
    if master.empty:
        return []
    excluded_emails = load_cross_town_exclusions()
    team_roles = {"host", "ambassador", "regional lead"}
    excl = master["email"].astype(str).str.strip().str.lower().isin(excluded_emails)
    df = master[~excl & (master["towns_visited_count"] >= 2)].copy()
    df["visits_n"] = pd.to_numeric(df["visits_region"], errors="coerce").fillna(0)
    # Sort by total visits (primary), towns visited count (secondary)
    df = df.sort_values(["visits_n", "towns_visited_count"], ascending=False).head(top_n)
    result = []
    for _, row in df.iterrows():
        name     = str(row.get("name", "") or "").strip().title()
        company  = _clean_company(str(row.get("company", "") or ""))
        role_raw = str(row.get("role_region", "") or "").strip().lower()
        role     = role_raw if role_raw in team_roles else ""
        result.append(dict(
            name=name,
            company=company,
            visits=int(row["visits_n"]),
            towns_count=int(row.get("towns_visited_count", 0)),
            towns=str(row.get("towns_visited", "") or ""),
            role=role,
        ))
    return result


ACTIVE_MEMBER_MONTHS = 6  # Regional active = visited in last 6 months (town packs use 3)

def build_region_totals(overview, master):
    total_events_12m = 0
    # Sum is per-town totals (one person visiting 2 towns = counted twice here)
    total_visits_12m = 0
    if not overview.empty:
        total_events_12m  = int(overview["num_events_12m"].sum()) if "num_events_12m" in overview.columns else 0
        total_visits_12m  = int(overview["total_unique_attendees_12m"].sum()) if "total_unique_attendees_12m" in overview.columns else 0

    multi_town    = 0
    total_people  = 0   # unduplicated region headcount
    active_members = 0  # visited in last 6 months
    if not master.empty:
        team_roles = {"host", "ambassador", "regional lead"}
        is_team = master["role_region"].astype(str).str.strip().str.lower().isin(team_roles)
        members = master[~is_team]
        total_people   = len(members)
        multi_town     = int((pd.to_numeric(members.get("towns_visited_count", pd.Series()), errors="coerce") >= 2).sum())
        # Active: used visits_region as proxy — 6+ visits in rolling 12m is a reasonable 6-month threshold
        active_members = int((pd.to_numeric(members.get("visits_region", pd.Series()), errors="coerce") >= ACTIVE_MEMBER_MONTHS).sum())

    return dict(
        total_events_12m=total_events_12m,
        total_visits_12m=total_visits_12m,
        total_people=total_people,
        multi_town=multi_town,
        active_members=active_members,
    )


# ══════════════════════════════════════════════════════════════════════════════
# HTML GENERATION
# ══════════════════════════════════════════════════════════════════════════════

import re as _re

def _clean_company(name):
    """Strip address fragments from company names (e.g. 'AM Joinery Ltd, Gilbert Avenue, Rugby')."""
    if not name or str(name).strip().lower() in ("", "nan"):
        return ""
    name = str(name).strip()
    # If the name contains a comma followed by what looks like a street/place, truncate there
    # Heuristic: more than one comma → likely address appended
    parts = name.split(",")
    if len(parts) > 2:
        # Keep only up to first comma
        return parts[0].strip()
    if len(parts) == 2:
        # Check if second part looks like an address (number + word, or known place names)
        second = parts[1].strip()
        if _re.match(r'^\d+\s', second) or len(second) > 30:
            return parts[0].strip()
    return name


def _rag_colour(rag):
    return {"green": LIME, "amber": ORANGE, "red": PINK}.get(rag, GREY)


def _rag_label(rag):
    return {"green": "Strong", "amber": "Steady", "red": "Needs attention"}.get(rag, "—")


def _host_amb_html(t):
    """Host name + ambassador list with training indicators for the town card."""
    parts = []
    hi = t.get("host_info", {})
    if hi:
        host_name = hi.get("name", "")
        if host_name:
            caretaker_tag = (
                f' <span style="background:{ORANGE}33;color:{ORANGE};padding:1px 5px;'
                f'border-radius:3px;font-size:9px;font-weight:700;">caretaker</span>'
                if hi.get("caretaker") else ""
            )
            parts.append(
                f'<span style="color:{GREY};">Host:</span> '
                f'<strong style="color:{DARK};">{host_name}</strong>{caretaker_tag}'
            )

    ambs = t.get("town_ambs", [])
    if ambs:
        amb_tags = []
        for a in ambs:
            tick = "✓" if a["trained"] else "·"
            col  = LIME if a["trained"] else GREY
            amb_tags.append(
                f'<span style="color:{col};font-weight:{"700" if a["trained"] else "400"};">'
                f'{tick} {a["name"]}</span>'
            )
        parts.append(
            f'<span style="color:{GREY};">Ambassadors:</span> ' + ", ".join(amb_tags)
        )

    if not parts:
        return ""
    inner = " &nbsp;|&nbsp; ".join(parts)
    return (
        f'<div style="margin-top:8px;padding:6px 8px;background:{LGREY};border-radius:5px;'
        f'font-size:10px;line-height:1.6;">{inner}</div>'
    )


def _pct(val):
    if val is None or (hasattr(val, '__class__') and val.__class__.__name__ == 'float' and pd.isna(val)):
        return "—"
    try:
        return f"{round(float(val) * 100)}%"
    except Exception:
        return "—"


def _n(val):
    if val is None:
        return "—"
    try:
        v = float(val)
        if pd.isna(v):
            return "—"
        return str(int(v))
    except Exception:
        return "—"


def snap_card(colour, value, label):
    return f"""
    <div style="flex:1;min-width:110px;background:#fff;border-top:4px solid {colour};
         border-radius:8px;padding:14px 12px;text-align:center;
         box-shadow:0 1px 4px rgba(0,0,0,0.07);">
      <div style="font-size:28px;font-weight:800;color:{colour};">{value}</div>
      <div style="font-size:11px;color:{GREY};margin-top:4px;font-weight:600;
           text-transform:uppercase;letter-spacing:0.5px;">{label}</div>
    </div>"""


def town_card(t):
    rag_col   = _rag_colour(t["rag"])
    rag_lbl   = _rag_label(t["rag"])
    trend_bar = ""
    if t["trend_vals"]:
        mx = max(t["trend_vals"]) or 1
        bars = ""
        for v in t["trend_vals"][-6:]:
            h = max(4, int((v / mx) * 40))
            bars += f'<div style="width:8px;height:{h}px;background:{t["colour"]};border-radius:2px 2px 0 0;"></div>'
        trend_bar = f'<div style="display:flex;align-items:flex-end;gap:3px;margin-top:4px;">{bars}</div>'

    # 3-2-1 row
    q_html = ""
    if t["q_interaction"] is not None:
        q_html = f"""
        <div style="margin-top:10px;padding:8px 10px;background:{LGREY};border-radius:6px;
             font-size:11px;display:flex;gap:12px;flex-wrap:wrap;">
          <span><strong style="color:{TEAL};">{_pct(t['q_interaction'])}</strong> interaction</span>
          <span><strong style="color:{ORANGE};">{_n(t['q_met3'])}</strong> met 3 new</span>
          <span><strong style="color:{PINK};">{_n(t['q_121s'])}</strong> 1-2-1s</span>
          <span><strong style="color:{LIME};">{_n(t['q_brought'])}</strong> brought someone</span>
        </div>"""

    # Sponsor flags
    sp_html = ""
    for s in t["sponsor_flags"]:
        flag_col = PINK if "risk" in s["flag"].lower() else (LIME if "excellent" in s["flag"].lower() else ORANGE)
        sp_html += f'<span style="background:{flag_col}22;color:{flag_col};padding:2px 6px;border-radius:4px;font-size:10px;font-weight:700;margin-right:4px;">{s["company"]} {s["pct"]}%</span>'

    # Star rating badge
    star_str  = t.get("ee_star", "—")
    star_desc = t.get("ee_star_desc", "")
    star_col  = LIME if "\u2605" in star_str or "\u2728" in star_str else (ORANGE if t.get("is_new") else GREY)
    star_badge = (
        f'<span style="font-size:13px;font-weight:800;color:{star_col};" title="{star_desc}">'
        f'{star_str}</span>'
        f'<span style="font-size:10px;color:{GREY};margin-left:5px;">{star_desc}</span>'
    )

    # EE criteria dots (4 criteria: paying guests, trained ambassadors, Buzz Plus, sponsors)
    ee_dots = f'<div style="margin-bottom:4px;">{star_badge}</div>'
    for met, lbl in [
        (t["ee_paying"],      f"Avg \u2265{EE_PAYING_TARGET} guests"),
        (t["ee_ambassadors"], f"{EE_AMBASSADOR_MIN} trained ambs"),
        (t["ee_combined"],    f"Plus+Sponsors \u2265{EE_COMBINED_MIN} combined"),
    ]:
        col = LIME if met else PINK
        ee_dots += f'<span style="color:{col};font-size:10px;margin-right:8px;">{"●" if met else "○"} {lbl}</span>'

    latest_label = ""
    if t["latest_month"] and t["latest_month"] != "—":
        try:
            latest_label = datetime.strptime(t["latest_month"] + "-01", "%Y-%m-%d").strftime("%b %Y")
        except Exception:
            latest_label = t["latest_month"]

    return f"""
    <div style="background:#fff;border-radius:10px;padding:16px 18px;
         box-shadow:0 2px 6px rgba(0,0,0,0.07);border-left:5px solid {t['colour']};">

      <div style="display:flex;justify-content:space-between;align-items:flex-start;margin-bottom:10px;">
        <div>
          <div style="font-size:15px;font-weight:800;color:{DARK};">{t['label']}</div>
          <div style="font-size:11px;color:{GREY};margin-top:2px;">{t['num_events']} events (12m) · {t['total_12m']} unique visitors</div>
        </div>
        <span style="background:{rag_col};color:#fff;padding:3px 10px;border-radius:20px;
              font-size:10px;font-weight:700;white-space:nowrap;">{rag_lbl}</span>
      </div>

      <div style="display:flex;gap:16px;align-items:flex-end;">
        <div>
          <div style="font-size:26px;font-weight:800;color:{t['colour']};">{t['latest_count']}</div>
          <div style="font-size:10px;color:{GREY};">Last event ({latest_label})</div>
          <div style="font-size:11px;margin-top:2px;">
            <span style="color:{t['arrow_col']};font-weight:700;">{t['arrow']}</span>
            <span style="color:{GREY};"> avg {t['avg']}</span>
          </div>
        </div>
        <div style="flex:1;">{trend_bar}</div>
      </div>

      <div style="display:grid;grid-template-columns:1fr 1fr;gap:6px;margin-top:10px;font-size:11px;">
        <div style="padding:6px 8px;background:{LGREY};border-radius:5px;">
          <span style="color:{GREY};">Sponsors</span><br>
          <strong style="color:{''+LIME+'' if t['sponsors']>=SPONSOR_TOTAL_SLOTS else ''+ORANGE+''};">{t['sponsors']}/{SPONSOR_TOTAL_SLOTS}</strong>
          <span style="color:{GREY};"> slots filled · EE needs {EE_COMBINED_MIN} combined (Plus+Sponsors)</span>
        </div>
        <div style="padding:6px 8px;background:{LGREY};border-radius:5px;">
          <span style="color:{GREY};">Buzz Plus</span><br>
          <strong style="color:{DARK};">{t['active_plus']} member{'s' if t['active_plus']!=1 else ''}</strong>
          <span style="color:{GREY};"> · {t['strong_plus']} strong prospects</span>
        </div>
      </div>

      {f'<div style="margin-top:8px;font-size:10px;">{sp_html}</div>' if sp_html else ''}

      {q_html}

      <div style="margin-top:10px;font-size:10px;color:{GREY};">{ee_dots}</div>

      {_host_amb_html(t)}

    </div>"""


def build_chart_js(towns, trend_df):
    """Multi-line Chart.js for all 5 towns."""
    if trend_df.empty:
        return ""

    # Build a unified set of months
    all_months = sorted(trend_df["event_month"].unique())
    labels = []
    for m in all_months:
        try:
            labels.append(datetime.strptime(m + "-01", "%Y-%m-%d").strftime("%b %Y"))
        except Exception:
            labels.append(m)

    datasets = []
    for t in towns:
        town_df = trend_df[trend_df["town_code"] == t["code"]].set_index("event_month")
        data_points = []
        for m in all_months:
            if m in town_df.index:
                data_points.append(int(town_df.loc[m, "unique_attendees"]))
            else:
                data_points.append("null")

        datasets.append({
            "label": t["label"],
            "data": data_points,
            "borderColor": t["colour"],
            "backgroundColor": t["colour"] + "22",
            "borderWidth": 2,
            "pointRadius": 3,
            "tension": 0.3,
            "spanGaps": True,
        })

    labels_js   = json.dumps(labels)
    datasets_js = json.dumps(datasets)

    return f"""
    <canvas id="trendChart" style="width:100%;max-height:260px;"></canvas>
    <script src="https://cdnjs.cloudflare.com/ajax/libs/Chart.js/4.4.0/chart.umd.min.js"></script>
    <script>
    new Chart(document.getElementById('trendChart'), {{
      type: 'line',
      data: {{ labels: {labels_js}, datasets: {datasets_js} }},
      options: {{
        responsive: true,
        plugins: {{
          legend: {{ position: 'bottom', labels: {{ font: {{ size: 11 }}, boxWidth: 12 }} }},
          tooltip: {{ mode: 'index', intersect: false }}
        }},
        scales: {{
          y: {{ beginAtZero: true, grid: {{ color: '#f0f0f0' }},
                ticks: {{ font: {{ size: 11 }} }} }},
          x: {{ grid: {{ display: false }},
                ticks: {{ font: {{ size: 11 }}, maxRotation: 45 }} }}
        }}
      }}
    }});
    </script>"""


def build_actions(towns):
    actions = []

    # Declining attendance — skip new events (not enough history) and flag caretaker towns differently
    for t in towns:
        if t.get("is_new"):
            continue  # New events won't have meaningful decline data yet
        vals = t["trend_vals"]
        if len(vals) >= 3 and vals[-1] < vals[-3] - 5:
            is_ct = t.get("host_info", {}).get("caretaker", False)
            if is_ct:
                actions.append((PINK, t["label"],
                    f"Attendance dropped from {vals[-3]} to {vals[-1]} over 3 events. "
                    f"Review attendee re-engagement plan and new face conversion — check lapsed list before next event."))
            else:
                host_name = t.get("host_info", {}).get("name", "the host")
                actions.append((PINK, t["label"],
                    f"Attendance dropped from {vals[-3]} to {vals[-1]} over 3 events. "
                    f"Connect with {host_name} — review lapsed re-engagement and new face conversion strategy."))

    # New events — positive note, not a flag
    new_towns = [t for t in towns if t.get("is_new")]
    if new_towns:
        names = ", ".join(t["label"] for t in new_towns)
        actions.append((LIME, f"New events — {names}",
            f"Still in early build phase. Focus on consistency, warm introductions, and first sponsor conversations. "
            f"Event Excellence scoring will reflect their share of the Nov–Oct conference period."))

    # Buzz Plus — show all 5 towns so Emma has the full picture
    plus_parts = []
    for t in towns:
        if t.get("is_new") and t["strong_plus"] == 0:
            plus_parts.append(f"{t['short']}: building (new event)")
        elif t["active_plus"] > 0:
            plus_parts.append(f"{t['short']}: {t['active_plus']} active member{'s' if t['active_plus']!=1 else ''}, {t['strong_plus']} strong prospects")
        else:
            plus_parts.append(f"{t['short']}: {t['strong_plus']} strong, {t['poss_plus']} possible, 0 active")
    actions.append((TEAL, "Buzz Plus position — all events",
        " · ".join(plus_parts) + ". "
        "Brief hosts on strong prospects before each event — warm introduction beats a cold ask."))

    # Sponsor position — show all 5 towns with their slot count
    sp_parts = []
    for t in towns:
        new_flag = " (new)" if t.get("is_new") else ""
        sp_parts.append(f"{t['short']}: {t['sponsors']}/{SPONSOR_TOTAL_SLOTS}{new_flag}")
    actions.append((ORANGE, "Sponsor position — all events",
        " · ".join(sp_parts) + f". "
        f"Event Excellence needs Buzz Plus + sponsors combined ≥{EE_COMBINED_MIN}. "
        f"Use existing regional sponsors as social proof when approaching new ones."))

    # At-risk sponsors
    for t in towns:
        for s in t["sponsor_flags"]:
            if "risk" in s["flag"].lower():
                actions.append((PINK, f"Sponsor at risk — {t['label']}",
                    f"{s['company']} attendance at {s['pct']}%. Review relationship and attendance expectation before renewal conversation."))

    # Steady towns — only add if we have space
    green = [t for t in towns if t["rag"] == "green"]
    if green and len(actions) < 4:
        names = ", ".join(t["label"] for t in green)
        actions.append((LIME, "Strong towns — keep the momentum",
            f"{names} are performing well. Focus energy on Event Excellence progression (sponsor and Buzz Plus) "
            f"rather than attendance rescue."))

    return actions[:8]  # Cap at 8 actions


def _regional_sponsors_html(rsp):
    """Regional sponsor + charity partner panel."""
    if rsp is None:
        rsp = {}
    sp_count = rsp.get("sp_count", 0)
    sp_max   = rsp.get("sp_max", 4)
    sp_names = rsp.get("sp_names", [])
    ch_name  = rsp.get("ch_name", "")

    sp_slots_html = ""
    for i in range(sp_max):
        if i < len(sp_names):
            sp_slots_html += (
                f'<span style="display:inline-block;padding:4px 10px;margin:3px;'
                f'background:{TEAL}22;color:{TEAL};border:1px solid {TEAL}44;'
                f'border-radius:5px;font-size:11px;font-weight:600;">{sp_names[i]}</span>'
            )
        else:
            sp_slots_html += (
                f'<span style="display:inline-block;padding:4px 10px;margin:3px;'
                f'background:#f9f9f9;color:{GREY};border:1px dashed #ddd;'
                f'border-radius:5px;font-size:11px;">Slot {i+1} — available</span>'
            )

    ch_html = (
        f'<span style="display:inline-block;padding:4px 10px;background:{LIME}22;'
        f'color:{LIME};border:1px solid {LIME}44;border-radius:5px;font-size:11px;'
        f'font-weight:600;">{ch_name}</span>'
    ) if ch_name else (
        f'<span style="display:inline-block;padding:4px 10px;background:#f9f9f9;'
        f'color:{GREY};border:1px dashed #ddd;border-radius:5px;font-size:11px;">'
        f'Regional charity partner — not yet confirmed</span>'
    )

    return f"""
  <div class="section">
    <div class="section-title">Regional sponsors &amp; charity partner</div>
    <div style="display:flex;gap:20px;flex-wrap:wrap;">
      <div style="flex:2;min-width:240px;">
        <div style="font-size:11px;color:{GREY};font-weight:700;text-transform:uppercase;
             letter-spacing:0.5px;margin-bottom:8px;">Sponsors ({sp_count}/{sp_max} slots)</div>
        <div>{sp_slots_html}</div>
      </div>
      <div style="flex:1;min-width:180px;">
        <div style="font-size:11px;color:{GREY};font-weight:700;text-transform:uppercase;
             letter-spacing:0.5px;margin-bottom:8px;">Charity partner</div>
        <div>{ch_html}</div>
      </div>
    </div>
  </div>"""


def render_html(towns, cross_town, totals, trend_df, generated, regional_sp=None):
    chart_html    = build_chart_js(towns, trend_df)
    actions       = build_actions(towns)
    town_cards_h  = "".join(town_card(t) for t in towns)

    # Snap cards
    snap_row = "".join([
        snap_card(TEAL,   totals["total_events_12m"],  "Events (12m)"),
        snap_card(ORANGE, totals["total_visits_12m"],  "Event visits (12m)"),   # per-town, not deduplicated
        snap_card(PINK,   totals["total_people"],       "People in region"),     # unduplicated headcount
        snap_card(LIME,   totals["multi_town"],         "Multi-town visitors"),
        snap_card(TEAL,   totals["active_members"],     f"Active members ({ACTIVE_MEMBER_MONTHS}+ visits)"),
    ])

    # Action items
    actions_html = ""
    for i, (col, title, body) in enumerate(actions, 1):
        # Stable key based on index + truncated title (survives regeneration as long as order holds)
        note_key = f"buzz_rl_action_{i}_{title[:24].replace(' ', '_').replace('/', '_')}"
        actions_html += f"""
        <div style="display:flex;gap:12px;align-items:flex-start;padding:12px 0;
             {'border-bottom:1px solid #f0f0f0;' if i < len(actions) else ''}">
          <div style="width:28px;height:28px;border-radius:50%;background:{col};color:#fff;
               display:flex;align-items:center;justify-content:center;font-weight:800;
               font-size:12px;flex-shrink:0;">{i}</div>
          <div style="flex:1;min-width:0;">
            <div style="font-weight:700;font-size:13px;color:{DARK};margin-bottom:3px;">{title}</div>
            <div style="font-size:12px;color:{GREY};line-height:1.5;">{body}</div>
            <textarea id="note_{i}" data-key="{note_key}"
              placeholder="Your notes — host chat outcome, next steps, anything relevant..."
              oninput="saveNote(this)"
              style="margin-top:8px;width:100%;min-height:38px;max-height:140px;
                     font-size:11px;font-family:inherit;color:{DARK};
                     background:#FAFBFC;border:1px solid #E5E7EB;border-radius:6px;
                     padding:6px 8px;resize:vertical;outline:none;line-height:1.5;
                     box-sizing:border-box;"></textarea>
          </div>
        </div>"""

    # Cross-town table
    def _role_badge(role):
        if not role:
            return ""
        label  = {"regional lead": "RL", "host": "Host", "ambassador": "Amb"}.get(role, role.title())
        colour = {"regional lead": TEAL, "host": ORANGE, "ambassador": LIME}.get(role, GREY)
        return (
            f'<span style="background:{colour}22;color:{colour};border:1px solid {colour}44;'
            f'padding:1px 5px;border-radius:3px;font-size:9px;font-weight:700;'
            f'margin-left:5px;vertical-align:middle;">{label}</span>'
        )

    ct_rows = ""
    for p in cross_town:
        town_badges = ""
        for tc in p["towns"].split(", "):
            code = next((c for c, l in TOWN_LABELS.items() if l == tc.strip()), None)
            col  = TOWN_COLOURS.get(code, GREY) if code else GREY
            short = TOWN_SHORT.get(code, tc[:2]) if code else tc[:2]
            town_badges += f'<span style="background:{col}22;color:{col};padding:1px 5px;border-radius:3px;font-size:10px;font-weight:700;margin-right:3px;">{short}</span>'
        name_cell = p['name'] + _role_badge(p.get('role', ''))
        ct_rows += f"""
        <tr style="border-bottom:1px solid #f0f0f0;">
          <td style="padding:8px 6px;font-weight:600;font-size:12px;">{name_cell}</td>
          <td style="padding:8px 6px;font-size:11px;color:{GREY};">{p['company']}</td>
          <td style="padding:8px 6px;text-align:center;font-weight:700;color:{DARK};">{p['visits']}</td>
          <td style="padding:8px 6px;text-align:center;">{p['towns_count']}</td>
          <td style="padding:8px 6px;">{town_badges}</td>
        </tr>"""

    # EE summary table
    def _dot(met):
        col = LIME if met else PINK
        sym = "●" if met else "○"
        return f'<span style="color:{col};font-size:14px;">{sym}</span>'

    ee_rows = ""
    for t in towns:
        avg_col = LIME if t["ee_paying"] else PINK
        # Ambassador training detail for EE table
        ambs = t.get("town_ambs", [])
        trained = t.get("trained_count", 0)
        amb_detail = f'{trained}/{len(ambs)} trained' if ambs else "—"
        star_str = t.get("ee_star", "—")
        star_col2 = LIME if "\u2605" in star_str or "\u2728" in star_str else (ORANGE if t.get("is_new") else GREY)
        ee_rows += f"""
        <tr style="border-bottom:1px solid #f0f0f0;">
          <td style="padding:8px 6px;font-weight:600;font-size:12px;">
            <span style="display:inline-block;width:10px;height:10px;border-radius:50%;
                 background:{t['colour']};margin-right:6px;"></span>{t['label']}</td>
          <td style="padding:8px 6px;text-align:center;font-size:12px;font-weight:700;
               color:{avg_col};">{t['avg']}</td>
          <td style="padding:8px 6px;text-align:center;">{_dot(t['ee_paying'])}</td>
          <td style="padding:8px 6px;text-align:center;font-size:11px;">{_dot(t['ee_ambassadors'])} <span style="color:{GREY};">{amb_detail}</span></td>
          <td style="padding:8px 6px;text-align:center;">{_dot(t['ee_combined'])} <span style="color:{GREY};font-size:11px;">{t['active_plus']}+{t['sponsors']}/{EE_COMBINED_MIN} combined</span></td>
          <td style="padding:8px 6px;text-align:center;font-size:14px;font-weight:800;color:{star_col2};">{star_str}</td>
        </tr>"""

    return f"""<!DOCTYPE html>
<html lang="en">
<head>
<meta charset="UTF-8">
<meta name="viewport" content="width=device-width, initial-scale=1.0">
<title>Business Buzz — Leicestershire &amp; Rutland Regional Dashboard</title>
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
        text-transform:uppercase; padding:6px 6px 10px; letter-spacing:0.5px; }}
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
       margin-bottom:20px;display:flex;justify-content:space-between;align-items:center;">
    <div>
      <div style="font-size:10px;font-weight:700;letter-spacing:2px;
           text-transform:uppercase;opacity:0.7;margin-bottom:4px;">Business Buzz</div>
      <div style="font-size:22px;font-weight:800;">Leicestershire &amp; Rutland</div>
      <div style="font-size:13px;opacity:0.8;margin-top:3px;">Regional Dashboard</div>
    </div>
    <div style="text-align:right;font-size:12px;opacity:0.8;">
      <div style="font-size:18px;font-weight:800;opacity:1;color:#fff;">5</div>
      <div>Active towns</div>
      <div style="margin-top:8px;">Generated {generated}</div>
    </div>
  </div>

  <!-- REGIONAL SNAPSHOT -->
  <div class="section">
    <div class="section-title">Regional snapshot</div>
    <div style="display:flex;gap:10px;flex-wrap:wrap;">{snap_row}</div>
  </div>

  <!-- ATTENDANCE TREND -->
  <div class="section">
    <div class="section-title">Attendance trend — all towns</div>
    {chart_html}
  </div>

  <!-- TOWN HEALTH CARDS -->
  <div class="section">
    <div class="section-title">Town health</div>
    <div style="display:grid;grid-template-columns:repeat(auto-fit,minmax(300px,1fr));gap:14px;">
      {town_cards_h}
    </div>
  </div>

  <!-- ACTION ITEMS -->
  <div class="section">
    <div class="section-title">This month — actions &amp; watch points</div>
    {actions_html}
  </div>

  <!-- EVENT EXCELLENCE TRACKER -->
  <div class="section">
    <div class="section-title">Event Excellence position — awards conference period (Nov–Oct)</div>
    <div style="font-size:11px;color:{GREY};margin-bottom:12px;">
      HQ criteria (4-5 star): ≥{EE_PAYING_TARGET} avg paying guests · {EE_AMBASSADOR_MIN}+ trained ambassadors ·
      Buzz Plus + sponsors combined ≥{EE_COMBINED_MIN} (any mix) · special prize: {EE_SPONSOR_FULL} sponsors AND {EE_PLUS_FULL} Buzz Plus.
      ✓ = ambassador training complete.
    </div>
    <table>
      <tr>
        <th>Town</th>
        <th style="text-align:center;">Avg guests</th>
        <th style="text-align:center;">≥{EE_PAYING_TARGET} guests</th>
        <th style="text-align:center;">{EE_AMBASSADOR_MIN} trained ambs</th>
        <th style="text-align:center;">Buzz Plus</th>
        <th style="text-align:center;">Sponsors</th>
        <th style="text-align:center;">Rating</th>
      </tr>
      {ee_rows}
    </table>
    <div style="margin-top:12px;font-size:11px;color:{GREY};">
      ● = criterion met &nbsp;|&nbsp; ○ = not yet met &nbsp;|&nbsp;
      Attendance figure used as proxy for paying guests until app export is available.
    </div>
  </div>

  <!-- REGIONAL SPONSORS & CHARITY -->
  {_regional_sponsors_html(regional_sp)}

  <!-- CROSS-TOWN CHAMPIONS -->
  <div class="section">
    <div class="section-title">Cross-town champions — most engaged across the region</div>
    <div style="font-size:11px;color:{GREY};margin-bottom:12px;">
      People attending 2+ towns — potential ambassadors, advocates, and Buzz Plus prospects.
      Blank company names mean the member hasn't completed their profile in the event app.
    </div>
    <table>
      <tr>
        <th>Name</th><th>Company</th>
        <th style="text-align:center;">Region visits</th>
        <th style="text-align:center;">Towns</th>
        <th>Where</th>
      </tr>
      {ct_rows}
    </table>
  </div>

  <!-- FOOTER -->
  <div style="text-align:center;font-size:11px;color:{GREY};padding:16px 0;">
    Business Buzz · Leicestershire &amp; Rutland · Regional Lead Dashboard ·
    Data from Buzz_Region_Curated · Generated {generated}
  </div>

</div>

<script>
// Save & restore notes in the Actions & Watchpoints section
function saveNote(el) {{
  try {{
    localStorage.setItem(el.dataset.key, el.value);
    el.style.borderColor = '#8CC63F';
    setTimeout(function() {{ el.style.borderColor = '#E5E7EB'; }}, 700);
  }} catch(e) {{}}
}}
document.addEventListener('DOMContentLoaded', function() {{
  document.querySelectorAll('textarea[data-key]').forEach(function(el) {{
    try {{
      var saved = localStorage.getItem(el.dataset.key);
      if (saved) {{
        el.value = saved;
        // Auto-expand to fit content
        el.style.height = 'auto';
        el.style.height = Math.min(el.scrollHeight, 140) + 'px';
      }}
    }} catch(e) {{}}
  }});
}});
</script>
</body>
</html>"""


# ══════════════════════════════════════════════════════════════════════════════
# MAIN
# ══════════════════════════════════════════════════════════════════════════════

def main():
    generated = f"{datetime.now().day} {datetime.now().strftime('%B %Y')}"
    print(f"[Regional Dashboard] Loading data from {CURATED}")

    overview    = load_town_overview()
    trend       = load_attendance_trend()
    sponsor_cap = load_sponsor_capacity()
    sponsor_att = load_sponsor_attendance()
    master      = load_master_people()
    data_321    = load_321_town_summary()
    plus_counts, plus_members, plus_strong, plus_possible = load_active_plus_by_town()
    ambassadors = load_ambassadors()
    hosts       = load_hosts()
    regional_sp = load_regional_sponsors()

    towns      = build_town_data(
        overview, trend, sponsor_cap, sponsor_att, data_321,
        plus_counts=plus_counts, plus_members=plus_members,
        plus_strong=plus_strong, plus_possible=plus_possible,
        ambassadors=ambassadors, hosts=hosts,
    )
    cross_town = build_cross_town(master)
    totals     = build_region_totals(overview, master)

    html = render_html(towns, cross_town, totals, trend, generated, regional_sp=regional_sp)

    OUT_HTML.parent.mkdir(parents=True, exist_ok=True)
    OUT_HTML.write_text(html, encoding="utf-8")
    print(f"[Regional Dashboard] Written to {OUT_HTML}")


if __name__ == "__main__":
    main()
