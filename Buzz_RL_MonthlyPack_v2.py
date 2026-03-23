"""
Buzz_RL_MonthlyPack_v2.py
=========================
Generates a branded HTML monthly pack for the Regional Lead.
Parallel to Buzz_Town_HostPack_v2.py — same pipeline, same trigger.

Output (in Buzz_Region_Curated/):
    RL_Monthly_Pack.html

Usage:
    python Buzz_RL_MonthlyPack_v2.py

Sections:
    1. 5-town health dashboard (attendance, sponsor, Plus, award)
    2. This month's actions — what actually needs doing
    3. Sponsor pipeline — slots, lockouts, renewals, at-risk
    4. Cross-town visitors — most engaged people across the region
    5. Host recruitment candidates — ready to step up
    6. Solo event readiness — which towns are ready
"""

from datetime import date
from pathlib import Path
from typing import Optional
import pandas as pd

# ------------------------------------------------------------------
# PATHS
# ------------------------------------------------------------------

def _detect_region_root(start: Path) -> Path:
    p = start
    for _ in range(6):
        if (p / "Buzz_Region_Curated").exists():
            return p
        if p.parent == p:
            break
        p = p.parent
    return start

SCRIPT_DIR     = Path(__file__).resolve().parent
REGION_ROOT    = _detect_region_root(SCRIPT_DIR)
REGION_CURATED = REGION_ROOT / "Buzz_Region_Curated"

DASHBOARD_FILE  = REGION_CURATED / "region_dashboard.xlsx"
SPONSOR_FILE    = REGION_CURATED / "sponsor_intelligence.xlsx"
BUZZPLUS_FILE   = REGION_CURATED / "buzzplus_intelligence.xlsx"
EXCELLENCE_FILE = REGION_CURATED / "region_event_excellence.xlsx"
MASTER_PEOPLE   = REGION_CURATED / "region_master_people.csv"

TODAY       = date.today()
MONTH_LABEL = TODAY.strftime("%B %Y")
GENERATED   = TODAY.strftime("%-d %b %Y")

TEAL   = "#00A19A"
ORANGE = "#F39200"
PINK   = "#D60B52"
LIME   = "#B6BD00"

TOWNS_ORDER = ["MarketHarborough","Leicester","Lutterworth","Hinckley","Loughborough"]
TOWN_LABELS = {
    "MarketHarborough": "Market Harborough",
    "Leicester":        "Leicester",
    "Lutterworth":      "Lutterworth",
    "Hinckley":         "Hinckley",
    "Loughborough":     "Loughborough",
}

# ------------------------------------------------------------------
# DATA HELPERS
# ------------------------------------------------------------------

def _safe_excel(path: Path, sheet: str) -> pd.DataFrame:
    if not path.exists():
        return pd.DataFrame()
    try:
        return pd.read_excel(path, sheet_name=sheet)
    except Exception:
        return pd.DataFrame()

def _nc(df: pd.DataFrame) -> pd.DataFrame:
    df = df.copy()
    df.columns = [str(c).strip().lower() for c in df.columns]
    return df

def _tc(s) -> str:
    s = str(s or "").strip()
    if not s or s.lower() == "nan":
        return ""
    if s == s.upper() or s == s.lower():
        return s.title()
    return s

def _initials(name: str) -> str:
    parts = str(name).strip().split()
    return "".join(p[0].upper() for p in parts[:2] if p)

# ------------------------------------------------------------------
# DATA LOADERS
# ------------------------------------------------------------------

def load_town_overview() -> pd.DataFrame:
    return _nc(_safe_excel(DASHBOARD_FILE, "Town_Overview"))

def load_trend() -> pd.DataFrame:
    return _nc(_safe_excel(DASHBOARD_FILE, "Attendance_Trend_12m"))

def load_sponsor_capacity() -> pd.DataFrame:
    return _nc(_safe_excel(SPONSOR_FILE, "Sponsor_Capacity"))

def load_sponsor_attendance() -> pd.DataFrame:
    return _nc(_safe_excel(SPONSOR_FILE, "Sponsor_Attendance"))

def load_excellence() -> pd.DataFrame:
    df = _nc(_safe_excel(EXCELLENCE_FILE, "Sheet1"))
    if df.empty:
        return df
    # Latest buzz year per town
    return df.sort_values("buzz_year_label").groupby("town_code").last().reset_index()

def load_master_people() -> pd.DataFrame:
    if not MASTER_PEOPLE.exists():
        return pd.DataFrame()
    try:
        return pd.read_csv(MASTER_PEOPLE)
    except Exception:
        return pd.DataFrame()

def load_buzzplus_prospects() -> pd.DataFrame:
    return _nc(_safe_excel(BUZZPLUS_FILE, "Plus_Prospects"))

# ------------------------------------------------------------------
# SECTION BUILDERS
# ------------------------------------------------------------------

def build_town_health(overview: pd.DataFrame, excellence: pd.DataFrame,
                      sponsor_cap: pd.DataFrame, trend: pd.DataFrame) -> list:
    """One dict per town with all health indicators."""
    towns = []
    for code in TOWNS_ORDER:
        label = TOWN_LABELS.get(code, code)
        ov  = overview[overview["town_code"] == code]
        ex  = excellence[excellence["town_code"] == code] if not excellence.empty else pd.DataFrame()
        sp  = sponsor_cap[sponsor_cap["town_code"] == code] if not sponsor_cap.empty else pd.DataFrame()
        tr  = trend[trend["town_code"] == code].sort_values("event_month") if not trend.empty else pd.DataFrame()

        def _n(df, col, default=0):
            if df.empty or col not in df.columns: return default
            v = pd.to_numeric(df.iloc[0][col], errors="coerce")
            return int(v) if pd.notna(v) else default

        def _f(df, col, default=0.0):
            if df.empty or col not in df.columns: return default
            v = pd.to_numeric(df.iloc[0][col], errors="coerce")
            return round(float(v), 1) if pd.notna(v) else default

        def _s(df, col, default=""):
            if df.empty or col not in df.columns: return default
            return str(df.iloc[0][col] or default).strip()

        avg   = _f(ov, "avg_unique_per_event_12m")
        events = _n(ov, "num_events_12m")
        strong = _n(ov, "buzzplus_strong_prospects")
        poss   = _n(ov, "buzzplus_possible_prospects")
        plus   = _n(ov, "active_plus_members")
        slots  = _n(ov, "capacity_left")
        lockouts = _s(ov, "industry_lockouts")
        att_focus = _s(ov, "attendance_focus")

        # Award stars from excellence
        stars   = _n(ex, "award_stars")
        rating  = _s(ex, "award_rating")
        host    = _s(ex, "host_name")
        ambassadors = _n(ex, "active_ambassadors")
        buzz_year   = _s(ex, "buzz_year_label")

        # Trend: last 3 months
        trend_vals = []
        if not tr.empty and "unique_attendees" in tr.columns:
            trend_vals = [int(v) for v in tr["unique_attendees"].tail(3).tolist()]

        # Traffic light
        if att_focus == "Strong" and slots < 4:
            rag = "green"
        elif att_focus == "Low":
            rag = "red"
        else:
            rag = "amber"

        # Solo readiness score (0-3)
        solo = (
            int(avg >= 25) +
            int(strong + poss >= 10) +
            int(events >= 9)
        )

        towns.append(dict(
            code=code, label=label, avg=avg, events=events,
            strong=strong, poss=poss, plus=plus, slots=slots,
            lockouts=lockouts, att_focus=att_focus,
            stars=stars, rating=rating, host=host,
            ambassadors=ambassadors, buzz_year=buzz_year,
            trend=trend_vals, rag=rag, solo=solo,
        ))
    return towns


def build_actions(towns: list) -> list:
    """Derive concrete monthly actions from town health data."""
    actions = []
    # Sponsor gaps
    no_sponsor = [t for t in towns if t["slots"] >= 4]
    some_gap   = [t for t in towns if 0 < t["slots"] < 4]
    if no_sponsor:
        names = ", ".join(t["label"] for t in no_sponsor)
        actions.append(dict(
            colour=ORANGE, num="1",
            title="Sponsor outreach — priority towns",
            body=f"{names} {'has' if len(no_sponsor)==1 else 'have'} no active sponsors. "
                 f"These are your highest-priority outreach targets this month. "
                 f"Check industry lockouts before approaching anyone.",
            tag="Sponsors"
        ))
    if some_gap:
        names = ", ".join(f"{t['label']} ({t['slots']} slot{'s' if t['slots']!=1 else ''})" for t in some_gap)
        actions.append(dict(
            colour=ORANGE, num="2" if no_sponsor else "1",
            title="Sponsor slots to fill",
            body=f"{names}. These towns have existing sponsors — use them as social proof in conversations.",
            tag="Sponsors"
        ))

    # Buzz Plus — strong prospects with zero conversions
    plus_towns = [t for t in towns if t["strong"] >= 3 and t["plus"] == 0]
    if plus_towns:
        names = ", ".join(f"{t['label']} ({t['strong']} strong)" for t in plus_towns)
        n = str(len(actions) + 1)
        actions.append(dict(
            colour=PINK, num=n,
            title="Buzz Plus conversion follow-up",
            body=f"{names}. Strong prospects identified but zero active Plus members. "
                 f"Brief the host specifically on having the conversation — don't just leave it in the pack.",
            tag="Buzz Plus"
        ))

    # Attendance concerns
    low_towns = [t for t in towns if t["att_focus"] == "Low" or (t["trend"] and len(t["trend"]) >= 2 and t["trend"][-1] < t["trend"][-2] - 5)]
    if low_towns:
        names = ", ".join(t["label"] for t in low_towns)
        n = str(len(actions) + 1)
        actions.append(dict(
            colour=TEAL, num=n,
            title="Attendance support needed",
            body=f"{names}. Numbers are either low or trending down. Worth a direct conversation with the host — what's happening in the room, are they re-engaging lapsed visitors?",
            tag="Attendance"
        ))

    # Missing hosts
    no_host = [t for t in towns if not t["host"] or t["host"].lower() in ("nan","none","")]
    if no_host:
        names = ", ".join(t["label"] for t in no_host)
        n = str(len(actions) + 1)
        actions.append(dict(
            colour=PINK, num=n,
            title="Host not recorded",
            body=f"{names} {'has' if len(no_host)==1 else 'have'} no host recorded in the event excellence data. Confirm the current host and update the roles file.",
            tag="Hosts"
        ))

    if not actions:
        actions.append(dict(
            colour=LIME, num="1",
            title="All towns stable this month",
            body="No immediate actions triggered by current thresholds. Use this month to get ahead on sponsor pipeline and host relationship-building.",
            tag="Monitor"
        ))

    return actions


def build_sponsor_pipeline(sponsor_cap: pd.DataFrame,
                            sponsor_att: pd.DataFrame) -> list:
    rows = []
    for code in TOWNS_ORDER:
        label = TOWN_LABELS.get(code, code)
        cap = sponsor_cap[sponsor_cap["town_code"] == code] if not sponsor_cap.empty else pd.DataFrame()
        att = sponsor_att[sponsor_att["town_code"] == code] if not sponsor_att.empty else pd.DataFrame()

        def _n(df, col, default=0):
            if df.empty or col not in df.columns: return default
            v = pd.to_numeric(df.iloc[0][col], errors="coerce")
            return int(v) if pd.notna(v) else default

        slots      = _n(cap, "capacity_left")
        current    = _n(cap, "current_sponsors")
        lockouts   = str(cap.iloc[0].get("industry_lockouts","") if not cap.empty else "") or "None"

        sponsor_rows = []
        if not att.empty:
            for _, row in att.iterrows():
                if str(row.get("town_code","")) != code:
                    continue
                flag = str(row.get("compliance_flag",""))
                pct  = row.get("attendance_pct", 0) or 0
                sponsor_rows.append(dict(
                    company=str(row.get("sponsor_company","")).strip(),
                    flag=flag,
                    pct=round(float(pct)*100),
                ))

        rows.append(dict(
            code=code, label=label, slots=slots,
            current=current, lockouts=lockouts,
            sponsors=sponsor_rows,
        ))
    return rows


def build_cross_town(master: pd.DataFrame, top_n: int = 12) -> list:
    if master.empty:
        return []
    team_roles = {"host","ambassador","regional lead"}
    df = master.copy()
    df.columns = [str(c).strip().lower() for c in df.columns]

    # Filter: 2+ towns, no current role, not team
    is_team = df["role_region"].astype(str).str.strip().str.lower().isin(team_roles)
    df = df[~is_team & (df["towns_visited_count"] >= 2)].copy()
    df = df.sort_values(["towns_visited_count","visits_region"], ascending=False).head(top_n)

    result = []
    for _, row in df.iterrows():
        name    = _tc(row.get("name",""))
        company = _tc(row.get("company",""))
        visits  = int(pd.to_numeric(row.get("visits_region",0), errors="coerce") or 0)
        towns_n = int(pd.to_numeric(row.get("towns_visited_count",0), errors="coerce") or 0)
        towns_v = str(row.get("towns_visited","") or "").strip()
        result.append(dict(
            name=name or row.get("email","").split("@")[0].replace("."," ").title(),
            company="" if company.lower() in ("nan","") else company,
            email=str(row.get("email","")),
            visits=visits,
            towns_count=towns_n,
            towns=towns_v,
            initials=_initials(name),
        ))
    return result


def build_host_candidates(master: pd.DataFrame, top_n: int = 10) -> list:
    if master.empty:
        return []
    team_roles = {"host","ambassador","regional lead"}
    df = master.copy()
    df.columns = [str(c).strip().lower() for c in df.columns]

    is_team = df["role_region"].astype(str).str.strip().str.lower().isin(team_roles)
    df = df[~is_team].copy()
    df["visits_n"] = pd.to_numeric(df["visits_region"], errors="coerce").fillna(0)
    df = df[df["visits_n"] >= 6].sort_values("visits_n", ascending=False).head(top_n)

    result = []
    for _, row in df.iterrows():
        name    = _tc(row.get("name",""))
        company = _tc(row.get("company",""))
        visits  = int(row["visits_n"])
        towns   = str(row.get("towns_visited","") or "")
        since   = str(row.get("first_seen_buzzyear_label","") or "")
        result.append(dict(
            name=name or row.get("email","").split("@")[0].replace("."," ").title(),
            company="" if company.lower() in ("nan","") else company,
            email=str(row.get("email","")),
            visits=visits,
            towns=towns,
            since=since,
            initials=_initials(name),
        ))
    return result

# ------------------------------------------------------------------
# HTML RENDERER
# ------------------------------------------------------------------

def _rag_colour(rag: str) -> str:
    return {"green": LIME, "amber": ORANGE, "red": PINK}.get(rag, "#9CA3AF")

def _rag_bg(rag: str) -> str:
    return {"green": "#F4F6CC", "amber": "#FEF0DA", "red": "#FCEEF4"}.get(rag, "#F3F4F6")

def _rag_fg(rag: str) -> str:
    return {"green": "#606500", "amber": "#9A5800", "red": "#8C0636"}.get(rag, "#6B7280")

def _stars(n: int) -> str:
    return ("★" * n + "☆" * (4 - n)) if n else "—"

def _mini_bar(vals: list) -> str:
    if not vals:
        return '<span style="color:#9CA3AF;font-size:11px">no trend data</span>'
    mx = max(vals) if max(vals) > 0 else 1
    bars = ""
    for v in vals:
        h = max(4, round(v / mx * 28))
        bars += (f'<div style="width:10px;height:{h}px;background:{TEAL};'
                 f'border-radius:2px;align-self:flex-end;opacity:0.8"></div>')
    return f'<div style="display:flex;gap:3px;align-items:flex-end;height:30px">{bars}</div>'

def _av(initials: str, bg: str, fg: str) -> str:
    return (f'<div style="width:34px;height:34px;border-radius:50%;background:{bg};'
            f'color:{fg};display:flex;align-items:center;justify-content:center;'
            f'font-size:11px;font-weight:700;flex-shrink:0">{initials}</div>')

def render_html(towns: list, actions: list, sponsor_pipeline: list,
                cross_town: list, candidates: list) -> str:

    # ── TOWN HEALTH CARDS ────────────────────────────────────────
    town_cards = ""
    for t in towns:
        rc  = _rag_colour(t["rag"])
        rbg = _rag_bg(t["rag"])
        rfg = _rag_fg(t["rag"])
        bar = _mini_bar(t["trend"])
        plus_total = t["strong"] + t["poss"]
        solo_label = ["Not ready", "Building", "Nearly there", "Ready"][min(t["solo"], 3)]
        solo_col   = [PINK, ORANGE, ORANGE, LIME][min(t["solo"], 3)]
        solo_bg    = ["#FCEEF4","#FEF0DA","#FEF0DA","#F4F6CC"][min(t["solo"], 3)]
        solo_fg    = ["#8C0636","#9A5800","#9A5800","#606500"][min(t["solo"], 3)]
        host_str   = t["host"].replace(" (MHBuzz Host)","").replace(" (MH Buzz Host)","") if t["host"] else "Not recorded"
        host_col   = "#111827" if t["host"] else PINK

        town_cards += f"""
        <div style="background:var(--color-background-primary);border:0.5px solid var(--color-border-tertiary);
                    border-top:4px solid {rc};border-radius:12px;padding:16px">
          <div style="display:flex;justify-content:space-between;align-items:flex-start;margin-bottom:10px">
            <div>
              <div style="font-size:15px;font-weight:700;color:var(--color-text-primary)">{t['label']}</div>
              <div style="font-size:11px;color:var(--color-text-secondary);margin-top:2px">{t['buzz_year']}</div>
            </div>
            <div style="background:{rbg};color:{rfg};font-size:10px;font-weight:700;
                        padding:3px 8px;border-radius:999px;letter-spacing:0.06em">{t['att_focus'].upper()}</div>
          </div>
          <div style="display:grid;grid-template-columns:1fr 1fr;gap:8px;margin-bottom:10px">
            <div style="background:var(--color-background-secondary);border-radius:8px;padding:8px 10px">
              <div style="font-size:10px;color:var(--color-text-secondary);margin-bottom:3px">Avg visitors</div>
              <div style="font-size:22px;font-weight:700;color:{rc}">{t['avg']}</div>
            </div>
            <div style="background:var(--color-background-secondary);border-radius:8px;padding:8px 10px">
              <div style="font-size:10px;color:var(--color-text-secondary);margin-bottom:3px">Buzz Plus</div>
              <div style="font-size:22px;font-weight:700;color:{PINK}">{plus_total}</div>
            </div>
          </div>
          <div style="margin-bottom:10px">{bar}</div>
          <div style="font-size:11px;color:var(--color-text-secondary);margin-bottom:6px;display:flex;justify-content:space-between">
            <span>Host: <span style="color:{host_col};font-weight:600">{host_str}</span></span>
            <span>Amb: {t['ambassadors']} &nbsp; Award: {_stars(t['stars'])}</span>
          </div>
          <div style="display:flex;gap:6px;flex-wrap:wrap">
            <span style="font-size:10px;font-weight:700;padding:2px 8px;border-radius:999px;
                         background:#E6F7F6;color:#007A74">{t['slots']} sponsor slot{'s' if t['slots']!=1 else ''} open</span>
            <span style="font-size:10px;font-weight:700;padding:2px 8px;border-radius:999px;
                         background:{solo_bg};color:{solo_fg}">Solo: {solo_label}</span>
          </div>
        </div>"""

    # ── ACTION CARDS ─────────────────────────────────────────────
    action_html = ""
    for a in actions:
        action_html += f"""
        <div style="display:flex;gap:12px;padding:12px 14px;background:var(--color-background-secondary);
                    border-radius:10px;border-left:3px solid {a['colour']}">
          <div style="width:24px;height:24px;border-radius:50%;background:{a['colour']};color:#fff;
                      display:flex;align-items:center;justify-content:center;
                      font-size:12px;font-weight:700;flex-shrink:0">{a['num']}</div>
          <div style="flex:1;min-width:0">
            <div style="font-size:13px;font-weight:700;color:var(--color-text-primary);margin-bottom:3px">{a['title']}</div>
            <div style="font-size:12px;color:var(--color-text-secondary);line-height:1.55">{a['body']}</div>
          </div>
          <div style="font-size:10px;font-weight:700;padding:2px 8px;height:fit-content;border-radius:999px;
                      background:var(--color-background-primary);color:var(--color-text-secondary);white-space:nowrap">{a['tag']}</div>
        </div>"""

    # ── SPONSOR PIPELINE ─────────────────────────────────────────
    spon_html = ""
    for s in sponsor_pipeline:
        slot_pips = ""
        for i in range(4):
            filled = i < (4 - s["slots"])
            bg = TEAL if filled else "var(--color-border-tertiary)"
            slot_pips += f'<div style="width:16px;height:8px;border-radius:2px;background:{bg}"></div>'

        sponsor_detail = ""
        for sp in s["sponsors"]:
            fc = {"Excellent": LIME, "Good": TEAL, "Needs Attention": ORANGE, "At Risk": PINK}.get(sp["flag"], "#9CA3AF")
            sponsor_detail += (f'<div style="display:flex;justify-content:space-between;align-items:center;'
                               f'padding:5px 0;border-top:0.5px solid var(--color-border-tertiary);font-size:12px">'
                               f'<span style="color:var(--color-text-primary)">{sp["company"]}</span>'
                               f'<span style="color:{fc};font-weight:700">{sp["flag"]}</span></div>')
        if not sponsor_detail:
            sponsor_detail = '<div style="font-size:12px;color:var(--color-text-secondary);padding:5px 0;font-style:italic">No active sponsors</div>'

        lockout_str = s["lockouts"] if s["lockouts"] and s["lockouts"].lower() not in ("nan","none","") else "None"

        spon_html += f"""
        <div style="background:var(--color-background-primary);border:0.5px solid var(--color-border-tertiary);
                    border-radius:10px;padding:14px 16px">
          <div style="display:flex;justify-content:space-between;align-items:center;margin-bottom:8px">
            <div style="font-size:14px;font-weight:700;color:var(--color-text-primary)">{s['label']}</div>
            <div style="display:flex;gap:3px">{slot_pips}</div>
          </div>
          {sponsor_detail}
          <div style="font-size:11px;color:var(--color-text-secondary);margin-top:8px">
            <span style="font-weight:600">Locked out:</span> {lockout_str}
          </div>
        </div>"""

    # ── CROSS-TOWN VISITORS ──────────────────────────────────────
    ct_html = ""
    for p in cross_town:
        ct_html += f"""
        <div style="display:flex;align-items:center;gap:10px;padding:8px 0;
                    border-bottom:0.5px solid var(--color-border-tertiary)">
          {_av(p['initials'], '#E6F7F6', '#007A74')}
          <div style="flex:1;min-width:0">
            <div style="font-size:13px;font-weight:600;color:var(--color-text-primary)">{p['name']}</div>
            <div style="font-size:11px;color:var(--color-text-secondary)">{p['company'] or p['email']}</div>
            <div style="font-size:11px;color:var(--color-text-secondary);margin-top:2px">{p['towns']}</div>
          </div>
          <div style="text-align:right;flex-shrink:0">
            <div style="font-size:13px;font-weight:700;color:{TEAL}">{p['towns_count']} towns</div>
            <div style="font-size:11px;color:var(--color-text-secondary)">{p['visits']} visits</div>
          </div>
        </div>"""
    if not ct_html:
        ct_html = '<p style="font-size:13px;color:var(--color-text-secondary);font-style:italic">No cross-town visitors found.</p>'

    # ── HOST CANDIDATES ──────────────────────────────────────────
    cand_html = ""
    for p in candidates:
        cand_html += f"""
        <div style="display:flex;align-items:center;gap:10px;padding:8px 0;
                    border-bottom:0.5px solid var(--color-border-tertiary)">
          {_av(p['initials'], '#FEF0DA', '#9A5800')}
          <div style="flex:1;min-width:0">
            <div style="font-size:13px;font-weight:600;color:var(--color-text-primary)">{p['name']}</div>
            <div style="font-size:11px;color:var(--color-text-secondary)">{p['company'] or p['email']}</div>
            <div style="font-size:11px;color:var(--color-text-secondary)">Attending since {p['since']} &middot; {p['towns']}</div>
          </div>
          <div style="text-align:right;flex-shrink:0">
            <div style="font-size:13px;font-weight:700;color:{ORANGE}">{p['visits']} visits</div>
          </div>
        </div>"""
    if not cand_html:
        cand_html = '<p style="font-size:13px;color:var(--color-text-secondary);font-style:italic">No candidates found.</p>'

    return f"""<!DOCTYPE html>
<html lang="en">
<head>
<meta charset="UTF-8">
<meta name="viewport" content="width=device-width,initial-scale=1.0">
<title>Business Buzz RL Pack &ndash; {MONTH_LABEL}</title>
<style>
*,*::before,*::after{{box-sizing:border-box;margin:0;padding:0}}
body{{font-family:'Century Gothic','Gill Sans',Calibri,sans-serif;font-size:14px;
      color:#111827;background:#F0F0F0;padding:24px 16px 48px}}
.pack{{max-width:800px;margin:0 auto;box-shadow:0 4px 24px rgba(0,0,0,.10)}}
.header{{background:{TEAL};border-radius:12px 12px 0 0;padding:28px 32px 22px;color:#fff;position:relative;overflow:hidden}}
.header::before{{content:'';position:absolute;right:-40px;top:-40px;width:220px;height:220px;border-radius:50%;background:rgba(255,255,255,.07)}}
.colour-bar{{display:flex;height:5px}}
.cb1{{flex:1;background:{TEAL}}}.cb2{{flex:1;background:{ORANGE}}}.cb3{{flex:1;background:{PINK}}}.cb4{{flex:1;background:{LIME}}}
.body{{background:#fff;border:1px solid #E5E7EB;border-top:none;border-radius:0 0 12px 12px;padding:28px 32px 32px}}
.section{{margin-bottom:28px}}
.section-hdr{{display:flex;align-items:center;gap:10px;padding-bottom:9px;
              border-bottom:2px solid #F3F4F6;margin-bottom:14px}}
.sec-num{{width:27px;height:27px;border-radius:50%;display:flex;align-items:center;
          justify-content:center;font-size:12px;font-weight:700;flex-shrink:0;color:#fff}}
.sec-title{{font-size:16px;font-weight:700;color:#111827}}
.sec-desc{{font-size:12px;color:#9CA3AF;margin-left:auto}}
.grid-5{{display:grid;grid-template-columns:repeat(5,minmax(0,1fr));gap:10px}}
.grid-2{{display:grid;grid-template-columns:repeat(2,minmax(0,1fr));gap:10px}}
.actions{{display:flex;flex-direction:column;gap:10px}}
.footer{{display:flex;justify-content:space-between;font-size:11px;color:#9CA3AF;
         margin-top:24px;padding-top:14px;border-top:1px solid #E5E7EB}}
@media print{{
  body{{background:#fff;padding:0}}
  .pack{{max-width:100%;box-shadow:none}}
  .header{{border-radius:0}}
  .body{{border-radius:0;border:none}}
  .section{{break-inside:avoid}}
}}
</style>
</head>
<body>
<div class="pack">
  <div class="header">
    <div style="position:relative;display:flex;justify-content:space-between;align-items:flex-start;margin-bottom:6px">
      <div>
        <div style="font-size:11px;letter-spacing:.12em;text-transform:uppercase;color:rgba(255,255,255,.65);margin-bottom:8px">
          Business Buzz &nbsp;&middot;&nbsp; Leicestershire &amp; Rutland &nbsp;&middot;&nbsp; Regional Lead
        </div>
        <div style="font-size:24px;font-weight:700;color:#fff;line-height:1.2">Monthly overview</div>
        <div style="font-size:13px;color:rgba(255,255,255,.75);margin-top:4px">{MONTH_LABEL}</div>
      </div>
      <div style="font-size:12px;color:rgba(255,255,255,.55);text-align:right">Generated {GENERATED}</div>
    </div>
    <div style="margin-top:16px;font-size:14px;color:rgba(255,255,255,.88);line-height:1.65;
                border-top:1px solid rgba(255,255,255,.2);padding-top:14px">
      Hi Emma &mdash; here&rsquo;s your regional picture for {MONTH_LABEL}.
      Actions, sponsor gaps, cross-town visitors, and host recruitment candidates &mdash; everything in one place.
    </div>
  </div>
  <div class="colour-bar"><div class="cb1"></div><div class="cb2"></div><div class="cb3"></div><div class="cb4"></div></div>
  <div class="body">

    <!-- SECTION 1: TOWN HEALTH -->
    <div class="section">
      <div class="section-hdr">
        <div class="sec-num" style="background:{TEAL}">1</div>
        <div class="sec-title">5-town health</div>
        <div class="sec-desc">Attendance &middot; Plus pipeline &middot; Sponsors &middot; Award</div>
      </div>
      <div class="grid-5">{town_cards}</div>
    </div>

    <!-- SECTION 2: ACTIONS -->
    <div class="section">
      <div class="section-hdr">
        <div class="sec-num" style="background:{ORANGE}">2</div>
        <div class="sec-title">This month&rsquo;s actions</div>
        <div class="sec-desc">Triggered by your data</div>
      </div>
      <div class="actions">{action_html}</div>
    </div>

    <!-- SECTION 3: SPONSOR PIPELINE -->
    <div class="section">
      <div class="section-hdr">
        <div class="sec-num" style="background:{LIME}">3</div>
        <div class="sec-title">Sponsor pipeline</div>
        <div class="sec-desc">Slots &middot; Lockouts &middot; Compliance</div>
      </div>
      <div class="grid-5">{spon_html}</div>
    </div>

    <!-- SECTION 4: CROSS-TOWN -->
    <div class="section">
      <div class="section-hdr">
        <div class="sec-num" style="background:{TEAL}">4</div>
        <div class="sec-title">Cross-town visitors</div>
        <div class="sec-desc">Your most engaged people across the region &middot; prime Buzz Plus targets</div>
      </div>
      {ct_html}
    </div>

    <!-- SECTION 5: HOST CANDIDATES -->
    <div class="section">
      <div class="section-hdr">
        <div class="sec-num" style="background:{ORANGE}">5</div>
        <div class="sec-title">Host &amp; ambassador candidates</div>
        <div class="sec-desc">High visit count &middot; no current role &middot; ready to step up</div>
      </div>
      {cand_html}
    </div>

    <div class="footer">
      <span>Business Buzz &nbsp;&middot;&nbsp; Leicestershire &amp; Rutland</span>
      <span>Regional Lead use only &nbsp;&middot;&nbsp; not for distribution</span>
    </div>
  </div>
</div>
</body>
</html>"""

# ------------------------------------------------------------------
# MAIN
# ------------------------------------------------------------------

def main() -> None:
    print("[INFO] Loading data...")
    overview    = load_town_overview()
    trend       = load_trend()
    sponsor_cap = load_sponsor_capacity()
    sponsor_att = load_sponsor_attendance()
    excellence  = load_excellence()
    master      = load_master_people()

    print("[INFO] Building sections...")
    towns    = build_town_health(overview, excellence, sponsor_cap, trend)
    actions  = build_actions(towns)
    pipeline = build_sponsor_pipeline(sponsor_cap, sponsor_att)
    cross    = build_cross_town(master)
    cands    = build_host_candidates(master)

    print("[INFO] Rendering HTML...")
    html = render_html(towns, actions, pipeline, cross, cands)

    out = REGION_CURATED / "RL_Monthly_Pack.html"
    out.write_text(html, encoding="utf-8")
    print(f"[OK] Wrote {out}")


if __name__ == "__main__":
    main()
