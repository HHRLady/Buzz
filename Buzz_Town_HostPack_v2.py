"""
Buzz_Town_HostPack_v2.py
========================
Generates a branded HTML host pack (+ plain-text email body) for each
Business Buzz town.

Drop-in replacement for Buzz_Town_HostPack.py.
Slot into Buzz_Region_AllBuild.py by swapping the script name in REGION_SCRIPTS.

Output per town  (in  Buzz_Region_Curated/host_packs/):
    HostPack_<TownName>.html          open in Chrome -> Ctrl+P -> Save as PDF
    HostPack_<TownName>_email.txt     paste as email body, attach the PDF

Usage:
    python Buzz_Town_HostPack_v2.py --town Loughborough
    python Buzz_Town_HostPack_v2.py --town ALL
    python Buzz_Town_HostPack_v2.py --town ALL --event-date "Thursday 20 March 2026"

Design rules:
    - Host first name pulled from roles_<TownCode>.csv (role == Host, no end_date)
    - Regulars  : 2+ months attended, excludes team roles, capped at 10
    - Lapsed    : 2+ months ever, last seen > 3 months ago, capped at 8
    - Prospects : sourced from regional buzzplus_intelligence.xlsx, strong first, capped at 8
    - Sponsors  : from sponsors_<TownCode>.csv
    - Snapshot  : from region_dashboard.xlsx Town_Overview row
    - All logic is attendance-based only (no payment fields used)
"""

import argparse
import re
from datetime import date
from pathlib import Path
from textwrap import dedent
from typing import Optional

import pandas as pd


# ------------------------------------------------------------------
# REGION ROOT DETECTION
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

TOWNS = {
    "MarketHarborough": ("Buzz_Event_Dashboard_MarketHarborough", "Market Harborough"),
    "Leicester":        ("Buzz_Event_Dashboard_Leicester",        "Leicester"),
    "Lutterworth":      ("Buzz_Event_Dashboard_Lutterworth",      "Lutterworth"),
    "Hinckley":         ("Buzz_Event_Dashboard_Hinckley",         "Hinckley"),
    "Loughborough":     ("Buzz_Event_Dashboard_Loughborough",     "Loughborough"),
}

BUZZPLUS_FILE  = REGION_CURATED / "buzzplus_intelligence.xlsx"
SPONSOR_FILE   = REGION_CURATED / "sponsor_intelligence.xlsx"
DASHBOARD_FILE = REGION_CURATED / "region_dashboard.xlsx"

MAX_REGULARS  = 10
MAX_LAPSED    =  8
MAX_PROSPECTS =  8

TODAY       = date.today()
MONTH_LABEL = TODAY.strftime("%B %Y")

# ------------------------------------------------------------------
# OFFICIAL BUSINESS BUZZ BRAND COLOURS (from portal.mybuzz.uk)
# ------------------------------------------------------------------
# "Blue" (teal): #00A19A  |  Orange: #F39200
# Pink:          #D60B52  |  Green (lime): #B6BD00
TEAL   = "#00A19A"
ORANGE = "#F39200"
PINK   = "#D60B52"
LIME   = "#B6BD00"

TEAL_BG  = "#E6F7F6"   # light teal for avatars/badges
TEAL_FG  = "#007A74"
ORG_BG   = "#FEF0DA"
ORG_FG   = "#9A5800"
PINK_BG  = "#FCEEF4"
PINK_FG  = "#8C0636"
LIME_BG  = "#F4F6CC"
LIME_FG  = "#606500"


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


def _first_int(df: pd.DataFrame, col: str) -> int:
    if df.empty or col not in df.columns:
        return 0
    try:
        return int(pd.to_numeric(df[col], errors="coerce").fillna(0).iloc[0])
    except Exception:
        return 0


def _parse_period(s) -> Optional[pd.Period]:
    try:
        return pd.Period(str(s).strip(), freq="M")
    except Exception:
        return None


def _filter_town(df: pd.DataFrame, code: str, label: str) -> pd.DataFrame:
    if df.empty:
        return df
    if "town_code" in df.columns:
        return df[df["town_code"].astype(str).str.strip().str.lower() == code.lower()].copy()
    if "town_name" in df.columns:
        return df[df["town_name"].astype(str).str.strip().str.lower() == label.lower()].copy()
    return df


def _initials(name: str) -> str:
    parts = str(name).strip().split()
    return "".join(p[0].upper() for p in parts[:2] if p)


def _title_case(name: str) -> str:
    if not name:
        return ""
    # Leave mixed-case as-is; fix ALL CAPS or all lower
    stripped = str(name).strip()
    if stripped == stripped.upper() or stripped == stripped.lower():
        return stripped.title()
    return stripped


# ------------------------------------------------------------------
# DATA LOADERS
# ------------------------------------------------------------------

def load_host_name(town_base: Path, town_code: str) -> str:
    path = town_base / "data_ref" / f"roles_{town_code}.csv"
    if not path.exists():
        return "Host"
    try:
        df = pd.read_csv(path)
        df.columns = [c.strip().lower() for c in df.columns]
        hosts = df[df["role"].astype(str).str.strip().str.lower() == "host"].copy()
        if "end_date" in hosts.columns:
            active = hosts[hosts["end_date"].astype(str).str.strip().isin(["", "nan", "NaT"])]
            hosts = active if not active.empty else hosts
        if not hosts.empty and "name" in hosts.columns:
            full = str(hosts.iloc[0]["name"]).strip()
            return full.split()[0] if full else "Host"
    except Exception:
        pass
    return "Host"


def load_team_emails(town_base: Path, town_code: str) -> set:
    path = town_base / "data_ref" / f"roles_{town_code}.csv"
    if not path.exists():
        return set()
    try:
        df = pd.read_csv(path)
        df.columns = [c.strip().lower() for c in df.columns]
        if "email" in df.columns:
            return set(df["email"].astype(str).str.strip().str.lower())
    except Exception:
        pass
    return set()


def load_attendance(monthly_dir: Path) -> pd.DataFrame:
    if not monthly_dir.exists():
        return pd.DataFrame()
    frames = []
    for f in sorted(monthly_dir.glob("*_attendance.xlsx")):
        try:
            frames.append(pd.read_excel(f))
        except Exception:
            continue
    if not frames:
        return pd.DataFrame()
    df = pd.concat(frames, ignore_index=True)
    df = _nc(df)
    if "email" in df.columns:
        df["email"] = df["email"].astype(str).str.strip().str.lower()
    if "event_month" in df.columns:
        df["event_month"] = df["event_month"].astype(str).str.strip()
    return df


def load_sponsors_csv(town_base: Path, town_code: str) -> pd.DataFrame:
    path = town_base / "data_ref" / f"sponsors_{town_code}.csv"
    if not path.exists():
        return pd.DataFrame()
    try:
        return _nc(pd.read_csv(path))
    except Exception:
        return pd.DataFrame()


# ------------------------------------------------------------------
# SECTION BUILDERS
# ------------------------------------------------------------------

def build_snapshot(town_code: str, town_label: str) -> dict:
    df  = _nc(_safe_excel(DASHBOARD_FILE, "Town_Overview"))
    row = _filter_town(df, town_code, town_label)
    if row.empty:
        return dict(avg_attendance=0, num_events=0, strong=0,
                    possible=0, total_prospects=0, active_plus=0, slots_left=0)
    r = row.iloc[0]
    def _n(col): return int(pd.to_numeric(r.get(col, 0), errors="coerce") or 0)
    strong   = _n("buzzplus_strong_prospects")
    possible = _n("buzzplus_possible_prospects")
    return dict(
        avg_attendance  = round(float(pd.to_numeric(r.get("avg_unique_per_event_12m", 0), errors="coerce") or 0)),
        num_events      = _n("num_events_12m"),
        strong          = strong,
        possible        = possible,
        total_prospects = strong + possible,
        active_plus     = _n("active_plus_members"),
        slots_left      = _n("capacity_left"),
    )


def build_regulars(attendance: pd.DataFrame, team_emails: set) -> list:
    if attendance.empty:
        return []
    if not {"email", "event_month"}.issubset(attendance.columns):
        return []

    BAD = {"test@test", "nan", ""}
    att = attendance[~attendance["email"].isin(team_emails | BAD)].copy()

    agg = {}
    if "name" in att.columns:
        agg["name"] = ("name", "last")
    if "company" in att.columns:
        agg["company"] = ("company", "last")
    agg["months_ever"] = ("event_month", "nunique")

    grp = att.groupby("email").agg(**{k: v for k, v in agg.items()}).reset_index()
    grp = grp[grp["months_ever"] >= 2].sort_values("months_ever", ascending=False).head(MAX_REGULARS)

    result = []
    for _, row in grp.iterrows():
        name    = _title_case(str(row.get("name", "") or ""))
        company = _title_case(str(row.get("company", "") or ""))
        if not name or name.lower() in ("nan", ""):
            name = row["email"].split("@")[0].replace(".", " ").title()
        result.append(dict(
            name        = name,
            company     = "" if company.lower() in ("nan", "") else company,
            email       = row["email"],
            months_ever = int(row["months_ever"]),
            initials    = _initials(name),
        ))
    return result


def build_lapsed(attendance: pd.DataFrame, team_emails: set) -> list:
    if attendance.empty:
        return []
    if not {"email", "event_month"}.issubset(attendance.columns):
        return []

    BAD = {"test@test", "nan", ""}
    att = attendance[~attendance["email"].isin(team_emails | BAD)].copy()

    periods = att["event_month"].apply(_parse_period).dropna()
    if periods.empty:
        return []
    latest   = periods.max()
    cutoff   = latest - 3   # lapsed = not seen in last 3 months
    too_old  = latest - 24  # ignore anyone gone more than 2 years

    agg = {"months_ever": ("event_month", "nunique"), "last_month": ("event_month", "max")}
    if "name" in att.columns:
        agg["name"] = ("name", "last")
    if "company" in att.columns:
        agg["company"] = ("company", "last")

    grp = att.groupby("email").agg(**agg).reset_index()
    grp["last_period"] = grp["last_month"].apply(_parse_period)
    lapsed = grp[
        (grp["months_ever"] >= 2) &
        (grp["last_period"].apply(lambda p: p is not None and too_old <= p <= cutoff))
    ].sort_values("last_month", ascending=False).head(MAX_LAPSED)

    result = []
    for _, row in lapsed.iterrows():
        name    = _title_case(str(row.get("name", "") or ""))
        company = _title_case(str(row.get("company", "") or ""))
        if not name or name.lower() in ("nan", ""):
            name = row["email"].split("@")[0].replace(".", " ").title()
        try:
            p = _parse_period(row["last_month"])
            last_seen = p.strftime("%b %Y") if p else str(row["last_month"])
        except Exception:
            last_seen = str(row["last_month"])
        result.append(dict(
            name      = name,
            company   = "" if company.lower() in ("nan", "") else company,
            email     = row["email"],
            last_seen = last_seen,
            initials  = _initials(name),
        ))
    return result


def build_prospects(town_code: str, town_label: str) -> list:
    df = _nc(_safe_excel(BUZZPLUS_FILE, "Plus_Prospects"))
    df = _filter_town(df, town_code, town_label)
    if df.empty:
        return []

    tier_order = {"buzz plus strong": 0, "buzz plus possible": 1}
    if "prospect_tier" in df.columns:
        df["_t"] = df["prospect_tier"].astype(str).str.strip().str.lower().map(tier_order).fillna(2)
    else:
        df["_t"] = 1
    if "visits_12m" in df.columns:
        df["visits_12m"] = pd.to_numeric(df["visits_12m"], errors="coerce").fillna(0)
    df = df.sort_values(["_t", "visits_12m"], ascending=[True, False]).head(MAX_PROSPECTS)

    result = []
    for _, row in df.iterrows():
        name    = _title_case(str(row.get("person_name", "") or ""))
        company = _title_case(str(row.get("company", "") or ""))
        tier    = str(row.get("prospect_tier", "")).strip()
        visits  = int(row.get("visits_12m", 0) or 0)
        if not name or name.lower() in ("nan", ""):
            name = str(row.get("email", "")).split("@")[0].replace(".", " ").title()
        result.append(dict(
            name      = name,
            company   = "" if company.lower() in ("nan", "") else company,
            email     = str(row.get("email", "")),
            visits    = visits,
            is_strong = "strong" in tier.lower(),
            initials  = _initials(name),
        ))
    return result


def build_sponsors(town_base: Path, town_code: str, town_label: str) -> tuple:
    df = load_sponsors_csv(town_base, town_code)
    cap_df   = _nc(_safe_excel(SPONSOR_FILE, "Sponsor_Capacity"))
    cap_row  = _filter_town(cap_df, town_code, town_label)
    slots_left = _first_int(cap_row, "capacity_left")

    if df.empty:
        return [], slots_left

    active = []
    for _, row in df.iterrows():
        end = str(row.get("end_date", "") or "").strip()
        if end and end.lower() not in ("nan", ""):
            try:
                end_dt = pd.to_datetime(end, dayfirst=True, errors="coerce")
                if pd.notna(end_dt) and end_dt.date() < TODAY:
                    continue
            except Exception:
                pass
        company = _title_case(str(row.get("company", "") or ""))
        contact = _title_case(str(row.get("primary_contact", "") or row.get("sponsor_name", "") or ""))
        renewal = ""
        if end and end.lower() not in ("nan", ""):
            try:
                rd = pd.to_datetime(end, dayfirst=True, errors="coerce")
                if pd.notna(rd):
                    renewal = rd.strftime("%-d %b %Y")
            except Exception:
                renewal = end
        if company and company.lower() not in ("nan", ""):
            active.append(dict(company=company, contact=contact, renewal=renewal))

    return active, slots_left


# ------------------------------------------------------------------
# HTML RENDERER
# ------------------------------------------------------------------

def _empty(msg: str) -> str:
    return f'<p class="empty-msg">{msg}</p>'


def render_html(town_label, host_name, snap, regulars, lapsed,
                prospects, sponsors, slots_left, event_date=None):

    event_line = f" &nbsp;&middot;&nbsp; Next event: {event_date}" if event_date else ""

    # Snapshot cards — each topped with its brand colour
    def sc(label, value, colour):
        return (f'<div class="snap-card" style="border-top:4px solid {colour}">'
                f'<div class="snap-label">{label}</div>'
                f'<div class="snap-value" style="color:{colour}">{value or "&#8212;"}</div>'
                f'</div>')

    snap_html = (
        sc("Avg visitors / event (12m)", snap["avg_attendance"], TEAL)
        + sc("Buzz Plus prospects",       snap["total_prospects"], ORANGE)
        + sc("Active Buzz Plus",          snap["active_plus"],     PINK)
        + sc("Sponsor slots open",        slots_left,               LIME)
    )

    # Section 1 – Regulars
    if regulars:
        rows1 = ""
        for p in regulars:
            ev = f"{p['months_ever']} event{'s' if p['months_ever'] != 1 else ''}"
            rows1 += (f'<div class="person-row">'
                      f'<div class="avatar" style="background:{TEAL_BG};color:{TEAL_FG}">{p["initials"]}</div>'
                      f'<div class="person-info">'
                      f'<div class="person-name">{p["name"]}</div>'
                      f'<div class="person-sub">{p["company"] or p["email"]}</div>'
                      f'</div>'
                      f'<div class="badge" style="background:{TEAL_BG};color:{TEAL_FG}">{ev}</div>'
                      f'</div>')
        s1 = rows1
        p1 = (f'<div class="prompt-box" style="border-left-color:{TEAL}">'
              f'<div class="prompt-label" style="color:{TEAL}">Suggested approach</div>'
              'These visitors keep coming back &#8212; a moment of genuine recognition goes a long way. '
              '&ldquo;Great to see you again, you&rsquo;re part of what makes this room.&rdquo;'
              '</div>')
    else:
        s1 = _empty("No regulars yet &#8212; this will grow over time.")
        p1 = ""

    # Section 2 – Lapsed
    if lapsed:
        rows2 = ""
        for p in lapsed:
            rows2 += (f'<div class="person-row">'
                      f'<div class="avatar" style="background:{ORG_BG};color:{ORG_FG}">{p["initials"]}</div>'
                      f'<div class="person-info">'
                      f'<div class="person-name">{p["name"]}</div>'
                      f'<div class="person-sub">{p["company"] or p["email"]}</div>'
                      f'</div>'
                      f'<div class="badge" style="background:#F3F4F6;color:#6B7280">Last seen {p["last_seen"]}</div>'
                      f'</div>')
        s2 = rows2
        p2 = (f'<div class="prompt-box" style="border-left-color:{ORANGE}">'
              f'<div class="prompt-label" style="color:{ORANGE}">Suggested opener</div>'
              '&ldquo;Hi [name] &#8212; it&rsquo;s been a while since we&rsquo;ve seen you at Buzz. '
              'We&rsquo;ve got a great crowd this month &#8212; would love to see you back.&rdquo;'
              '</div>')
    else:
        s2 = _empty("No lapsed visitors to re-engage this month.")
        p2 = ""

    # Section 3 – Buzz Plus
    if prospects:
        rows3 = ""
        for p in prospects:
            av_bg = PINK_BG if p["is_strong"] else ORG_BG
            av_fg = PINK_FG if p["is_strong"] else ORG_FG
            b_bg  = PINK_BG if p["is_strong"] else ORG_BG
            b_fg  = PINK_FG if p["is_strong"] else ORG_FG
            vis   = f"{p['visits']} visit{'s' if p['visits'] != 1 else ''} this year"
            rows3 += (f'<div class="person-row">'
                      f'<div class="avatar" style="background:{av_bg};color:{av_fg}">{p["initials"]}</div>'
                      f'<div class="person-info">'
                      f'<div class="person-name">{p["name"]}</div>'
                      f'<div class="person-sub">{p["company"] or p["email"]}</div>'
                      f'</div>'
                      f'<div class="badge" style="background:{b_bg};color:{b_fg}">{vis}</div>'
                      f'</div>')
        s3 = rows3
        p3 = (f'<div class="prompt-box" style="border-left-color:{PINK}">'
              f'<div class="prompt-label" style="color:{PINK}">Suggested approach</div>'
              'These visitors clearly love Buzz &#8212; they&rsquo;ve been coming regularly. '
              'A natural, low-pressure mention is all it takes: '
              '&ldquo;Have you heard about Buzz Plus? Given how often you&rsquo;re here, '
              'it might be worth a look.&rdquo;'
              '</div>')
    else:
        s3 = _empty("No Buzz Plus prospects yet &#8212; check back as attendance builds.")
        p3 = ""

    # Section 4 – Sponsors
    spon_rows = ""
    for s in sponsors:
        ren = f" &nbsp;&middot;&nbsp; renews {s['renewal']}" if s["renewal"] else ""
        spon_rows += (f'<div class="sponsor-row">'
                      f'<div><span class="sponsor-name">{s["company"]}</span>'
                      f'<span class="sponsor-contact">{s["contact"]}{ren}</span></div>'
                      f'<div class="badge" style="background:{LIME_BG};color:{LIME_FG}">Active</div>'
                      f'</div>')
    for _ in range(slots_left):
        spon_rows += ('<div class="sponsor-row">'
                      '<div class="sponsor-name" style="font-style:italic;opacity:0.45">Slot available</div>'
                      '<div class="badge" style="background:#F3F4F6;color:#6B7280">Open</div>'
                      '</div>')
    if not spon_rows:
        spon_rows = _empty("No sponsor data available.")

    sec4_desc = (f"{slots_left} slot{'s' if slots_left != 1 else ''} still available"
                 if slots_left > 0 else "All slots filled")

    generated = TODAY.strftime("%-d %b %Y")

    return f"""<!DOCTYPE html>
<html lang="en">
<head>
<meta charset="UTF-8">
<meta name="viewport" content="width=device-width,initial-scale=1.0">
<title>Business Buzz &ndash; {town_label} Host Pack &ndash; {MONTH_LABEL}</title>
<style>
*,*::before,*::after{{box-sizing:border-box;margin:0;padding:0}}
body{{font-family:'Century Gothic','Gill Sans',Calibri,sans-serif;font-size:14px;color:#111827;background:#F0F0F0;padding:24px 16px 40px}}
.pack{{max-width:720px;margin:0 auto;box-shadow:0 4px 24px rgba(0,0,0,.10)}}
.header{{background:{TEAL};border-radius:12px 12px 0 0;padding:28px 32px 22px;color:#fff;position:relative;overflow:hidden}}
.header::before{{content:'';position:absolute;right:-40px;top:-40px;width:220px;height:220px;border-radius:50%;background:rgba(255,255,255,.08)}}
.header::after{{content:'';position:absolute;right:60px;bottom:-60px;width:160px;height:160px;border-radius:50%;background:rgba(255,255,255,.05)}}
.header-top{{display:flex;justify-content:space-between;align-items:flex-start;margin-bottom:6px;position:relative}}
.brand-dots{{display:flex;gap:6px;margin-bottom:12px}}
.dot{{width:12px;height:12px;border-radius:50%;border:2px solid rgba(255,255,255,.35)}}
.brand-label{{font-size:11px;letter-spacing:.12em;text-transform:uppercase;color:rgba(255,255,255,.7);margin-bottom:8px}}
.pack-title{{font-size:24px;font-weight:700;color:#fff;line-height:1.2}}
.pack-subtitle{{font-size:13px;color:rgba(255,255,255,.75);margin-top:4px}}
.header-date{{font-size:12px;color:rgba(255,255,255,.6);text-align:right;white-space:nowrap;position:relative}}
.greeting{{margin-top:16px;font-size:14px;color:rgba(255,255,255,.9);line-height:1.65;border-top:1px solid rgba(255,255,255,.2);padding-top:14px;position:relative}}
.colour-bar{{display:flex;height:5px}}
.cb1{{flex:1;background:{TEAL}}}.cb2{{flex:1;background:{ORANGE}}}.cb3{{flex:1;background:{PINK}}}.cb4{{flex:1;background:{LIME}}}
.body{{background:#fff;border:1px solid #E5E7EB;border-top:none;border-radius:0 0 12px 12px;padding:24px 32px 28px}}
.snapshot{{display:grid;grid-template-columns:repeat(4,minmax(0,1fr));gap:10px;margin-bottom:26px}}
.snap-card{{background:#FAFAFA;border:1px solid #E5E7EB;border-radius:8px;padding:12px 12px 10px}}
.snap-label{{font-size:11px;color:#6B7280;line-height:1.35;margin-bottom:6px}}
.snap-value{{font-size:28px;font-weight:700;line-height:1}}
.section{{margin-bottom:22px}}
.section-header{{display:flex;align-items:center;gap:10px;padding-bottom:9px;border-bottom:2px solid #F3F4F6;margin-bottom:11px}}
.sec-num{{width:27px;height:27px;border-radius:50%;display:flex;align-items:center;justify-content:center;font-size:12px;font-weight:700;flex-shrink:0;color:#fff}}
.sec-title{{font-size:15px;font-weight:700;color:#111827}}
.sec-desc{{font-size:12px;color:#9CA3AF;margin-left:auto}}
.person-row{{display:flex;align-items:center;gap:10px;padding:8px 0;border-bottom:1px solid #F9F9F9}}
.person-row:last-of-type{{border-bottom:none}}
.avatar{{width:36px;height:36px;border-radius:50%;display:flex;align-items:center;justify-content:center;font-size:11px;font-weight:700;flex-shrink:0}}
.person-info{{flex:1;min-width:0}}
.person-name{{font-size:13px;font-weight:600;color:#111827;white-space:nowrap;overflow:hidden;text-overflow:ellipsis}}
.person-sub{{font-size:12px;color:#6B7280;white-space:nowrap;overflow:hidden;text-overflow:ellipsis}}
.badge{{font-size:11px;padding:3px 10px;border-radius:999px;white-space:nowrap;flex-shrink:0;font-weight:600}}
.prompt-box{{border-left:3px solid;border-radius:0 6px 6px 0;padding:9px 13px;margin-top:11px;font-size:12px;color:#374151;line-height:1.6;background:#FAFAFA}}
.prompt-label{{font-size:10px;font-weight:700;letter-spacing:.08em;text-transform:uppercase;margin-bottom:3px}}
.sponsor-row{{display:flex;justify-content:space-between;align-items:center;padding:8px 0;border-bottom:1px solid #F9F9F9;font-size:13px}}
.sponsor-row:last-of-type{{border-bottom:none}}
.sponsor-name{{font-weight:600;color:#111827;display:block}}
.sponsor-contact{{font-size:12px;color:#6B7280;margin-top:1px;display:block}}
.empty-msg{{font-size:13px;color:#9CA3AF;padding:8px 0;font-style:italic}}
.footer{{display:flex;justify-content:space-between;font-size:11px;color:#9CA3AF;margin-top:22px;padding-top:14px;border-top:1px solid #E5E7EB}}
@media print{{
  body{{background:#fff;padding:0}}
  .pack{{max-width:100%;box-shadow:none}}
  .header{{border-radius:0}}
  .body{{border-radius:0;border:none;padding:20px 28px 24px}}
  .prompt-box,.section{{break-inside:avoid}}
}}
</style>
</head>
<body>
<div class="pack">
  <div class="header">
    <div class="header-top">
      <div>
        <div class="brand-dots">
          <div class="dot" style="background:#fff"></div>
          <div class="dot" style="background:{ORANGE}"></div>
          <div class="dot" style="background:{PINK}"></div>
          <div class="dot" style="background:{LIME}"></div>
        </div>
        <div class="brand-label">Business Buzz &nbsp;&middot;&nbsp; Leicestershire &amp; Rutland</div>
        <div class="pack-title">{town_label} host pack</div>
        <div class="pack-subtitle">{MONTH_LABEL}{event_line}</div>
      </div>
      <div class="header-date">Generated {generated}</div>
    </div>
    <div class="greeting">Hi {host_name} &#8212; here&rsquo;s your monthly briefing. Three things to do before the event, and a few faces to know on the day. This pack is for your eyes only &#8212; please don&rsquo;t share or forward it.</div>
  </div>
  <div class="colour-bar"><div class="cb1"></div><div class="cb2"></div><div class="cb3"></div><div class="cb4"></div></div>
  <div class="body">
    <div class="snapshot">{snap_html}</div>
    <div class="section">
      <div class="section-header">
        <div class="sec-num" style="background:{TEAL}">1</div>
        <div class="sec-title">Faces to recognise this month</div>
        <div class="sec-desc">Regulars worth a warm welcome</div>
      </div>
      {s1}{p1}
    </div>
    <div class="section">
      <div class="section-header">
        <div class="sec-num" style="background:{ORANGE}">2</div>
        <div class="sec-title">Worth a call before the event</div>
        <div class="sec-desc">Lapsed visitors to re-engage</div>
      </div>
      {s2}{p2}
    </div>
    <div class="section">
      <div class="section-header">
        <div class="sec-num" style="background:{PINK}">3</div>
        <div class="sec-title">Buzz Plus conversations to have on the night</div>
        <div class="sec-desc">{snap['strong']} strong &nbsp;&middot;&nbsp; {snap['possible']} possible</div>
      </div>
      {s3}{p3}
    </div>
    <div class="section">
      <div class="section-header">
        <div class="sec-num" style="background:{LIME}">4</div>
        <div class="sec-title">Sponsors this month</div>
        <div class="sec-desc">{sec4_desc}</div>
      </div>
      {spon_rows}
    </div>
    <div class="footer">
      <span>Business Buzz &nbsp;&middot;&nbsp; Leicestershire &amp; Rutland region</span>
      <span>For host use only &nbsp;&middot;&nbsp; not for distribution</span>
    </div>
  </div>
</div>
</body>
</html>"""


# ------------------------------------------------------------------
# EMAIL TEXT RENDERER
# ------------------------------------------------------------------

def render_email(town_label, host_name, snap, regulars, lapsed, prospects, slots_left):
    lines = [
        f"Hi {host_name},",
        "",
        f"Your {town_label} host pack for {MONTH_LABEL} is attached.",
        "Here's the short version:",
        "",
        "YOUR MONTH AT A GLANCE",
        f"  Avg visitors per event (12 months):  {snap['avg_attendance']}",
        f"  Buzz Plus prospects:                  {snap['total_prospects']} ({snap['strong']} strong, {snap['possible']} possible)",
        f"  Active Buzz Plus members:             {snap['active_plus']}",
        f"  Sponsor slots available:              {slots_left}",
        "",
        "1. FACES TO RECOGNISE THIS MONTH",
    ]
    if regulars:
        for p in regulars:
            co = f" ({p['company']})" if p["company"] else ""
            lines.append(f"  {p['name']}{co} — {p['months_ever']} events")
    else:
        lines.append("  No regulars yet.")

    lines += ["", "2. WORTH A CALL BEFORE THE EVENT"]
    if lapsed:
        for p in lapsed:
            co = f" ({p['company']})" if p["company"] else ""
            lines.append(f"  {p['name']}{co} — last seen {p['last_seen']}")
        lines += [
            "",
            '  Suggested opener: "Hi [name] — it\'s been a while since we\'ve seen',
            '  you at Buzz. We\'ve got a great crowd this month — would love to see you back."',
        ]
    else:
        lines.append("  No lapsed visitors to re-engage this month.")

    lines += ["", "3. BUZZ PLUS CONVERSATIONS"]
    if prospects:
        for p in prospects:
            co   = f" ({p['company']})" if p["company"] else ""
            tier = "Strong" if p["is_strong"] else "Possible"
            lines.append(f"  {p['name']}{co} — {p['visits']} visits ({tier})")
        lines += [
            "",
            '  Suggested approach: "Have you heard about Buzz Plus? Given how often',
            '  you\'re here, it might be worth a look."',
        ]
    else:
        lines.append("  No Buzz Plus prospects yet.")

    lines += [
        "",
        "Open the attached PDF for the full pack.",
        "Any questions, just reply to this email.",
        "",
        "Thanks,",
        "Emma",
        "",
        "--",
        "Business Buzz Leicestershire & Rutland",
        "This email and its attachment are for host use only. Please do not forward.",
    ]
    return "\n".join(lines)


# ------------------------------------------------------------------
# MAIN BUILD FUNCTION
# ------------------------------------------------------------------

def build_host_pack(town_code: str, event_date: Optional[str] = None) -> None:
    if town_code not in TOWNS:
        raise ValueError(f"Unknown town: {town_code!r}. Valid: {list(TOWNS)}")

    folder_name, town_label = TOWNS[town_code]
    town_base   = REGION_ROOT / folder_name
    monthly_dir = town_base / "data_curated" / "monthly"

    host_name   = load_host_name(town_base, town_code)
    team_emails = load_team_emails(town_base, town_code)
    attendance  = load_attendance(monthly_dir)

    snap               = build_snapshot(town_code, town_label)
    regulars           = build_regulars(attendance, team_emails)
    lapsed             = build_lapsed(attendance, team_emails)
    prospects          = build_prospects(town_code, town_label)
    sponsors, slots    = build_sponsors(town_base, town_code, town_label)

    # Prefer slots_left from dashboard snapshot over sponsors CSV count
    slots_left = snap["slots_left"] if snap["slots_left"] > 0 else slots

    html_str  = render_html(town_label, host_name, snap, regulars, lapsed,
                            prospects, sponsors, slots_left, event_date)
    email_str = render_email(town_label, host_name, snap, regulars, lapsed,
                             prospects, slots_left)

    out_dir = REGION_CURATED / "host_packs"
    out_dir.mkdir(parents=True, exist_ok=True)

    safe = town_label.replace(" ", "_")
    html_path  = out_dir / f"HostPack_{safe}.html"
    email_path = out_dir / f"HostPack_{safe}_email.txt"

    html_path.write_text(html_str,   encoding="utf-8")
    email_path.write_text(email_str, encoding="utf-8")

    print(f"[OK] {town_label}")
    print(f"     HTML  -> {html_path}")
    print(f"     Email -> {email_path}")


# ------------------------------------------------------------------
# CLI
# ------------------------------------------------------------------

def main() -> None:
    ap = argparse.ArgumentParser(description="Build HTML host packs for Business Buzz towns (v2).")
    ap.add_argument("--town", default="ALL",
                    help="Town code (e.g. Loughborough) or ALL")
    ap.add_argument("--event-date", default=None,
                    help="Optional next event date shown in header, e.g. 'Thursday 20 March 2026'")
    args = ap.parse_args()

    town = args.town.strip()
    if town.upper() == "ALL":
        for code in TOWNS:
            try:
                build_host_pack(code, event_date=args.event_date)
            except Exception as exc:
                print(f"[FAIL] {code}: {exc}")
    else:
        build_host_pack(town, event_date=args.event_date)


if __name__ == "__main__":
    main()
