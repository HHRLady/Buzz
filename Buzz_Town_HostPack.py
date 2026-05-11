"""
Buzz Town Host Pack Generator
Generates a branded HTML host pack (with Chart.js visuals) and plain-text
email body for each town.

Usage:
    python Buzz_Town_HostPack.py --town ALL
    python Buzz_Town_HostPack.py --town MarketHarborough
    python Buzz_Town_HostPack.py --town ALL --event-date "Wednesday 16 April 2026"
"""

import os
import argparse
import json
from datetime import datetime, date
from dateutil.relativedelta import relativedelta
import pandas as pd

def _read_csv_safe(path, **kwargs):
    """Read a CSV trying UTF-8 first, then latin-1 as fallback.
    Handles files saved by Windows/Excel in either encoding."""
    try:
        return __import__('pandas').read_csv(path, encoding='utf-8-sig', **kwargs)
    except (UnicodeDecodeError, Exception):
        return __import__('pandas').read_csv(path, encoding='latin-1', **kwargs)


# ── Brand colours ──────────────────────────────────────────────────────────────
TEAL   = "#00A19A"
ORANGE = "#F39200"
PINK   = "#D60B52"
LIME   = "#B6BD00"
DARK   = "#222222"
GREY   = "#cccccc"

# ── Town config ────────────────────────────────────────────────────────────────
TOWNS = [
    "MarketHarborough",
    "Leicester",
    "Lutterworth",
    "Hinckley",
    "Loughborough",
]

TOWN_DISPLAY = {
    "MarketHarborough": "Market Harborough",
    "Leicester":        "Leicester",
    "Lutterworth":      "Lutterworth",
    "Hinckley":         "Hinckley",
    "Loughborough":     "Loughborough",
}

# ── Thresholds ─────────────────────────────────────────────────────────────────
REGULAR_MIN_VISITS    = 3
REGULAR_MONTHS        = 12
LAPSED_MONTHS         = 3
LAPSED_MAX_MONTHS     = 6    # months shown in the pack (display window)
LAPSED_FULL_MAX_MONTHS = 24  # full look-back window (counted but not listed)
BUZZPLUS_STRONG_MIN   = 6
BUZZPLUS_POSSIBLE_MIN = 3
NEW_TO_REGION_MONTHS  = 2
STREAK_MONTHS         = 12   # how many months of dots to show per regular
RETURN_WINDOW_MONTHS  = 3    # look back this many months for first-timer return rate
# Event excellence: paying guest target (excludes hosts, RL, and ambassadors at their own event)
PAYING_GUEST_TARGET   = 25   # minimum for event excellence


# ══════════════════════════════════════════════════════════════════════════════
# DATA LOADING
# ══════════════════════════════════════════════════════════════════════════════

def get_base_folder():
    return os.path.dirname(os.path.abspath(__file__))


def get_town_folder(base, town_code):
    return os.path.join(base, f"Buzz_Event_Dashboard_{town_code}")


def load_roles(base, town_code):
    path = os.path.join(get_town_folder(base, town_code), "data_ref",
                        f"roles_{town_code}.csv")
    if not os.path.exists(path):
        return pd.DataFrame(columns=["email", "role", "name", "end_date"])
    df = _read_csv_safe(path)
    df.columns = [c.strip().lower() for c in df.columns]
    df["end_date"] = df["end_date"].astype(str).str.strip()
    df["is_current"] = df["end_date"].isin(["", "nan", "NaT", "None"])
    return df[df["is_current"]]


def load_sponsors(base, town_code):
    path = os.path.join(get_town_folder(base, town_code), "data_ref",
                        f"sponsors_{town_code}.csv")
    if not os.path.exists(path):
        return pd.DataFrame()
    df = _read_csv_safe(path)
    df.columns = [c.strip().lower() for c in df.columns]
    today_str = date.today().strftime("%Y-%m-%d")
    if "end_date" in df.columns:
        df["end_date"] = pd.to_datetime(df["end_date"], dayfirst=True, errors="coerce")
        df = df[df["end_date"].isna() | (df["end_date"] >= pd.Timestamp(today_str))]
    return df


def load_exclusions(base):
    path = os.path.join(base, "exclude_region.csv")
    if not os.path.exists(path):
        return pd.DataFrame(columns=["email", "name", "company", "reason",
                                     "exclude_regulars", "exclude_lapsed",
                                     "exclude_buzzplus", "exclude_by_company"])
    df = _read_csv_safe(path)
    df.columns = [c.strip().lower() for c in df.columns]
    for col in ["exclude_regulars", "exclude_lapsed",
                "exclude_buzzplus", "exclude_by_company"]:
        if col in df.columns:
            df[col] = df[col].astype(str).str.upper().isin(["TRUE", "YES", "1"])
        else:
            df[col] = False
    if "company" not in df.columns:
        df["company"] = ""
    return df


def load_region_master_people(base):
    path = os.path.join(base, "Buzz_Region_Curated", "region_master_people.csv")
    if not os.path.exists(path):
        return pd.DataFrame()
    df = _read_csv_safe(path)
    df.columns = [c.strip().lower() for c in df.columns]
    return df


def load_region_team_emails(base):
    """Return set of all team member emails across the region from buzzplus_intelligence."""
    path = os.path.join(base, "Buzz_Region_Curated", "buzzplus_intelligence.xlsx")
    if not os.path.exists(path):
        return set()
    try:
        df = pd.read_excel(path, sheet_name="Excluded_Team_Roles")
        df.columns = [c.strip().lower() for c in df.columns]
        if "email" not in df.columns:
            return set()
        return set(df["email"].astype(str).str.strip().str.lower())
    except Exception:
        return set()


def load_active_plus_emails(base):
    """Return set of current Buzz Plus member emails (to exclude from prospects).
    Reads from Buzz_Region_Ref/buzz_plus_members.xlsx — active members only."""
    path = os.path.join(base, "Buzz_Region_Ref", "buzz_plus_members.xlsx")
    if not os.path.exists(path):
        return set()
    try:
        df = pd.read_excel(path)
        df.columns = [c.strip().lower() for c in df.columns]
        if "email" not in df.columns:
            return set()
        # Only exclude active members (ignore expired/cancelled)
        if "status" in df.columns:
            df = df[df["status"].astype(str).str.strip().str.lower() == "active"]
        return set(df["email"].astype(str).str.strip().str.lower())
    except Exception:
        return set()


def load_declined_plus_emails(base):
    """Return set of emails for people who have explicitly declined Buzz Plus."""
    path = os.path.join(base, "Buzz_Region_Ref", "buzz_plus_declined.xlsx")
    if not os.path.exists(path):
        return set()
    try:
        df = pd.read_excel(path)
        df.columns = [c.strip().lower() for c in df.columns]
        if "email" not in df.columns:
            return set()
        return set(df["email"].astype(str).str.strip().str.lower())
    except Exception:
        return set()


def load_attendance_history(base, town_code):
    """
    Read all monthly _attendance.xlsx files.
    Returns (person_df, all_attendance_df) or None.
    """
    monthly_dir = os.path.join(get_town_folder(base, town_code),
                               "data_curated", "monthly")
    if not os.path.exists(monthly_dir):
        return None

    frames = []
    for fname in sorted(os.listdir(monthly_dir)):
        if "_attendance" not in fname or not fname.endswith(".xlsx"):
            continue
        month_str = None
        for part in fname.replace(".xlsx", "").split("_"):
            if len(part) == 7 and part[4] == "-":
                month_str = part
                break
        if not month_str:
            continue
        try:
            df = pd.read_excel(os.path.join(monthly_dir, fname))
            df.columns = [c.strip().lower() for c in df.columns]
            if "email" not in df.columns:
                continue
            df["event_month"] = month_str
            frames.append(df)
        except Exception:
            continue

    if not frames:
        return None

    all_df = pd.concat(frames, ignore_index=True)
    all_df["email"] = all_df["email"].astype(str).str.strip().str.lower()

    def first_non_empty(s):
        vals = s.dropna().astype(str).str.strip()
        vals = vals[vals != ""]
        return vals.iloc[0] if len(vals) else ""

    import re as _re
    def clean_company(name):
        """Strip address fragments appended to company names (e.g. 'AM Joinery Ltd, Gilbert Avenue, Rugby')."""
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

    agg = {"first_seen_town": ("event_month", "min"),
           "last_seen_town":  ("event_month", "max"),
           "visits_total":    ("event_month", "count")}
    if "name" in all_df.columns:
        agg["name"] = ("name", first_non_empty)
    if "company" in all_df.columns:
        agg["company"] = ("company", first_non_empty)

    person_df = all_df.groupby("email").agg(**agg).reset_index()
    if "name" not in person_df.columns:
        person_df["name"] = person_df["email"]
    if "company" not in person_df.columns:
        person_df["company"] = ""
    # Strip address fragments from company names
    person_df["company"] = person_df["company"].apply(clean_company)

    # Rolling 12-month window ending at the latest event month on record
    latest_event = all_df["event_month"].max()
    rolling_start = (datetime.strptime(latest_event + "-01", "%Y-%m-%d")
                     - relativedelta(months=11)).strftime("%Y-%m")
    year_counts = (all_df[all_df["event_month"] >= rolling_start]
                   .groupby("email").size()
                   .rename("visits_this_year")
                   .reset_index())
    person_df = person_df.merge(year_counts, on="email", how="left")
    person_df["visits_this_year"] = (person_df["visits_this_year"]
                                     .fillna(0).astype(int))
    return person_df, all_df


def load_321_data(base, town_code, latest_month=None):
    """Return a dict of 3-2-1 data for the most recent month available.

    Reads Town_Month_Summary (aggregated) first; falls back to Event_Level for
    notes. Filters to latest_month if provided, otherwise uses the most recent
    row for this town.
    """
    path = os.path.join(base, "Buzz_Region_Curated", "buzz_321_intelligence.xlsx")
    if not os.path.exists(path):
        return None
    try:
        sheets = pd.read_excel(path, sheet_name=None)
        # Normalise column names across all sheets
        for name in sheets:
            sheets[name].columns = [c.strip().lower() for c in sheets[name].columns]

        def filter_town(df):
            col = next((c for c in ["town_code", "town"] if c in df.columns), None)
            if not col:
                return pd.DataFrame()
            return df[df[col].astype(str).str.lower() == town_code.lower()]

        # --- primary: Town_Month_Summary ---
        summary = None
        if "Town_Month_Summary" in sheets or "town_month_summary" in [s.lower() for s in sheets]:
            sname = next(s for s in sheets if s.lower() == "town_month_summary")
            town_rows = filter_town(sheets[sname])
            if not town_rows.empty:
                if latest_month and "event_month_key" in town_rows.columns:
                    # Normalise key to YYYY-MM before comparing — avoids wrong-month fallback
                    norm_keys = town_rows["event_month_key"].astype(str).str[:7]
                    exact = town_rows[norm_keys == str(latest_month)[:7]]
                    summary = exact.iloc[-1].to_dict() if not exact.empty else None
                    # Do NOT fall back to the latest row — that silently shows stale data
                else:
                    summary = town_rows.iloc[-1].to_dict()

        # --- secondary: Event_Level (for notes) ---
        notes_val = None
        if "Event_Level" in sheets or "event_level" in [s.lower() for s in sheets]:
            ename = next(s for s in sheets if s.lower() == "event_level")
            ev_rows = filter_town(sheets[ename])
            if not ev_rows.empty:
                if latest_month and "event_month_key" in ev_rows.columns:
                    exact = ev_rows[ev_rows["event_month_key"] == latest_month]
                    ev_rows = exact if not exact.empty else ev_rows
                ev_row = ev_rows.iloc[-1]
                notes_val = ev_row.get("notes") if pd.notna(ev_row.get("notes", float("nan"))) else None

        if summary is None:
            return None
        if notes_val:
            summary["notes"] = notes_val
        return summary

    except Exception:
        pass
    return None


# ══════════════════════════════════════════════════════════════════════════════
# ANALYTICS
# ══════════════════════════════════════════════════════════════════════════════

def build_attendance_trend(all_attendance, latest_month, n_months=13):
    """Return ordered list of (label, count) for last n_months events."""
    months = []
    for i in range(n_months - 1, -1, -1):
        m = (datetime.strptime(latest_month + "-01", "%Y-%m-%d")
             - relativedelta(months=i)).strftime("%Y-%m")
        months.append(m)

    # Only months that actually have data
    month_counts = (all_attendance.groupby("event_month").size()
                    .reindex(months, fill_value=0))

    labels = []
    for m in months:
        try:
            labels.append(datetime.strptime(m + "-01", "%Y-%m-%d").strftime("%b %y"))
        except Exception:
            labels.append(m)

    return labels, month_counts.tolist()


def build_streaks(all_attendance, regulars_df, latest_month):
    """
    For each regular, return a list of booleans (True=attended) for the last
    STREAK_MONTHS months, most recent last.
    Returns dict: email -> [bool, bool, ...]
    """
    months = []
    for i in range(STREAK_MONTHS - 1, -1, -1):
        m = (datetime.strptime(latest_month + "-01", "%Y-%m-%d")
             - relativedelta(months=i)).strftime("%Y-%m")
        months.append(m)

    attended_set = (all_attendance.groupby("email")["event_month"]
                    .apply(set).to_dict())

    streaks = {}
    for _, row in regulars_df.iterrows():
        email = str(row["email"]).strip().lower()
        s = attended_set.get(email, set())
        streaks[email] = [m in s for m in months]
    return streaks


def build_firsttimer_return_rate(all_attendance, latest_month):
    """
    Look at first-timers from RETURN_WINDOW_MONTHS ago.
    Return (rate_pct, returned_count, total_count, cohort_month_label).
    """
    cohort_month = (
        datetime.strptime(latest_month + "-01", "%Y-%m-%d")
        - relativedelta(months=RETURN_WINDOW_MONTHS)
    ).strftime("%Y-%m")

    # People whose first ever visit was in cohort_month
    first_visits = (all_attendance.groupby("email")["event_month"].min()
                    .reset_index(drop=False))
    first_visits.columns = ["email", "first_seen"]
    cohort = first_visits[first_visits["first_seen"] == cohort_month]["email"].tolist()

    if not cohort:
        return None, 0, 0, cohort_month

    # How many came back at least once after cohort_month?
    after = all_attendance[
        (all_attendance["email"].isin(cohort)) &
        (all_attendance["event_month"] > cohort_month)
    ]
    returned = after["email"].nunique()
    total    = len(cohort)
    rate     = round((returned / total) * 100) if total else 0
    label    = datetime.strptime(cohort_month + "-01",
                                 "%Y-%m-%d").strftime("%B %Y")
    return rate, returned, total, label


# ══════════════════════════════════════════════════════════════════════════════
# CATEGORISATION
# ══════════════════════════════════════════════════════════════════════════════

def is_excluded(row, exclusions, category):
    col_map = {"regulars": "exclude_regulars",
               "lapsed":   "exclude_lapsed",
               "buzzplus": "exclude_buzzplus"}
    col     = col_map.get(category, "")
    email   = str(row.get("email", "")).strip().lower()
    company = str(row.get("company", "")).strip().lower()

    email_match = exclusions[
        exclusions["email"].astype(str).str.strip().str.lower() == email
    ]
    if not email_match.empty:
        excl = email_match.iloc[0]
        if col and excl.get(col, False):
            return True
        if excl.get("exclude_by_company", False):
            return True

    if company:
        co_excl = exclusions[
            exclusions["exclude_by_company"].astype(bool) &
            exclusions["company"].astype(str).str.strip().str.lower().apply(
                lambda c: bool(c) and c in company
            )
        ]
        if not co_excl.empty:
            return True
    return False


def enrich_with_new_flags(person_df, region_master, latest_month):
    person_df = person_df.copy()
    person_df["new_to_event"] = (person_df["first_seen_town"] == latest_month)

    # "New to Region" = their very first Buzz event anywhere in the region was THIS month.
    # Using a rolling window caused active people (e.g. 5 visits since joining 2 months ago)
    # to be incorrectly badged. Exact match on latest_month is the correct intent.
    if (region_master is not None and not region_master.empty
            and "first_seen_month_region" in region_master.columns
            and "email" in region_master.columns):
        new_emails = set(
            region_master[region_master["first_seen_month_region"] == latest_month]
            ["email"].astype(str).str.strip().str.lower()
        )
        person_df["new_to_region"] = (
            person_df["email"].astype(str).str.strip().str.lower()
            .isin(new_emails)
        )
    else:
        person_df["new_to_region"] = False
    return person_df


def build_new_faces(person_df, exclusions, latest_month, region_team_emails=None):
    new_faces = person_df[person_df["new_to_event"]].copy()
    if region_team_emails:
        new_faces = new_faces[
            ~new_faces["email"].astype(str).str.strip().str.lower().isin(region_team_emails)
        ]
    if not new_faces.empty:
        new_faces = new_faces[
            ~new_faces.apply(lambda r: is_excluded(r, exclusions, "regulars"), axis=1)
        ]
    return new_faces.sort_values("name")


def build_regulars(person_df, exclusions, latest_month, region_team_emails=None):
    cutoff = (
        datetime.strptime(latest_month + "-01", "%Y-%m-%d")
        - relativedelta(months=REGULAR_MONTHS)
    ).strftime("%Y-%m")
    regs = person_df[
        (person_df["visits_this_year"] >= REGULAR_MIN_VISITS) &
        (person_df["last_seen_town"] >= cutoff)
    ].copy()
    if region_team_emails:
        regs = regs[
            ~regs["email"].astype(str).str.strip().str.lower().isin(region_team_emails)
        ]
    if not regs.empty:
        regs = regs[~regs.apply(lambda r: is_excluded(r, exclusions, "regulars"), axis=1)]
    return regs.sort_values("visits_this_year", ascending=False)


def build_lapsed(person_df, exclusions, latest_month, region_team_emails=None):
    """
    Returns (display_df, hidden_count).
    display_df  — lapsed visitors last seen within LAPSED_MAX_MONTHS (6) months
    hidden_count — additional people lapsed 6–LAPSED_FULL_MAX_MONTHS months ago
    """
    lapsed_cutoff = (
        datetime.strptime(latest_month + "-01", "%Y-%m-%d")
        - relativedelta(months=LAPSED_MONTHS)
    ).strftime("%Y-%m")
    too_old_display = (
        datetime.strptime(latest_month + "-01", "%Y-%m-%d")
        - relativedelta(months=LAPSED_MAX_MONTHS)
    ).strftime("%Y-%m")
    too_old_full = (
        datetime.strptime(latest_month + "-01", "%Y-%m-%d")
        - relativedelta(months=LAPSED_FULL_MAX_MONTHS)
    ).strftime("%Y-%m")

    # Display list: 3–6 months
    display = person_df[
        (person_df["last_seen_town"] < lapsed_cutoff) &
        (person_df["last_seen_town"] >= too_old_display)
    ].copy()
    # Exclude region-wide team members (hosts/ambassadors from other towns etc.)
    if region_team_emails and not display.empty:
        display = display[
            ~display["email"].astype(str).str.strip().str.lower().isin(region_team_emails)
        ]
    if not display.empty:
        display = display[
            ~display.apply(lambda r: is_excluded(r, exclusions, "lapsed"), axis=1)
        ]

    # Full list: 3–24 months (count only)
    full = person_df[
        (person_df["last_seen_town"] < lapsed_cutoff) &
        (person_df["last_seen_town"] >= too_old_full)
    ]
    hidden_count = max(0, len(full) - len(display))

    return display.sort_values("last_seen_town", ascending=False), hidden_count


def build_buzzplus(person_df, exclusions, roles_df,
                   region_team_emails=None, active_plus_emails=None,
                   declined_plus_emails=None, sponsor_companies=None):
    # Combine exclusion sets: current-town team + region-wide team + existing Plus members
    # + people who have explicitly declined
    team_emails = set(roles_df["email"].astype(str).str.strip().str.lower().tolist())
    if region_team_emails:
        team_emails |= region_team_emails
    if active_plus_emails:
        team_emails |= active_plus_emails
    if declined_plus_emails:
        team_emails |= declined_plus_emails

    # Normalised sponsor company names for company-level exclusion
    sp_names = [s.strip().lower() for s in (sponsor_companies or []) if s and s.strip()]

    strong   = person_df[person_df["visits_this_year"] >= BUZZPLUS_STRONG_MIN].copy()
    possible = person_df[
        (person_df["visits_this_year"] >= BUZZPLUS_POSSIBLE_MIN) &
        (person_df["visits_this_year"] < BUZZPLUS_STRONG_MIN)
    ].copy()
    for df in [strong, possible]:
        team_mask = df["email"].astype(str).str.strip().str.lower().isin(team_emails)
        df.drop(df[team_mask].index, inplace=True)
        if not df.empty:
            excl_mask = df.apply(
                lambda r: is_excluded(r, exclusions, "buzzplus"), axis=1
            )
            df.drop(df[excl_mask].index, inplace=True)
        # Exclude anyone whose company name matches a current sponsor
        if sp_names and not df.empty:
            sp_mask = df["company"].astype(str).str.strip().str.lower().apply(
                lambda c: any(bool(sp) and sp in c for sp in sp_names)
            )
            df.drop(df[sp_mask].index, inplace=True)
    return (strong.sort_values("visits_this_year", ascending=False),
            possible.sort_values("visits_this_year", ascending=False))


def format_last_seen(month_str):
    try:
        return datetime.strptime(month_str + "-01", "%Y-%m-%d").strftime("%B %Y")
    except Exception:
        return month_str


# ══════════════════════════════════════════════════════════════════════════════
# HTML HELPERS
# ══════════════════════════════════════════════════════════════════════════════

def badge(label, colour):
    return (
        f'<span style="display:inline-block;background:{colour};color:#fff;'
        f'font-size:10px;font-weight:700;padding:2px 8px;border-radius:10px;'
        f'margin-left:6px;vertical-align:middle;letter-spacing:0.4px;">'
        f'{label}</span>'
    )


def streak_dots_html(attended_list):
    """Return 12 coloured dot spans. Filled = attended, grey = missed."""
    dots = []
    for attended in attended_list:
        colour = TEAL if attended else GREY
        dots.append(
            f'<span style="display:inline-block;width:10px;height:10px;'
            f'border-radius:50%;background:{colour};margin:0 1px;'
            f'vertical-align:middle;" title="{"attended" if attended else "absent"}"></span>'
        )
    return '<span style="margin-left:8px;">' + "".join(dots) + '</span>'


def person_row_html(row, extra_text="", streak_dots=""):
    name    = str(row.get("name", "")).strip() or str(row.get("email", ""))
    company = str(row.get("company", "")).strip()
    co_html = (f'<span style="font-size:13px;color:#666;">{company}</span>'
               if company else "")
    extra_html = (f'<span style="font-size:12px;color:#888;margin-left:4px;">'
                  f'{extra_text}</span>' if extra_text else "")
    region_badge = badge("NEW TO REGION", PINK) if row.get("new_to_region") else ""
    return (
        f'<div style="padding:8px 0;border-bottom:1px solid #eee;">'
        f'<div style="display:flex;align-items:center;flex-wrap:wrap;gap:4px;">'
        f'<strong style="font-size:14px;">{name}</strong>'
        f'{(" &nbsp;" + co_html) if co_html else ""}'
        f'{extra_html}{region_badge}'
        f'</div>'
        f'{streak_dots}'
        f'</div>'
    )


def person_card_html(row, extra_text="", streak_dots="", first_visit=False):
    """Card-style person entry for two-column grid layout."""
    name    = str(row.get("name", "")).strip() or str(row.get("email", ""))
    company = str(row.get("company", "")).strip()
    co_html = (f'<div style="font-size:12px;color:#666;margin-top:1px;">{company}</div>'
               if company else "")
    extra_html = (f'<div style="font-size:11px;color:#888;margin-top:2px;">{extra_text}</div>'
                  if extra_text else "")
    region_badge = badge("NEW TO REGION", PINK) if row.get("new_to_region") else ""
    first_badge  = badge("FIRST VISIT", ORANGE) if first_visit else ""
    return (
        f'<div style="padding:8px 10px;background:#f8f9fa;border-radius:5px;'
        f'border:1px solid #e8e8e8;min-width:0;">'
        f'<div style="display:flex;flex-wrap:wrap;gap:3px;align-items:baseline;">'
        f'<strong style="font-size:13px;">{name}</strong>'
        f'{first_badge}{region_badge}'
        f'</div>'
        f'{co_html}{extra_html}{streak_dots}'
        f'</div>'
    )


def two_col_grid(items_html):
    """Wrap a list of HTML strings in a two-column grid."""
    return (
        '<div style="display:grid;grid-template-columns:1fr 1fr;gap:6px;'
        'margin-top:4px;">'
        + "".join(items_html)
        + '</div>'
    )


def section_html(number, colour, title, body, subtitle=""):
    sub = (f'<p style="font-size:12px;color:#666;margin:0 0 12px;">{subtitle}</p>'
           if subtitle else "")
    return (
        f'<div style="margin-bottom:22px;border:1px solid #e8e8e8;'
        f'border-radius:6px;overflow:hidden;">'
        f'<div style="background:{colour};padding:12px 18px;'
        f'display:flex;align-items:center;gap:12px;">'
        f'<span style="background:rgba(255,255,255,0.25);color:#fff;'
        f'font-size:13px;font-weight:800;width:26px;height:26px;border-radius:50%;'
        f'display:inline-flex;align-items:center;justify-content:center;">'
        f'{number}</span>'
        f'<span style="color:#fff;font-size:15px;font-weight:700;">{title}</span>'
        f'</div>'
        f'<div style="padding:16px 18px;">{sub}{body}</div>'
        f'</div>'
    )


def snap_card_html(colour, value, label):
    return (
        f'<div style="flex:1;min-width:100px;background:#fff;'
        f'border-top:4px solid {colour};border-radius:6px;padding:12px 10px;'
        f'text-align:center;box-shadow:0 1px 4px rgba(0,0,0,0.07);">'
        f'<div style="font-size:26px;font-weight:800;color:{colour};">{value}</div>'
        f'<div style="font-size:11px;color:#666;margin-top:4px;">{label}</div>'
        f'</div>'
    )


def def_item_html(term, definition):
    return (
        f'<div style="margin-bottom:9px;">'
        f'<strong style="color:{DARK};">{term}:</strong> '
        f'<span style="color:#555;">{definition}</span></div>'
    )


# ══════════════════════════════════════════════════════════════════════════════
# HTML GENERATION
# ══════════════════════════════════════════════════════════════════════════════

def generate_html(town_code, host_name, event_date_str, latest_month,
                  new_faces, regulars, lapsed, strong, possible,
                  sponsors, data_321, total_visitors, avg_visitors,
                  trend_labels, trend_data, streaks,
                  return_rate, returned_count, total_firsttimers, cohort_label,
                  lapsed_hidden=0):

    town_display = TOWN_DISPLAY.get(town_code, town_code)
    generated_date = datetime.now().strftime("%d %B %Y")

    # ── Snapshot cards ─────────────────────────────────────────────────────────
    # Event excellence: paying guests proxy from 321 in_room data (until app exports this)
    paying_guests = None
    if data_321:
        for k in data_321:
            if "in_room" in k.lower() or "total_in_room" in k.lower():
                v = data_321[k]
                if pd.notna(v):
                    try:
                        paying_guests = int(float(v))
                    except (ValueError, TypeError):
                        pass
                    break
    excellence_colour = LIME if (paying_guests is not None and paying_guests >= PAYING_GUEST_TARGET) else PINK
    excellence_label  = f"In room (target {PAYING_GUEST_TARGET})"
    excellence_val    = paying_guests if paying_guests is not None else "—"

    snapshot = (
        snap_card_html(TEAL,              total_visitors or "—",         "Booked this month") +
        snap_card_html(ORANGE,            avg_visitors   or "—",         "Monthly avg (booked)") +
        snap_card_html(excellence_colour, excellence_val,                 excellence_label) +
        snap_card_html(PINK,              len(new_faces),                 "New faces") +
        snap_card_html(LIME,              len(regulars),                  "Regulars") +
        snap_card_html(TEAL,              len(lapsed),                    "Lapsed to re-engage") +
        snap_card_html(ORANGE,            len(strong) + len(possible),    "Buzz Plus prospects")
    )

    # ── Attendance trend chart (Chart.js) ──────────────────────────────────────
    chart_labels_js = json.dumps(trend_labels)
    chart_data_js   = json.dumps(trend_data)
    # Highlight the latest month bar
    chart_colours_js = json.dumps(
        [TEAL if i < len(trend_data) - 1 else ORANGE
         for i in range(len(trend_data))]
    )

    trend_chart_html = f"""
<div style="margin-bottom:22px;border:1px solid #e8e8e8;border-radius:6px;overflow:hidden;">
  <div style="background:{TEAL};padding:12px 18px;">
    <span style="color:#fff;font-size:15px;font-weight:700;">Bookings — last 13 months</span>
  </div>
  <div style="padding:16px 18px;">
    <canvas id="trendChart_{town_code}" height="80"></canvas>
  </div>
</div>
<script>
(function() {{
  var ctx = document.getElementById('trendChart_{town_code}').getContext('2d');
  new Chart(ctx, {{
    type: 'bar',
    data: {{
      labels: {chart_labels_js},
      datasets: [{{
        label: 'Visitors',
        data: {chart_data_js},
        backgroundColor: {chart_colours_js},
        borderRadius: 4,
        borderSkipped: false
      }}]
    }},
    options: {{
      responsive: true,
      plugins: {{
        legend: {{ display: false }},
        tooltip: {{
          callbacks: {{
            label: function(ctx) {{ return ctx.parsed.y + ' booked'; }}
          }}
        }}
      }},
      scales: {{
        y: {{
          beginAtZero: true,
          ticks: {{ precision: 0 }},
          grid: {{ color: '#f0f0f0' }}
        }},
        x: {{ grid: {{ display: false }} }}
      }}
    }}
  }});
}})();
</script>"""

    # ── First-timer return rate ────────────────────────────────────────────────
    if return_rate is not None and total_firsttimers > 0:
        rate_colour = LIME if return_rate >= 60 else (ORANGE if return_rate >= 35 else PINK)
        return_rate_html = f"""
<div style="margin-bottom:22px;border:1px solid #e8e8e8;border-radius:6px;overflow:hidden;">
  <div style="background:{ORANGE};padding:12px 18px;">
    <span style="color:#fff;font-size:15px;font-weight:700;">First-timer return rate</span>
  </div>
  <div style="padding:16px 18px;display:flex;align-items:center;gap:28px;flex-wrap:wrap;">
    <div style="text-align:center;">
      <div style="font-size:52px;font-weight:800;color:{rate_colour};line-height:1;">
        {return_rate}%
      </div>
      <div style="font-size:12px;color:#666;margin-top:4px;">returned</div>
    </div>
    <div style="font-size:14px;color:#444;line-height:1.6;">
      Of the <strong>{total_firsttimers} first-timers</strong> who came in
      <strong>{cohort_label}</strong>,<br>
      <strong>{returned_count}</strong> have been back at least once since.
      <br><br>
      <span style="font-size:12px;color:#888;">
        Target: 60%+ is strong &nbsp;·&nbsp; 35–59% is healthy &nbsp;·&nbsp; under 35% needs attention
      </span>
    </div>
    <div style="flex:1;min-width:160px;">
      <canvas id="returnChart_{town_code}" height="120"></canvas>
    </div>
  </div>
</div>
<script>
(function() {{
  var ctx = document.getElementById('returnChart_{town_code}').getContext('2d');
  new Chart(ctx, {{
    type: 'doughnut',
    data: {{
      labels: ['Returned', 'Did not return'],
      datasets: [{{
        data: [{returned_count}, {total_firsttimers - returned_count}],
        backgroundColor: ['{rate_colour}', '#eeeeee'],
        borderWidth: 0
      }}]
    }},
    options: {{
      responsive: true,
      cutout: '72%',
      plugins: {{
        legend: {{ position: 'bottom', labels: {{ font: {{ size: 11 }} }} }}
      }}
    }}
  }});
}})();
</script>"""
    else:
        return_rate_html = ""

    # ── Section 1: New faces ───────────────────────────────────────────────────
    if new_faces.empty:
        new_faces_body = ('<p style="color:#888;font-style:italic;">'
                         'No first-time visitors this month.</p>')
    else:
        new_faces_body = two_col_grid([
            person_card_html(r, first_visit=True)
            for _, r in new_faces.iterrows()
        ])
    new_faces_section = section_html(
        "1", ORANGE, "New faces this month", new_faces_body,
        "Attending this event for the very first time — give them a warm welcome "
        "and make sure they know what to expect from Buzz."
    )

    # ── Section 2: Regulars with streak dots ───────────────────────────────────
    if regulars.empty:
        regs_body = ('<p style="color:#888;font-style:italic;">'
                     'No regulars to show this month.</p>')
    else:
        streak_legend = (
            f'<p style="font-size:11px;color:#888;margin:0 0 10px;">'
            f'Dots show attendance over the last {STREAK_MONTHS} months '
            f'(left = oldest &nbsp;·&nbsp; '
            f'<span style="color:{TEAL};font-weight:700;">&#9679;</span> attended &nbsp;·&nbsp; '
            f'<span style="color:{GREY};font-weight:700;">&#9679;</span> absent)</p>'
        )
        reg_cards = []
        for _, r in regulars.iterrows():
            email = str(r["email"]).strip().lower()
            dots  = streak_dots_html(streaks.get(email, [False] * STREAK_MONTHS))
            reg_cards.append(person_card_html(
                r,
                extra_text=f"{int(r['visits_this_year'])} visits (rolling 12 months)",
                streak_dots=dots
            ))
        regs_body = streak_legend + two_col_grid(reg_cards)

    reg_subtitle = (
        f"Visited {REGULAR_MIN_VISITS}+ times in the last {REGULAR_MONTHS} months — "
        f"make them feel recognised."
    )
    regs_section = section_html("2", TEAL, "Faces to recognise this month",
                                regs_body, reg_subtitle)

    # ── Section 3: Lapsed ─────────────────────────────────────────────────────
    if lapsed.empty:
        lapsed_body = ('<p style="color:#888;font-style:italic;">'
                       'No lapsed visitors in the last 6 months.</p>')
    else:
        lapsed_cards = [person_card_html(
            r,
            extra_text=f"Last seen: {format_last_seen(str(r.get('last_seen_town', '')))}"
        ) for _, r in lapsed.iterrows()]
        lapsed_body = two_col_grid(lapsed_cards)
        if lapsed_hidden > 0:
            lapsed_body += (
                f'<p style="font-size:12px;color:#888;margin-top:10px;font-style:italic;">'
                f'Plus {lapsed_hidden} more people last seen 6+ months ago — '
                f'contact your regional lead for the full list.</p>'
            )

    lapsed_sub = (
        f"Haven't attended for {LAPSED_MONTHS}+ months — "
        f"a quick message before the event could bring them back. "
        'Try: <em>"Haven\'t seen you at Buzz for a while — '
        "we've got a great crowd this month, would love to see you back."
        '</em>'
    )
    lapsed_section = section_html("3", PINK,
                                  "People worth a call before the event",
                                  lapsed_body, lapsed_sub)

    # ── Section 4: Buzz Plus ───────────────────────────────────────────────────
    if strong.empty and possible.empty:
        plus_body = ('<p style="color:#888;font-style:italic;">'
                     'No Buzz Plus prospects this month.</p>')
    else:
        plus_body = ""
        if not strong.empty:
            plus_body += ('<p style="font-size:12px;font-weight:700;color:#888;'
                          'margin:0 0 6px;">STRONG PROSPECTS</p>')
            plus_body += two_col_grid([
                person_card_html(r,
                    extra_text=f"{int(r['visits_this_year'])} visits (rolling 12 months)")
                for _, r in strong.iterrows()
            ])
        if not possible.empty:
            plus_body += ('<p style="font-size:12px;font-weight:700;color:#888;'
                          'margin:12px 0 6px;">POSSIBLE PROSPECTS</p>')
            plus_body += two_col_grid([
                person_card_html(r,
                    extra_text=f"{int(r['visits_this_year'])} visits (rolling 12 months)")
                for _, r in possible.iterrows()
            ])

    plus_sub = (
        "Catch them during the event — a natural, low-pressure mention "
        "between 10am and noon is all it takes. "
        'Try: <em>"You clearly love Buzz — have you heard about Buzz Plus?"</em>'
    )
    plus_section = section_html("4", LIME,
                                "Buzz Plus conversations to have on the day",
                                plus_body, plus_sub)

    # ── Section 5: Sponsors ────────────────────────────────────────────────────
    if sponsors.empty:
        sponsors_body = ('<p style="color:#888;font-style:italic;">'
                         'No current sponsors.</p>')
    else:
        def _fmt_date(d):
            """Format a date value as 'd Mon YYYY', handling NaT/NaN gracefully."""
            try:
                if pd.isna(d) or str(d).strip() in ("", "nan", "NaT", "None"):
                    return ""
                return pd.Timestamp(d).strftime("%d %b %Y").lstrip("0")
            except Exception:
                raw = str(d).strip()
                return raw[:10] if raw else ""

        sp_rows = []
        for _, s in sponsors.iterrows():
            company = str(s.get("company", "")).strip()
            contact = str(s.get("primary_contact",
                         s.get("sponsor_name", ""))).strip()
            level   = str(s.get("package", s.get("sponsor_level", ""))).strip()
            display = company or contact
            contact_line = (
                f'<div style="font-size:12px;color:#666;">{contact}</div>'
                if contact and contact != company else ""
            )
            level_line = (
                f'<div style="font-size:11px;color:#888;">{level}</div>'
                if level and level not in ("nan", "") else ""
            )
            sd = _fmt_date(s.get("start_date", ""))
            ed = _fmt_date(s.get("end_date",   ""))
            date_parts = []
            if sd: date_parts.append(f"From {sd}")
            if ed: date_parts.append(f"Until {ed}")
            date_line = (
                f'<div style="font-size:11px;color:#888;">'
                f'{"&nbsp;&nbsp;·&nbsp;&nbsp;".join(date_parts)}</div>'
                if date_parts else ""
            )
            sp_rows.append(
                f'<div style="padding:8px 0;border-bottom:1px solid #eee;">'
                f'<strong>{display}</strong>'
                f'{contact_line}{level_line}{date_line}'
                f'</div>'
            )
        sponsors_body = "".join(sp_rows)
    sponsors_section = section_html("5", TEAL, "Current sponsors", sponsors_body)

    # ── Section 6: 3-2-1 insights ─────────────────────────────────────────────
    insights_section = ""
    if data_321:
        def val(key):
            for k in data_321:
                if key.lower() in k.lower():
                    v = data_321[k]
                    if pd.notna(v) and str(v).strip() not in ("", "nan"):
                        return str(v)
            return "—"

        def stat_card(colour, value, label):
            return (
                f'<div style="flex:1;min-width:100px;background:#fff;'
                f'border-top:3px solid {colour};border-radius:6px;padding:10px 8px;'
                f'text-align:center;box-shadow:0 1px 3px rgba(0,0,0,0.06);">'
                f'<div style="font-size:22px;font-weight:800;color:{colour};">{value}</div>'
                f'<div style="font-size:11px;color:#666;margin-top:3px;">{label}</div></div>'
            )

        def fmt_rate(raw):
            """Convert a decimal rate (0.46) to a percentage string (46%)."""
            if raw == "—":
                return "—"
            try:
                f = float(raw)
                return f"{round(f * 100)}%"
            except (ValueError, TypeError):
                return raw

        def count_val(key):
            """Prefer total_ prefixed count columns over rate columns."""
            # Try total_ prefix first (Town_Month_Summary)
            for k in data_321:
                if k.startswith("total_") and key.lower() in k.lower():
                    v = data_321[k]
                    if pd.notna(v) and str(v).strip() not in ("", "nan"):
                        return str(int(float(v)))
            return val(key)

        note = val("notes") or val("note")
        note_html = (
            f'<div style="margin-top:12px;padding:10px 14px;background:#f0faf9;'
            f'border-left:3px solid {TEAL};border-radius:4px;font-style:italic;'
            f'font-size:13px;color:#444;">"{note}"</div>'
            if note and note != "—" else ""
        )
        stats_html = (
            stat_card(TEAL,   fmt_rate(val("avg_interaction_rate") or val("interaction_rate")), "Interaction rate") +
            stat_card(ORANGE, count_val("met_3_new") or count_val("met_3"),                     "Met 3 new people") +
            stat_card(PINK,   count_val("one_2_ones"),                                           "1-2-1s booked") +
            stat_card(LIME,   count_val("brought_some1") or count_val("brought_someone"),        "Brought someone")
        )
        insights_body = (f'<div style="display:flex;gap:10px;flex-wrap:wrap;">'
                         f'{stats_html}</div>{note_html}')
        insights_section = section_html("6", LIME,
                                        "Event quality — 3-2-1 insights",
                                        insights_body)

    # ── Definitions footer ─────────────────────────────────────────────────────
    defs = (
        def_item_html("New faces",
            "Attending this town for the very first time. "
            "May also carry a <em>New to Region</em> badge if this is their "
            "first Business Buzz event anywhere in the region.")
        + def_item_html("New to Region",
            "Their very first Business Buzz event anywhere in this region — "
            "this month is month one for them.")
        + def_item_html("Streak dots",
            f"The {STREAK_MONTHS} coloured dots next to each regular show their "
            f"attendance over the last {STREAK_MONTHS} months (left = oldest). "
            f"Teal = attended, grey = absent.")
        + def_item_html("Faces to recognise",
            f"Attended {REGULAR_MIN_VISITS}+ times in the rolling 12 months "
            f"and visited within the last {REGULAR_MONTHS} months.")
        + def_item_html("Lapsed visitors",
            f"Not attended for {LAPSED_MONTHS}+ months, but visited within "
            f"the last {LAPSED_MAX_MONTHS} months. Realistic re-engagement conversations.")
        + def_item_html("First-timer return rate",
            f"Percentage of first-timers from {RETURN_WINDOW_MONTHS} months ago "
            f"who have been back at least once. 60%+ is strong.")
        + def_item_html("Buzz Plus",
            f"A marketing subscription. <strong>Strong prospects:</strong> "
            f"{BUZZPLUS_STRONG_MIN}+ visits this year. "
            f"<strong>Possible:</strong> {BUZZPLUS_POSSIBLE_MIN}–{BUZZPLUS_STRONG_MIN - 1} visits.")
        + def_item_html("Rolling 12 months", "Visit counts use the 12 months up to the most recent event, not a fixed calendar year.")
    )

    definitions_block = (
        f'<div style="margin-top:28px;padding:20px 24px;background:#f7f7f7;'
        f'border-top:3px solid {TEAL};border-radius:6px;">'
        f'<p style="font-size:12px;font-weight:800;color:{TEAL};margin:0 0 14px;'
        f'text-transform:uppercase;letter-spacing:1px;">How to read this pack</p>'
        f'{defs}'
        f'</div>'
    )

    # ── Assemble ───────────────────────────────────────────────────────────────
    html = f"""<!DOCTYPE html>
<html lang="en">
<head>
<meta charset="UTF-8">
<meta name="viewport" content="width=device-width,initial-scale=1.0">
<title>Host Pack – {town_display} – {event_date_str}</title>
<script src="https://cdnjs.cloudflare.com/ajax/libs/Chart.js/3.9.1/chart.min.js"></script>
<style>
  * {{ box-sizing:border-box; margin:0; padding:0; }}
  body {{
    font-family: 'Century Gothic', 'Gill Sans', Calibri, sans-serif;
    background: #fff;
    color: {DARK};
    max-width: 820px;
    margin: 0 auto;
    padding: 24px;
  }}
  @media print {{
    body {{ padding: 8px; max-width: 100%; }}
    script {{ display: none; }}
  }}
</style>
</head>
<body>

<div style="background:{TEAL};border-radius:8px 8px 0 0;padding:24px 28px;color:#fff;">
  <div style="font-size:11px;letter-spacing:2px;text-transform:uppercase;
    opacity:0.8;margin-bottom:6px;">Business Buzz Host Pack</div>
  <div style="font-size:26px;font-weight:800;">{town_display}</div>
 <div style="font-size:15px;opacity:0.9;margin-top:4px;">
    Hi {host_name} — {event_date_str}</div>
  <div style="font-size:13px;opacity:0.85;margin-top:14px;padding-top:12px;
    border-top:1px solid rgba(255,255,255,0.25);line-height:1.6;">
    Here&rsquo;s your monthly briefing &mdash; a few faces to recognise on the day,
    anyone worth a call before the event, and the Buzz Plus conversations to have
    on the night. This pack is for your eyes only &mdash; please don&rsquo;t share
    or forward it.
  </div>
</div>

<div style="display:flex;height:6px;margin-bottom:20px;">
  <div style="flex:1;background:{TEAL};"></div>
  <div style="flex:1;background:{ORANGE};"></div>
  <div style="flex:1;background:{PINK};"></div>
  <div style="flex:1;background:{LIME};"></div>
</div>

<div style="display:flex;gap:10px;flex-wrap:wrap;margin-bottom:24px;">
  {snapshot}
</div>

{trend_chart_html}
{return_rate_html}
{new_faces_section}
{regs_section}
{lapsed_section}
{plus_section}
{sponsors_section}
{insights_section}

{definitions_block}

<div style="margin-top:20px;padding-top:14px;border-top:1px solid #eee;
  font-size:11px;color:#bbb;text-align:center;">
  Generated {generated_date} · Business Buzz Leics &amp; Rutland
</div>

</body>
</html>"""

    return html


# ══════════════════════════════════════════════════════════════════════════════
# EMAIL TEXT
# ══════════════════════════════════════════════════════════════════════════════

def generate_email_text(town_code, host_name, event_date_str,
                        new_faces, regulars, lapsed, strong, possible,
                        return_rate, returned_count, total_firsttimers, cohort_label):
    town_display = TOWN_DISPLAY.get(town_code, town_code)

    def fmt(row, note=""):
        name    = str(row.get("name", "")).strip() or str(row.get("email", ""))
        company = str(row.get("company", "")).strip()
        flags   = []
        if row.get("new_to_event"):
            flags.append("FIRST VISIT")
        if row.get("new_to_region"):
            flags.append("NEW TO REGION")
        flag_str = f" [{', '.join(flags)}]" if flags else ""
        return (f"  • {name}"
                + (f" – {company}" if company else "")
                + (f" — {note}" if note else "")
                + flag_str)

    lines = [
        f"BUSINESS BUZZ HOST PACK — {town_display.upper()}",
        f"Hi {host_name} — {event_date_str}",
        "=" * 60,
        "",
    ]

    if return_rate is not None and total_firsttimers > 0:
        lines += [
            f"FIRST-TIMER RETURN RATE: {return_rate}%",
            f"  {returned_count} of {total_firsttimers} first-timers from {cohort_label} have returned.",
            "",
        ]

    lines += [
        "1. NEW FACES THIS MONTH",
        "   (Attending this town for the very first time)",
        "",
    ]
    if new_faces.empty:
        lines.append("   No first-time visitors this month.")
    else:
        lines += [fmt(r) for _, r in new_faces.iterrows()]

    lines += [
        "", f"2. FACES TO RECOGNISE",
        f"   ({REGULAR_MIN_VISITS}+ visits in rolling 12 months, seen in last {REGULAR_MONTHS} months)",
        "",
    ]
    if regulars.empty:
        lines.append("   No regulars to show this month.")
    else:
        lines += [fmt(r, f"{int(r['visits_this_year'])} visits this year")
                  for _, r in regulars.iterrows()]

    lines += [
        "", "3. PEOPLE WORTH A CALL BEFORE THE EVENT",
        f"   (Lapsed {LAPSED_MONTHS}+ months, within last {LAPSED_MAX_MONTHS} months)",
        "",
    ]
    if lapsed.empty:
        lines.append("   No lapsed visitors to re-engage this month.")
    else:
        lines += [fmt(r, f"last seen {format_last_seen(str(r['last_seen_town']))}")
                  for _, r in lapsed.iterrows()]

    lines += [
        "", "4. BUZZ PLUS CONVERSATIONS TO HAVE ON THE NIGHT",
        "   (10am–noon — keep it natural)",
        "",
    ]
    all_p = list(strong.iterrows()) + list(possible.iterrows())
    if not all_p:
        lines.append("   No Buzz Plus prospects this month.")
    else:
        lines += [
            fmt(r, f"{int(r['visits_this_year'])} visits — "
                   f"{'Strong' if int(r['visits_this_year']) >= BUZZPLUS_STRONG_MIN else 'Possible'}")
            for _, r in all_p
        ]

    lines += [
        "",
        "=" * 60,
        "KEY: [FIRST VISIT] = first time at this town",
        f"     [NEW TO REGION] = very first Buzz event anywhere in this region",
        "",
        f"Generated {datetime.now().strftime('%d %B %Y')} · Business Buzz Leics & Rutland",
    ]
    return "\n".join(lines)


# ══════════════════════════════════════════════════════════════════════════════
# MAIN
# ══════════════════════════════════════════════════════════════════════════════

def process_town(base, town_code, event_date_arg=None):
    print(f"\n[{town_code}] Building host pack...")

    roles_df           = load_roles(base, town_code)
    sponsors_df        = load_sponsors(base, town_code)
    exclusions         = load_exclusions(base)
    region_master        = load_region_master_people(base)
    region_team_emails   = load_region_team_emails(base)
    active_plus_emails   = load_active_plus_emails(base)
    declined_plus_emails = load_declined_plus_emails(base)

    result = load_attendance_history(base, town_code)
    if result is None:
        print("  No attendance data — skipping.")
        return

    person_df, all_attendance = result
    if person_df.empty:
        print("  Empty attendance data — skipping.")
        return

    latest_month = person_df["last_seen_town"].max()
    person_df    = enrich_with_new_flags(person_df, region_master, latest_month)

    # Host name
    host_rows = roles_df[
        roles_df["role"].astype(str).str.strip().str.lower() == "host"
    ]
    host_name = "Host"
    if not host_rows.empty:
        full = str(host_rows.iloc[-1]["name"]).strip()
        host_name = full.split()[0] if full else "Host"

    # Attendance stats
    total_visitors = len(all_attendance[all_attendance["event_month"] == latest_month])
    cutoff_12 = (
        datetime.strptime(latest_month + "-01", "%Y-%m-%d")
        - relativedelta(months=12)
    ).strftime("%Y-%m")
    recent = all_attendance[all_attendance["event_month"] >= cutoff_12]
    avg_visitors = (round(recent.groupby("event_month").size().mean())
                    if not recent.empty else None)

    # Sponsor company names — load from the RAW (unfiltered) file so that past sponsors
    # like TaxAssist are still excluded from Buzz Plus even after their contract expires.
    sponsor_companies = []
    raw_sp_path = os.path.join(get_town_folder(base, town_code), "data_ref",
                               f"sponsors_{town_code}.csv")
    try:
        raw_sp = pd.read_csv(raw_sp_path)
        raw_sp.columns = [c.strip().lower() for c in raw_sp.columns]
        if "company" in raw_sp.columns:
            sponsor_companies = [
                str(n).strip() for n in raw_sp["company"]
                if str(n).strip() and str(n).strip().lower() != "nan"
            ]
    except Exception:
        pass

    # Categorise
    new_faces           = build_new_faces(person_df, exclusions, latest_month, region_team_emails)
    regulars            = build_regulars(person_df, exclusions, latest_month, region_team_emails)
    lapsed, lapsed_hidden = build_lapsed(person_df, exclusions, latest_month, region_team_emails)
    strong, possible    = build_buzzplus(person_df, exclusions, roles_df,
                                         region_team_emails, active_plus_emails,
                                         declined_plus_emails,
                                         sponsor_companies=sponsor_companies)
    data_321         = load_321_data(base, town_code, latest_month)

    # Analytics
    trend_labels, trend_data = build_attendance_trend(all_attendance, latest_month)
    streaks = build_streaks(all_attendance, regulars, latest_month)
    return_rate, returned_count, total_ft, cohort_label = \
        build_firsttimer_return_rate(all_attendance, latest_month)

    event_date_str = event_date_arg or datetime.strptime(
        latest_month + "-01", "%Y-%m-%d"
    ).strftime("%B %Y")

    html = generate_html(
        town_code, host_name, event_date_str, latest_month,
        new_faces, regulars, lapsed, strong, possible,
        sponsors_df, data_321, total_visitors, avg_visitors,
        trend_labels, trend_data, streaks,
        return_rate, returned_count, total_ft, cohort_label,
        lapsed_hidden=lapsed_hidden
    )
    email_txt = generate_email_text(
        town_code, host_name, event_date_str,
        new_faces, regulars, lapsed, strong, possible,
        return_rate, returned_count, total_ft, cohort_label
    )

    out_dir = os.path.join(base, "Buzz_Region_Curated", "host_packs")
    os.makedirs(out_dir, exist_ok=True)

    with open(os.path.join(out_dir, f"HostPack_{town_code}.html"),
              "w", encoding="utf-8") as f:
        f.write(html)
    with open(os.path.join(out_dir, f"HostPack_{town_code}_email.txt"),
              "w", encoding="utf-8") as f:
        f.write(email_txt)

    print(f"  New faces: {len(new_faces)}  Regulars: {len(regulars)}  "
          f"Lapsed: {len(lapsed)}  Plus: {len(strong)+len(possible)}")
    if return_rate is not None:
        print(f"  First-timer return rate: {return_rate}% "
              f"({returned_count}/{total_ft} from {cohort_label})")
    print(f"  Saved -> Buzz_Region_Curated/host_packs/HostPack_{town_code}.html")


def main():
    parser = argparse.ArgumentParser(description="Generate Buzz host packs.")
    parser.add_argument("--town", default="ALL", help="Town code or ALL")
    parser.add_argument("--event-date", default=None,
                        help="Event date for header e.g. 'Wednesday 16 April 2026'")
    args = parser.parse_args()

    base  = get_base_folder()
    towns = TOWNS if args.town.upper() == "ALL" else [args.town]

    print("=" * 60)
    print("Business Buzz — Town Host Pack Generator")
    print("=" * 60)

    for town in towns:
        process_town(base, town, event_date_arg=args.event_date)

    print("\nAll done.")


if __name__ == "__main__":
    main()
