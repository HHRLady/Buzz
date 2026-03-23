import argparse
from pathlib import Path
from typing import Optional

import pandas as pd

# ==========================================================
# Buzz – Town Host Pack Builder (UPDATED – Option A)
#
# Option A rules
# - No stamping, no month label on filenames or snapshot.
# - Host Packs always use FULL HISTORY for attendance-driven tabs.
#
# Data governance
# - Email is the primary identifier (compulsory in bookings).
# - Payment logic is not used.
# - Buzz Plus prospects are sourced from regional Buzz Plus Intelligence.
# - Team roles are excluded region-wide from Buzz Plus prospects upstream.
#
# Output location (authoritative)
#   <RegionRoot>\Buzz_Region_Curated\host_packs\
# ==========================================================


def _detect_region_root(start: Path) -> Path:
    p = start
    for _ in range(6):
        if (p / "Buzz_Region_Curated").exists():
            return p
        if p.parent == p:
            break
        p = p.parent
    return start


SCRIPT_DIR = Path(__file__).resolve().parent
REGION_ROOT = _detect_region_root(SCRIPT_DIR)
REGION_CURATED = REGION_ROOT / "Buzz_Region_Curated"

TOWNS = {
    "MarketHarborough": ("Buzz_Event_Dashboard_MarketHarborough", "Market Harborough"),
    "Leicester":        ("Buzz_Event_Dashboard_Leicester",        "Leicester"),
    "Lutterworth":      ("Buzz_Event_Dashboard_Lutterworth",      "Lutterworth"),
    "Hinckley":         ("Buzz_Event_Dashboard_Hinckley",         "Hinckley"),
    "Loughborough":     ("Buzz_Event_Dashboard_Loughborough",     "Loughborough"),
}

BUZZPLUS_FILE = REGION_CURATED / "buzzplus_intelligence.xlsx"
SPONSOR_FILE  = REGION_CURATED / "sponsor_intelligence.xlsx"


def _safe_read_excel(path: Path, sheet_name: str) -> pd.DataFrame:
    if not path.exists():
        return pd.DataFrame()
    try:
        return pd.read_excel(path, sheet_name=sheet_name)
    except Exception:
        return pd.DataFrame()


def _normalise_cols(df: pd.DataFrame) -> pd.DataFrame:
    df = df.copy()
    df.columns = [str(c).strip().lower() for c in df.columns]
    return df


def _write_definitions_tab(writer: pd.ExcelWriter) -> None:
    rows = [
        ["Buzz Plus Strong Prospect",
         "Attendance-based signal indicating a strong opportunity for a Buzz Plus conversation. "
         "Rule: attended 6+ events in the last 12 months. "
         "Hosts and Ambassadors are excluded region-wide."],
        ["Buzz Plus Possible Prospect",
         "Attendance-based signal indicating an emerging opportunity for a Buzz Plus conversation. "
         "Rule: attended 4–5 events in the last 12 months. "
         "Hosts and Ambassadors are excluded region-wide."],
        ["Regulars and Fans",
         "Attendance-based only. Regulars are people who have attended in 2+ distinct months historically. "
         "Derived from curated monthly attendance files."],
        ["Lapsed and Reactivation",
         "Attendance-based only. Lapsed indicates 2+ months ever but not attended in the latest available month. "
         "This is a prompt for thoughtful re-engagement, not a judgement."],
        ["Data Notes",
         "Email is the primary identifier (compulsory in the booking system). Payment type is not used in any logic."],
        ["Folder Truth",
         "Authoritative outputs live under the Region root: Buzz_Region_Curated. Do not use any pack stored inside a town folder."],
    ]
    pd.DataFrame(rows, columns=["Term", "Meaning"]).to_excel(writer, sheet_name="Definitions", index=False)


def _parse_month_yyyy_mm(s: str) -> Optional[pd.Period]:
    s = (s or "").strip()
    if not s:
        return None
    try:
        return pd.Period(s, freq="M")
    except Exception:
        return None


def _load_attendance(monthly_dir: Path) -> pd.DataFrame:
    if not monthly_dir.exists():
        return pd.DataFrame()

    files = sorted(monthly_dir.glob("*_attendance.xlsx"))
    frames: list[pd.DataFrame] = []
    for f in files:
        try:
            frames.append(pd.read_excel(f))
        except Exception:
            continue

    if not frames:
        return pd.DataFrame()

    df = pd.concat(frames, ignore_index=True)
    df = _normalise_cols(df)

    # expected
    if "email" in df.columns:
        df["email"] = df["email"].astype(str).str.strip().str.lower()

    if "event_month" in df.columns:
        df["event_month"] = df["event_month"].astype(str).str.strip()

    return df


def _filter_by_town_code(df: pd.DataFrame, town_code: str) -> pd.DataFrame:
    if df.empty:
        return df
    if "town_code" in df.columns:
        return df[df["town_code"].astype(str).str.strip().str.lower() == town_code.lower()].copy()
    # fallback: town_name match if town_code missing
    if "town_name" in df.columns:
        town_label = TOWNS[town_code][1]
        return df[df["town_name"].astype(str).str.strip().str.lower() == town_label.lower()].copy()
    return df


def build_host_pack(town_code: str) -> Path:
    if town_code not in TOWNS:
        raise ValueError(f"Unknown town: {town_code}")

    town_folder, town_label = TOWNS[town_code]
    town_base = REGION_ROOT / town_folder
    monthly_dir = town_base / "data_curated" / "monthly"

    attendance = _load_attendance(monthly_dir)

    # Regional intelligence sources
    plus_summary   = _normalise_cols(_safe_read_excel(BUZZPLUS_FILE, "Town_Plus_Summary"))
    plus_members   = _normalise_cols(_safe_read_excel(BUZZPLUS_FILE, "Active_Plus_Members"))
    plus_prospects = _normalise_cols(_safe_read_excel(BUZZPLUS_FILE, "Plus_Prospects"))

    sponsor_capacity = _normalise_cols(_safe_read_excel(SPONSOR_FILE, "Sponsor_Capacity"))

    plus_summary_town = _filter_by_town_code(plus_summary, town_code)
    plus_members_town = _filter_by_town_code(plus_members, town_code)
    plus_prospects_town = _filter_by_town_code(plus_prospects, town_code)
    sponsor_capacity_town = _filter_by_town_code(sponsor_capacity, town_code)

    def _first_int(df: pd.DataFrame, col: str) -> int:
        if df.empty or col not in df.columns:
            return 0
        try:
            return int(pd.to_numeric(df[col], errors="coerce").fillna(0).iloc[0])
        except Exception:
            return 0

    snapshot = pd.DataFrame([{
        "Town": town_label,
        "Buzz Plus Strong Prospects": _first_int(plus_summary_town, "buzzplus_strong_prospects"),
        "Buzz Plus Possible Prospects": _first_int(plus_summary_town, "buzzplus_possible_prospects"),
        "Active Buzz Plus Members": _first_int(plus_summary_town, "active_plus_members"),
        "Sponsor Capacity Left": _first_int(sponsor_capacity_town, "capacity_left"),
        "Attendance files loaded": "all history",
    }])

    # Regulars & Lapsed (attendance-based)
    regulars = pd.DataFrame()
    lapsed = pd.DataFrame()

    if not attendance.empty and "email" in attendance.columns and "event_month" in attendance.columns:
        att = attendance.dropna(subset=["email"]).copy()

        # Months attended per email
        grp = (att.groupby("email")["event_month"].nunique().reset_index(name="months_attended_ever"))
        regulars = grp[grp["months_attended_ever"] >= 2].sort_values("months_attended_ever", ascending=False)

        # Lapsed: attended 2+ months ever, but not in latest available month
        try:
            months = att["event_month"].dropna().unique().tolist()
            periods = [p for p in (_parse_month_yyyy_mm(m) for m in months) if p is not None]
            latest = max(periods) if periods else None

            if latest is not None:
                last_by = (att.groupby("email")["event_month"].max().reset_index(name="last_seen_month"))
                last_by["last_seen_period"] = last_by["last_seen_month"].apply(lambda x: _parse_month_yyyy_mm(str(x)))
                merged = grp.merge(last_by, on="email", how="left")
                merged["is_lapsed"] = (merged["months_attended_ever"] >= 2) & (merged["last_seen_period"].apply(lambda p: p is not None and p < latest))
                lapsed = merged[merged["is_lapsed"]].sort_values("last_seen_month")
        except Exception:
            lapsed = pd.DataFrame()

    # Sponsors sheet (town-level ref if present)
    sponsors = pd.DataFrame()
    sponsor_ref = town_base / "data_ref" / f"sponsors_{town_code}.csv"
    if sponsor_ref.exists():
        try:
            sponsors = pd.read_csv(sponsor_ref)
        except Exception:
            sponsors = pd.DataFrame()

    out_dir = REGION_CURATED / "host_packs"
    out_dir.mkdir(parents=True, exist_ok=True)

    out_path = out_dir / f"HostPack_{town_label}.xlsx"

    with pd.ExcelWriter(out_path, engine="openpyxl") as writer:
        snapshot.to_excel(writer, sheet_name="Snapshot", index=False)
        regulars.to_excel(writer, sheet_name="Regulars_and_Fans", index=False)
        lapsed.to_excel(writer, sheet_name="Lapsed_and_Reactivation", index=False)
        plus_members_town.to_excel(writer, sheet_name="Buzz Plus Members", index=False)
        plus_prospects_town.to_excel(writer, sheet_name="Buzz Plus Prospects", index=False)
        sponsors.to_excel(writer, sheet_name="Sponsors", index=False)
        _write_definitions_tab(writer)

    print(f"[OK] Wrote {out_path}")
    return out_path


def main() -> None:
    parser = argparse.ArgumentParser(description="Build Host Packs for one town or all towns (full history, no stamping).")
    parser.add_argument("--town", default="ALL", help="Town code (e.g. MarketHarborough) or ALL")
    args = parser.parse_args()

    town = args.town.strip()

    if town.upper() == "ALL":
        for town_code in TOWNS.keys():
            build_host_pack(town_code)
    else:
        build_host_pack(town)


if __name__ == "__main__":
    main()
