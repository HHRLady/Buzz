import argparse
from pathlib import Path
from typing import Optional

import pandas as pd

# ==========================================================
# Buzz – Town Host Pack Builder (UPDATED – Option A + 3-2-1)
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
# - 3-2-1 insight is sourced from regional 3-2-1 intelligence.
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
INTEL_321_FILE = REGION_CURATED / "buzz_321_intelligence.xlsx"


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
        ["3-2-1 Insights",
         "Event-level networking behaviour signal drawn from the 3-2-1 tracker. "
         "It shows tracker participation, new connections, 1:1 conversations and bring-someone activity."],
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
    if "town_name" in df.columns:
        town_label = TOWNS[town_code][1]
        return df[df["town_name"].astype(str).str.strip().str.lower() == town_label.lower()].copy()
    return df


def _first_int(df: pd.DataFrame, col: str) -> int:
    if df.empty or col not in df.columns:
        return 0
    try:
        return int(pd.to_numeric(df[col], errors="coerce").fillna(0).iloc[0])
    except Exception:
        return 0


def _first_str(df: pd.DataFrame, col: str, default: str = "") -> str:
    if df.empty or col not in df.columns:
        return default
    try:
        val = df[col].iloc[0]
        return default if pd.isna(val) else str(val)
    except Exception:
        return default


def _build_321_sheet_data(town_code: str) -> tuple[pd.DataFrame, pd.DataFrame, pd.DataFrame]:
    event_level = _normalise_cols(_safe_read_excel(INTEL_321_FILE, "Event_Level"))
    town_month = _normalise_cols(_safe_read_excel(INTEL_321_FILE, "Town_Month_Summary"))

    event_level = _filter_by_town_code(event_level, town_code)
    town_month = _filter_by_town_code(town_month, town_code)

    latest_summary = pd.DataFrame()
    if not town_month.empty:
        town_month["event_month_key"] = town_month["event_month_key"].astype(str)
        latest_summary = town_month.sort_values(by=["event_month_key"], ascending=[False]).head(1).copy()

    recent_trend = pd.DataFrame()
    if not event_level.empty:
        if "event_date" in event_level.columns:
            event_level["event_date"] = pd.to_datetime(event_level["event_date"], errors="coerce", dayfirst=True)
        keep_cols = [
            "event_date", "event_month", "in_room", "tracker_users",
            "interaction_rate", "met_3_new", "one_2_ones", "brought_some1",
            "coverage_flag", "quality_flag"
        ]
        for col in keep_cols:
            if col not in event_level.columns:
                event_level[col] = pd.NA
        recent_trend = event_level[keep_cols].sort_values(by=["event_date"], ascending=[False]).head(3).copy()

    notes = pd.DataFrame()
    if not event_level.empty:
        note_cols = ["event_date", "notes", "submitted_by", "coverage_flag", "quality_flag"]
        for col in note_cols:
            if col not in event_level.columns:
                event_level[col] = pd.NA
        notes = event_level[note_cols].copy()
        notes = notes[(notes["notes"].fillna("").astype(str).str.strip() != "") | (notes["quality_flag"].fillna("").astype(str).str.strip() != "")]
        notes = notes.sort_values(by=["event_date"], ascending=[False]).head(10)

    return latest_summary, recent_trend, notes


def build_host_pack(town_code: str) -> Path:
    if town_code not in TOWNS:
        raise ValueError(f"Unknown town: {town_code}")

    town_folder, town_label = TOWNS[town_code]
    town_base = REGION_ROOT / town_folder
    monthly_dir = town_base / "data_curated" / "monthly"

    attendance = _load_attendance(monthly_dir)

    plus_summary   = _normalise_cols(_safe_read_excel(BUZZPLUS_FILE, "Town_Plus_Summary"))
    plus_members   = _normalise_cols(_safe_read_excel(BUZZPLUS_FILE, "Active_Plus_Members"))
    plus_prospects = _normalise_cols(_safe_read_excel(BUZZPLUS_FILE, "Plus_Prospects"))
    sponsor_capacity = _normalise_cols(_safe_read_excel(SPONSOR_FILE, "Sponsor_Capacity"))

    plus_summary_town = _filter_by_town_code(plus_summary, town_code)
    plus_members_town = _filter_by_town_code(plus_members, town_code)
    plus_prospects_town = _filter_by_town_code(plus_prospects, town_code)
    sponsor_capacity_town = _filter_by_town_code(sponsor_capacity, town_code)

    latest_321, trend_321, notes_321 = _build_321_sheet_data(town_code)

    snapshot = pd.DataFrame([{
        "Town": town_label,
        "Buzz Plus Strong Prospects": _first_int(plus_summary_town, "buzzplus_strong_prospects"),
        "Buzz Plus Possible Prospects": _first_int(plus_summary_town, "buzzplus_possible_prospects"),
        "Active Buzz Plus Members": _first_int(plus_summary_town, "active_plus_members"),
        "Sponsor Capacity Left": _first_int(sponsor_capacity_town, "capacity_left"),
        "Latest 3-2-1 Coverage": _first_str(latest_321, "coverage_bucket", "No 3-2-1 data"),
        "Latest 3-2-1 Interaction Rate": _first_str(latest_321, "avg_interaction_rate", ""),
        "Attendance files loaded": "all history",
    }])

    regulars = pd.DataFrame()
    lapsed = pd.DataFrame()

    if not attendance.empty and "email" in attendance.columns and "event_month" in attendance.columns:
        att = attendance.dropna(subset=["email"]).copy()

        grp = (att.groupby("email")["event_month"].nunique().reset_index(name="months_attended_ever"))
        regulars = grp[grp["months_attended_ever"] >= 2].sort_values("months_attended_ever", ascending=False)

        try:
            months = att["event_month"].dropna().unique().tolist()
            periods = [p for p in (_parse_month_yyyy_mm(m) for m in months) if p is not None]
            latest = max(periods) if periods else None

            if latest is not None:
                last_by = (att.groupby("email")["event_month"].max().reset_index(name="last_seen_month"))
                last_by["last_seen_period"] = last_by["last_seen_month"].apply(lambda x: _parse_month_yyyy_mm(str(x)))
                merged = grp.merge(last_by, on="email", how="left")
                merged["is_lapsed"] = (
                    (merged["months_attended_ever"] >= 2)
                    & (merged["last_seen_period"].apply(lambda p: p is not None and p < latest))
                )
                lapsed = merged[merged["is_lapsed"]].sort_values("last_seen_month")
        except Exception:
            lapsed = pd.DataFrame()

    sponsors = pd.DataFrame()
    sponsor_ref = town_base / "data_ref" / f"sponsors_{town_code}.csv"
    if sponsor_ref.exists():
        try:
            sponsors = pd.read_csv(sponsor_ref)
        except Exception:
            sponsors = pd.DataFrame()

    coaching_line = "No 3-2-1 data available yet."
    if not latest_321.empty:
        coverage = _first_str(latest_321, "coverage_bucket", "")
        rate = _first_str(latest_321, "avg_interaction_rate", "")
        if coverage == "Strong":
            coaching_line = f"Strong 3-2-1 participation in the latest tracked month. Keep reinforcing what is working in-room. Interaction rate: {rate}"
        elif coverage == "Moderate":
            coaching_line = f"Moderate 3-2-1 participation in the latest tracked month. Prompt the room more deliberately to complete the tracker. Interaction rate: {rate}"
        elif coverage == "Low":
            coaching_line = f"Low 3-2-1 participation in the latest tracked month. Encourage stronger completion and more visible networking prompts. Interaction rate: {rate}"
        else:
            coaching_line = "3-2-1 data exists but room-count coverage needs checking."

    three_two_one_summary = pd.DataFrame([{"Coaching note": coaching_line}])

    out_dir = REGION_CURATED / "host_packs"
    out_dir.mkdir(parents=True, exist_ok=True)

    out_path = out_dir / f"HostPack_{town_label}.xlsx"

    with pd.ExcelWriter(out_path, engine="openpyxl", datetime_format="DD/MM/YYYY") as writer:
        snapshot.to_excel(writer, sheet_name="Snapshot", index=False)
        regulars.to_excel(writer, sheet_name="Regulars_and_Fans", index=False)
        lapsed.to_excel(writer, sheet_name="Lapsed_and_Reactivation", index=False)
        plus_members_town.to_excel(writer, sheet_name="Buzz Plus Members", index=False)
        plus_prospects_town.to_excel(writer, sheet_name="Buzz Plus Prospects", index=False)
        sponsors.to_excel(writer, sheet_name="Sponsors", index=False)
        three_two_one_summary.to_excel(writer, sheet_name="3-2-1 Insights", index=False, startrow=0)
        latest_321.to_excel(writer, sheet_name="3-2-1 Insights", index=False, startrow=3)
        trend_321.to_excel(writer, sheet_name="3-2-1 Insights", index=False, startrow=8)
        notes_321.to_excel(writer, sheet_name="3-2-1 Insights", index=False, startrow=14)
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
