import argparse
from pathlib import Path
import pandas as pd

# ==========================================================
# REGIONAL DASHBOARD (UPDATED)
# ==========================================================
# Changes represented in this version:
# - Removes payment-type dependency; all attendance metrics are attendance-based
# - Uses unique attendees (person_key/email) per event month
# - Reads Buzz Plus prospect counts using explicit Buzz Plus column names
# ==========================================================

BASE = Path(__file__).resolve().parent
REGION_CURATED = BASE / "Buzz_Region_Curated"
BUZZPLUS_FILE = REGION_CURATED / "buzzplus_intelligence.xlsx"
SPONSOR_FILE = REGION_CURATED / "sponsor_intelligence.xlsx"

TOWNS = [
    ("MarketHarborough", "Market Harborough"),
    ("Leicester", "Leicester"),
    ("Lutterworth", "Lutterworth"),
    ("Hinckley", "Hinckley"),
    ("Loughborough", "Loughborough"),
]


def parse_event_month_str(m: str):
    if not isinstance(m, str):
        m = str(m)
    m = m.strip()
    if not m or len(m) < 7 or "-" not in m:
        return None, None, None
    try:
        year = int(m[:4])
        month = int(m[5:7])
        idx = year * 12 + month
        return year, month, idx
    except Exception:
        return None, None, None


def _email_key(email: str) -> str:
    return (email or "").strip().lower()


def load_region_attendance() -> pd.DataFrame:
    frames = []
    for town_code, town_label in TOWNS:
        monthly_dir = BASE / f"Buzz_Event_Dashboard_{town_code}" / "data_curated" / "monthly"
        if not monthly_dir.exists():
            continue
        for fpath in monthly_dir.glob("*_attendance.xlsx"):
            try:
                df = pd.read_excel(fpath)
            except Exception:
                continue
            if df is None or df.empty:
                continue
            df["town_code"] = town_code
            df["town_name"] = town_label
            for col in ["event_month", "email", "person_key"]:
                if col not in df.columns:
                    df[col] = ""
            df["email"] = df["email"].fillna("").astype(str).str.strip()
            df["person_key"] = df["person_key"].fillna("").astype(str).str.strip()
            missing_pk = df["person_key"] == ""
            if missing_pk.any():
                df.loc[missing_pk, "person_key"] = df.loc[missing_pk, "email"].map(_email_key)
            frames.append(df)

    if not frames:
        return pd.DataFrame(columns=["town_code", "town_name", "event_month", "email", "person_key"])

    all_att = pd.concat(frames, ignore_index=True)

    years, months, idxs = [], [], []
    for m in all_att["event_month"].astype(str):
        y, mm, idx = parse_event_month_str(m)
        years.append(y)
        months.append(mm)
        idxs.append(idx)
    all_att["event_year"] = years
    all_att["event_month_num"] = months
    all_att["month_index"] = idxs
    all_att = all_att[all_att["month_index"].notna()].copy()
    return all_att


def load_buzzplus_summary() -> pd.DataFrame:
    if not BUZZPLUS_FILE.exists():
        return pd.DataFrame(columns=[
            "town_code", "town_name",
            "active_plus_members",
            "buzzplus_strong_prospects",
            "buzzplus_possible_prospects",
            "unique_recent_nonplus_visitors",
        ])
    try:
        return pd.read_excel(BUZZPLUS_FILE, sheet_name="Town_Plus_Summary")
    except Exception:
        return pd.DataFrame()


def load_sponsor_capacity() -> pd.DataFrame:
    if not SPONSOR_FILE.exists():
        return pd.DataFrame(columns=["town_code", "town_name", "active_sponsors", "capacity_left"])
    try:
        return pd.read_excel(SPONSOR_FILE, sheet_name="Sponsor_Capacity")
    except Exception:
        return pd.DataFrame(columns=["town_code", "town_name", "active_sponsors", "capacity_left"])


def compute_attendance_metrics(region_att: pd.DataFrame):
    """Compute per-town attendance metrics for the last ~12 months relative to latest data month.

    Metrics:
      - num_events_12m
      - total_unique_attendees_12m (sum of unique attendees per event month)
      - avg_unique_per_event_12m
    """
    if region_att.empty:
        town_metrics = pd.DataFrame(columns=[
            "town_code", "town_name",
            "num_events_12m",
            "total_unique_attendees_12m",
            "avg_unique_per_event_12m",
        ])
        trend = pd.DataFrame(columns=[
            "town_code", "town_name", "event_month", "unique_attendees"
        ])
        return town_metrics, trend

    max_idx = region_att["month_index"].max()
    cutoff = max_idx - 11 if pd.notna(max_idx) else None
    recent = region_att if cutoff is None else region_att[region_att["month_index"] >= cutoff].copy()

    trend = (
        recent.groupby(["town_code", "town_name", "event_month"])["person_key"]
        .nunique()
        .reset_index()
        .rename(columns={"person_key": "unique_attendees"})
    )

    town_group = trend.groupby(["town_code", "town_name"])
    town_metrics = town_group.agg(
        num_events_12m=("event_month", "nunique"),
        total_unique_attendees_12m=("unique_attendees", "sum"),
        avg_unique_per_event_12m=("unique_attendees", "mean"),
    ).reset_index()

    town_metrics["avg_unique_per_event_12m"] = town_metrics["avg_unique_per_event_12m"].round(2)
    town_metrics["num_events_12m"] = town_metrics["num_events_12m"].astype(int)
    town_metrics["total_unique_attendees_12m"] = town_metrics["total_unique_attendees_12m"].astype(int)

    return town_metrics, trend


def build_town_overview(town_metrics: pd.DataFrame,
                        plus_summary: pd.DataFrame,
                        sponsor_capacity: pd.DataFrame) -> pd.DataFrame:
    # Merge
    df = town_metrics.merge(plus_summary, on=["town_code", "town_name"], how="left")
    df = df.merge(sponsor_capacity, on=["town_code", "town_name"], how="left")

    # Fill
    for c in ["active_plus_members", "buzzplus_strong_prospects", "buzzplus_possible_prospects", "unique_recent_nonplus_visitors",
              "active_sponsors", "capacity_left"]:
        if c in df.columns:
            df[c] = df[c].fillna(0).astype(int)

    # Focus flags (editable thresholds)
    df["attendance_focus"] = df["avg_unique_per_event_12m"].apply(
        lambda v: "Strong" if v >= 25 else ("OK" if v >= 20 else "Low")
    )
    df["sponsor_focus"] = df["capacity_left"].apply(lambda v: "Can recruit" if v > 0 else "At capacity")
    df["plus_focus"] = df.apply(
        lambda r: "Plus opportunity" if (r["buzzplus_strong_prospects"] > 0 and r["active_plus_members"] < 2) else "",
        axis=1
    )

    # Priority score (simple heuristic)
    df["priority_score"] = 0
    df.loc[df["attendance_focus"] == "Low", "priority_score"] += 2
    df.loc[df["sponsor_focus"] == "Can recruit", "priority_score"] += 1
    df.loc[df["plus_focus"] == "Plus opportunity", "priority_score"] += 1

    return df


def main():
    ap = argparse.ArgumentParser(description="Build Business Buzz Regional Dashboard workbook (UPDATED).")
    ap.add_argument("--out", default=str(REGION_CURATED / "region_dashboard.xlsx"))
    args = ap.parse_args()

    region_att = load_region_attendance()
    plus_summary = load_buzzplus_summary()
    sponsor_capacity = load_sponsor_capacity()

    town_metrics, trend = compute_attendance_metrics(region_att)
    overview = build_town_overview(town_metrics, plus_summary, sponsor_capacity)

    out = Path(args.out)
    out.parent.mkdir(parents=True, exist_ok=True)

    with pd.ExcelWriter(out, engine="xlsxwriter") as writer:
        overview.to_excel(writer, sheet_name="Town_Overview", index=False)
        trend.to_excel(writer, sheet_name="Attendance_Trend_12m", index=False)

    print(f"[OK] Wrote {out}")


if __name__ == "__main__":
    main()
