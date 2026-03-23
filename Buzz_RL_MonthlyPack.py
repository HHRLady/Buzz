import argparse
from pathlib import Path
import pandas as pd

# ==========================================================
# REGIONAL LEAD MONTHLY PACK (UPDATED)
# ==========================================================
# Changes represented in this version:
# - Aligns to attendance-based (no payment) metrics used in updated regional dashboard
# - Uses explicit Buzz Plus prospect column naming
# - Includes consistent RL_Priorities and RL_Actions derived from focus flags
# ==========================================================

BASE = Path(__file__).resolve().parent
REGION_CURATED = BASE / "Buzz_Region_Curated"

REGION_DASHBOARD = REGION_CURATED / "region_dashboard.xlsx"
BUZZPLUS_INTEL = REGION_CURATED / "buzzplus_intelligence.xlsx"
SPONSOR_INTEL = REGION_CURATED / "sponsor_intelligence.xlsx"
EVENT_EXCELLENCE = REGION_CURATED / "region_event_excellence.xlsx"

OUT_FILE = REGION_CURATED / "RL_Monthly_Pack.xlsx"


def _read_excel_safe(path: Path, sheet: str) -> pd.DataFrame:
    if not path.exists():
        return pd.DataFrame()
    try:
        return pd.read_excel(path, sheet_name=sheet)
    except Exception:
        return pd.DataFrame()


def build_priorities(town_overview: pd.DataFrame) -> pd.DataFrame:
    if town_overview is None or town_overview.empty:
        return pd.DataFrame(columns=["town_code", "town_name", "priority_score", "priority_bucket"])

    df = town_overview.copy()
    df["priority_bucket"] = df["priority_score"].apply(
        lambda s: "High" if s >= 3 else ("Medium" if s == 2 else ("Low" if s == 1 else "Monitor"))
    )
    df = df.sort_values(by=["priority_score", "town_name"], ascending=[False, True])
    return df[["town_code", "town_name", "priority_score", "priority_bucket",
               "attendance_focus", "sponsor_focus", "plus_focus"]]


def build_actions(town_overview: pd.DataFrame) -> pd.DataFrame:
    if town_overview is None or town_overview.empty:
        return pd.DataFrame(columns=["town_code", "town_name", "action_area", "recommended_action", "trigger"])

    rows = []
    for _, r in town_overview.iterrows():
        town_code = r.get("town_code", "")
        town_name = r.get("town_name", "")
        triggers = []

        if r.get("attendance_focus") == "Low":
            triggers.append("Attendance: Low")
            rows.append({
                "town_code": town_code,
                "town_name": town_name,
                "action_area": "Attendance",
                "recommended_action": "Review last 3 months attendance trend; agree a short-term host support plan and visitor reactivation activity.",
                "trigger": f"avg_unique_per_event_12m={r.get('avg_unique_per_event_12m', '')}"
            })

        if r.get("sponsor_focus") == "Can recruit":
            triggers.append("Sponsorship: Capacity")
            rows.append({
                "town_code": town_code,
                "town_name": town_name,
                "action_area": "Sponsorship",
                "recommended_action": "Confirm target sectors and begin sponsor outreach aligned to local lockouts and capacity.",
                "trigger": f"capacity_left={r.get('capacity_left', '')}"
            })

        if r.get("plus_focus") == "Plus opportunity":
            triggers.append("Buzz Plus: Opportunity")
            rows.append({
                "town_code": town_code,
                "town_name": town_name,
                "action_area": "Buzz Plus",
                "recommended_action": "Prioritise Buzz Plus conversations with Strong Prospects; track outcomes and conversion intent.",
                "trigger": f"strong_prospects={r.get('buzzplus_strong_prospects', '')}, active_plus={r.get('active_plus_members', '')}"
            })

        if not triggers:
            rows.append({
                "town_code": town_code,
                "town_name": town_name,
                "action_area": "Maintain",
                "recommended_action": "No immediate action triggered by current thresholds. Monitor and maintain cadence.",
                "trigger": ""
            })

    return pd.DataFrame(rows)


def main():
    ap = argparse.ArgumentParser(description="Build Business Buzz RL Monthly Pack (UPDATED).")
    ap.add_argument("--out", default=str(OUT_FILE))
    args = ap.parse_args()

    town_overview = _read_excel_safe(REGION_DASHBOARD, "Town_Overview")
    trend = _read_excel_safe(REGION_DASHBOARD, "Attendance_Trend_12m")

    plus_summary = _read_excel_safe(BUZZPLUS_INTEL, "Town_Plus_Summary")
    plus_members = _read_excel_safe(BUZZPLUS_INTEL, "Active_Plus_Members")
    plus_prospects = _read_excel_safe(BUZZPLUS_INTEL, "Plus_Prospects")

    sponsor_capacity = _read_excel_safe(SPONSOR_INTEL, "Sponsor_Capacity")
    sponsor_opps = _read_excel_safe(SPONSOR_INTEL, "Sponsor_Opportunities")

    event_excellence = _read_excel_safe(EVENT_EXCELLENCE, "Event_Excellence") if EVENT_EXCELLENCE.exists() else pd.DataFrame()

    priorities = build_priorities(town_overview)
    actions = build_actions(town_overview)

    out = Path(args.out)
    out.parent.mkdir(parents=True, exist_ok=True)

    with pd.ExcelWriter(out, engine="xlsxwriter") as writer:
        town_overview.to_excel(writer, sheet_name="Town_Overview", index=False)
        trend.to_excel(writer, sheet_name="Attendance_Trend_12m", index=False)
        plus_summary.to_excel(writer, sheet_name="BuzzPlus_Summary", index=False)
        plus_members.to_excel(writer, sheet_name="BuzzPlus_Members", index=False)
        plus_prospects.to_excel(writer, sheet_name="BuzzPlus_Prospects", index=False)
        sponsor_capacity.to_excel(writer, sheet_name="Sponsor_Capacity", index=False)
        sponsor_opps.to_excel(writer, sheet_name="Sponsor_Opportunities", index=False)
        if not event_excellence.empty:
            event_excellence.to_excel(writer, sheet_name="Event_Excellence", index=False)
        priorities.to_excel(writer, sheet_name="RL_Priorities", index=False)
        actions.to_excel(writer, sheet_name="RL_Actions", index=False)

    print(f"[OK] Wrote {out}")


if __name__ == "__main__":
    main()
