import argparse
from pathlib import Path
import pandas as pd

BASE = Path(__file__).resolve().parent
REGION_REF = BASE / "Buzz_Region_Ref"
REGION_CURATED = BASE / "Buzz_Region_Curated"

PREFERRED_INPUTS = [
    REGION_REF / "Buzz_321_Tracker.xlsx",
    REGION_REF / "LRBuzz 321 Tracker.xlsx",
]

OUT_XLSX = REGION_CURATED / "buzz_321_intelligence.xlsx"
OUT_CSV = REGION_CURATED / "buzz_321_event_level.csv"

TOWN_MAP = {
    "marketharborough": ("MarketHarborough", "Market Harborough"),
    "market harborough": ("MarketHarborough", "Market Harborough"),
    "leicester": ("Leicester", "Leicester"),
    "lutterworth": ("Lutterworth", "Lutterworth"),
    "hinckley": ("Hinckley", "Hinckley"),
    "loughborough": ("Loughborough", "Loughborough"),
}


def _find_input_file() -> Path | None:
    for path in PREFERRED_INPUTS:
        if path.exists():
            return path
    return None


def _normalise_cols(df: pd.DataFrame) -> pd.DataFrame:
    df = df.copy()
    df.columns = [str(c).strip() for c in df.columns]
    return df


def _canonicalise_town(value: str) -> tuple[str, str]:
    raw = str(value or "").strip()
    key = raw.lower()
    if key in TOWN_MAP:
        return TOWN_MAP[key]
    compact = key.replace(" ", "")
    if compact in TOWN_MAP:
        return TOWN_MAP[compact]
    return ("", raw)


def _to_numeric(df: pd.DataFrame, cols: list[str]) -> pd.DataFrame:
    for col in cols:
        if col not in df.columns:
            df[col] = 0
        df[col] = pd.to_numeric(df[col], errors="coerce").fillna(0)
    return df


def _coverage_bucket(rate: float) -> str:
    if pd.isna(rate):
        return "No room count"
    if rate >= 0.40:
        return "Strong"
    if rate >= 0.20:
        return "Moderate"
    return "Low"


def _empty_event_cols() -> list[str]:
    return [
        "town_code", "town_name", "venue_raw", "event_date", "event_month_key", "event_month",
        "in_room", "tracker_users", "met_3_new", "one_2_ones", "brought_some1",
        "hosting_team", "buzz_in", "buzz_out", "net_room_flow",
        "interaction_rate", "met_new_rate", "one_to_one_rate", "brought_someone_rate",
        "notes", "submitted_by", "status", "coverage_flag", "quality_flag"
    ]


def _empty_summary_cols() -> list[str]:
    return [
        "town_code", "town_name", "event_month_key", "event_month", "latest_event_date",
        "events_with_321_data", "total_in_room", "total_tracker_users",
        "total_met_3_new", "total_one_2_ones", "total_brought_some1",
        "avg_interaction_rate", "avg_met_new_rate", "avg_one_to_one_rate", "avg_brought_someone_rate",
        "coverage_bucket"
    ]


def _make_empty_outputs(reason: str) -> None:
    REGION_CURATED.mkdir(parents=True, exist_ok=True)
    event_cols = _empty_event_cols()
    summary_cols = _empty_summary_cols()
    quality = pd.DataFrame([{"message": reason}])

    with pd.ExcelWriter(OUT_XLSX, engine="xlsxwriter", datetime_format="dd/mm/yyyy") as writer:
        pd.DataFrame(columns=event_cols).to_excel(writer, sheet_name="Event_Level", index=False)
        pd.DataFrame(columns=summary_cols).to_excel(writer, sheet_name="Town_Month_Summary", index=False)
        pd.DataFrame(columns=[c for c in summary_cols if c not in {"event_month_key", "event_month"}]).to_excel(
            writer, sheet_name="Town_Summary", index=False
        )
        quality.to_excel(writer, sheet_name="Data_Quality", index=False)

    pd.DataFrame(columns=event_cols).to_csv(OUT_CSV, index=False)
    print(f"[INFO] {reason}")
    print(f"[OK] Wrote {OUT_XLSX}")
    print(f"[OK] Wrote {OUT_CSV}")


def build_321_intelligence(input_file: Path) -> None:
    REGION_CURATED.mkdir(parents=True, exist_ok=True)

    try:
        xls = pd.ExcelFile(input_file)
        sheet = "321_Tracker" if "321_Tracker" in xls.sheet_names else xls.sheet_names[0]
        raw = pd.read_excel(input_file, sheet_name=sheet)
    except Exception as exc:
        _make_empty_outputs(f"Could not read 3-2-1 tracker: {exc}")
        return

    raw = _normalise_cols(raw)
    source_rows = len(raw)
    if raw.empty:
        _make_empty_outputs("3-2-1 tracker is empty.")
        return

    raw["status"] = raw.get("status", "").astype(str).str.strip()
    submitted = raw[raw["status"].str.lower() == "submitted"].copy()
    submitted_rows = len(submitted)

    submitted["venue_raw"] = submitted.get("venue", "").astype(str).str.strip()
    submitted["date"] = pd.to_datetime(submitted.get("date"), errors="coerce", dayfirst=True)
    submitted["timestamp"] = pd.to_datetime(submitted.get("timestamp"), errors="coerce", dayfirst=True)

    required_valid = submitted[(submitted["venue_raw"] != "") & submitted["date"].notna()].copy()

    numeric_cols = [
        "met_3_new", "one_2_ones", "brought_some1", "in_room",
        "hosting_team", "buzz_in", "buzz_out", "321tracker_users"
    ]
    working = _to_numeric(required_valid.copy(), numeric_cols)

    town_pairs = working["venue_raw"].apply(_canonicalise_town)
    working["town_code"] = town_pairs.apply(lambda x: x[0])
    working["town_name"] = town_pairs.apply(lambda x: x[1])
    working = working[working["town_code"] != ""].copy()
    mapped_rows = len(working)

    working = working.sort_values(by=["town_code", "date", "timestamp"], ascending=[True, True, True])
    before_dupes = len(working)
    working = working.drop_duplicates(subset=["town_code", "date"], keep="last").copy()
    deduped_count = before_dupes - len(working)

    working["event_date"] = working["date"].dt.normalize()
    working["event_month_key"] = working["date"].dt.strftime("%Y-%m")
    working["event_month"] = working["date"].dt.strftime("%B %Y")
    working["tracker_users"] = working["321tracker_users"]
    working["net_room_flow"] = working["buzz_in"] - working["buzz_out"]

    def _safe_rate(num_col: str) -> pd.Series:
        return working.apply(
            lambda r: (float(r[num_col]) / float(r["in_room"])) if pd.notna(r["in_room"]) and float(r["in_room"]) > 0 else pd.NA,
            axis=1
        )

    working["interaction_rate"] = _safe_rate("tracker_users")
    working["met_new_rate"] = _safe_rate("met_3_new")
    working["one_to_one_rate"] = _safe_rate("one_2_ones")
    working["brought_someone_rate"] = _safe_rate("brought_some1")
    working["coverage_flag"] = working["interaction_rate"].apply(_coverage_bucket)

    def _quality_flag(row) -> str:
        flags = []
        if row["in_room"] <= 0:
            flags.append("Missing room count")
        if row["tracker_users"] > row["in_room"] and row["in_room"] > 0:
            flags.append("Tracker users exceed room count")
        if row["met_3_new"] > row["in_room"] and row["in_room"] > 0:
            flags.append("Met 3 new exceeds room count")
        if row["one_2_ones"] > row["in_room"] and row["in_room"] > 0:
            flags.append("1:1s exceed room count")
        if row["brought_some1"] > row["in_room"] and row["in_room"] > 0:
            flags.append("Brought someone exceeds room count")
        return "OK" if not flags else " | ".join(flags)

    working["quality_flag"] = working.apply(_quality_flag, axis=1)

    event_level = working[_empty_event_cols()].copy()

    town_month_summary = (
        event_level.groupby(["town_code", "town_name", "event_month_key", "event_month"], dropna=False)
        .agg(
            latest_event_date=("event_date", "max"),
            events_with_321_data=("event_date", "count"),
            total_in_room=("in_room", "sum"),
            total_tracker_users=("tracker_users", "sum"),
            total_met_3_new=("met_3_new", "sum"),
            total_one_2_ones=("one_2_ones", "sum"),
            total_brought_some1=("brought_some1", "sum"),
            avg_interaction_rate=("interaction_rate", "mean"),
            avg_met_new_rate=("met_new_rate", "mean"),
            avg_one_to_one_rate=("one_to_one_rate", "mean"),
            avg_brought_someone_rate=("brought_someone_rate", "mean"),
        )
        .reset_index()
    )
    town_month_summary["coverage_bucket"] = town_month_summary["avg_interaction_rate"].apply(_coverage_bucket)
    town_month_summary = town_month_summary[_empty_summary_cols()].copy()

    town_summary = (
        event_level.groupby(["town_code", "town_name"], dropna=False)
        .agg(
            latest_event_date=("event_date", "max"),
            events_with_321_data=("event_date", "count"),
            total_in_room=("in_room", "sum"),
            total_tracker_users=("tracker_users", "sum"),
            total_met_3_new=("met_3_new", "sum"),
            total_one_2_ones=("one_2_ones", "sum"),
            total_brought_some1=("brought_some1", "sum"),
            avg_interaction_rate=("interaction_rate", "mean"),
            avg_met_new_rate=("met_new_rate", "mean"),
            avg_one_to_one_rate=("one_to_one_rate", "mean"),
            avg_brought_someone_rate=("brought_someone_rate", "mean"),
        )
        .reset_index()
    )
    town_summary["coverage_bucket"] = town_summary["avg_interaction_rate"].apply(_coverage_bucket)

    quality = pd.DataFrame([
        {"metric": "source_rows", "value": source_rows},
        {"metric": "submitted_rows", "value": submitted_rows},
        {"metric": "rows_with_required_fields", "value": len(required_valid)},
        {"metric": "mapped_town_rows", "value": mapped_rows},
        {"metric": "valid_rows_used", "value": len(event_level)},
        {"metric": "dropped_missing_required", "value": submitted_rows - len(required_valid)},
        {"metric": "dropped_unknown_or_unmapped_town", "value": len(required_valid) - mapped_rows},
        {"metric": "deduplicated_rows_removed", "value": deduped_count},
        {"metric": "input_file", "value": str(input_file)},
    ])

    with pd.ExcelWriter(OUT_XLSX, engine="xlsxwriter", datetime_format="dd/mm/yyyy") as writer:
        event_level.to_excel(writer, sheet_name="Event_Level", index=False)
        town_month_summary.to_excel(writer, sheet_name="Town_Month_Summary", index=False)
        town_summary.to_excel(writer, sheet_name="Town_Summary", index=False)
        quality.to_excel(writer, sheet_name="Data_Quality", index=False)

        workbook = writer.book
        pct_fmt = workbook.add_format({"num_format": "0%"})
        date_fmt = workbook.add_format({"num_format": "dd/mm/yyyy"})

        for sheet_name, df in {
            "Event_Level": event_level,
            "Town_Month_Summary": town_month_summary,
            "Town_Summary": town_summary,
        }.items():
            ws = writer.sheets[sheet_name]
            for i, col in enumerate(df.columns):
                width = max(len(str(col)), 14)
                if col in {"notes", "quality_flag"}:
                    width = 42
                elif col in {"submitted_by", "venue_raw"}:
                    width = 24
                ws.set_column(i, i, width)
                if col in {"interaction_rate", "met_new_rate", "one_to_one_rate", "brought_someone_rate", "avg_interaction_rate", "avg_met_new_rate", "avg_one_to_one_rate", "avg_brought_someone_rate"}:
                    ws.set_column(i, i, 14, pct_fmt)
                if col in {"event_date", "latest_event_date"}:
                    ws.set_column(i, i, 14, date_fmt)

    event_level.to_csv(OUT_CSV, index=False)

    print(f"[INFO] Read 3-2-1 tracker from {input_file}")
    print(f"[INFO] Valid event rows used: {len(event_level)}")
    print(f"[OK] Wrote {OUT_XLSX}")
    print(f"[OK] Wrote {OUT_CSV}")


def main() -> None:
    parser = argparse.ArgumentParser(description="Build Business Buzz 3-2-1 Intelligence.")
    parser.add_argument("--infile", default="", help="Optional path to tracker file")
    args = parser.parse_args()

    input_file = Path(args.infile) if args.infile else _find_input_file()
    if input_file is None or not input_file.exists():
        _make_empty_outputs(
            "No 3-2-1 tracker file found in Buzz_Region_Ref. Expected Buzz_321_Tracker.xlsx or LRBuzz 321 Tracker.xlsx."
        )
        return

    build_321_intelligence(input_file)


if __name__ == "__main__":
    main()
