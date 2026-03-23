import re
from pathlib import Path
from datetime import datetime
import pandas as pd

# ==========================================================
# Buzz Plus Intelligence Engine (UPDATED)
#
# Changes implemented
# - Region-wide exclusion of active Hosts and Ambassadors from Buzz Plus prospecting
#   (based on email in roles_<Town>.csv under each town's data_ref folder).
# - Prospect tier rules updated to match Emma's policy:
#     * Strong = visits_12m >= 6
#     * Possible = visits_12m in [4, 5]
#   No payment logic used.
# - Town summary columns standardised for downstream packs:
#     buzzplus_strong_prospects, buzzplus_possible_prospects
#
# Outputs (written to Region Curated):
# - Buzz_Region_Curated/buzzplus_intelligence.xlsx
#   Sheets:
#     Town_Plus_Summary
#     Active_Plus_Members
#     Plus_Prospects
#     Excluded_Team_Roles
# ==========================================================

BASE = Path(__file__).resolve().parent

TOWNS = [
    ("MarketHarborough", "Market Harborough"),
    ("Leicester", "Leicester"),
    ("Lutterworth", "Lutterworth"),
    ("Hinckley", "Hinckley"),
    ("Loughborough", "Loughborough"),
]

REGION_CURATED = BASE / "Buzz_Region_Curated"
REGION_REF = BASE / "Buzz_Region_Ref"

OUTPUT_XLSX = REGION_CURATED / "buzzplus_intelligence.xlsx"

# Buzz Plus membership source (existing regional reference file)
PLUS_MEMBERS_FILE = REGION_REF / "buzzplus_members.xlsx"

# Attendance inputs: town curated monthly attendance files
TOWN_FOLDERS = {
    code: BASE / f"Buzz_Event_Dashboard_{code}"
    for code, _ in TOWNS
}

TEAM_ROLES_FILENAME_TEMPLATE = "roles_{town_code}.csv"


# ---------------------------
# Helpers
# ---------------------------

def _norm_text(x: str) -> str:
    if pd.isna(x):
        return ""
    x = str(x).strip()
    x = re.sub(r"\s+", " ", x)
    return x


def _norm_email(x: str) -> str:
    x = _norm_text(x).lower()
    return x


def _norm_role(x: str) -> str:
    x = _norm_text(x).lower()
    if x == "host":
        return "Host"
    if x == "ambassador":
        return "Ambassador"
    # allow synonyms/variants if they appear
    if "host" in x:
        return "Host"
    if "ambass" in x:
        return "Ambassador"
    return _norm_text(x)


def _read_excel_any(path: Path) -> pd.DataFrame:
    if not path.exists():
        return pd.DataFrame()
    try:
        return pd.read_excel(path)
    except Exception:
        return pd.DataFrame()


def _safe_read_excel_sheet(path: Path, sheet_name: str) -> pd.DataFrame:
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


def _parse_yyyy_mm(x: str) -> pd.Period | None:
    x = _norm_text(x)
    if not x:
        return None
    try:
        return pd.Period(x, freq="M")
    except Exception:
        return None


# ---------------------------
# Load region attendance (last 12 months window calculations need month_index)
# ---------------------------

def load_region_attendance() -> pd.DataFrame:
    frames: list[pd.DataFrame] = []

    for town_code, town_name in TOWNS:
        town_base = TOWN_FOLDERS[town_code]
        monthly_dir = town_base / "data_curated" / "monthly"
        if not monthly_dir.exists():
            continue

        for fp in sorted(monthly_dir.glob("*_attendance.xlsx")):
            try:
                df = pd.read_excel(fp)
            except Exception:
                continue

            df = _normalise_cols(df)

            # required fields
            # Expect at least: email, event_month
            if "email" not in df.columns or "event_month" not in df.columns:
                continue

            df["email"] = df["email"].apply(_norm_email)
            df["event_month"] = df["event_month"].astype(str).str.strip()

            # Attach town columns if not present
            if "town_code" not in df.columns:
                df["town_code"] = town_code
            if "town_name" not in df.columns:
                df["town_name"] = town_name

            # Month index for recency calcs
            df["period"] = df["event_month"].apply(_parse_yyyy_mm)
            df = df[df["period"].notna()].copy()
            if df.empty:
                continue

            # month_index as integer offset from 1970-01 for comparisons
            df["month_index"] = df["period"].apply(lambda p: int(p.year) * 12 + int(p.month))

            # optional display fields
            if "name" not in df.columns and "person_name" in df.columns:
                df["name"] = df["person_name"]

            frames.append(df)

    if not frames:
        return pd.DataFrame()

    region_att = pd.concat(frames, ignore_index=True)
    return region_att


# ---------------------------
# Load Buzz Plus members (reference)
# ---------------------------

def load_plus_members() -> pd.DataFrame:
    plus = _read_excel_any(PLUS_MEMBERS_FILE)
    if plus.empty:
        return pd.DataFrame()

    plus = _normalise_cols(plus)

    # Expected fields (best effort): member_name, member_company, town_code, start_date, end_date, status
    # Map common variants
    col_map = {
        "name": "member_name",
        "member": "member_name",
        "company": "member_company",
        "organisation": "member_company",
        "org": "member_company",
    }
    for src, dst in col_map.items():
        if src in plus.columns and dst not in plus.columns:
            plus[dst] = plus[src]

    if "town_code" not in plus.columns and "town" in plus.columns:
        # attempt to transform town names to codes (best effort)
        town_name_to_code = {name.lower(): code for code, name in TOWNS}
        plus["town_code"] = plus["town"].astype(str).str.strip().str.lower().map(town_name_to_code).fillna("")

    # Normalise key fields
    if "member_name" in plus.columns:
        plus["member_name"] = plus["member_name"].astype(str).str.strip()
    else:
        plus["member_name"] = ""

    if "member_company" in plus.columns:
        plus["member_company"] = plus["member_company"].astype(str).str.strip()
    else:
        plus["member_company"] = ""

    if "town_code" in plus.columns:
        plus["town_code"] = plus["town_code"].astype(str).str.strip()

    # Dates
    for dc in ["start_date", "end_date"]:
        if dc in plus.columns:
            plus[dc] = pd.to_datetime(plus[dc], errors="coerce")
        else:
            plus[dc] = pd.NaT

    # Status (optional)
    status_col = "status" if "status" in plus.columns else None
    plus["status_norm"] = plus[status_col].fillna("").astype(str).str.strip().str.lower() if status_col else ""

    # Active now = within start/end and not explicitly cancelled
    today = pd.Timestamp.today().normalize()
    mask_period = plus["start_date"].notna() & (plus["start_date"] <= today)
    mask_end = plus["end_date"].isna() | (plus["end_date"] >= today)

    # Treat "cancel" / "inactive" statuses as inactive
    bad_status = plus["status_norm"].str.contains("cancel|inactive|ended|lapsed", regex=True, na=False)
    plus["is_active_now"] = mask_period & mask_end & (~bad_status)

    return plus


# ---------------------------
# Load region-wide team roles for exclusion
# ---------------------------

def load_team_role_exclusions() -> tuple[set[str], pd.DataFrame]:
    """
    Returns:
      - excluded_emails: set of emails to exclude from Buzz Plus prospecting
      - excluded_df: detailed rows for auditing
    """
    rows: list[dict] = []

    for town_code, town_name in TOWNS:
        roles_fp = (TOWN_FOLDERS[town_code] / "data_ref" / TEAM_ROLES_FILENAME_TEMPLATE.format(town_code=town_code))
        if not roles_fp.exists():
            continue

        try:
            df = pd.read_csv(roles_fp)
        except Exception:
            continue

        df = _normalise_cols(df)

        # required
        if "email" not in df.columns or "role" not in df.columns:
            continue

        df["email"] = df["email"].apply(_norm_email)
        df["role"] = df["role"].apply(_norm_role)

        # Optional end_date
        if "end_date" in df.columns:
            df["end_date"] = pd.to_datetime(df["end_date"], errors="coerce")
        else:
            df["end_date"] = pd.NaT

        # active = no end_date OR end_date in future
        today = pd.Timestamp.today().normalize()
        is_active = df["end_date"].isna() | (df["end_date"] >= today)

        df = df[is_active].copy()
        if df.empty:
            continue

        df["town_code"] = town_code
        df["town_name"] = town_name

        # Exclude only Host & Ambassador
        df = df[df["role"].isin(["Host", "Ambassador"])].copy()
        if df.empty:
            continue

        for _, r in df.iterrows():
            rows.append({
                "email": r.get("email", ""),
                "role": r.get("role", ""),
                "town_code": town_code,
                "town_name": town_name,
            })

    excluded_df = pd.DataFrame(rows)
    excluded_emails = set(excluded_df["email"].dropna().astype(str).str.strip().str.lower().tolist()) if not excluded_df.empty else set()
    return excluded_emails, excluded_df


# ---------------------------
# Build intelligence (summary + prospects)
# ---------------------------

def build_buzzplus_intelligence(region_att: pd.DataFrame, plus_members: pd.DataFrame, excluded_emails: set[str]) -> tuple[pd.DataFrame, pd.DataFrame, pd.DataFrame]:
    """
    Returns:
      - town_summary_df
      - active_plus_members_df
      - prospects_df
    """
    if region_att.empty:
        cols_summary = [
            "town_code", "town_name",
            "active_plus_members",
            "buzzplus_strong_prospects",
            "buzzplus_possible_prospects",
            "unique_recent_nonplus_visitors",
        ]
        return (
            pd.DataFrame(columns=cols_summary),
            pd.DataFrame(columns=[
                "town_code", "town_name",
                "member_name", "member_company",
                "start_date", "end_date", "status_norm"
            ]),
            pd.DataFrame(columns=[
                "town_code", "town_name",
                "email", "person_name", "company",
                "visits_12m", "last_event_month",
                "prospect_tier"
            ]),
        )

    # Filter to last 12 months based on month_index
    max_idx = region_att["month_index"].max()
    if pd.isna(max_idx):
        recent_att = region_att.copy()
    else:
        cutoff = int(max_idx) - 11  # last 12 calendar months
        recent_att = region_att[region_att["month_index"] >= cutoff].copy()

    # Exclude team roles from the prospect pool (region-wide)
    if excluded_emails:
        recent_att = recent_att[~recent_att["email"].isin(excluded_emails)].copy()

    # Prepare Plus membership info
    if plus_members is None or plus_members.empty:
        active_plus_df = pd.DataFrame(columns=[
            "town_code", "town_name",
            "member_name", "member_company",
            "start_date", "end_date", "status_norm"
        ])
        active_plus_emails = set()  # not available from this reference
    else:
        active_plus = plus_members[plus_members["is_active_now"]].copy()
        town_name_map = {code: name for code, name in TOWNS}
        active_plus["town_name"] = active_plus["town_code"].map(town_name_map).fillna("")
        active_plus_df = active_plus[[
            "town_code", "town_name",
            "member_name", "member_company",
            "start_date", "end_date", "status_norm"
        ]].copy()

        # Note: membership file does not reliably include emails; we do not use it for exclusion.
        active_plus_emails = set()

    # Prospect logic:
    # - Build per town, per email, based on visits in last 12 months
    prospect_rows: list[dict] = []

    # Ensure optional columns exist
    if "name" not in recent_att.columns:
        recent_att["name"] = ""
    if "company" not in recent_att.columns:
        recent_att["company"] = ""

    for town_code, town_name in TOWNS:
        town_att = recent_att[recent_att["town_code"] == town_code].copy()
        if town_att.empty:
            continue

        grouped = town_att.groupby("email", dropna=True)
        for email, g in grouped:
            if not email:
                continue

            # Determine visits & last attendance
            visits_12m = int(g["event_month"].nunique())

            # Tier rules (Emma's policy)
            if visits_12m >= 6:
                prospect_tier = "Buzz Plus Strong"
            elif visits_12m in (4, 5):
                prospect_tier = "Buzz Plus Possible"
            else:
                continue

            latest = g.sort_values("month_index").iloc[-1]
            last_month_str = str(latest.get("event_month", "")).strip()

            person_name = str(latest.get("name", "")).strip()
            company = str(latest.get("company", "")).strip()

            prospect_rows.append({
                "town_code": town_code,
                "town_name": town_name,
                "email": email,
                "person_name": person_name,
                "company": company,
                "visits_12m": visits_12m,
                "last_event_month": last_month_str,
                "prospect_tier": prospect_tier,
            })

    prospects_df = pd.DataFrame(prospect_rows)
    if not prospects_df.empty:
        prospects_df = prospects_df.sort_values(["town_name", "prospect_tier", "visits_12m"], ascending=[True, True, False])

    # Town summaries
    summary_rows: list[dict] = []
    for town_code, town_name in TOWNS:
        town_prospects = prospects_df[prospects_df["town_code"] == town_code] if not prospects_df.empty else pd.DataFrame()
        strong_count = int((town_prospects["prospect_tier"] == "Buzz Plus Strong").sum()) if not town_prospects.empty else 0
        possible_count = int((town_prospects["prospect_tier"] == "Buzz Plus Possible").sum()) if not town_prospects.empty else 0

        active_plus_count = int((active_plus_df["town_code"] == town_code).sum()) if not active_plus_df.empty else 0

        # "unique recent non-plus visitors": total unique emails in last 12 months for that town (after team exclusion)
        unique_recent = int(recent_att[recent_att["town_code"] == town_code]["email"].nunique()) if not recent_att.empty else 0

        summary_rows.append({
            "town_code": town_code,
            "town_name": town_name,
            "active_plus_members": active_plus_count,
            "buzzplus_strong_prospects": strong_count,
            "buzzplus_possible_prospects": possible_count,
            "unique_recent_nonplus_visitors": unique_recent,
        })

    town_summary_df = pd.DataFrame(summary_rows)
    return town_summary_df, active_plus_df, prospects_df


def main() -> None:
    REGION_CURATED.mkdir(parents=True, exist_ok=True)

    print("[INFO] Loading regional attendance from town monthly files...")
    region_att = load_region_attendance()

    print("[INFO] Loading Buzz Plus members from Buzz_Region_Ref...")
    plus_members = load_plus_members()

    print("[INFO] Loading region-wide team roles for exclusion...")
    excluded_emails, excluded_df = load_team_role_exclusions()
    if excluded_emails:
        print(f"[INFO] Excluding {len(excluded_emails)} unique team emails from prospecting.")
    else:
        print("[WARN] No team role exclusions loaded (roles files missing or empty).")

    print("[INFO] Building Buzz Plus Intelligence (summary + prospects)...")
    town_summary_df, active_plus_df, prospects_df = build_buzzplus_intelligence(region_att, plus_members, excluded_emails)

    print(f"[INFO] Writing {OUTPUT_XLSX}...")
    with pd.ExcelWriter(OUTPUT_XLSX, engine="openpyxl") as writer:
        town_summary_df.to_excel(writer, sheet_name="Town_Plus_Summary", index=False)
        active_plus_df.to_excel(writer, sheet_name="Active_Plus_Members", index=False)
        prospects_df.to_excel(writer, sheet_name="Plus_Prospects", index=False)
        excluded_df.to_excel(writer, sheet_name="Excluded_Team_Roles", index=False)

    print("[OK] Buzz Plus Intelligence build completed.")


if __name__ == "__main__":
    main()
