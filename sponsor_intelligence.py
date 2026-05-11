import pandas as pd
from pathlib import Path


# ==========================================================
# CONFIGURATION
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
REGION_CURATED.mkdir(exist_ok=True)


# ==========================================================
# INDUSTRY LOOKUP / CATEGORISATION
# ==========================================================

def categorise_company(name: str) -> str:
    """
    Industry classifier used for sponsor lock-out mapping.

    Categories (checked in priority order — first match wins):
      Accountancy | Legal | Insurance | Mortgage & Lending |
      Financial Services | HR & Recruitment | IT & Technology |
      Marketing, Design & PR | Property & Construction |
      Health & Wellbeing | Print, Signs & Promotional |
      Photography & Media | Business Coaching & Consulting |
      General Business

    To add or adjust a category, extend the relevant keyword list below.
    Order matters — more specific categories appear before broader ones
    (e.g. Insurance before Financial Services).
    """
    if not isinstance(name, str):
        return "Unknown"
    n = name.lower()

    # Accountancy
    if any(k in n for k in ("account", "bookkeep", "chartered", "tax advisor", "tax adviser", "payroll")):
        return "Accountancy"

    # Legal
    if any(k in n for k in ("solicitor", "legal", "law firm", "barrister", "conveyancing", "notary")):
        return "Legal"

    # Insurance (before Financial Services — more specific)
    if any(k in n for k in ("insur", "underwriter", "protection specialist")):
        return "Insurance"

    # Mortgage & Lending (before Financial Services)
    if any(k in n for k in ("mortgage", "lending", "bridging", "remortgage")):
        return "Mortgage & Lending"

    # Financial Services / IFA / Wealth
    if any(k in n for k in ("financ", "ifa", "wealth", "investment", "pension", "asset manag", "financial plann")):
        return "Financial Services"

    # HR & Recruitment
    if any(k in n for k in ("recruit", "staffing", "talent", "human resource", " hr ", "hr consult", "people consult", "employment")):
        return "HR & Recruitment"

    # IT & Technology
    if any(k in n for k in ("software", "cyber", "digital", "cloud", " it ", "i.t.", "technology", "tech ", "systems", "web design", "app dev", "data consult", "managed service")):
        return "IT & Technology"

    # Marketing, Design & PR
    if any(k in n for k in ("marketing", "brand", " pr ", "public relation", "design", "creative", "agency", "advertising", "social media", "copywriter", "content")):
        return "Marketing, Design & PR"

    # Property & Construction
    if any(k in n for k in ("property", "estate agent", "surveyor", "architect", "construction", "building", "developer", "plumber", "electrician", "roofing", "flooring", "landscap")):
        return "Property & Construction"

    # Health & Wellbeing
    if any(k in n for k in ("health", "wellbeing", "wellness", "therapy", "therapist", "physiother", "dental", "dentist", "nutrition", "fitness", "gym", "osteo", "chiropract", "counsell", "mental health")):
        return "Health & Wellbeing"

    # Print, Signs & Promotional
    if any(k in n for k in ("print", "signage", "signs", "promotional", "merchandise", "embroid", "workwear", "banner")):
        return "Print, Signs & Promotional"

    # Photography & Media
    if any(k in n for k in ("photo", "videograph", "filmmaker", "media", " film ", "podcast", "broadcast")):
        return "Photography & Media"

    # Business Coaching & Consulting (broad — after sector-specific above)
    if any(k in n for k in ("coach", "consult", "mentor", "training", "business advisor", "business adviser", "facilitator")):
        return "Business Coaching & Consulting"

    return "General Business"


# ==========================================================
# LOAD ATTENDANCE ACROSS TOWNS (FROM MONTHLY FILES)
# ==========================================================

def load_region_attendance() -> pd.DataFrame:
    """
    Load all *_attendance.xlsx files for all towns in the region.
    Mirrors the approach used in Buzz_Region_EventExcellence.py.
    """
    frames = []

    for town_code, town_label in TOWNS:
        monthly_dir = BASE / f"Buzz_Event_Dashboard_{town_code}" / "data_curated" / "monthly"
        if not monthly_dir.exists():
            print(f"[INFO] No monthly folder for {town_code} at {monthly_dir}")
            continue

        for fpath in monthly_dir.glob("*_attendance.xlsx"):
            try:
                df = pd.read_excel(fpath)
            except Exception as exc:
                print(f"[WARN] Failed to read {fpath}: {exc}")
                continue

            if df.empty:
                continue

            df["town_code"] = town_code
            df["town_name"] = town_label

            # Ensure required columns exist
            for col in ["event_month", "role", "company", "sponsor_company", "is_sponsor_contact"]:
                if col not in df.columns:
                    if col == "is_sponsor_contact":
                        df[col] = False
                    else:
                        df[col] = ""

            # Normalise sponsor contact flag
            df["is_sponsor_contact"] = (
                df["is_sponsor_contact"]
                .astype(str)
                .str.lower()
                .isin(["true", "1", "yes", "y"])
            )

            frames.append(df)

    if not frames:
        print("[WARN] No attendance files found across towns.")
        return pd.DataFrame()

    all_att = pd.concat(frames, ignore_index=True)
    return all_att


# ==========================================================
# LOAD SPONSOR FILES (FROM data_ref, TOWN BY TOWN)
# ==========================================================

def load_sponsors_for_town(town_code: str) -> pd.DataFrame:
    """
    Load sponsors_<Town>.csv from the data_ref folder for each town.
    We do NOT move files – we respect the existing structure.
    """
    folder = BASE / f"Buzz_Event_Dashboard_{town_code}" / "data_ref"
    f = folder / f"sponsors_{town_code}.csv"

    if not f.exists():
        return pd.DataFrame()

    df = pd.read_csv(f)
    df.columns = [c.strip() for c in df.columns]

    # Normalise
    if "sponsor_name" in df.columns:
        df["sponsor_name"] = df["sponsor_name"].astype(str).str.strip()
    else:
        # Fallback: if no sponsor_name column, try "primary_contact" as a label
        df["sponsor_name"] = df.get("primary_contact", "").astype(str).str.strip()

    df["company"] = df.get("company", "").astype(str).str.strip()
    df["category"] = df["company"].apply(categorise_company)

    df["start_date"] = pd.to_datetime(df.get("start_date"), errors="coerce", dayfirst=True)
    df["end_date"] = pd.to_datetime(df.get("end_date"), errors="coerce", dayfirst=True)

    # Active if no end date, or end date in the future
    df["active"] = df["end_date"].isna() | (df["end_date"] >= pd.Timestamp.today())

    df["town_code"] = town_code

    return df


# ==========================================================
# ATTENDANCE ANALYSIS FOR SPONSORS
# ==========================================================

def compute_sponsor_attendance(region_att: pd.DataFrame) -> pd.DataFrame:
    """
    Count how many events each sponsor contact booked for.
    Uses is_sponsor_contact + sponsor_company from the attendance files.
    """
    if region_att.empty:
        return pd.DataFrame()

    sponsor_rows = region_att[region_att["is_sponsor_contact"] == True].copy()

    if sponsor_rows.empty:
        return pd.DataFrame(columns=[
            "town_code",
            "town_name",
            "sponsor_company",
            "events_booked",
            "expected_events",
            "attendance_pct",
            "compliance_flag",
        ])

    grouped = sponsor_rows.groupby(
        ["town_code", "town_name", "sponsor_company"]
    ).agg(
        events_booked=("event_month", "nunique")
    ).reset_index()

    grouped["expected_events"] = 12
    grouped["attendance_pct"] = grouped["events_booked"] / grouped["expected_events"]

    def flag(val: float) -> str:
        if val >= 1.0:
            return "Excellent"
        if val >= 0.85:
            return "Good"
        if val >= 0.6:
            return "Needs Attention"
        return "At Risk"

    grouped["compliance_flag"] = grouped["attendance_pct"].apply(flag)

    return grouped


# ==========================================================
# INDUSTRY LOCK-OUT MAP
# ==========================================================

def compute_lockout_map(all_sponsors: pd.DataFrame,
                        region_att: pd.DataFrame) -> pd.DataFrame:
    """
    Produce town-level lockout lists (sponsor, host, ambassador).
    Industry is determined by company name via categorise_company().
    """
    rows = []

    for code, label in TOWNS:
        # Sponsor categories (active only)
        sponsors = all_sponsors[
            (all_sponsors["town_code"] == code) & (all_sponsors["active"])
        ]
        sponsor_cats = sorted(sponsors["category"].unique())

        # Host categories
        host_rows = region_att[
            (region_att["town_code"] == code)
            & (region_att["role"].astype(str).str.lower() == "host")
        ].copy()
        host_cats = sorted(
            host_rows["company"].apply(categorise_company).unique()
        ) if not host_rows.empty else []

        # Ambassador categories
        amb_rows = region_att[
            (region_att["town_code"] == code)
            & (region_att["role"].astype(str).str.lower() == "ambassador")
        ].copy()
        amb_cats = sorted(
            amb_rows["company"].apply(categorise_company).unique()
        ) if not amb_rows.empty else []

        combined = sorted(set(sponsor_cats + host_cats + amb_cats))

        rows.append({
            "town_code": code,
            "town_name": label,
            "sponsor_categories_locked": ", ".join(sponsor_cats),
            "host_categories_locked": ", ".join(host_cats),
            "amb_categories_locked": ", ".join(amb_cats),
            "combined_lockout": ", ".join(combined),
        })

    return pd.DataFrame(rows)


# ==========================================================
# SPONSOR CAPACITY PER TOWN
# ==========================================================

def compute_sponsor_capacity(all_sponsors: pd.DataFrame,
                             lockout_map: pd.DataFrame) -> pd.DataFrame:
    rows = []

    for code, label in TOWNS:
        sponsors = all_sponsors[
            (all_sponsors["town_code"] == code) & (all_sponsors["active"])
        ]
        current_count = sponsors["sponsor_name"].nunique()
        capacity_left = max(0, 4 - current_count)
        status = "Full" if capacity_left == 0 else "Open"

        lockouts_row = lockout_map[lockout_map["town_code"] == code]
        lockouts = ""
        if not lockouts_row.empty:
            lockouts = lockouts_row.iloc[0]["combined_lockout"]

        rows.append({
            "town_code": code,
            "town_name": label,
            "current_sponsors": current_count,
            "capacity_left": capacity_left,
            "status": status,
            "industry_lockouts": lockouts,
        })

    return pd.DataFrame(rows)


# ==========================================================
# SPONSOR OPPORTUNITY REPORT
# ==========================================================

def compute_sponsor_opportunities(capacity: pd.DataFrame,
                                  lockouts: pd.DataFrame) -> pd.DataFrame:
    rows = []

    for code, label in TOWNS:
        cap_row = capacity[capacity["town_code"] == code]
        lock_row = lockouts[lockouts["town_code"] == code]

        if cap_row.empty:
            capacity_left = 0
            status = "Unknown"
        else:
            capacity_left = int(cap_row.iloc[0]["capacity_left"])
            status = cap_row.iloc[0]["status"]

        lockout_categories = ""
        if not lock_row.empty:
            lockout_categories = lock_row.iloc[0]["combined_lockout"]

        can_sell = capacity_left > 0

        rows.append({
            "town_code": code,
            "town_name": label,
            "can_sell_sponsorship": "Yes" if can_sell else "No",
            "capacity_left": capacity_left,
            "status": status,
            "lockout_categories": lockout_categories,
        })

    return pd.DataFrame(rows)


# ==========================================================
# MAIN
# ==========================================================

def main():
    print("[INFO] Loading regional attendance from town monthly files...")
    region_att = load_region_attendance()
    if region_att.empty:
        print("[WARN] No attendance data found. Nothing to do.")
        return

    print("[INFO] Loading all sponsor files from data_ref...")
    sponsor_frames = []
    for code, label in TOWNS:
        df = load_sponsors_for_town(code)
        if not df.empty:
            sponsor_frames.append(df)

    if sponsor_frames:
        all_sponsors = pd.concat(sponsor_frames, ignore_index=True)
    else:
        print("[WARN] No sponsor files found in any data_ref folders.")
        all_sponsors = pd.DataFrame(columns=[
            "sponsor_name", "company", "category",
            "start_date", "end_date", "active", "town_code"
        ])

    print("[INFO] Computing sponsor attendance...")
    sponsor_attendance = compute_sponsor_attendance(region_att)

    print("[INFO] Computing industry lock-out map...")
    lockout_map = compute_lockout_map(all_sponsors, region_att)

    print("[INFO] Computing sponsor capacity...")
    capacity = compute_sponsor_capacity(all_sponsors, lockout_map)

    print("[INFO] Computing sponsor opportunities...")
    opportunities = compute_sponsor_opportunities(capacity, lockout_map)

    out = REGION_CURATED / "sponsor_intelligence.xlsx"
    print(f"[INFO] Writing {out}...")

    with pd.ExcelWriter(out, engine="xlsxwriter") as writer:
        sponsor_attendance.to_excel(writer, sheet_name="Sponsor_Attendance", index=False)
        lockout_map.to_excel(writer, sheet_name="Industry_Lockout_Map", index=False)
        capacity.to_excel(writer, sheet_name="Sponsor_Capacity", index=False)
        opportunities.to_excel(writer, sheet_name="Sponsor_Opportunities", index=False)

    print("[OK] Sponsor Intelligence Engine completed.")


if __name__ == "__main__":
    main()
