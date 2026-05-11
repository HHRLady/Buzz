import re
from pathlib import Path
from datetime import datetime

import pandas as pd


# ------------- CONFIG -------------
BASE = Path(__file__).resolve().parent

# Towns in your region – keep in sync with your dashboards
TOWNS = [
    {
        "code": "MarketHarborough",
        "label": "Market Harborough",
        "folder": "Buzz_Event_Dashboard_MarketHarborough",
    },
    {
        "code": "Leicester",
        "label": "Leicester",
        "folder": "Buzz_Event_Dashboard_Leicester",
    },
    {
        "code": "Lutterworth",
        "label": "Lutterworth",
        "folder": "Buzz_Event_Dashboard_Lutterworth",
    },
    {
        "code": "Hinckley",
        "label": "Hinckley",
        "folder": "Buzz_Event_Dashboard_Hinckley",
    },
    {
        "code": "Loughborough",
        "label": "Loughborough",
        "folder": "Buzz_Event_Dashboard_Loughborough",
    },
]

REGION_CURATED = BASE / "Buzz_Region_Curated"
REGION_CURATED.mkdir(parents=True, exist_ok=True)

REGION_REF = BASE / "Buzz_Region_Ref"
REGION_REF.mkdir(parents=True, exist_ok=True)


# ------------- HELPERS -------------
def norm_text(x: str) -> str:
    if pd.isna(x):
        return ""
    x = str(x).strip().lower()
    x = re.sub(r"\s+", " ", x)
    return x


def norm_person_name(x: str) -> str:
    if pd.isna(x):
        return ""
    x = str(x)
    x = re.sub(r"\(.*?\)", "", x)
    return norm_text(x)


def clean_company(c: str) -> str:
    c = norm_text(c)
    c = re.sub(r"\b(ltd|limited)\b", "limited", c)
    c = c.replace("&", "and")
    c = re.sub(r"[^a-z0-9\s]", "", c)
    c = re.sub(r"\s+", " ", c).strip()
    return c


def make_person_key(name: str, company: str, email: str = "") -> str:
    email_n = norm_text(email)
    if email_n:
        return f"email||{email_n}"
    return f"nameco||{norm_person_name(name)}||{clean_company(company)}"


def month_to_buzzyear(month: str):
    """
    Convert 'YYYY-MM' to (buzz_year_start, buzz_year_label, buzz_month_index)
    Buzz year runs Nov (1) to Oct (12).
    """
    if not isinstance(month, str) or len(month) < 7 or "-" not in month:
        return None, None, None
    try:
        year = int(month[:4])
        mm = int(month[5:7])
    except ValueError:
        return None, None, None

    if mm >= 11:
        buzz_start = year
        buzz_month_index = mm - 10  # Nov=1, Dec=2
    else:
        buzz_start = year - 1
        buzz_month_index = mm + 2   # Jan=3, ... Oct=12

    label = f"{buzz_start}-{buzz_start + 1}"
    return buzz_start, label, buzz_month_index


def infer_month_from_filename(path: Path) -> str:
    m = re.search(r"_(\d{4}-\d{2})_attendance\.xlsx$", path.name)
    if not m:
        return ""
    return m.group(1)


def _pick_column(df: pd.DataFrame, tokens):
    for c in df.columns:
        low = str(c).lower()
        if any(t in low for t in tokens):
            return c
    return None


# ------------- LOAD ATTENDANCE ACROSS TOWNS -------------
def load_region_attendance() -> pd.DataFrame:
    """
    Load all *_attendance.xlsx files for all towns in the region.
    Rebuilds person_key so we can identify people across towns.
    """
    frames = []
    for town in TOWNS:
        monthly_dir = BASE / town["folder"] / "data_curated" / "monthly"
        if not monthly_dir.exists():
            print(f"[INFO] No monthly folder for {town['code']} at {monthly_dir}")
            continue

        for fpath in monthly_dir.glob("*_attendance.xlsx"):
            try:
                df = pd.read_excel(fpath)
            except Exception as exc:
                print(f"[WARN] Failed to read {fpath}: {exc}")
                continue

            if df.empty:
                continue

            df["town"] = town["label"]
            df["town_code"] = town["code"]

            # Ensure event_month exists
            if "event_month" not in df.columns:
                month = infer_month_from_filename(fpath)
                df["event_month"] = month

            # Ensure key columns exist
            for col in ["name", "company", "email", "payment_type", "role",
                        "sponsor_company", "is_sponsor_contact"]:
                if col not in df.columns:
                    if col == "is_sponsor_contact":
                        df[col] = False
                    else:
                        df[col] = ""

            # Rebuild person_key to be safe
            df["person_key"] = df.apply(
                lambda r: make_person_key(r["name"], r["company"], r["email"]),
                axis=1,
            )

            frames.append(df)

    if not frames:
        print("[WARN] No attendance files found across towns.")
        return pd.DataFrame()

    all_att = pd.concat(frames, ignore_index=True)

    # Normalise is_sponsor_contact to boolean
    if all_att["is_sponsor_contact"].dtype == "object":
        all_att["is_sponsor_contact"] = (
            all_att["is_sponsor_contact"]
            .astype(str)
            .str.lower()
            .isin(["true", "1", "yes", "y"])
        )
    else:
        all_att["is_sponsor_contact"] = all_att["is_sponsor_contact"].astype(bool)

    # Attach buzz-year columns
    by_start = []
    by_label = []
    by_idx = []
    for m in all_att["event_month"]:
        s, lbl, idx = month_to_buzzyear(str(m))
        by_start.append(s)
        by_label.append(lbl)
        by_idx.append(idx)

    all_att["buzz_year_start"] = by_start
    all_att["buzz_year_label"] = by_label
    all_att["buzz_month_index"] = by_idx

    return all_att


# ------------- LOAD BUZZ PLUS MEMBERS -------------
def load_buzz_plus_members() -> pd.DataFrame:
    """
    Load buzz_plus_members.[csv/xlsx/xls] from Buzz_Region_Ref.

    Expected logical columns:
      - member_name
      - member_company
      - town_code
      - start_date
      - end_date (optional)
      - status (optional)
    """
    path = None
    for ext in (".csv", ".xlsx", ".xls"):
        candidate = REGION_REF / f"buzz_plus_members{ext}"
        if candidate.exists():
            path = candidate
            break

    if path is None:
        print("[INFO] No buzz_plus_members file found. Buzz Plus score will be 0.")
        return pd.DataFrame(columns=[
            "member_name",
            "member_company",
            "town_code",
            "start_date",
            "end_date",
            "status_norm",
        ])

    if path.suffix.lower() == ".csv":
        df = pd.read_csv(path)
    else:
        df = pd.read_excel(path)

    if df.empty:
        return pd.DataFrame(columns=[
            "member_name",
            "member_company",
            "town_code",
            "start_date",
            "end_date",
            "status_norm",
        ])

    df.columns = [str(c).strip() for c in df.columns]

    name_col = _pick_column(df, ["member", "name"])
    comp_col = _pick_column(df, ["company", "business", "organisation", "organization", "org"])
    town_col = _pick_column(df, ["town", "event", "location"])
    start_col = _pick_column(df, ["start"])
    end_col = _pick_column(df, ["end", "until", "to", "finish"])
    status_col = _pick_column(df, ["status"])

    if not (name_col and town_col and start_col):
        print("[WARN] buzz_plus_members file missing key columns – ignoring")
        return pd.DataFrame(columns=[
            "member_name",
            "member_company",
            "town_code",
            "start_date",
            "end_date",
            "status_norm",
        ])

    plus = pd.DataFrame()
    plus["member_name"] = df[name_col].fillna("").astype(str)
    if comp_col:
        plus["member_company"] = df[comp_col].fillna("").astype(str)
    else:
        plus["member_company"] = ""

    plus["town_code"] = df[town_col].fillna("").astype(str)

    plus["start_date"] = pd.to_datetime(df[start_col], errors="coerce", dayfirst=True)
    if end_col:
        plus["end_date"] = pd.to_datetime(df[end_col], errors="coerce", dayfirst=True)
    else:
        plus["end_date"] = pd.NaT

    if status_col:
        plus["status_norm"] = df[status_col].fillna("").astype(str).str.lower().str.strip()
    else:
        plus["status_norm"] = ""

    # Drop rows without a town or start date
    plus = plus[plus["town_code"].astype(str).str.strip() != ""]
    plus = plus[plus["start_date"].notna()]

    return plus


# ------------- AWARD RATING LOGIC -------------
def compute_award_rating(
    avg_best10: float,
    ambassadors: int,
    plus_members: int,
    sponsors: int,
):
    """
    Implement the Event Excellence award logic:

      - All awards (3/4/5-star) require avg_best10 >= 25.
      - 3-Star: attendance only.
      - 4-Star: attendance + 2 ambassadors.
      - 5-Star: attendance + 2 ambassadors + combined total of
                Buzz Plus members + sponsors >= 4 (any mix).
                e.g. 4 sponsors, 3+1, 2+2, 1+3, 0+4 all qualify.
      - Special prize: as 5-Star plus sponsors >= 4 AND Buzz Plus >= 5.
    """

    # Default
    stars = 0
    rating = "Below 3-Star Criteria"
    special = False

    if avg_best10 < 25:
        return stars, rating, special

    # Attendance criterion met – at least 3 stars
    # 5-Star logic: combined Buzz Plus + sponsors >= 4
    if ambassadors >= 2 and (plus_members + sponsors) >= 4:
        stars = 5
        rating = "5-Star Award"
        if sponsors >= 4 and plus_members >= 5:
            special = True
            rating = "5-Star Award (Special Prize Eligible)"
        return stars, rating, special

    # 4-Star logic
    if ambassadors >= 2:
        stars = 4
        rating = "4-Star Award"
        return stars, rating, special

    # 3-Star logic (attendance only)
    stars = 3
    rating = "3-Star Award"
    return stars, rating, special


# ------------- BUILD EVENT EXCELLENCE TABLE -------------
def build_event_excellence(region_att: pd.DataFrame, plus_members: pd.DataFrame) -> pd.DataFrame:
    """
    Build one row per town per buzz year, aligned to the Event Excellence table.

    Columns:
      - buzz_year_label
      - town_code
      - event_name
      - host_name
      - avg_unique_attendees_allmonths
      - avg_unique_attendees_best10
      - active_ambassadors
      - active_plus_members
      - local_sponsors
      - award_stars
      - award_rating
      - special_prize_eligible
    """
    if region_att.empty:
        cols = [
            "buzz_year_label",
            "town_code",
            "event_name",
            "host_name",
            "avg_unique_attendees_allmonths",
            "avg_unique_attendees_best10",
            "active_ambassadors",
            "active_plus_members",
            "local_sponsors",
            "award_stars",
            "award_rating",
            "special_prize_eligible",
        ]
        return pd.DataFrame(columns=cols)

    region_att = region_att.copy()
    region_att = region_att[region_att["buzz_year_label"].notna()]

    summaries = []

    for (year_label, town_code), g in region_att.groupby(["buzz_year_label", "town_code"]):
        if not isinstance(year_label, str) or year_label.strip() == "":
            continue

        event_name = g["town"].iloc[0] if "town" in g.columns else town_code

        # Host name – most frequent Host in that town/year
        host_mask = g["role"].astype(str).str.lower() == "host"
        host_name = ""
        if host_mask.any():
            host_counts = (
                g.loc[host_mask, "name"]
                .astype(str)
                .str.strip()
                .value_counts()
            )
            if not host_counts.empty:
                host_name = host_counts.idxmax()

        # Months where the event ran (any attendance)
        months = sorted({
            str(m) for m in g["event_month"].astype(str)
            if str(m).strip() not in ("", "nan")
        })

        avg_allmonths = 0.0
        avg_best10 = 0.0

        if months:
            # Unique attendees per month (attendance-based; no payment-type dependency)
            attendees_per_month = (
                g.groupby("event_month")["person_key"]
                .nunique()
                .to_dict()
            )

            monthly_counts = [int(attendees_per_month.get(m, 0)) for m in months]
            total_attendees = sum(monthly_counts)

            # Average across all months that ran
            avg_allmonths = total_attendees / len(months)

            # "Best ten months" average
            top_counts = sorted(monthly_counts, reverse=True)[: min(10, len(monthly_counts))]
            if top_counts:
                avg_best10 = sum(top_counts) / len(top_counts)

        # Active ambassadors (approximation for "trained & active 3+ months")
        amb_mask = g["role"].astype(str).str.lower() == "ambassador"
        active_ambassadors = g.loc[amb_mask, "person_key"].nunique()

        # Local sponsors (distinct sponsor companies with a sponsor contact)
        s_mask = g["is_sponsor_contact"].fillna(False)
        local_sponsors = (
            g.loc[s_mask, "sponsor_company"]
            .replace("", pd.NA)
            .dropna()
            .nunique()
        )

        # Active Buzz Plus members for this town & year
        active_plus_members = 0
        if not plus_members.empty:
            try:
                year_start = int(str(year_label).split("-")[0])
                year_start_date = datetime(year_start, 11, 1)
                year_end_date = datetime(year_start + 1, 10, 31)
            except Exception:
                year_start_date = None
                year_end_date = None

            if year_start_date is not None:
                pm = plus_members[plus_members["town_code"] == town_code].copy()
                if not pm.empty:
                    mask_period = pm["start_date"] <= year_end_date
                    mask_period &= pm["end_date"].isna() | (pm["end_date"] >= year_start_date)

                    if "status_norm" in pm.columns:
                        ok_status = pm["status_norm"].isin(["", "active", "current", "live"])
                        mask_period &= ok_status

                    active_plus_members = int(mask_period.sum())

        # Compute award rating
        stars, award_rating, special_flag = compute_award_rating(
            avg_best10=avg_best10,
            ambassadors=active_ambassadors,
            plus_members=active_plus_members,
            sponsors=int(local_sponsors),
        )

        summaries.append(
            {
                "buzz_year_label": year_label,
                "town_code": town_code,
                "event_name": event_name,
                "host_name": host_name,
                "avg_unique_attendees_allmonths": round(avg_allmonths, 1),
                "avg_unique_attendees_best10": round(avg_best10, 1),
                "active_ambassadors": int(active_ambassadors),
                "active_plus_members": active_plus_members,
                "local_sponsors": int(local_sponsors),
                "award_stars": int(stars),
                "award_rating": award_rating,
                "special_prize_eligible": "Yes" if special_flag else "No",
            }
        )

    df = pd.DataFrame(summaries)
    if not df.empty:
        df = df.sort_values(["buzz_year_label", "event_name"])

    return df


# ------------- MAIN RUNNER -------------
def main():
    print("[INFO] Loading regional attendance for Event Excellence...")
    region_att = load_region_attendance()

    if region_att.empty:
        print("[WARN] No attendance data found. Nothing to do.")
        return

    print("[INFO] Loading Buzz Plus members...")
    plus_members = load_buzz_plus_members()

    print("[INFO] Building Event Excellence table...")
    ee_df = build_event_excellence(region_att, plus_members)

    out_path = REGION_CURATED / "region_event_excellence.xlsx"
    ee_df.to_excel(out_path, index=False)
    print(f"[OK] Wrote {out_path} ({len(ee_df)} rows)")


if __name__ == "__main__":
    main()
