import argparse
import re
from pathlib import Path

import pandas as pd


# ------------- CONFIG -------------
BASE = Path(__file__).resolve().parent

# One row per town in the region
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
        buzz_month_index = mm + 2  # Jan=3, ... Oct=12

    label = f"{buzz_start}-{buzz_start + 1}"
    return buzz_start, label, buzz_month_index


def first_non_empty(series: pd.Series) -> str:
    for v in series:
        if isinstance(v, str) and v.strip():
            return v
    return ""


def most_common_non_empty(series: pd.Series) -> str:
    s = series.dropna().astype(str).str.strip()
    s = s[s != ""]
    if s.empty:
        return ""
    return s.value_counts().idxmax()


# ------------- LOAD PER-TOWN MASTER PEOPLE -------------
def load_town_master_people(town_conf):
    town_code = town_conf["code"]
    folder = BASE / town_conf["folder"] / "data_curated"
    path = folder / f"master_people_{town_code}.csv"

    if not path.exists():
        print(f"[WARN] No master people file for {town_code} at {path}")
        return pd.DataFrame()

    df = pd.read_csv(path)
    if df.empty:
        return df

    df["town_source"] = town_conf["label"]
    return df


def build_region_master_people():
    frames = []
    for town in TOWNS:
        df = load_town_master_people(town)
        if not df.empty:
            frames.append(df)

    if not frames:
        print("[WARN] No town master people files found. Region master will be empty.")
        cols = [
            "person_key",
            "name",
            "company",
            "email",
            "first_seen_month_region",
            "last_seen_month_region",
            "visits_region",
            "last_event_date_region",
            "role_region",
            "sponsor_company_region",
            "sponsor_level_region",
            "towns_visited",
            "towns_visited_count",
            "first_seen_buzzyear_label",
        ]
        return pd.DataFrame(columns=cols)

    all_people = pd.concat(frames, ignore_index=True)

    # Ensure expected columns exist even if older versions are missing some
    for col in [
        "person_key",
        "name",
        "company",
        "email",
        "first_seen_month",
        "last_seen_month",
        "visits",
        "last_event_date",
        "role",
        "sponsor_company",
        "sponsor_level",
    ]:
        if col not in all_people.columns:
            all_people[col] = ""

    # Aggregate to one row per person_key across the region
    def agg_towns(series):
        s = series.dropna().astype(str).str.strip()
        s = s[s != ""]
        if s.empty:
            return ""
        unique = sorted(set(s))
        return ", ".join(unique)

    # Role priority across towns
    role_priority_map = {
        "regional lead": 3,
        "host": 2,
        "ambassador": 1,
    }

    def pick_role(series):
        s = series.dropna().astype(str).str.strip()
        if s.empty:
            return ""
        # map to priority
        lower = s.str.lower()
        pr = lower.map(role_priority_map).fillna(0)
        # pick the highest priority role if any
        idx_max = pr.idxmax()
        if pr.loc[idx_max] == 0:
            # no known "special" roles, just pick most common
            return most_common_non_empty(s)
        return s.loc[idx_max]

    grouped = all_people.groupby("person_key", as_index=False).agg(
        name=("name", first_non_empty),
        company=("company", most_common_non_empty),
        email=("email", first_non_empty),
        first_seen_month_region=("first_seen_month", "min"),
        last_seen_month_region=("last_seen_month", "max"),
        visits_region=("visits", "sum"),
        last_event_date_region=("last_event_date", "max"),
        role_region=("role", pick_role),
        sponsor_company_region=("sponsor_company", most_common_non_empty),
        sponsor_level_region=("sponsor_level", most_common_non_empty),
        towns_visited=("town_source", agg_towns),
    )

    grouped["towns_visited_count"] = grouped["towns_visited"].apply(
        lambda x: 0 if not isinstance(x, str) or not x.strip() else len({t.strip() for t in x.split(",")})
    )

    # Buzz-year of first ever visit in region
    by_start_list = []
    by_label_list = []
    by_month_index_list = []
    for m in grouped["first_seen_month_region"]:
        _, label, _idx = month_to_buzzyear(m)
        by_start_list.append(_)
        by_label_list.append(label)
        by_month_index_list.append(_idx)

    grouped["first_seen_buzzyear_label"] = by_label_list

    # Define column order
    cols_order = [
        "person_key",
        "name",
        "company",
        "email",
        "first_seen_month_region",
        "last_seen_month_region",
        "visits_region",
        "last_event_date_region",
        "role_region",
        "sponsor_company_region",
        "sponsor_level_region",
        "towns_visited",
        "towns_visited_count",
        "first_seen_buzzyear_label",
    ]

    grouped = grouped[cols_order]
    return grouped


# ------------- LOAD PER-TOWN MONTHLY ATTENDANCE -------------
def infer_month_from_filename(path: Path) -> str:
    m = re.search(r"_(\d{4}-\d{2})_attendance\.xlsx$", path.name)
    if not m:
        return ""
    return m.group(1)


def load_region_attendance():
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

            # Make sure name/company/email exist
            for col in ["name", "company", "email"]:
                if col not in df.columns:
                    df[col] = ""

            # Rebuild person_key so we can deduplicate across towns
            df["person_key"] = df.apply(
                lambda r: make_person_key(r["name"], r["company"], r["email"]),
                axis=1,
            )

            frames.append(df)

    if not frames:
        print("[WARN] No attendance files found across towns.")
        return pd.DataFrame()

    all_att = pd.concat(frames, ignore_index=True)

    # Normalise some expected columns
    for col in ["role", "sponsor_company", "sponsor_level", "is_sponsor_contact"]:
        if col not in all_att.columns:
            all_att[col] = "" if col != "is_sponsor_contact" else False

    return all_att


# ------------- REGION MONTHLY & YEARLY SUMMARIES -------------
def build_region_monthly_summary(region_att: pd.DataFrame, region_master: pd.DataFrame):
    if region_att.empty:
        cols = [
            "event_month",
            "buzz_year_label",
            "total_bookings",
            "unique_attendees_region",
            "new_visitors_region",
            "host_attendance",
            "ambassador_attendance",
            "sponsor_contacts_present",
            "sponsor_companies_represented",
            "towns_running",
        ]
        return pd.DataFrame(columns=cols)

    # Buzz-year columns on attendance
    by_start = []
    by_label = []
    by_idx = []
    for m in region_att["event_month"]:
        s, lbl, idx = month_to_buzzyear(str(m))
        by_start.append(s)
        by_label.append(lbl)
        by_idx.append(idx)

    region_att["buzz_year_start"] = by_start
    region_att["buzz_year_label"] = by_label
    region_att["buzz_month_index"] = by_idx

    # Map: month -> how many NEW people regionally first seen that month
    if not region_master.empty:
        new_counts = (
            region_master.groupby("first_seen_month_region")["person_key"]
            .nunique()
            .to_dict()
        )
    else:
        new_counts = {}

    # Build monthly summary
    summaries = []
    for month, g in region_att.groupby("event_month"):
        month = str(month)

        # ignore any weird blanks
        if month == "" or month.lower() == "nan":
            continue

        # Basic info
        buzz_start, buzz_label, buzz_idx = month_to_buzzyear(month)

        total_bookings = len(g)
        unique_attendees = g["person_key"].nunique()

        host_att = (g["role"].astype(str).str.lower() == "host").sum()
        amb_att = (g["role"].astype(str).str.lower() == "ambassador").sum()

        # sponsor metrics
        sponsor_contacts = g["is_sponsor_contact"].fillna(False)
        # If column is stored as text "TRUE"/"FALSE", normalise
        if sponsor_contacts.dtype == "object":
            sponsor_contacts_bool = sponsor_contacts.astype(str).str.lower().isin(["true", "1", "yes"])
        else:
            sponsor_contacts_bool = sponsor_contacts.astype(bool)

        sponsor_contacts_present = int(sponsor_contacts_bool.sum())
        sponsor_companies = (
            g.loc[sponsor_contacts_bool, "sponsor_company"]
            .replace("", pd.NA)
            .dropna()
            .nunique()
        )

        towns_running = g["town_code"].nunique()

        new_visitors_region = int(new_counts.get(month, 0))

        summaries.append(
            {
                "event_month": month,
                "buzz_year_label": buzz_label,
                "buzz_month_index": buzz_idx,
                "total_bookings": total_bookings,
                "unique_attendees_region": unique_attendees,
                "new_visitors_region": new_visitors_region,
                "host_attendance": int(host_att),
                "ambassador_attendance": int(amb_att),
                "sponsor_contacts_present": sponsor_contacts_present,
                "sponsor_companies_represented": int(sponsor_companies),
                "towns_running": int(towns_running),
            }
        )

    monthly_df = pd.DataFrame(summaries)
    if not monthly_df.empty:
        monthly_df = monthly_df.sort_values(["buzz_year_label", "buzz_month_index"])

    return monthly_df


def build_region_buzzyear_summary(region_att: pd.DataFrame, region_master: pd.DataFrame):
    if region_att.empty:
        cols = [
            "buzz_year_label",
            "total_bookings",
            "unique_attendees_region",
            "new_visitors_region",
            "host_attendance",
            "ambassador_attendance",
            "sponsor_contacts_present",
            "sponsor_companies_represented",
        ]
        return pd.DataFrame(columns=cols)

    # Make sure buzz_year_label exists
    if "buzz_year_label" not in region_att.columns:
        by_start = []
        by_label = []
        by_idx = []
        for m in region_att["event_month"]:
            s, lbl, idx = month_to_buzzyear(str(m))
            by_start.append(s)
            by_label.append(lbl)
            by_idx.append(idx)
        region_att["buzz_year_label"] = by_label

    # For new visitors per buzz year, use the master file's first_seen_buzzyear_label
    if not region_master.empty and "first_seen_buzzyear_label" in region_master.columns:
        new_year_counts = (
            region_master.groupby("first_seen_buzzyear_label")["person_key"]
            .nunique()
            .to_dict()
        )
    else:
        new_year_counts = {}

    summaries = []
    for year_label, g in region_att.groupby("buzz_year_label"):
        if not isinstance(year_label, str) or year_label.strip() == "":
            continue

        total_bookings = len(g)
        unique_attendees = g["person_key"].nunique()

        host_att = (g["role"].astype(str).str.lower() == "host").sum()
        amb_att = (g["role"].astype(str).str.lower() == "ambassador").sum()

        sponsor_contacts = g["is_sponsor_contact"].fillna(False)
        if sponsor_contacts.dtype == "object":
            sponsor_contacts_bool = sponsor_contacts.astype(str).str.lower().isin(["true", "1", "yes"])
        else:
            sponsor_contacts_bool = sponsor_contacts.astype(bool)

        sponsor_contacts_present = int(sponsor_contacts_bool.sum())
        sponsor_companies = (
            g.loc[sponsor_contacts_bool, "sponsor_company"]
            .replace("", pd.NA)
            .dropna()
            .nunique()
        )

        new_visitors_region = int(new_year_counts.get(year_label, 0))

        summaries.append(
            {
                "buzz_year_label": year_label,
                "total_bookings": total_bookings,
                "unique_attendees_region": unique_attendees,
                "new_visitors_region": new_visitors_region,
                "host_attendance": int(host_att),
                "ambassador_attendance": int(amb_att),
                "sponsor_contacts_present": sponsor_contacts_present,
                "sponsor_companies_represented": int(sponsor_companies),
            }
        )

    year_df = pd.DataFrame(summaries)
    if not year_df.empty:
        year_df = year_df.sort_values("buzz_year_label")

    return year_df


# ------------- MAIN RUNNER -------------
def main():
    parser = argparse.ArgumentParser(description="Build regional Buzz roll-up (Leics & Rutland).")
    args = parser.parse_args()

    print("[INFO] Building region master people...")
    region_master = build_region_master_people()
    master_path = REGION_CURATED / "region_master_people.csv"
    region_master.to_csv(master_path, index=False)
    print(f"[OK] Wrote {master_path} ({len(region_master)} people)")

    print("[INFO] Loading regional attendance...")
    region_att = load_region_attendance()

    if region_att.empty:
        print("[WARN] No attendance data found. Skipping summaries.")
        return

    print("[INFO] Building regional monthly summary...")
    monthly_df = build_region_monthly_summary(region_att, region_master)
    monthly_path = REGION_CURATED / "region_monthly_summary.xlsx"
    monthly_df.to_excel(monthly_path, index=False)
    print(f"[OK] Wrote {monthly_path} ({len(monthly_df)} rows)")

    print("[INFO] Building regional buzz-year summary...")
    year_df = build_region_buzzyear_summary(region_att, region_master)
    year_path = REGION_CURATED / "region_buzzyear_summary.xlsx"
    year_df.to_excel(year_path, index=False)
    print(f"[OK] Wrote {year_path} ({len(year_df)} rows)")


if __name__ == "__main__":
    main()
