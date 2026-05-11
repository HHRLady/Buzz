"""
Buzz_Region_SyncAmbassadors.py
──────────────────────────────────────────────────────────────────────────────
Rebuilds Buzz_Region_Ref/buzz_ambassadors.xlsx from the five town roles CSVs.

Run this whenever you update a roles_{Town}.csv. It is also called automatically
by Buzz_Region_AllBuild.py (first regional step, before Event Excellence).

What it does
────────────
  • Reads every Ambassador row from all five town roles CSVs.
  • Derives has_left automatically:
      – end_date set and in the past → Yes
      – end_date blank or in the future → No
  • Preserves ambassador_training_complete and notes from the existing
    buzz_ambassadors.xlsx so you only need to maintain training status
    in one place. New people default to training_complete = No.
  • Writes the result back to buzz_ambassadors.xlsx, sorted by town then name.

Columns in buzz_ambassadors.xlsx
─────────────────────────────────
  email | name | town_code | town_name | ambassador_training_complete | has_left | notes
"""

from pathlib import Path
from datetime import date
import pandas as pd

# ── Config ────────────────────────────────────────────────────────────────────

BASE = Path(__file__).resolve().parent

TOWNS = [
    {"code": "MarketHarborough", "name": "Market Harborough",
     "folder": "Buzz_Event_Dashboard_MarketHarborough"},
    {"code": "Leicester",        "name": "Leicester",
     "folder": "Buzz_Event_Dashboard_Leicester"},
    {"code": "Lutterworth",      "name": "Lutterworth",
     "folder": "Buzz_Event_Dashboard_Lutterworth"},
    {"code": "Hinckley",         "name": "Hinckley",
     "folder": "Buzz_Event_Dashboard_Hinckley"},
    {"code": "Loughborough",     "name": "Loughborough",
     "folder": "Buzz_Event_Dashboard_Loughborough"},
]

REGION_REF   = BASE / "Buzz_Region_Ref"
OUTPUT_PATH  = REGION_REF / "buzz_ambassadors.xlsx"

TODAY = date.today()


# ── Helpers ───────────────────────────────────────────────────────────────────

def _read_csv_safe(path: Path) -> pd.DataFrame:
    """Read CSV trying UTF-8 first, then latin-1 fallback."""
    try:
        return pd.read_csv(path, encoding="utf-8-sig")
    except UnicodeDecodeError:
        return pd.read_csv(path, encoding="latin-1")


def _has_left(end_date_val) -> str:
    """Return 'Yes' if end_date is set and in the past, else 'No'."""
    if pd.isna(end_date_val):
        return "No"
    try:
        d = pd.to_datetime(end_date_val, dayfirst=True, errors="coerce")
        if pd.isna(d):
            return "No"
        return "Yes" if d.date() < TODAY else "No"
    except Exception:
        return "No"


# ── Load existing file to preserve training status and notes ──────────────────

def load_existing() -> dict:
    """
    Returns a dict keyed by (email_lower, town_code) →
    {'ambassador_training_complete': ..., 'notes': ...}
    """
    existing = {}
    if not OUTPUT_PATH.exists():
        return existing
    try:
        df = pd.read_excel(OUTPUT_PATH)
        df.columns = [str(c).strip() for c in df.columns]
        for _, row in df.iterrows():
            key = (str(row.get("email", "")).strip().lower(),
                   str(row.get("town_code", "")).strip())
            existing[key] = {
                "ambassador_training_complete": str(row.get("ambassador_training_complete", "No")).strip(),
                "notes": str(row.get("notes", "")).strip() if not pd.isna(row.get("notes", "")) else "",
            }
    except Exception as exc:
        print(f"[WARN] Could not read existing buzz_ambassadors.xlsx: {exc}")
    return existing


# ── Read all town roles CSVs ──────────────────────────────────────────────────

def load_all_ambassadors() -> pd.DataFrame:
    rows = []
    for town in TOWNS:
        csv_path = BASE / town["folder"] / "data_ref" / f"roles_{town['code']}.csv"
        if not csv_path.exists():
            print(f"[WARN] No roles file found for {town['code']} at {csv_path}")
            continue

        try:
            df = _read_csv_safe(csv_path)
        except Exception as exc:
            print(f"[WARN] Could not read {csv_path}: {exc}")
            continue

        df.columns = [str(c).strip() for c in df.columns]

        # Find role column (flexible naming)
        role_col = next((c for c in df.columns if "role" in c.lower()), None)
        if role_col is None:
            print(f"[WARN] No role column in {csv_path}")
            continue

        # Filter ambassadors only
        ambs = df[df[role_col].astype(str).str.strip().str.lower() == "ambassador"].copy()
        if ambs.empty:
            continue

        # Normalise columns
        email_col    = next((c for c in df.columns if "email" in c.lower()), None)
        name_col     = next((c for c in df.columns if "name" in c.lower()), None)
        end_col      = next((c for c in df.columns if "end" in c.lower()), None)

        for _, row in ambs.iterrows():
            email    = str(row[email_col]).strip() if email_col else ""
            name     = str(row[name_col]).strip()  if name_col  else ""
            end_val  = row[end_col] if end_col and end_col in row.index else None

            rows.append({
                "email":     email,
                "name":      name,
                "town_code": town["code"],
                "town_name": town["name"],
                "_end_val":  end_val,
            })

    if not rows:
        print("[WARN] No ambassador rows found across any town roles files.")
        return pd.DataFrame()

    return pd.DataFrame(rows)


# ── Main ──────────────────────────────────────────────────────────────────────

def main():
    print("[INFO] Loading existing buzz_ambassadors.xlsx for training/notes preservation...")
    existing = load_existing()
    print(f"[INFO] Found {len(existing)} existing records to preserve.")

    print("[INFO] Reading ambassador rows from all town roles CSVs...")
    df = load_all_ambassadors()

    if df.empty:
        print("[WARN] Nothing to write.")
        return

    print(f"[INFO] Found {len(df)} ambassador role entries across all towns.")

    # Derive has_left
    df["has_left"] = df["_end_val"].apply(_has_left)

    # Merge in training status and notes from existing file
    def _get_existing(email, town_code, field, default=""):
        key = (email.strip().lower(), town_code.strip())
        return existing.get(key, {}).get(field, default)

    df["ambassador_training_complete"] = df.apply(
        lambda r: _get_existing(r["email"], r["town_code"], "ambassador_training_complete", "No"),
        axis=1
    )
    df["notes"] = df.apply(
        lambda r: _get_existing(r["email"], r["town_code"], "notes", ""),
        axis=1
    )

    # Preserve email casing from the roles file but normalise lookup
    # Final column order
    out = df[["email", "name", "town_code", "town_name",
              "ambassador_training_complete", "has_left", "notes"]].copy()

    # Sort: town_name, then has_left (No first), then name
    out = out.sort_values(
        ["town_name", "has_left", "name"],
        ascending=[True, True, True]
    ).reset_index(drop=True)

    REGION_REF.mkdir(parents=True, exist_ok=True)
    out.to_excel(OUTPUT_PATH, index=False)

    # Summary
    active  = (out["has_left"] == "No").sum()
    left    = (out["has_left"] == "Yes").sum()
    trained = ((out["has_left"] == "No") & (out["ambassador_training_complete"].str.lower() == "yes")).sum()
    new_keys = set(
        (r["email"].strip().lower(), r["town_code"]) for _, r in out.iterrows()
    ) - set(existing.keys())

    print(f"[OK] buzz_ambassadors.xlsx written — {active} active, {left} left, {trained} trained.")
    if new_keys:
        print(f"[INFO] {len(new_keys)} new ambassador(s) added (training defaulted to No):")
        for email, town in sorted(new_keys):
            print(f"       {email} ({town})")


if __name__ == "__main__":
    main()
