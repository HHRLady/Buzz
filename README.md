# Business Buzz — Leicestershire & Rutland

**Regional data pipeline and host pack automation for the Business Buzz Leicestershire & Rutland region.**

This repo contains the Python scripts that process monthly attendance data, generate branded host packs for each town, and produce the Regional Lead's monthly overview. It does not contain any raw data — all source files and curated outputs live locally on the RL's machine.

---

## What this does

Each month, the pipeline:

1. Processes raw attendance CSVs from the Buzz app into curated monthly files (run locally)
2. Builds regional intelligence outputs — Buzz Plus prospects, sponsor capacity, event excellence
3. Generates a branded HTML host pack for each of the 5 towns
4. Generates a branded HTML Regional Lead monthly overview
5. (Via GitHub Actions) Publishes the HTML packs as web links and emails each host automatically

---

## Towns covered

| Town | Status |
|------|--------|
| Market Harborough | Active |
| Leicester | Active |
| Lutterworth | Active |
| Hinckley | Active |
| Loughborough | Active |

---

## Repo structure

```
/
├── Buzz_Region_AllBuild.py          # Master runner — kicks off the full pipeline
├── Buzz_Town_HostPack_v2.py         # Generates HTML host pack per town
├── Buzz_RL_MonthlyPack_v2.py        # Generates HTML RL monthly overview
├── Buzz_Region_Rollup.py            # Regional attendance rollup
├── Buzz_Region_EventExcellence.py   # Event Excellence scoring
├── Buzz_Region_Dashboard.py         # Regional dashboard (feeds RL pack)
├── buzzplus_intelligence.py         # Buzz Plus prospect engine
├── sponsor_intelligence.py          # Sponsor capacity and pipeline
├── send_packs.py                    # Emails host packs via SendGrid (CI only)
├── requirements.txt                 # Python dependencies
├── .gitignore                       # Keeps data files off GitHub
├── .github/
│   └── workflows/
│       └── build-host-packs.yml     # GitHub Actions workflow
└── Buzz_Region_Curated/
    └── host_packs/
        ├── HostPack_Market_Harborough.html
        ├── HostPack_Leicester.html
        ├── HostPack_Lutterworth.html
        ├── HostPack_Hinckley.html
        ├── HostPack_Loughborough.html
        └── RL_Monthly_Pack.html
```

> **Data files are not in this repo.** The `Buzz_Event_Dashboard_*` folders and all `.xlsx` / `.csv` outputs live locally only. See `.gitignore` for the full exclusion list.

---

## Running locally

### Requirements

- Python 3.12+
- Dependencies: `pip install -r requirements.txt`

### Run the full pipeline

```bash
python Buzz_Region_AllBuild.py
```

This runs all town builds followed by all regional scripts in the correct order. HTML packs are written to `Buzz_Region_Curated/host_packs/`.

### Run a single host pack

```bash
python Buzz_Town_HostPack_v2.py --town Loughborough
python Buzz_Town_HostPack_v2.py --town ALL
python Buzz_Town_HostPack_v2.py --town ALL --event-date "Thursday 17 April 2026"
```

### Run the RL pack only

```bash
python Buzz_RL_MonthlyPack_v2.py
```

### Send packs manually (dry run)

```bash
python send_packs.py --dry-run
```

---

## GitHub Actions (automated delivery)

The workflow in `.github/workflows/build-host-packs.yml` runs when triggered manually from the Actions tab.

**To trigger:**
1. Go to the repo on GitHub
2. Click the **Actions** tab
3. Click **Build and Send Host Packs**
4. Click **Run workflow**
5. Optionally enter the next event date and tick dry run for a test

**What it does:**
- Installs Python and dependencies
- Runs `Buzz_Town_HostPack_v2.py --town ALL`
- Publishes HTML files to GitHub Pages
- Runs `send_packs.py` to email each host their pack

### Required GitHub Secrets

Set these under Settings → Secrets and variables → Actions:

| Secret | Description |
|--------|-------------|
| `SENDGRID_API_KEY` | SendGrid API key for email delivery |
| `SENDER_EMAIL` | From address (e.g. emma@thehorseyhrlady.co.uk) |
| `SENDER_NAME` | Display name (e.g. Emma Smith - Business Buzz) |
| `PAGES_BASE_URL` | GitHub Pages base URL for web links |
| `HOST_EMAIL_MARKETHARBOROUGH` | Host's email address |
| `HOST_EMAIL_LEICESTER` | Host's email address |
| `HOST_EMAIL_LUTTERWORTH` | Host's email address |
| `HOST_EMAIL_HINCKLEY` | Host's email address |
| `HOST_EMAIL_LOUGHBOROUGH` | Host's email address |

---

## Host pack logic

All logic is attendance-based. No payment data is used.

| Section | Logic |
|---------|-------|
| Regulars | 2+ months attended, excludes team roles, top 10 |
| Lapsed | 2+ months ever, last seen 3–24 months ago, top 8 |
| Buzz Plus prospects | From regional intelligence, strong first, top 8 |
| Sponsors | From `sponsors_<TownCode>.csv` in `data_ref` |
| Snapshot | From `region_dashboard.xlsx` Town_Overview |

**Host personalisation** — host first name is pulled from `roles_<TownCode>.csv` (role = Host, no end date).

---

## RL pack logic

| Section | Source |
|---------|--------|
| 5-town health | `region_dashboard.xlsx` + `region_event_excellence.xlsx` |
| Actions | Derived from health thresholds |
| Sponsor pipeline | `sponsor_intelligence.xlsx` |
| Cross-town visitors | `region_master_people.csv` (towns_visited_count ≥ 2) |
| Host candidates | `region_master_people.csv` (visits ≥ 6, no current role) |

---

## Brand

Business Buzz official brand colours:

| Name | Hex |
|------|-----|
| Teal ("Blue") | `#00A19A` |
| Orange | `#F39200` |
| Pink | `#D60B52` |
| Lime ("Green") | `#B6BD00` |

Primary typeface: Century Gothic. Secondary: Gill Sans.

---

## Data governance

- Email addresses are the primary visitor identifier (compulsory in the booking system)
- No financial or payment data is used in any logic
- HTML outputs published to GitHub Pages contain no email addresses — names and companies only
- Raw data and curated Excel files never leave the RL's local machine

---

## Maintained by

Emma Smith — Regional Lead, Business Buzz Leicestershire & Rutland  
[portal.mybuzz.uk](https://portal.mybuzz.uk)
