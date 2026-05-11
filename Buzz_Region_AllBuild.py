import subprocess
import sys
from pathlib import Path

try:
    import yaml
except ImportError:
    yaml = None

# ==========================================================
# Business Buzz – Region All Build (UPDATED – Option A)
#
# Option A rules
# - No Host Pack stamping. Host Packs always build full history.
#
# Outputs are written under:
#   C:\Users\EmmaSmith\OneDrive - The Horsey HR Lady\Documents\Business Buzz\BusinessBuzz_Region\Buzz_Region_Curated\
#
# Pipeline behaviour is controlled by buzz_config.yml:
#   allbuild.stop_on_failure   — halt immediately if any step fails (default: true)
#   allbuild.town_target_month — set to YYYY-MM to rebuild only that month per town
# ==========================================================

BASE = Path(__file__).resolve().parent
REGION_CURATED = BASE / "Buzz_Region_Curated"

# ── Load config ────────────────────────────────────────────────────────────────
_cfg_path = BASE / "buzz_config.yml"
_allbuild_cfg: dict = {}
if yaml and _cfg_path.exists():
    try:
        with open(_cfg_path, encoding="utf-8") as _f:
            _full_cfg = yaml.safe_load(_f) or {}
        _allbuild_cfg = _full_cfg.get("allbuild", {})
    except Exception as _exc:
        print(f"[WARN] Could not read buzz_config.yml: {_exc}. Using defaults.")
elif not yaml:
    print("[WARN] PyYAML not installed — buzz_config.yml not read. Using defaults.")
    print("[WARN] Install with: pip install pyyaml --break-system-packages")

# Town build month from config
TOWN_TARGET_MONTH: str = _allbuild_cfg.get("town_target_month", "") or ""

# Stop immediately if any step fails
STOP_ON_FAILURE: bool = bool(_allbuild_cfg.get("stop_on_failure", True))


TOWN_BUILDS = {
    "MarketHarborough": ("Buzz_Event_Dashboard_MarketHarborough", "buzz_mh_build.py"),
    "Leicester":        ("Buzz_Event_Dashboard_Leicester",        "buzz_le_build.py"),
    "Lutterworth":      ("Buzz_Event_Dashboard_Lutterworth",      "buzz_lw_build.py"),
    "Hinckley":         ("Buzz_Event_Dashboard_Hinckley",         "buzz_hi_build.py"),
    "Loughborough":     ("Buzz_Event_Dashboard_Loughborough",     "buzz_lb_build.py"),
}

REGION_SCRIPTS = [
    ("Sync ambassadors",       "Buzz_Region_SyncAmbassadors.py",  []),  # must run first — rebuilds buzz_ambassadors.xlsx from town roles CSVs
    ("Region rollup",          "Buzz_Region_Rollup.py",           []),
    ("Event Excellence",       "Buzz_Region_EventExcellence.py",  []),
    ("Sponsor intelligence",   "sponsor_intelligence.py",         []),
    ("Buzz Plus intelligence", "buzzplus_intelligence.py",        []),
    ("3-2-1 intelligence",     "buzz_321_intelligence.py",        []),  # must run before host packs
    ("Regional dashboard",     "Buzz_Region_Dashboard.py",        []),
    ("RL Monthly Combined",    "Buzz_Region_Monthly_Combined.py", []),  # merged pack — must run after Dashboard + 321
    ("RL Quarterly Dashboard",  "Buzz_Region_QuarterlyDashboard.py", []),  # current quarter, in-progress aware

    # Host Packs: FULL HISTORY, no stamping
    ("Town host packs",        "Buzz_Town_HostPack.py",           ["--town", "ALL"]),
]


def _run_step(label: str, cmd: list[str], cwd: Path) -> bool:
    print(f"\n[STEP] {label}")
    print(f"[CMD ] {cmd} (cwd={cwd})")

    try:
        result = subprocess.run(cmd, cwd=str(cwd), check=False, text=True)
    except Exception as exc:
        print(f"[FAIL] {label} – error running command: {exc}")
        return False

    if result.returncode == 0:
        print(f"[OK  ] {label} completed.")
        return True

    print(f"[FAIL] {label} returned code {result.returncode}")
    return False


def main() -> None:
    print("=" * 43)
    print("   Business Buzz – Region All Build")
    print("=" * 43)
    print(f"Base folder:        {BASE}")
    print(f"Town target month:  {TOWN_TARGET_MONTH or '(none — full run)'}")
    print(f"Stop on failure:    {STOP_ON_FAILURE}")

    REGION_CURATED.mkdir(parents=True, exist_ok=True)

    failed_steps: list[str] = []

    # 1) Town builds
    print("\n---------- TOWN BUILDS ----------")
    for town_name, (folder_name, script_name) in TOWN_BUILDS.items():
        town_folder = BASE / folder_name
        script_path = town_folder / script_name
        label = f"{town_name} – {script_name}"

        if not town_folder.exists():
            print(f"\n[STEP] {label}")
            print(f"[WARN] Town folder not found: {town_folder}")
            continue

        if not script_path.exists():
            print(f"\n[STEP] {label}")
            print(f"[WARN] Town build script not found: {script_path}")
            continue

        cmd = [sys.executable, script_name]
        if TOWN_TARGET_MONTH:
            cmd += ["--month", TOWN_TARGET_MONTH]

        ok = _run_step(label=label, cmd=cmd, cwd=town_folder)
        if not ok:
            failed_steps.append(label)
            if STOP_ON_FAILURE:
                print(f"\n[HALT] stop_on_failure=true — aborting after failed step: {label}")
                print(f"[HALT] Fix the error above and re-run AllBuild.")
                _print_summary(failed_steps, aborted=True)
                sys.exit(1)

    # 2) Regional scripts
    print("\n---------- REGIONAL SCRIPTS ----------")
    for label, script_name, extra_args in REGION_SCRIPTS:
        script_path = BASE / script_name
        if not script_path.exists():
            print(f"\n[STEP] {label}")
            print(f"[WARN] Regional script not found: {script_path}")
            continue

        cmd = [sys.executable, script_name] + (extra_args or [])
        ok = _run_step(label=label, cmd=cmd, cwd=BASE)
        if not ok:
            failed_steps.append(label)
            if STOP_ON_FAILURE:
                print(f"\n[HALT] stop_on_failure=true — aborting after failed step: {label}")
                print(f"[HALT] Fix the error above and re-run AllBuild.")
                _print_summary(failed_steps, aborted=True)
                sys.exit(1)

    _print_summary(failed_steps, aborted=False)


def _print_summary(failed_steps: list[str], aborted: bool) -> None:
    print("\n" + "=" * 43)
    if aborted:
        print("   AllBuild ABORTED.")
    elif failed_steps:
        print("   AllBuild complete — WITH ERRORS.")
    else:
        print("   AllBuild sequence complete.")
    print(f"   Check {REGION_CURATED} for outputs.")
    if failed_steps:
        print(f"\n   Failed steps ({len(failed_steps)}):")
        for s in failed_steps:
            print(f"     ✗  {s}")
    print("=" * 43)


if __name__ == "__main__":
    main()
