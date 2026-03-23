import subprocess
import sys
from pathlib import Path

# ==========================================================
# Business Buzz – Region All Build (UPDATED – Option A)
#
# Option A rules
# - No Host Pack stamping. Host Packs always build full history.
#
# Outputs are written under:
#   C:\BusinessBuzz_Region\Buzz_Region_Curated\
# ==========================================================

BASE = Path(__file__).resolve().parent
REGION_CURATED = BASE / "Buzz_Region_Curated"

# Town build month:
# - Set to "YYYY-MM" to rebuild just that month per town (fast).
# - Leave blank to run town scripts with no --month argument.
TOWN_TARGET_MONTH = ""  # e.g. "2026-01" or ""

TOWN_BUILDS = {
    "MarketHarborough": ("Buzz_Event_Dashboard_MarketHarborough", "buzz_mh_build.py"),
    "Leicester":        ("Buzz_Event_Dashboard_Leicester",        "buzz_le_build.py"),
    "Lutterworth":      ("Buzz_Event_Dashboard_Lutterworth",      "buzz_lw_build.py"),
    "Hinckley":         ("Buzz_Event_Dashboard_Hinckley",         "buzz_hi_build.py"),
    "Loughborough":     ("Buzz_Event_Dashboard_Loughborough",     "buzz_lb_build.py"),
}

REGION_SCRIPTS = [
    ("Region rollup",          "Buzz_Region_Rollup.py",           []),
    ("Event Excellence",       "Buzz_Region_EventExcellence.py",  []),
    ("Sponsor intelligence",   "sponsor_intelligence.py",         []),
    ("Buzz Plus intelligence", "buzzplus_intelligence.py",        []),
    ("Regional dashboard",     "Buzz_Region_Dashboard.py",        []),

    # Host Packs: FULL HISTORY, no stamping
    ("Town host packs",        "Buzz_Town_HostPack.py",           ["--town", "ALL"]),

    ("RL Monthly Pack",        "Buzz_RL_MonthlyPack.py",          []),
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
    print(f"Base folder: {BASE}")
    print(f"Town target month: {TOWN_TARGET_MONTH or '(none)'}")

    REGION_CURATED.mkdir(parents=True, exist_ok=True)

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

        _run_step(label=label, cmd=cmd, cwd=town_folder)

    # 2) Regional scripts
    print("\n---------- REGIONAL SCRIPTS ----------")
    for label, script_name, extra_args in REGION_SCRIPTS:
        script_path = BASE / script_name
        if not script_path.exists():
            print(f"\n[STEP] {label}")
            print(f"[WARN] Regional script not found: {script_path}")
            continue

        cmd = [sys.executable, script_name] + (extra_args or [])
        _run_step(label=label, cmd=cmd, cwd=BASE)

    print("\n" + "=" * 43)
    print("   AllBuild sequence complete.")
    print(f"   Check {REGION_CURATED} for outputs.")
    print("=" * 43)


if __name__ == "__main__":
    main()
