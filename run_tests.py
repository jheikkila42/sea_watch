#!/usr/bin/env python3
"""CLI runner for project tests.

Examples:
  python run_tests.py all
  python run_tests.py stcw-rest
  python run_tests.py stcw-split
  python run_tests.py stcw-long-rest
  python run_tests.py stcw-status
  python run_tests.py daily-hours
  python run_tests.py special-ops
"""

import argparse
import subprocess
import sys

GROUP_TO_ARGS = {
    "all": ["-q"],
    "stcw-rest": ["-q", "-m", "stcw_rest"],
    "stcw-split": ["-q", "-m", "stcw_split"],
    "stcw-long-rest": ["-q", "-m", "stcw_long_rest"],
    "stcw-status": ["-q", "-m", "stcw_status"],
    "daily-hours": ["-q", "-m", "daily_hours"],
    "special-ops": ["-q", "-m", "special_ops"],
}


def main() -> int:
    parser = argparse.ArgumentParser(description="Run sea_watch tests by group")
    parser.add_argument(
        "group",
        choices=GROUP_TO_ARGS.keys(),
        help="Test group to run (or 'all').",
    )
    args = parser.parse_args()

    cmd = [sys.executable, "-m", "pytest", *GROUP_TO_ARGS[args.group]]
    print("Running:", " ".join(cmd))
    return subprocess.call(cmd)


if __name__ == "__main__":
    raise SystemExit(main())
