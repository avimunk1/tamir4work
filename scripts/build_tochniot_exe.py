#!/usr/bin/env python3
"""
Build Windows exe for tochniot_le_mosad.

Run on Windows:  uv run python scripts/build_tochniot_exe.py

Output: dist/TochniotLeMosad.exe (single file, no console window)
"""

import subprocess
import sys
from pathlib import Path

PROJECT_ROOT = Path(__file__).resolve().parent.parent
SPEC_FILE = PROJECT_ROOT / "TochniotLeMosad_win.spec"
SCRIPT = PROJECT_ROOT / "src" / "excel_extractor" / "tochniot_le_mosad.py"


def main() -> None:
    use_console = "--console" in sys.argv

    if sys.platform != "win32":
        print("Note: For a Windows .exe, run this script on a Windows machine.")
        print("On macOS/Linux, this builds a native executable (not Windows).")
        print()

    # On Windows, use the spec file (no BUNDLE). Else use direct command.
    if sys.platform == "win32" and SPEC_FILE.exists():
        cmd = [
            sys.executable,
            "-m",
            "PyInstaller",
            "--clean",
            str(SPEC_FILE),
        ]
    else:
        cmd = [
            sys.executable,
            "-m",
            "PyInstaller",
            "--onefile",
            "--name",
            "TochniotLeMosad",
            "--clean",
        ]
        if not use_console:
            cmd.append("--noconsole")
        cmd.append(str(SCRIPT))

    print("Running:", " ".join(cmd))
    subprocess.run(cmd, cwd=PROJECT_ROOT, check=True)
    print()
    print("Done! Output: dist/TochniotLeMosad.exe" if sys.platform == "win32" else "Done! Output: dist/TochniotLeMosad")


if __name__ == "__main__":
    main()
