#!/usr/bin/env python3
"""Convert PPT/PPTX/DOCX to PDF via Microsoft Office on macOS."""
from __future__ import annotations

import subprocess
import sys
import time
from pathlib import Path


def applescript(script: str) -> subprocess.CompletedProcess[str]:
    return subprocess.run(
        ["osascript", "-e", script],
        capture_output=True,
        text=True,
    )


def convert_powerpoint(src: Path, dst: Path) -> None:
    dst.parent.mkdir(parents=True, exist_ok=True)
    if dst.exists():
        dst.unlink()
    src_posix = str(src.resolve())
    dst_posix = str(dst.resolve())
    script = f'''
tell application "Microsoft PowerPoint"
    activate
    open POSIX file "{src_posix}"
    delay 1
    set thePres to active presentation
    save thePres in POSIX file "{dst_posix}" as save as PDF
    close thePres saving no
end tell
'''
    r = applescript(script)
    if r.returncode != 0 or not dst.exists():
        raise RuntimeError(
            f"PowerPoint failed for {src.name}: {r.stderr.strip() or r.stdout.strip() or 'no output'}"
        )


def convert_word(src: Path, dst: Path) -> None:
    dst.parent.mkdir(parents=True, exist_ok=True)
    if dst.exists():
        dst.unlink()
    src_posix = str(src.resolve())
    dst_posix = str(dst.resolve())
    script = f'''
tell application "Microsoft Word"
    activate
    set theDoc to open POSIX file "{src_posix}"
    delay 1
    save as theDoc file name POSIX file "{dst_posix}" file format format PDF
    close theDoc saving no
end tell
'''
    r = applescript(script)
    if r.returncode != 0 or not dst.exists():
        raise RuntimeError(
            f"Word failed for {src.name}: {r.stderr.strip() or r.stdout.strip() or 'no output'}"
        )


if __name__ == "__main__":
    src = Path(sys.argv[1])
    dst = Path(sys.argv[2])
    ext = src.suffix.lower()
    t0 = time.time()
    if ext in {".ppt", ".pptx"}:
        convert_powerpoint(src, dst)
    elif ext in {".doc", ".docx"}:
        convert_word(src, dst)
    else:
        raise SystemExit(f"Unsupported: {src}")
    print(f"OK {src.name} -> {dst} ({dst.stat().st_size} bytes, {time.time()-t0:.1f}s)")
