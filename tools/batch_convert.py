#!/usr/bin/env python3
"""Convert PPT/PPTX/DOCX to PDF via Microsoft Office into /tmp, then copy."""
from __future__ import annotations

import shutil
import subprocess
import time
from pathlib import Path

ROOT = Path("/Users/zfeng/Github/ECON-UNO")
SRC = Path("/Users/zfeng/Library/CloudStorage/OneDrive-Personal/Teachings/UNO")
TMP = Path("/tmp/econ-uno-pdfs")
TMP.mkdir(parents=True, exist_ok=True)

SKIP_NAMES = {
    "Macroeconomics_Stephen D_ Williamson.pdf",
    "Williamson_Macro6e_ppt_01.pptx",
}


def osascript(script: str, timeout: int = 300) -> subprocess.CompletedProcess[str]:
    return subprocess.run(["osascript", "-e", script], capture_output=True, text=True, timeout=timeout)


def quit_app(name: str) -> None:
    osascript(f'tell application "{name}" to if it is running then quit', timeout=45)
    time.sleep(1)


def ppt_to_tmp(src: Path, tmp_pdf: Path) -> None:
    if tmp_pdf.exists():
        tmp_pdf.unlink()
    wait = 2
    script = f'''
with timeout of 240 seconds
  tell application "Microsoft PowerPoint"
    activate
    open POSIX file "{src}"
    delay {wait}
    save active presentation in POSIX file "{tmp_pdf}" as save as PDF
    close active presentation saving no
  end tell
end timeout
'''
    r = osascript(script, timeout=260)
    if r.returncode != 0 or not tmp_pdf.exists() or tmp_pdf.stat().st_size < 1000:
        raise RuntimeError(r.stderr.strip() or r.stdout.strip() or "PowerPoint produced no PDF")


def word_to_tmp(src: Path, tmp_pdf: Path) -> None:
    if tmp_pdf.exists():
        tmp_pdf.unlink()
    script = f'''
with timeout of 120 seconds
  tell application "Microsoft Word"
    activate
    open POSIX file "{src}"
    delay 1.5
    set theDoc to active document
    save as theDoc file name POSIX file "{tmp_pdf}" file format format PDF
    close theDoc saving no
  end tell
end timeout
'''
    r = osascript(script, timeout=140)
    if r.returncode != 0 or not tmp_pdf.exists() or tmp_pdf.stat().st_size < 1000:
        raise RuntimeError(r.stderr.strip() or r.stdout.strip() or "Word produced no PDF")


JOBS: list[tuple[Path, Path]] = []
for p in sorted((SRC / "ECON2220/Notes").glob("Lecture*.ppt*")):
    JOBS.append((p, ROOT / "econ2220/notes" / (p.stem + ".pdf")))
for p in sorted((SRC / "ECON2220/Tests").glob("*.pdf")):
    JOBS.append((p, ROOT / "econ2220/tests" / p.name.replace(" ", "_")))
for p in sorted((SRC / "ECON3220/Notes").iterdir()):
    if p.name in SKIP_NAMES or p.name.startswith("."):
        continue
    if p.suffix.lower() in {".ppt", ".pptx"}:
        JOBS.append((p, ROOT / "econ3220/notes" / (p.stem.replace(" ", "_") + ".pdf")))
for p in sorted((SRC / "ECON4660/Notes").iterdir()):
    if p.name.startswith(".") or p.suffix.lower() not in {".ppt", ".pptx"}:
        continue
    JOBS.append((p, ROOT / "econ4660/notes" / (p.stem + ".pdf")))
for p in sorted((SRC / "ECON4660").glob("PS_*.docx")):
    JOBS.append((p, ROOT / "econ4660/homework" / (p.stem + ".pdf")))
JOBS.append((SRC / "ECON3220/Syllabus/ECON3220.docx", ROOT / "econ3220/syllabus.pdf"))
JOBS.append((SRC / "ECON4660/ECON4660.docx", ROOT / "econ4660/syllabus.pdf"))


def main() -> int:
    failed: list[tuple[str, str]] = []
    quit_app("Microsoft PowerPoint")
    quit_app("Microsoft Word")
    for i, (src, dst) in enumerate(JOBS, 1):
        if dst.exists() and dst.stat().st_size > 1000:
            print(f"[{i}/{len(JOBS)}] skip {dst.relative_to(ROOT)}", flush=True)
            continue
        print(f"[{i}/{len(JOBS)}] {src.name}", flush=True)
        t0 = time.time()
        tmp = TMP / dst.name
        try:
            ext = src.suffix.lower()
            if ext == ".pdf":
                dst.parent.mkdir(parents=True, exist_ok=True)
                shutil.copy2(src, dst)
            elif ext in {".ppt", ".pptx"}:
                try:
                    ppt_to_tmp(src, tmp)
                except Exception as e1:
                    print(f"    retry after quit ({e1})", flush=True)
                    quit_app("Microsoft PowerPoint")
                    time.sleep(2)
                    ppt_to_tmp(src, tmp)
                dst.parent.mkdir(parents=True, exist_ok=True)
                shutil.copy2(tmp, dst)
            else:
                try:
                    word_to_tmp(src, tmp)
                except Exception as e1:
                    print(f"    retry after quit ({e1})", flush=True)
                    quit_app("Microsoft Word")
                    time.sleep(2)
                    word_to_tmp(src, tmp)
                dst.parent.mkdir(parents=True, exist_ok=True)
                shutil.copy2(tmp, dst)
            print(f"    OK {dst.stat().st_size/1e6:.1f} MB in {time.time()-t0:.1f}s", flush=True)
        except Exception as e:
            print(f"    FAIL {e}", flush=True)
            failed.append((src.name, str(e)))
            quit_app("Microsoft PowerPoint")
            quit_app("Microsoft Word")
    print(f"--- failed {len(failed)}/{len(JOBS)} ---")
    for name, err in failed:
        print(f"  {name}: {err}")
    return 1 if failed else 0


if __name__ == "__main__":
    raise SystemExit(main())
