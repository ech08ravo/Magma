#!/usr/bin/env python3
"""
Run after adding a new CALM-dashboard-YYYYMMDD folder:

    python3 update-dashboard-data.py            # SEBE
    python3 update-dashboard-data.py --no-open  # don't open Finder at the end

It
  1. finds every CALM-dashboard-YYYYMMDD folder in the iCloud and repo Magma
     folders, and copies new or changed files across so both sides match
     (old/ is never touched);
  2. checks each folder has the three CALM exports;
  3. runs the SEBE extract (extract-faculty.py) for any folder whose SEBE/
     output is missing or older than its exports, and (re)builds
     SEBE-statistics.xlsx beside it (make-statistics.py);
  4. prints the SEBE folder to drop into the dashboard and opens it in Finder.
Uses only the Python standard library.
"""
import argparse
import importlib.util
import re
import shutil
import subprocess
import sys
from pathlib import Path

ICLOUD = Path.home() / 'Library/Mobile Documents/com~apple~CloudDocs/Debusseycat/Magma'
REPO = Path.home() / 'GitHubRepos/Magma'
FOLDER_RE = re.compile(r'^CALM-dashboard-(\d{8})$')
FACULTY = 'SEBE'
SKIP = {'.DS_Store'}

sys.dont_write_bytecode = True   # don't leave __pycache__ in the Magma folders
here = Path(__file__).resolve().parent
spec = importlib.util.spec_from_file_location('extract_faculty', here / 'extract-faculty.py')
xf = importlib.util.module_from_spec(spec)
spec.loader.exec_module(xf)
REQUIRED = [f'{name}.xlsx' for name in (xf.PATHWAYS_FILE, xf.TASKS_FILE, xf.COURSES_FILE)]
spec = importlib.util.spec_from_file_location('make_statistics', here / 'make-statistics.py')
ms = importlib.util.module_from_spec(spec)
spec.loader.exec_module(ms)


def calm_folders(root):
    if not root.is_dir():
        return {}
    return {p.name: p for p in root.iterdir() if p.is_dir() and FOLDER_RE.match(p.name)}


def mirror(src, dst):
    """Copy files in src that are missing from dst or newer than dst's copy. Returns count."""
    copied = 0
    for f in src.rglob('*'):
        if not f.is_file() or f.name in SKIP:
            continue
        target = dst / f.relative_to(src)
        if not target.exists() or f.stat().st_mtime > target.stat().st_mtime + 1:
            target.parent.mkdir(parents=True, exist_ok=True)
            shutil.copy2(f, target)
            copied += 1
    return copied


def validate_export(path, required_headers, minimum_rows=2):
    """Return (ok, detail) for a CALM export workbook. A file is usable only if it
    is readable, non-empty, and contains the required schema fields."""
    if not path.exists():
        return False, f'missing {path.name}'
    try:
        rows = xf.read_xlsx(path)
    except Exception as exc:  # pragma: no cover - defensive validation path
        return False, f'could not parse {path.name}: {exc}'

    rows = [r for r in rows if any(str(cell).strip() for cell in r)]
    if len(rows) < minimum_rows:
        return False, f'{path.name} is empty or header-only'

    header = [str(cell).strip() for cell in rows[0]]
    missing = [h for h in required_headers if h not in header]
    if missing:
        return False, f'{path.name} missing headers: {", ".join(missing)}'

    data_rows = [r for r in rows[1:] if any(str(cell).strip() for cell in r)]
    if not data_rows:
        return False, f'{path.name} has no data rows'

    return True, None


def validate_required_exports(folder):
    results = {}
    for name in (xf.PATHWAYS_FILE, xf.TASKS_FILE, xf.COURSES_FILE):
        path = folder / f'{name}.xlsx'
        if name == xf.PATHWAYS_FILE:
            required = ['Student Pathway', 'Unit Code']
            ok, detail = validate_export(path, required)
        elif name == xf.TASKS_FILE:
            required = ['Unit Code', 'Task Number', 'Assessment Title']
            ok, detail = validate_export(path, required)
        else:
            required = ['Course', 'Faculty', 'Course Name']
            ok, detail = validate_export(path, required)
        results[name] = (ok, detail)
    return results


def needs_extract(folder):
    out = folder / FACULTY
    outputs = [out / f'{Path(n).stem}.csv' for n in REQUIRED]
    if not all(o.exists() for o in outputs):
        return True
    newest_input = max((folder / n).stat().st_mtime for n in REQUIRED)
    return min(o.stat().st_mtime for o in outputs) < newest_input


def needs_stats(folder):
    stats = folder / f'{FACULTY}-statistics.xlsx'
    if not stats.exists():
        return True
    inputs = [folder / FACULTY / f'{Path(n).stem}.csv' for n in REQUIRED]
    return stats.stat().st_mtime < max(i.stat().st_mtime for i in inputs)


def latest_valid_folder(folders):
    candidates = []
    for _, folder in folders:
        missing = [n for n in REQUIRED if not (folder / n).exists()]
        if missing:
            continue
        validation = validate_required_exports(folder)
        if any(not ok for ok, _ in validation.values()):
            continue
        if needs_extract(folder) or needs_stats(folder):
            continue
        candidates.append(folder)
    return max(candidates, default=None, key=lambda p: p.name)


def archive_current_sebe(folder, faculty=FACULTY):
    """Move the existing current folder aside as a dated snapshot before rebuilding it."""
    current = folder / faculty
    if not current.is_dir():
        return None
    date_stamp = folder.name.replace('CALM-dashboard-', '') if folder.name.startswith('CALM-dashboard-') else 'snapshot'
    archive = folder / f'{faculty}-{date_stamp}'
    idx = 1
    while archive.exists():
        archive = folder / f'{faculty}-{date_stamp}-{idx}'
        idx += 1
    shutil.move(str(current), str(archive))
    return archive


def main():
    parser = argparse.ArgumentParser(description='Sync CALM exports and build the SEBE dashboard data.')
    parser.add_argument('--allow-missing', action='store_true',
                        help='Warn instead of skipping folders missing required CALM export files; keeps scanning the newest complete folder.')
    parser.add_argument('--no-open', action='store_true', help='Do not open the SEBE folder in Finder at the end.')
    parser.add_argument('--archive-sebe', action='store_true',
                        help='Keep the previous active SEBE folder as a dated snapshot before rebuilding it (e.g. SEBE-20261102).')
    args = parser.parse_args()

    roots = [r for r in (ICLOUD, REPO) if r.is_dir()]
    if not roots:
        sys.exit(f'Neither Magma folder found:\n  {ICLOUD}\n  {REPO}')

    # 1. Sync CALM folders between the two Magma copies (both directions, newest wins)
    if len(roots) == 2:
        for a, b in ((ICLOUD, REPO), (REPO, ICLOUD)):
            for name, folder in sorted(calm_folders(a).items()):
                n = mirror(folder, b / name)
                if n:
                    print(f'Synced {n} file(s): {a.parent.name}/{a.name}/{name} → {b.parent.name}/{b.name}')
    else:
        print(f'Only one Magma folder found ({roots[0]}); skipping sync.')

    # 2–3. Check and extract, working in the first available root
    main_root = roots[0]
    folders = sorted(calm_folders(main_root).items())
    if not folders:
        sys.exit(f'No CALM-dashboard-YYYYMMDD folders in {main_root}')

    for name, folder in folders:
        missing = [n for n in REQUIRED if not (folder / n).exists()]
        if missing:
            if args.allow_missing:
                print(f'\n{name}: missing {", ".join(missing)}; continuing in permissive mode')
            else:
                print(f'\n{name}: SKIPPED, missing {", ".join(missing)}')
                continue

        validation = validate_required_exports(folder)
        invalid = [(file_name, detail) for file_name, (ok, detail) in validation.items() if not ok]
        if invalid:
            print(f'\n{name}: INVALID export set')
            for file_name, detail in invalid:
                print(f'  - {file_name}: {detail}')
            continue

        if not needs_extract(folder) and not needs_stats(folder):
            print(f'\n{name}: {FACULTY} files and statistics up to date')
            continue
        if needs_extract(folder):
            if args.archive_sebe:
                archived = archive_current_sebe(folder, FACULTY)
                if archived:
                    print(f'  archived previous {FACULTY} folder to {archived.name}')
            print(f'\n{name}: extracting {FACULTY}…')
            rc = subprocess.run(['python3', str(here / 'extract-faculty.py'), FACULTY, str(folder)], capture_output=True, text=True)
            if rc.returncode != 0:
                print(f'  extract-faculty.py failed for {name}:')
                if rc.stderr.strip():
                    print(rc.stderr.strip())
                if rc.stdout.strip():
                    print(rc.stdout.strip())
                continue
        else:
            print(f'\n{name}:')
        print(f'  building {FACULTY}-statistics.xlsx…')
        rc = subprocess.run(['python3', str(here / 'make-statistics.py'), str(folder), FACULTY], capture_output=True, text=True)
        if rc.returncode != 0:
            print(f'  make-statistics.py failed for {name}:')
            if rc.stderr.strip():
                print(rc.stderr.strip())
            if rc.stdout.strip():
                print(rc.stdout.strip())
            continue
        if len(roots) == 2:
            other = (REPO if main_root == ICLOUD else ICLOUD) / name
            mirror(folder, other)
            print(f'  copied {FACULTY}/ and {FACULTY}-statistics.xlsx to {other.parent}')

    # 4. Point at the newest folder whose extract is complete, valid and current
    latest = latest_valid_folder(folders)
    if latest:
        latest_out = latest / FACULTY
        print(f'\nDrop these 3 files into the dashboard:\n  {latest_out}')
        print(f'Statistics workbook:\n  {latest / (FACULTY + "-statistics.xlsx")}')
        if not args.no_open and '--no-open' not in sys.argv:
            subprocess.run(['open', str(latest_out)])
    else:
        print('\nNo usable dashboard-ready folder was found: check the export files for missing or invalid CALM data.')


if __name__ == '__main__':
    main()
