#!/usr/bin/env python3
"""
Cut the CALM exports down to one faculty's courses, for the curriculum dashboard.

Reads the three CALM exports from a CALM-dashboard-YYYYMMDD folder:
    Student Pathways by Units.xlsx
    Units by Assessment Tasks.xlsx
    Course Data.xlsx
and writes the same three files as .csv into a sub-folder named after the
faculty (e.g. CALM-dashboard-20261027/SEBE/), keeping only:
    - courses in the chosen faculty (from Course Data)
    - every pathway of those courses (all pathway types: Core, MJ, ST, SP …)
    - the assessment tasks of every unit in those pathways
Drop the three .csv files into the dashboard as usual.

Usage:
    python3 extract-faculty.py                      # SEBE, newest CALM-dashboard-* folder
    python3 extract-faculty.py SEBE path/to/folder
Uses only the Python standard library (no openpyxl needed).
"""
import csv
import re
import sys
import zipfile
import xml.etree.ElementTree as ET
from pathlib import Path

FACULTIES = {
    'SEBE':   'Faculty Of Science, Engineering And Built Environment',
    'AE':     'Faculty Of Arts And Education',
    'BL':     'Faculty Of Business And Law',
    'HEALTH': 'Faculty Of Health',
}
PATHWAYS_FILE = 'Student Pathways by Units'
TASKS_FILE = 'Units by Assessment Tasks'
COURSES_FILE = 'Course Data'
FOOTER = 'no filters applied'   # CALM appends this row to every export

NS = {'m': 'http://schemas.openxmlformats.org/spreadsheetml/2006/main',
      'r': 'http://schemas.openxmlformats.org/officeDocument/2006/relationships'}


def col_index(ref):
    n = 0
    for ch in re.match(r'[A-Z]+', ref).group():
        n = n * 26 + ord(ch) - 64
    return n - 1


def read_xlsx(path):
    """Return rows (lists of strings) from the 'export' sheet, or the first sheet."""
    z = zipfile.ZipFile(path)
    strings = []
    if 'xl/sharedStrings.xml' in z.namelist():
        for si in ET.fromstring(z.read('xl/sharedStrings.xml')).findall('m:si', NS):
            strings.append(''.join(t.text or '' for t in si.iter('{%s}t' % NS['m'])))
    wb = ET.fromstring(z.read('xl/workbook.xml'))
    rels = {r.get('Id'): r.get('Target') for r in ET.fromstring(z.read('xl/_rels/workbook.xml.rels'))}
    sheets = wb.find('m:sheets', NS)
    sheet = next((s for s in sheets if s.get('name', '').lower() == 'export'), sheets[0])
    target = rels[sheet.get('{%s}id' % NS['r'])].lstrip('/')
    if not target.startswith('xl/'):
        target = 'xl/' + target
    rows = []
    for r in ET.fromstring(z.read(target)).iter('{%s}row' % NS['m']):
        cells, pos = {}, -1
        for c in r.findall('m:c', NS):
            pos = col_index(c.get('r')) if c.get('r') else pos + 1
            v, inline = c.find('m:v', NS), c.find('m:is', NS)
            if c.get('t') == 's' and v is not None:
                val = strings[int(v.text)]
            elif c.get('t') == 'inlineStr' and inline is not None:
                val = ''.join(t.text or '' for t in inline.iter('{%s}t' % NS['m']))
            else:
                val = v.text if v is not None else ''
            cells[pos] = val or ''
        rows.append([cells.get(i, '') for i in range(max(cells) + 1)] if cells else [])
    return rows


def load(folder, name):
    path = folder / f'{name}.xlsx'
    if not path.exists():
        sys.exit(f'Missing {path.name} in {folder}')
    rows = [r for r in read_xlsx(path) if any(r)]
    header, body = rows[0], rows[1:]
    body = [r + [''] * (len(header) - len(r)) for r in body
            if r and r[0].strip().lower() != FOOTER]
    return header, body


def course_code(pathway):
    m = re.match(r'[A-Za-z0-9]+', pathway or '')
    return m.group().upper() if m else ''


def write_csv(path, header, rows):
    with open(path, 'w', newline='', encoding='utf-8') as f:
        w = csv.writer(f)
        w.writerow(header)
        w.writerows(rows)


def main():
    here = Path(__file__).resolve().parent
    key = (sys.argv[1] if len(sys.argv) > 1 else 'SEBE').upper()
    if key not in FACULTIES:
        sys.exit(f'Unknown faculty {key!r}. Choose one of: {", ".join(FACULTIES)}')
    if len(sys.argv) > 2:
        folder = Path(sys.argv[2]).expanduser().resolve()
    else:
        found = sorted(p for p in here.glob('CALM-dashboard-*') if p.is_dir())
        if not found:
            sys.exit(f'No CALM-dashboard-* folder found next to {Path(__file__).name}')
        folder = found[-1]
    return extract(folder, key)


def extract(folder, key='SEBE'):
    """Write <folder>/<key>/*.csv for one faculty; returns the output folder."""
    faculty = FACULTIES[key]
    print(f'Source:  {folder}')
    print(f'Faculty: {faculty}')

    c_head, c_rows = load(folder, COURSES_FILE)
    c_code, c_fac = c_head.index('Course'), c_head.index('Faculty')
    courses = [r for r in c_rows if r[c_fac].strip() == faculty]
    codes = {r[c_code].strip().upper() for r in courses}

    p_head, p_rows = load(folder, PATHWAYS_FILE)
    p_path, p_unit = p_head.index('Student Pathway'), p_head.index('Unit Code')
    pathways = [r for r in p_rows if course_code(r[p_path]) in codes]
    units = {r[p_unit].strip() for r in pathways}

    t_head, t_rows = load(folder, TASKS_FILE)
    t_unit = t_head.index('Unit Code')
    tasks = [r for r in t_rows if r[t_unit].strip() in units]

    out = folder / key
    out.mkdir(exist_ok=True)
    write_csv(out / f'{COURSES_FILE}.csv', c_head, courses)
    write_csv(out / f'{PATHWAYS_FILE}.csv', p_head, pathways)
    write_csv(out / f'{TASKS_FILE}.csv', t_head, tasks)

    print(f'\nKept {len(codes)} of {len(c_rows)} courses, '
          f'{len({r[p_path] for r in pathways})} pathways, {len(units)} units, '
          f'{len(tasks)} of {len(t_rows)} assessment tasks.')
    print(f'Wrote 3 files to {out}')
    return out


if __name__ == '__main__':
    main()
