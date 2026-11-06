#!/usr/bin/env python3
"""
Build a statistics workbook from a faculty extract (the SEBE/*.csv files).

    python3 make-statistics.py                       # newest CALM-dashboard-*/SEBE
    python3 make-statistics.py path/to/CALM-dashboard-YYYYMMDD [SEBE]

Writes CALM-dashboard-YYYYMMDD/SEBE-statistics.xlsx (beside the SEBE folder,
not in it, so it never gets dropped into the dashboard by mistake).

Every unit appears ONCE in the Units sheet. The rules match the dashboard
(curriculum-dashboard-v5.html): GLO cell vocabulary, AoL task rule, hurdle
rule, year from unit code, 2026 teaching-period lookup (marked *), and course
blocks (Core Only + "<X> Only"; Core + X = Core ∪ X).
Uses only the Python standard library.
"""
import csv
import html
import re
import sys
import zipfile
from collections import defaultdict
from datetime import date
from pathlib import Path
from xml.sax.saxutils import escape

HERE = Path(__file__).resolve().parent
GLO_NAMES = ['Discipline-specific knowledge', 'Communication', 'Digital literacy', 'Critical thinking',
             'Problem solving', 'Self-management', 'Teamwork', 'Global citizenship']
NOT_AOL_TASK_TYPES = {'', 'na', 'n/a', 'not applicable', 'tbc', 'no', 'not aol task'}
YEAR_LABELS = {0: 'Non-major', 1: 'Y1', 2: 'Y2', 3: 'Y3', 4: 'Hons', 7: 'PG', 8: 'MRes', 9: 'PhD', -1: '?'}
YEAR_ORDER = {0: -1, 7: 50, 8: 60, 9: 70, -1: 99}
STATUS_LABELS = {'hurdle': 'Yes', 'nonhurdle': 'Non-hurdle Yes', 'no': 'No', 'notassessed': 'Not Assessed',
                 'dataerror': 'Data Error', 'tbc': 'TBC', 'missing': 'Form not received', 'blank': ''}


# ── Rules (mirroring the dashboard) ─────────────────────────────────────────
def text(v):
    return html.unescape(str(v or '')).replace('\xa0', ' ').strip()


def glo_state(v):
    t = text(v).lower()
    if not t:
        return 'blank'
    if t == 'tbc' or 'to be confirmed' in t:
        return 'tbc'
    if 'data error' in t:
        return 'dataerror'
    if 'form not received' in t:
        return 'missing'
    if 'not assessed' in t:
        return 'notassessed'
    if t == 'no':
        return 'no'
    if 'non-hurdle' in t or 'non hurdle' in t:
        return 'nonhurdle'
    if 'yes' in t:
        return 'hurdle'
    return 'blank'


def assessed(state):
    return state in ('hurdle', 'nonhurdle')


def is_hurdle(v):
    t = text(v).lower()
    return bool(t) and 'non' not in t and 'yes' in t


def task_tier(v):
    t = text(v).lower()
    if t in NOT_AOL_TASK_TYPES:
        return 'developing'
    if re.search(r'\bclo\b', t) and 'aol' not in t:
        return 'clo'
    return 'aol'


def year_of(code):
    m = re.search(r'[A-Za-z]+\s*(\d{3,})', code or '')
    if not m:
        return -1
    d = int(m.group(1)[0])
    return d if d in (0, 1, 2, 3, 4, 7, 8, 9) else -1


def year_key(y):
    return YEAR_ORDER.get(y, y)


def weighting_percent(v):
    t = text(v)
    m = re.search(r'(\d+(?:\.\d+)?)\s*(?:%|percent|per cent)', t, re.I) or re.fullmatch(r'(\d+(?:\.\d+)?)', t)
    return float(m.group(1)) if m else None


def short_at(n):
    m = re.search(r'AT-?\s*(\d+)\s*$', n or '', re.I)
    return f'AT{m.group(1)}' if m else (n or '')


def course_code(pathway):
    m = re.match(r'[A-Za-z0-9]+', pathway or '')
    return m.group().upper() if m else ''


def pathway_suffix(pathway):
    m = re.search(r'[_\-\s/]', pathway or '')
    return pathway[m.start() + 1:].strip() if m else pathway


def teaching_periods():
    """The dashboard's built-in Apr 2026 lookup (APR2026_TEACHING_PERIODS)."""
    page = HERE / 'curriculum-dashboard-v5.html'
    if not page.exists():
        return {}
    m = re.search(r'const APR2026_TEACHING_PERIODS = \{(.*?)\};', page.read_text(encoding='utf-8'), re.S)
    return dict(re.findall(r'(\w+):"([^"]*)"', m.group(1))) if m else {}


# ── Minimal .xlsx writer (inline strings, bold header, frozen row, filter) ──
_BAD_XML = re.compile('[\x00-\x08\x0b\x0c\x0e-\x1f]')


def _col(i):
    s = ''
    i += 1
    while i:
        i, r = divmod(i - 1, 26)
        s = chr(65 + r) + s
    return s


def _cell(ref, v, style=0):
    st = f' s="{style}"' if style else ''
    if v is None or v == '':
        return f'<c r="{ref}"{st}/>' if style else ''
    if isinstance(v, bool):
        v = 'Y' if v else 'N'
    if isinstance(v, (int, float)):
        return f'<c r="{ref}"{st}><v>{v}</v></c>'
    s = escape(_BAD_XML.sub('', str(v)))[:32000]
    return f'<c r="{ref}" t="inlineStr"{st}><is><t xml:space="preserve">{s}</t></is></c>'


def _sheet_xml(rows, widths, header_rows):
    ncols = max((len(r) for r in rows), default=1)
    cols = ''.join(f'<col min="{i+1}" max="{i+1}" width="{w}" customWidth="1"/>' for i, w in enumerate(widths))
    body = []
    for ri, row in enumerate(rows, 1):
        style = 1 if ri <= header_rows and header_rows else 0
        cells = ''.join(_cell(f'{_col(ci)}{ri}', v, style if v not in (None, '') else 0) for ci, v in enumerate(row))
        body.append(f'<row r="{ri}">{cells}</row>')
    pane = (f'<sheetViews><sheetView workbookViewId="0"><pane ySplit="{header_rows}" topLeftCell="A{header_rows+1}" '
            f'activePane="bottomLeft" state="frozen"/></sheetView></sheetViews>') if header_rows else ''
    filt = f'<autoFilter ref="A{header_rows}:{_col(ncols-1)}{len(rows)}"/>' if header_rows and len(rows) > header_rows else ''
    return ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
            '<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">'
            f'{pane}<cols>{cols}</cols><sheetData>{"".join(body)}</sheetData>{filt}</worksheet>')


def write_xlsx(path, sheets):
    """sheets: list of (name, rows, widths, header_rows)."""
    ct = ''.join(f'<Override PartName="/xl/worksheets/sheet{i}.xml" '
                 'ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>'
                 for i in range(1, len(sheets) + 1))
    with zipfile.ZipFile(path, 'w', zipfile.ZIP_DEFLATED) as z:
        z.writestr('[Content_Types].xml',
                   '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
                   '<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">'
                   '<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>'
                   '<Default Extension="xml" ContentType="application/xml"/>'
                   '<Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/>'
                   '<Override PartName="/xl/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml"/>'
                   f'{ct}</Types>')
        z.writestr('_rels/.rels',
                   '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
                   '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'
                   '<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="xl/workbook.xml"/>'
                   '</Relationships>')
        z.writestr('xl/workbook.xml',
                   '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
                   '<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" '
                   'xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><sheets>'
                   + ''.join(f'<sheet name="{escape(n)}" sheetId="{i}" r:id="rId{i}"/>' for i, (n, *_ ) in enumerate(sheets, 1))
                   + '</sheets></workbook>')
        z.writestr('xl/_rels/workbook.xml.rels',
                   '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
                   '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'
                   + ''.join(f'<Relationship Id="rId{i}" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet{i}.xml"/>'
                             for i in range(1, len(sheets) + 1))
                   + f'<Relationship Id="rId{len(sheets)+1}" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/>'
                   '</Relationships>')
        z.writestr('xl/styles.xml',
                   '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
                   '<styleSheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">'
                   '<fonts count="2"><font><sz val="11"/><name val="Calibri"/></font><font><b/><sz val="11"/><name val="Calibri"/></font></fonts>'
                   '<fills count="3"><fill><patternFill patternType="none"/></fill><fill><patternFill patternType="gray125"/></fill>'
                   '<fill><patternFill patternType="solid"><fgColor rgb="FFEDEDED"/></patternFill></fill></fills>'
                   '<borders count="1"><border><left/><right/><top/><bottom/><diagonal/></border></borders>'
                   '<cellStyleXfs count="1"><xf numFmtId="0" fontId="0" fillId="0" borderId="0"/></cellStyleXfs>'
                   '<cellXfs count="2"><xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0"/>'
                   '<xf numFmtId="0" fontId="1" fillId="2" borderId="0" xfId="0" applyFont="1" applyFill="1"/></cellXfs>'
                   '</styleSheet>')
        for i, (_, rows, widths, header_rows) in enumerate(sheets, 1):
            z.writestr(f'xl/worksheets/sheet{i}.xml', _sheet_xml(rows, widths, header_rows))


# ── Data model ───────────────────────────────────────────────────────────────
def read_csv(path):
    with open(path, encoding='utf-8', newline='') as f:
        return list(csv.DictReader(f))


def glo_col(headers, i):
    return next((h for h in headers if re.match(rf'glo\s*{i}\s*/\s*clo\s*{i}', h, re.I)), None)


def load(extract):
    paths = read_csv(extract / 'Student Pathways by Units.csv')
    tasks_raw = read_csv(extract / 'Units by Assessment Tasks.csv')
    courses_raw = read_csv(extract / 'Course Data.csv')
    periods = teaching_periods()

    courses = {r['Course'].strip().upper(): r for r in courses_raw if r.get('Course')}
    p_gcols = [glo_col(list(paths[0].keys()), i) for i in range(1, 9)] if paths else []
    t_gcols = [glo_col(list(tasks_raw[0].keys()), i) for i in range(1, 9)] if tasks_raw else []

    tasks = defaultdict(list)
    for r in tasks_raw:
        code = r['Unit Code'].strip()
        if not code or code.lower() == 'no filters applied':
            continue
        states = [glo_state(r.get(c)) if c else 'blank' for c in t_gcols]
        tasks[code].append({
            'number': text(r.get('Task Number')), 'title': text(r.get('Assessment Title')),
            'type': text(r.get('AoL Task Type')), 'tier': task_tier(r.get('AoL Task Type')),
            'hurdle_raw': text(r.get('Hurdle')), 'hurdle': is_hurdle(r.get('Hurdle')),
            'weighting': text(r.get('Weighting Information')), 'weight_pct': weighting_percent(r.get('Weighting Information')),
            'output': text(r.get('Student Output')), 'mode': text(r.get('Mode')),
            'participation': text(r.get('Participation')), 'states': states,
        })

    units, memberships = {}, []
    for r in paths:
        code, pw = r['Unit Code'].strip(), r['Student Pathway'].strip()
        if not code:
            continue
        if code not in units:
            units[code] = {'code': code, 'year': year_of(code), 'states': [glo_state(r.get(c)) if c else 'blank' for c in p_gcols],
                           'pathways': set(), 'tasks': tasks.get(code, [])}
        units[code]['pathways'].add(pw)
        memberships.append((code, pw))
    for code, ts in tasks.items():   # units only in the tasks file
        units.setdefault(code, {'code': code, 'year': year_of(code), 'states': ['blank'] * 8, 'pathways': set(), 'tasks': ts})

    for u in units.values():
        ts = u['tasks']
        u['tier'] = 'aol' if any(t['tier'] == 'aol' for t in ts) else 'clo' if any(t['tier'] == 'clo' for t in ts) else 'developing'
        u['any_hurdle'] = any(t['hurdle'] for t in ts)
        if code_period := periods.get(u['code']):
            u['period'], u['period_source'] = code_period + '*', 'Apr 2026 export'
        else:
            u['period'], u['period_source'] = 'T0', 'None (T0)'
        u['courses'] = sorted({course_code(p) for p in u['pathways']})
    return units, memberships, courses


def blocks_by_course(units):
    pathways = sorted({p for u in units.values() for p in u['pathways']})
    by_course = defaultdict(list)
    for p in pathways:
        by_course[course_code(p)].append(p)
    out = {}
    for c, ps in by_course.items():
        core = [p for p in ps if pathway_suffix(p).lower() == 'core only']
        only = [p for p in ps if re.search(r'\sonly$', pathway_suffix(p), re.I) and p not in core]
        chosen = core + only if (core or only) else ps
        out[c] = [(p, 'Core' if p in core else re.sub(r'\s+only$', '', pathway_suffix(p), flags=re.I)) for p in chosen]
    return out


def coverage(unit_list):
    """Per GLO: (assessed, assured, aol) task counts across the units' tasks."""
    res = []
    for g in range(8):
        a = s = l = 0
        for u in unit_list:
            for t in u['tasks']:
                st = t['states'][g]
                if assessed(st):
                    a += 1
                    s += st == 'hurdle'
                    l += t['tier'] == 'aol'
        res.append((a, s, l))
    return res


def build(calm_folder, key='SEBE'):
    extract = calm_folder / key
    units, memberships, courses = load(extract)
    blocks = blocks_by_course(units)
    order = sorted(units.values(), key=lambda u: (year_key(u['year']), u['code']))
    years = sorted({u['year'] for u in order}, key=year_key)
    gl = [f'GLO{i}' for i in range(1, 9)]

    # README
    readme = [
        [f'{key} statistics — generated {date.today().isoformat()} from {calm_folder.name}/{key}'], [],
        ['Sheet', 'One row per…', 'Use for'],
        ['Summary', '—', 'Headline numbers: totals, by year, GLO coverage by year'],
        ['Units', 'unit (each unit ONCE)', 'Unit-level stats: AoL / hurdle rates, GLOs per unit, tasks per unit'],
        ['Tasks', 'assessment task', 'Task-level stats: weighting %, task types, hurdle vs non-hurdle'],
        ['Unit-Pathway', 'unit in a pathway', 'Which courses / pathways / blocks each unit sits in'],
        ['GLO long', 'unit × GLO', 'PivotTables of GLO × year × status without formulas'],
        ['Blocks', 'course block (Core, majors…)', 'Tasks assessing each GLO, per block'],
        ['Combinations', 'Core, or Core + one block', 'GLO gaps per CALM pathway; filter "Any gap" = Y'], [],
        ['Rules (same as the dashboard)'],
        ['Assessed', 'A task rates the GLO Yes or Non-hurdle Yes'],
        ['Assured', 'A task rates the GLO a full Yes'],
        ['AoL', 'Assessed in an AoL task: AoL Task Type is anything other than NA / N/A / Not applicable / TBC / No / blank'],
        ['Hurdle', 'Hurdle column contains "yes" and not "non" (e.g. "Yes Part A" counts; "YNo", "Not assured" do not)'],
        ['Year', 'First digit of the unit code number: 0 Non-major, 1–3 Y1–Y3, 4 Hons, 7 PG, 8 MRes, 9 PhD'],
        ['Teaching period', '* = from the Apr 2026 export (CALM has no teaching period column); T0 = unknown'],
        ['Weighting %', 'First "N percent" / "N%" in Weighting Information (or a bare number); blank when none could be read'],
        ['Blocks', 'Core Only and each "<X> Only" pathway; CALM\'s "Core + X" pathway = Core ∪ X'],
    ]

    # Units
    u_rows = [['Unit', 'Year', 'AoL unit', 'Any hurdle', 'Tasks', 'AoL tasks', 'Hurdle tasks', 'Pathways', 'Courses',
               'Course codes', 'Teaching period', 'Period source', 'GLOs assessed (unit level)']
              + [f'{g} status' for g in gl] + [f'{g} in AoL task' for g in gl]]
    for u in order:
        ts = u['tasks']
        u_rows.append([u['code'], YEAR_LABELS.get(u['year'], '?'), u['tier'] == 'aol', u['any_hurdle'], len(ts),
                       sum(t['tier'] == 'aol' for t in ts), sum(t['hurdle'] for t in ts), len(u['pathways']), len(u['courses']),
                       ', '.join(u['courses']), u['period'], u['period_source'], sum(assessed(s) for s in u['states'])]
                      + [STATUS_LABELS[s] for s in u['states']]
                      + [any(t['tier'] == 'aol' and assessed(t['states'][g]) for t in ts) for g in range(8)])

    # Tasks
    t_rows = [['Unit', 'Year', 'Task', 'Task number', 'Title', 'AoL Task Type', 'Is AoL', 'Hurdle', 'Hurdle (raw)',
               'Weighting %', 'Weighting text', 'Student Output', 'Mode', 'Participation', 'GLOs assessed'] + gl]
    unread = 0
    for u in order:
        for t in u['tasks']:
            unread += t['weight_pct'] is None
            t_rows.append([u['code'], YEAR_LABELS.get(u['year'], '?'), short_at(t['number']), t['number'], t['title'], t['type'],
                           t['tier'] == 'aol', t['hurdle'], t['hurdle_raw'], t['weight_pct'], t['weighting'], t['output'],
                           t['mode'], t['participation'], sum(assessed(s) for s in t['states'])]
                          + [STATUS_LABELS[s] for s in t['states']])

    # Unit-Pathway
    block_label = {p: lbl for bl in blocks.values() for p, lbl in bl}
    up_rows = [['Unit', 'Year', 'Course', 'Course name', 'School', 'Pathway', 'Pathway type', 'Block']]
    for code, pw in sorted(set(memberships), key=lambda m: (course_code(m[1]), m[1], m[0])):
        c = courses.get(course_code(pw), {})
        sfx = pathway_suffix(pw)
        ptype = 'Core Only' if sfx.lower() == 'core only' else re.sub(r'(MJ|ST|SP)-\S+ : .*', r'\1', sfx)
        up_rows.append([code, YEAR_LABELS.get(units[code]['year'], '?'), course_code(pw), c.get('Course Name', ''),
                        c.get('School', ''), pw, ptype, block_label.get(pw, '')])

    # GLO long
    gl_rows = [['Unit', 'Year', 'GLO', 'GLO name', 'Unit status', 'Assessed (unit)', 'Assured (unit)',
                'Tasks assessing', 'AoL tasks assessing']]
    for u in order:
        for g in range(8):
            st = u['states'][g]
            gl_rows.append([u['code'], YEAR_LABELS.get(u['year'], '?'), f'GLO{g+1}', GLO_NAMES[g], STATUS_LABELS[st],
                            assessed(st), st == 'hurdle', sum(assessed(t['states'][g]) for t in u['tasks']),
                            sum(t['tier'] == 'aol' and assessed(t['states'][g]) for t in u['tasks'])])

    # Blocks
    def block_units(pws):
        pws = set(pws)
        return [u for u in order if u['pathways'] & pws]

    b_rows = [['Course', 'Course name', 'Block', 'Units', 'Tasks'] + [f'{g} assessed' for g in gl] + [f'{g} AoL' for g in gl]]
    c_rows = [['Course', 'Course name', 'Combination', 'Units', 'Tasks', 'Assessed met (of 8)', 'Assured met', 'AoL met',
               'Assessed gaps', 'Assured gaps', 'AoL gaps', 'Any gap']]
    for cc in sorted(blocks, key=lambda c: (courses.get(c, {}).get('Course Name', c), c)):
        name = courses.get(cc, {}).get('Course Name', '')
        bl = blocks[cc]
        for p, lbl in bl:
            us = block_units([p])
            cov = coverage(us)
            b_rows.append([cc, name, lbl, len(us), sum(len(u['tasks']) for u in us)] + [c[0] for c in cov] + [c[2] for c in cov])
        core = [p for p, lbl in bl if lbl == 'Core']
        combos = [('Core', core)] if core else []
        combos += [(f'Core + {lbl}' if core else lbl, core + [p]) for p, lbl in bl if lbl != 'Core']
        for label, pws in combos:
            us = block_units(pws)
            cov = coverage(us)
            gaps = [[f'GLO{g+1}' for g in range(8) if cov[g][k] == 0] for k in range(3)]
            c_rows.append([cc, name, label, len(us), sum(len(u['tasks']) for u in us),
                           8 - len(gaps[0]), 8 - len(gaps[1]), 8 - len(gaps[2]),
                           ', '.join(gaps[0]), ', '.join(gaps[1]), ', '.join(gaps[2]), any(gaps)])

    # Summary
    def pct(a, b):
        return round(100 * a / b, 1) if b else None

    all_tasks = [t for u in order for t in u['tasks']]
    s_rows = [[f'{key} summary — {calm_folder.name}'], [],
              ['Measure', 'Value'],
              ['Units', len(order)], ['Assessment tasks', len(all_tasks)],
              ['Courses', len(blocks)], ['Pathways', len({p for u in order for p in u['pathways']})],
              ['AoL units', sum(u['tier'] == 'aol' for u in order)],
              ['Units with a hurdle', sum(u['any_hurdle'] for u in order)],
              ['Units with teaching period from 2026 data (*)', sum(u['period_source'].startswith('Apr') for u in order)],
              ['Units with teaching period T0', sum(u['period'] == 'T0' for u in order)],
              ['Tasks with weighting % read', len(all_tasks) - unread], ['Tasks with weighting % not read', unread],
              ['Combinations with any GLO gap', sum(r[-1] is True for r in c_rows[1:])], [],
              ['By year', 'Units', 'AoL units', '% AoL', 'Hurdle units', '% hurdle', 'Tasks', 'Tasks per unit',
               'Mean weighting % (hurdle tasks)', 'Mean weighting % (other tasks)']]
    for y in years:
        us = [u for u in order if u['year'] == y]
        ts = [t for u in us for t in u['tasks']]
        hw = [t['weight_pct'] for t in ts if t['hurdle'] and t['weight_pct'] is not None]
        ow = [t['weight_pct'] for t in ts if not t['hurdle'] and t['weight_pct'] is not None]
        aol = sum(u['tier'] == 'aol' for u in us)
        hur = sum(u['any_hurdle'] for u in us)
        s_rows.append([YEAR_LABELS.get(y, '?'), len(us), aol, pct(aol, len(us)), hur, pct(hur, len(us)), len(ts),
                       round(len(ts) / len(us), 2) if us else None,
                       round(sum(hw) / len(hw), 1) if hw else None, round(sum(ow) / len(ow), 1) if ow else None])
    for title, fn in (('% of units assessing each GLO (unit level), by year', lambda u, g: assessed(u['states'][g])),
                      ('% of units with an AoL task assessing each GLO, by year',
                       lambda u, g: any(t['tier'] == 'aol' and assessed(t['states'][g]) for t in u['tasks']))):
        s_rows += [[], [title] + gl]
        for y in years + ['All']:
            us = order if y == 'All' else [u for u in order if u['year'] == y]
            s_rows.append([y if y == 'All' else YEAR_LABELS.get(y, '?')] + [pct(sum(fn(u, g) for u in us), len(us)) for g in range(8)])

    out = calm_folder / f'{key}-statistics.xlsx'
    write_xlsx(out, [
        ('README', readme, [18, 30, 90], 0),
        ('Summary', s_rows, [46, 10, 10, 10, 12, 10, 10, 12, 16, 16], 0),
        ('Units', u_rows, [10, 8, 9, 10, 7, 9, 11, 9, 8, 28, 22, 16, 12] + [14] * 8 + [10] * 8, 1),
        ('Tasks', t_rows, [10, 8, 7, 12, 40, 26, 7, 7, 14, 11, 50, 26, 18, 14, 9] + [14] * 8, 1),
        ('Unit-Pathway', up_rows, [10, 8, 8, 40, 36, 50, 12, 36], 1),
        ('GLO long', gl_rows, [10, 8, 7, 26, 14, 12, 12, 12, 14], 1),
        ('Blocks', b_rows, [8, 40, 40, 7, 7] + [10] * 16, 1),
        ('Combinations', c_rows, [8, 40, 50, 7, 7, 12, 11, 9, 24, 24, 24, 9], 1),
    ])
    gaps = sum(r[-1] is True for r in c_rows[1:])
    print(f'Statistics: {len(order)} units, {len(all_tasks)} tasks, {len(c_rows) - 1} combinations '
          f'({gaps} with a GLO gap) → {out.name}')
    return out


def main():
    key = (sys.argv[2] if len(sys.argv) > 2 else 'SEBE').upper()
    if len(sys.argv) > 1:
        folder = Path(sys.argv[1]).expanduser().resolve()
    else:
        found = sorted(p for p in HERE.glob('CALM-dashboard-*') if (p / key).is_dir())
        if not found:
            sys.exit(f'No CALM-dashboard-*/{key} folder found next to {Path(__file__).name}')
        folder = found[-1]
    if not (folder / key).is_dir():
        sys.exit(f'No {key} folder in {folder.name}. Run first:  python3 extract-faculty.py {key} "{folder}"')
    build(folder, key)


if __name__ == '__main__':
    main()
