#!/usr/bin/env python3
"""
Build a compact, filterable curriculum map workbook from a dashboard
timeline export (Assessment Timeline → "Export Excel" or "Export CSV").

    python3 make-curriculum-map.py "S320 Bachelor Of Science timeline_091126.csv"
    python3 make-curriculum-map.py export.xlsx -o "S320 map.xlsx"

Writes "<input name> map.xlsx" beside the input unless -o is given.

Sheets
  Summary  Year/trimester band → unit row → its tasks, with GLO1–8 as coloured
           codes. LIVE: it rebuilds from whatever is visible on the Tasks sheet,
           so filter or slice Tasks and the Summary follows.
  Tasks    The export, one row per task, as an Excel Table. Every row is
           complete, so filtering and sorting can never split a unit from its
           tasks. One ● column per pathway replaces the "Pathways" lookup.
  Outline  (hidden) The fixed unit/task skeleton the Summary filters.

Needs Excel 365 or Excel 2021+ (FILTER, MAXIFS). Requires: openpyxl (only for
.xlsx input) and XlsxWriter.
"""
import argparse
import csv
import re
import sys
from collections import OrderedDict
from pathlib import Path

try:
    import xlsxwriter
    from xlsxwriter.utility import xl_col_to_name
except ImportError:
    sys.exit('XlsxWriter is required:  python3 -m pip install XlsxWriter')

# ── GLO cell vocabulary (labels written by the dashboard's stateLabel()) ──
# code shown in the map, rank used to pick a unit's / period's strongest status
GLO_CODES = OrderedDict([
    ('yes',               ('H', 5)),   # hurdle Yes
    ('hurdle yes',        ('H', 5)),
    ('non-hurdle yes',    ('Y', 4)),
    ('data error',        ('E', 3)),
    ('tbc',               ('?', 2)),
    ('form not received', ('?', 2)),
    ('no',                ('·', 1)),
    ('not assessed',      ('·', 1)),
])
RANK_TO_CODE = ['', '·', '?', 'E', 'Y', 'H']          # index = rank
LEGEND = [
    ('H', 'Assessed – hurdle'),
    ('Y', 'Assessed – non-hurdle'),
    ('·', 'Not assessed / No'),
    ('E', 'Data error'),
    ('?', 'TBC'),
]
# fills mirror the dashboard chips: filled / outlined / faded / red / grey-blue
CODE_STYLE = {
    'H': {'bg_color': '#1F6F66', 'font_color': '#FFFFFF', 'bold': True},
    'Y': {'bg_color': '#BFE3DE', 'font_color': '#0F4D46'},
    '·': {'font_color': '#B5B9BF'},
    'E': {'bg_color': '#E05252', 'font_color': '#FFFFFF', 'bold': True},
    '?': {'bg_color': '#C9D3E0', 'font_color': '#33475B'},
}
STATE_STYLE = {   # same colours for the full-text GLO cells on the Tasks sheet
    'Yes': 'H', 'Hurdle Yes': 'H', 'Non-hurdle Yes': 'Y', 'Data Error': 'E',
    'TBC': '?', 'Form not received': '?',
}

YEAR_ORDER = {'Non-major': 0, 'Y1': 1, 'Y2': 2, 'Y3': 3, 'Hons': 4, 'PG': 7, 'MRes': 8, 'PhD': 9}


def read_rows(path):
    if path.suffix.lower() == '.csv':
        with open(path, newline='', encoding='utf-8-sig') as fh:
            return [dict(r) for r in csv.DictReader(fh)]
    import openpyxl
    wb = openpyxl.load_workbook(path, read_only=True, data_only=True)
    ws = wb.worksheets[0]
    it = ws.iter_rows(values_only=True)
    header = [str(h or '').strip() for h in next(it)]
    rows = []
    for vals in it:
        if not any(v not in (None, '') for v in vals):
            continue
        rows.append({h: ('' if v is None else str(v)) for h, v in zip(header, vals)})
    return rows


def find_col(columns, *names):
    low = {c.lower(): c for c in columns}
    for n in names:
        if n.lower() in low:
            return low[n.lower()]
    for n in names:                                   # prefix match ("Assessment Title (…)")
        for c in columns:
            if c.lower().startswith(n.lower()):
                return c
    return None


def is_hurdle_yes(v):          # same rule as the dashboard's isHurdleYes()
    t = str(v or '').strip().lower()
    return bool(t) and 'non' not in t and 'yes' in t


def glo_code(v):
    return GLO_CODES.get(str(v or '').strip().lower(), ('', 0))


def short_weight(text):
    """Best-effort headline weighting as a fraction (0.3), or None."""
    t = str(text or '').lower().replace('per cent', 'percent')
    pct = r'(\d+(?:\.\d+)?)\s*(?:%|percen)'
    m = re.search(r'total\s*[:=]?\s*' + pct, t)
    if not m:
        lead = re.match(r'\s*(\d+)\s*[x×]\s*(\d+(?:\.\d+)?)\s*(?:%|percen)', t)
        if lead:
            return float(lead.group(1)) * float(lead.group(2)) / 100
        m = re.search(pct, t)
    return float(m.group(1)) / 100 if m else None


def natural_key(s):
    return [int(p) if p.isdigit() else p.lower() for p in re.split(r'(\d+)', str(s))]


def year_key(label):
    return YEAR_ORDER.get(label, 99), label


def tri_key(label):
    m = re.search(r'(\d)', label or '')
    return (int(m.group(1)) if m else 9), label


def build(in_path, out_path):
    rows = read_rows(in_path)
    if not rows:
        sys.exit(f'No rows found in {in_path}')
    columns = list(rows[0].keys())

    c_year = find_col(columns, 'Study Year')
    c_tri = find_col(columns, 'Trimester')
    c_unit = find_col(columns, 'Unit Code')
    c_at = find_col(columns, 'AT Number', 'Task Number')
    c_title = find_col(columns, 'Assessment Title')
    c_aol = find_col(columns, 'AoL')
    c_hurdle = find_col(columns, 'Hurdle')
    c_weight = find_col(columns, 'Weighting')
    c_paths = find_col(columns, 'Pathways')
    glo_cols = [c for c in columns if re.match(r'GLO\d', c)]
    missing = [n for n, c in [('Study Year', c_year), ('Trimester', c_tri), ('Unit Code', c_unit),
                              ('AT Number', c_at)] if not c]
    if missing or len(glo_cols) != 8:
        sys.exit(f'This does not look like a dashboard timeline export (missing: '
                 f'{", ".join(missing) or "GLO1–GLO8 columns"}).')

    # ── pathways → flag columns ──
    pathways = []
    for r in rows:
        for p in str(r.get(c_paths, '') or '').split(';'):
            p = p.strip()
            if p and p not in pathways:
                pathways.append(p)
    courses = {re.match(r'^([A-Za-z0-9]+)', p).group(1) for p in pathways if re.match(r'^[A-Za-z0-9]+', p)}

    def path_header(p):
        label = re.sub(r'^[A-Za-z0-9]+_', '', p) if len(courses) == 1 else p
        return label.replace(' : ', ' ')

    def sort_core_first(p):
        s = path_header(p).lower()
        return (0 if s.startswith('core only') else 1 if s.startswith('core') else 2, s)
    pathways.sort(key=sort_core_first)
    flag_headers = []
    for p in pathways:                                 # Table headers must be unique
        h = path_header(p)[:250]
        while h in flag_headers or h in columns:
            h += ' '
        flag_headers.append(h)

    def key_of(r):
        return '|'.join(str(r[c]).strip() for c in (c_year, c_tri, c_unit, c_at))

    # ── outline: period → unit → task (deduped: one task row per key) ──
    tasks = OrderedDict()
    for r in rows:
        k = key_of(r)
        if k not in tasks:
            tasks[k] = r
    ordered = sorted(tasks.items(), key=lambda kv: (
        year_key(kv[1][c_year]), tri_key(kv[1][c_tri]), kv[1][c_unit], natural_key(kv[1][c_at])))

    outline = []        # dicts: type, label cells, codes, ranks, key
    period = unit = None
    for k, r in ordered:
        p = (r[c_year], r[c_tri])
        if p != period:
            period, unit = p, None
            outline.append({'type': 'Period', 'year': p[0], 'tri': p[1]})
        if r[c_unit] != unit:
            unit = r[c_unit]
            outline.append({'type': 'Unit', 'year': p[0], 'tri': p[1], 'unit': unit})
        codes = [glo_code(r[c]) for c in glo_cols]
        outline.append({
            'type': 'Task', 'key': k, 'task': r[c_at], 'title': r.get(c_title, ''),
            'aol': 'Yes' if str(r.get(c_aol, '')).strip().lower() == 'yes' else 'No',
            'hurdle': 'Yes' if is_hurdle_yes(r.get(c_hurdle, '')) else 'No',
            'weight': short_weight(r.get(c_weight, '')),
            'codes': [c for c, _ in codes], 'ranks': [n for _, n in codes],
        })

    wb = xlsxwriter.Workbook(str(out_path), {'use_future_functions': True,
                                             'strings_to_numbers': False,
                                             'strings_to_formulas': False})
    wb.set_properties({'title': f'Curriculum map – {in_path.stem}',
                       'comments': 'Built by make-curriculum-map.py from a dashboard timeline export'})
    F = lambda **kw: wb.add_format({'font_name': 'Calibri', 'font_size': 10, 'valign': 'vcenter', **kw})
    code_fmt = {c: F(align='center', **s) for c, s in CODE_STYLE.items()}

    ws_sum = wb.add_worksheet('Summary')
    ws_tasks = wb.add_worksheet('Tasks')
    ws_out = wb.add_worksheet('Outline')

    # ════════ Tasks (Excel Table) ════════
    n = len(rows)
    t_cols = columns + flag_headers + ['Map Key', 'Visible']
    data = []
    for r in rows:
        own = {p.strip() for p in str(r.get(c_paths, '') or '').split(';')}
        data.append([r[c] for c in columns] + ['●' if p in own else '' for p in pathways]
                    + [key_of(r), None])
    hdr = F(bold=True, text_wrap=True, valign='bottom', bg_color='#2B2F3A', font_color='#FFFFFF')
    hdr_rot = F(bold=True, rotation=90, align='center', valign='bottom',
                bg_color='#3B4252', font_color='#FFFFFF')
    cell = F()
    flag_fmt = F(align='center', font_color='#1F6F66', bold=True)
    tbl_cols = []
    for i, h in enumerate(t_cols):
        spec = {'header': h, 'header_format': hdr}
        if h in flag_headers:
            spec.update(header_format=hdr_rot, format=flag_fmt)
        elif h == 'Visible':
            pass                                        # per-row formulas written below
        else:
            spec['format'] = cell
        tbl_cols.append(spec)
    ws_tasks.add_table(0, 0, n, len(t_cols) - 1, {
        'name': 'Tasks', 'style': 'Table Style Light 1', 'columns': tbl_cols,
        'data': [row[:-1] for row in data]})
    ws_tasks.set_row(0, 150 if flag_headers else 30)
    widths = {c_year: 6, c_tri: 11, 'AT Group': 7, c_unit: 9, c_at: 10, c_title: 34,
              'Course Code': 8, 'Course': 22, c_paths: 30, 'Teaching Period Source': 12,
              'AoL Task Type': 22, c_aol: 5, c_hurdle: 8, c_weight: 18, 'Mode': 12,
              'Participation': 11, 'Student Output': 30}
    for i, h in enumerate(t_cols):
        if h in flag_headers:
            ws_tasks.set_column(i, i, 3.2)
        elif h in glo_cols:
            ws_tasks.set_column(i, i, 9)
        else:
            ws_tasks.set_column(i, i, widths.get(h, 12))
    key_col, vis_col = len(t_cols) - 2, len(t_cols) - 1
    # 1 while the row is shown, 0 once a filter/slicer hides it (plain A1 refs: robust everywhere)
    unit_letter = xl_col_to_name(columns.index(c_unit))
    for r in range(1, n + 1):
        ws_tasks.write_formula(r, vis_col, f'=SUBTOTAL(103,{unit_letter}{r + 1})', cell, 1)
    ws_tasks.set_column(key_col, vis_col, 8, None, {'hidden': True})
    ws_tasks.freeze_panes(1, columns.index(c_at) + 1)
    first_glo, last_glo = columns.index(glo_cols[0]), columns.index(glo_cols[-1])
    for label, code in STATE_STYLE.items():
        ws_tasks.conditional_format(1, first_glo, n, last_glo, {
            'type': 'cell', 'criteria': '==', 'value': f'"{label}"',
            'format': wb.add_format({k: v for k, v in CODE_STYLE[code].items()})})
    for faded in ('No', 'Not Assessed'):
        ws_tasks.conditional_format(1, first_glo, n, last_glo, {
            'type': 'cell', 'criteria': '==', 'value': f'"{faded}"',
            'format': wb.add_format({'font_color': '#9AA0A6'})})

    key_rng = f"Tasks!${xl_col_to_name(key_col)}$2:${xl_col_to_name(key_col)}${n + 1}"
    vis_rng = f"Tasks!${xl_col_to_name(vis_col)}$2:${xl_col_to_name(vis_col)}${n + 1}"

    # ════════ Outline (hidden) ════════
    # A Show | B Type | C Year | D Trimester | E Unit | F Task | G Title | H AoL | I Hurdle
    # J Weight | K–R GLO1–8 code | S Key | T–AA GLO1–8 rank
    out_hdr = ['Show', 'Type', 'Year', 'Trimester', 'Unit', 'Task', 'Assessment title', 'AoL',
               'Hurdle', 'Weight'] + [f'G{i}' for i in range(1, 9)] + ['Key'] + [f'R{i}' for i in range(1, 9)]
    ws_out.write_row(0, 0, out_hdr)
    pct = F(num_format='0%')
    # ranges of children for each Period / Unit row (outline rows are contiguous)
    m = len(outline)
    ends = [None] * m
    for i, o in enumerate(outline):
        if o['type'] == 'Task':
            continue
        j = i + 1
        while j < m and not (outline[j]['type'] == 'Period' or
                             (o['type'] == 'Unit' and outline[j]['type'] == 'Unit')):
            j += 1
        ends[i] = j - 1                                 # last child index
    choose = 'CHOOSE({}+1,"","·","?","E","Y","H")'
    for i, o in enumerate(outline):
        r = i + 1                                       # sheet row (0-based); Excel row = r + 1
        if o['type'] == 'Task':
            ws_out.write_formula(r, 0, f'=IF(COUNTIFS({key_rng},S{r + 1},{vis_rng},1)>0,1,0)', None, 1)
            ws_out.write_row(r, 1, ['Task', '', '', '', o['task'], o['title'], o['aol'], o['hurdle']])
            if o['weight'] is not None:
                ws_out.write_number(r, 9, o['weight'], pct)
            ws_out.write_row(r, 10, o['codes'])
            ws_out.write_string(r, 18, o['key'])
            ws_out.write_row(r, 19, o['ranks'])
            continue
        a, b = r + 2, ends[i] + 2                       # Excel rows of first/last child
        show = f'A{a}:A{b}'
        ws_out.write_formula(r, 0, f'=IF(SUM({show})>0,1,0)', None, 1)
        ws_out.write_string(r, 1, o['type'])
        is_task = f'B{a}:B{b},"Task",{show},1'
        n_tasks = f'COUNTIFS({is_task})'
        if o['type'] == 'Period':
            ws_out.write_row(r, 2, [o['year'], o['tri']])
            ws_out.write_formula(r, 4, f'=COUNTIFS(B{a}:B{b},"Unit",{show},1)&" units"')
            ws_out.write_formula(r, 5, f'={n_tasks}&" tasks"')
        else:
            ws_out.write_row(r, 2, [o['year'], o['tri'], o['unit']])
            ws_out.write_formula(r, 5, f'={n_tasks}&IF({n_tasks}=1," task"," tasks")')
            ws_out.write_formula(r, 7, f'=COUNTIFS(H{a}:H{b},"Yes",{show},1)')
            ws_out.write_formula(r, 8, f'=COUNTIFS(I{a}:I{b},"Yes",{show},1)')
            ws_out.write_formula(r, 9, f'=SUMIFS(J{a}:J{b},{show},1)', pct)
        for g in range(8):
            rc = xl_col_to_name(19 + g)
            ws_out.write_formula(r, 10 + g, '=' + choose.format(f'MAXIFS({rc}{a}:{rc}{b},{show},1)'))
    ws_out.hide()

    # ════════ Summary ════════
    last = m + 1
    ws_sum.hide_gridlines(2)
    ws_sum.set_column('A:A', 2, None, {'hidden': True})
    ws_sum.set_column('B:B', 9)       # Year
    ws_sum.set_column('C:C', 12)      # Trimester
    ws_sum.set_column('D:D', 9)       # Unit
    ws_sum.set_column('E:E', 10)      # Task
    ws_sum.set_column('F:F', 44)      # Title
    ws_sum.set_column('G:H', 7)       # AoL, Hurdle
    ws_sum.set_column('I:I', 7, F(num_format='0%;;'))
    ws_sum.set_column('J:Q', 4.2)     # GLO codes

    course_names = sorted({r.get('Course', '') for r in rows if r.get('Course')})
    ws_sum.write('B1', ' · '.join(course_names) or in_path.stem, F(bold=True, font_size=14))
    ws_sum.write_formula('B2', f'="Showing "&COUNTIFS(Outline!A2:A{last},1,Outline!B2:B{last},"Task")'
                               f'&" of {sum(1 for o in outline if o["type"] == "Task")} tasks · '
                               f'filter or slice the Tasks sheet to change this view"',
                         F(italic=True, font_color='#6B7280'))
    col = 1
    for code, text in LEGEND:
        ws_sum.write(2, col, code, code_fmt[code])
        ws_sum.write(2, col + 1, text, F(font_color='#4B5563'))
        col += 2 if len(text) < 14 else 3
    ws_sum.write(3, 1, 'Unit rows: G1–G8 = strongest status among the unit’s visible tasks; '
                       'Weight = their total (amber when it isn’t 100%). Full weighting text is on Tasks.',
                 F(font_color='#6B7280', font_size=9))
    heads = ['', 'Year', 'Trimester', 'Unit', 'Task', 'Assessment title', 'AoL', 'Hurdle', 'Weight']
    sh = F(bold=True, bottom=1, valign='bottom')
    ws_sum.write_row(4, 0, heads, sh)
    for g in range(8):
        ws_sum.write(4, 9 + g, f'G{g + 1}', F(bold=True, bottom=1, align='center', valign='bottom'))
    ws_sum.set_row(4, 18)
    for g, name in enumerate(['Discipline-specific knowledge', 'Communication', 'Digital literacy',
                              'Critical thinking', 'Problem solving', 'Self-management', 'Teamwork',
                              'Global citizenship']):
        ws_sum.write_comment(4, 9 + g, f'GLO{g + 1} {name}', {'x_scale': 1.4, 'y_scale': 0.4})
    src = f'Outline!B2:R{last}'
    ws_sum.write_dynamic_array_formula(
        'A6', f'=FILTER(IF({src}="","",{src}),Outline!A2:A{last}=1,"No tasks match the filters on the Tasks sheet")')
    ws_sum.freeze_panes(5, 0)

    rng = (5, 0, 5 + m, 16)
    for code, fmt in code_fmt.items():                  # GLO codes first: they win the fill
        ws_sum.conditional_format(5, 9, 5 + m, 16, {
            'type': 'cell', 'criteria': '==', 'value': f'"{code}"', 'format': fmt})
    # a unit whose (visible) task weights don't come to 100%: data or parsing worth a look
    ws_sum.conditional_format(5, 8, 5 + m, 8, {
        'type': 'formula', 'criteria': '=AND($A6="Unit",ABS($I6-1)>0.005)',
        'format': wb.add_format({'bold': True, 'bg_color': '#FDE7B0', 'font_color': '#8A5A00'})})
    ws_sum.conditional_format(*rng, {'type': 'formula', 'criteria': '=$A6="Period"',
                                     'format': wb.add_format({'bold': True, 'bg_color': '#E5E7EB',
                                                              'top': 1, 'top_color': '#6B7280'})})
    ws_sum.conditional_format(*rng, {'type': 'formula', 'criteria': '=$A6="Unit"',
                                     'format': wb.add_format({'bold': True, 'top': 1, 'top_color': '#D1D5DB'})})
    ws_sum.conditional_format(5, 5, 5 + m, 8, {'type': 'formula', 'criteria': '=$A6="Task"',
                                               'format': wb.add_format({'font_color': '#374151'})})
    ws_sum.activate()
    wb.close()
    return len(rows), sum(1 for o in outline if o['type'] == 'Task'), len(pathways)


def main():
    ap = argparse.ArgumentParser(description=__doc__.split('\n\n')[0])
    ap.add_argument('export', type=Path, help='timeline export (.csv or .xlsx)')
    ap.add_argument('-o', '--output', type=Path)
    a = ap.parse_args()
    if not a.export.exists():
        sys.exit(f'Not found: {a.export}')
    out = a.output or a.export.with_name(re.sub(r'\s*timeline_', ' map_', a.export.stem) + '.xlsx')
    if out.resolve() == a.export.resolve():
        out = a.export.with_name(a.export.stem + ' map.xlsx')
    rows, tasks, paths = build(a.export, out)
    print(f'Wrote {out}\n  {rows} export rows → {tasks} tasks in the Summary, {paths} pathway columns')


if __name__ == '__main__':
    main()
