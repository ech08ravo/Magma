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
  Timeline Like the dashboard's Trimester view: a band per study year, columns
           Trimester × AT1/AT2/AT3/Other, one task card per cell, coloured by
           alignment (⚑ AoL + hurdle, ● AoL, ◆ hurdle, ○ neither), with GLO dots
           (one cell per GLO under each card: the task's ⚑/●/◆/○ in that GLO's
           dashboard colour) and the guide week when the export has one. Also LIVE.
  Weeks    Like the dashboard's Week view: Trimester × week columns. Guide weeks
           from dashboard v7.2+ exports; other tasks by AT number (AT1 → Week 1),
           dashed and marked "≈ Week from AT number". Also LIVE.
  AOL type balance  GLO × year level: every GLO should have at least 2 AoL tasks (⚑/●) at
           every level. Marks per task in GLO colours, AoL count against target,
           and each AoL task's unit · AT · assessment type · weight. Also LIVE.
  All tasks Same layout as AOL type balance, but every task (not just AoL) under each GLO it
           addresses: unit-guide alignment (dashboard v7.4+ export) where available,
           otherwise CALM. No target; empty cells are flagged. Also LIVE.
  Tasks    The export, one row per task, as an Excel Table, plus planning columns
           (Tools, Assessment type) added empty for the team to fill in; filled Tools show on cards. Every row is
           complete, so filtering and sorting can never split a unit from its
           tasks. One ● column per pathway replaces the "Pathways" lookup.
  Outline  (hidden) The fixed unit/task skeleton the Summary filters.
  WeekRows (hidden) One row per task per week, feeding the Weeks sheet.
  Legend   Quick guide: what each sheet shows, the marks, GLO colours, type codes,
           filters and planning columns.

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
    from xlsxwriter.utility import xl_col_to_name, xl_rowcol_to_cell, xl_cell_to_rowcol as cell_to_rowcol
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

# Columns the team fills in during planning (e.g. Tools: FFIO, OT). Added empty when the
# export doesn't have them; shown on the cards when filled.
PLANNING_COLUMNS = ['Tools', 'Assessment type', 'Planned Weeks']

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


def mode_label(text):
    t = str(text or '').lower()
    in_person = bool(re.search(r'campus|in[- ]person|field|placement|site|face[- ]to[- ]face|lab', t))
    online = 'online' in t
    return 'Both' if in_person and online else 'In person' if in_person else 'Online' if online else 'Other'


def participation_label(text):
    t = str(text or '').lower()
    ind, grp = 'individual' in t, ('group' in t or 'team' in t)
    return 'Both' if ind and grp else 'Ind' if ind else 'Grp' if grp else '—'


def aol_type_short(text):
    t = re.sub(r'\s+', ' ', str(text or '')).strip()
    if t.lower() in ('', 'na', 'n/a', 'not applicable', 'no', 'tbc', 'not aol task'):
        return ''
    return t if len(t) <= 34 else t[:33].rstrip() + '…'


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
    # Planning columns typed in Excel (not in the dashboard export yet): add them
    # empty after the export columns so they are part of the Tasks table
    for planning_col in PLANNING_COLUMNS:
        if planning_col not in columns:
            columns.append(planning_col)
            for r in rows:
                r[planning_col] = ''

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
            'extra': [mode_label(r.get('Mode', '')), participation_label(r.get('Participation', '')),
                      str(r.get('Tools', '') or '').strip(), aol_type_short(r.get('AoL Task Type', ''))],
        })

    wb = xlsxwriter.Workbook(str(out_path), {'use_future_functions': True,
                                             'strings_to_numbers': False,
                                             'strings_to_formulas': False})
    wb.set_properties({'title': f'Curriculum map – {in_path.stem}',
                       'comments': 'Built by make-curriculum-map.py from a dashboard timeline export'})
    F = lambda **kw: wb.add_format({'font_name': 'Calibri', 'font_size': 10, 'valign': 'vcenter', **kw})
    code_fmt = {c: F(align='center', **s) for c, s in CODE_STYLE.items()}

    ws_tl = wb.add_worksheet('Timeline')
    ws_wk = wb.add_worksheet('Weeks')
    ws_bal = wb.add_worksheet('AOL type balance')
    ws_all = wb.add_worksheet('All tasks')
    ws_sum = wb.add_worksheet('Summary')
    ws_tasks = wb.add_worksheet('Tasks')
    ws_out = wb.add_worksheet('Outline')
    ws_wr = wb.add_worksheet('WeekRows')
    ws_leg = wb.add_worksheet('Legend')                 # quick guide, last tab

    # Timeline card text, one per task row, mirroring the dashboard card:
    #   ⚑ SLE111 · SPB-AT-2          alignment mark (dashboard icon shape), unit · AT
    #   Practical exercises          title
    #   35% · On-campus (· S466)     weight · mode (· course when the export has several)
    #   🔴1 🟠2 🟢4                   GLO dots in GLO colours
    #   Guide: Wk 4 (Quiz 2)         week line (guide weeks from v7.2 exports)
    multi_course = len({r.get('Course Code', '') for r in rows}) > 1
    has_guides = 'Guide Weeks' in columns

    def marker_of(r):
        if r.get('Alignment Mark'):                    # v7.2 export
            return r['Alignment Mark']
        aol = str(r.get(c_aol, '')).strip().lower() == 'yes'
        hurdle = is_hurdle_yes(r.get(c_hurdle, ''))
        return '⚑' if aol and hurdle else '●' if aol else '◆' if hurdle else '○'

    def glo_ids(r):
        if r.get('Aligned GLOs') is not None and 'Aligned GLOs' in columns:
            return [int(g) for g in re.findall(r'GLO(\d)', r['Aligned GLOs'])]
        return [i + 1 for i, c in enumerate(glo_cols) if glo_code(r[c])[0] in ('H', 'Y')]

    def card_text(r, week_line=None):
        title = str(r.get(c_title, '') or '(no title)').strip()
        title = title if len(title) <= 46 else title[:45].rstrip() + '…'
        w = short_weight(r.get(c_weight, ''))
        mode = str(r.get('Mode', '') or '').strip()
        mode = mode if len(mode) <= 22 else mode[:21].rstrip() + '…'
        bits = [f'{w * 100:g}%' if w is not None else '—', mode or '—']
        if multi_course:
            bits.append(r.get('Course Code', ''))
        lines = [f'{marker_of(r)} {r[c_unit]} · {r[c_at]}', title, ' · '.join(bits)]
        if not glo_ids(r):
            lines.append('No GLO alignment')
        if str(r.get('Assessment type', '') or '').strip():
            lines.append('Type: ' + str(r['Assessment type']).strip())
        if str(r.get('Tools', '') or '').strip():
            lines.append('Tools: ' + str(r['Tools']).strip())
        if week_line is None and r.get('Guide Week Detail'):
            week_line = f'Guide: {r["Guide Week Detail"]}'
        if week_line:
            lines.append(week_line)
        return '\n'.join(lines)

    def features_of(r):
        # filter helpers: in person (1/0), online (1/0), headline weight in % (blank if unreadable)
        mode = str(r.get('Mode', '') or '').lower()
        in_person = int(bool(re.search(r'campus|in[- ]person|field|placement|site|face[- ]to[- ]face|lab', mode)))
        online = int('online' in mode)
        other = int(not in_person and not online)
        w = short_weight(r.get(c_weight, ''))
        return [in_person, online, other, round(w * 100, 2) if w is not None else '']

    def type_icon(r):
        # primary task type as a short code: the type mentioned FIRST in Assessment
        # type (planning) or, until that's filled in, CALM Student Output
        text = (str(r.get('Assessment type', '') or '').strip()
                or str(r.get('Student Output', '') or '')).lower()
        hits = [(m.start(), code) for code, _, pattern in TASK_TYPES
                for m in [re.search(pattern, text)] if m]
        return min(hits)[1] if hits else TYPE_OTHER[0]

    def balance_line(r):
        # Balance sheet line: "⚑ SEE210 · SPB-AT-3 · Oral · 30%". Type is the planning
        # column 'Assessment type'; until it's filled, CALM Student Output marked with *
        planned = str(r.get('Assessment type', '') or '').strip()
        if planned:
            kind = planned
        else:
            so = re.sub(r'\s+', ' ', str(r.get('Student Output', '') or '')).strip() or '—'
            kind = (so if len(so) <= 26 else so[:25].rstrip() + '…') + '*'
        w = short_weight(r.get(c_weight, ''))
        return (f'{marker_of(r)} {r[c_unit]} · {r[c_at]} · {kind} · '
                + (f'{w * 100:g}%' if w is not None else '—'))

    has_guide_glos = 'Guide GLOs' in columns

    def any_glo_ids(r):
        # every task's GLOs: the unit guide's alignment where the export has it
        # (dashboard v7.4+), otherwise CALM (which only aligns AoL tasks)
        guide = [int(g) for g in re.findall(r'GLO(\d)', str(r.get('Guide GLOs', '') or ''))]
        return guide or glo_ids(r)

    def any_marks_of(r):
        ids, mk = set(any_glo_ids(r)), marker_of(r)
        return [mk if g in ids else '' for g in range(1, 9)]

    def all_line(r):
        source = '' if str(r.get('Guide GLOs', '') or '').strip() else ' · CALM'
        return balance_line(r) + source

    def sort_key(r):
        # unit code, then AT number with numbers zero-padded (AT-2 before AT-10)
        pad = lambda text: re.sub(r'\d+', lambda m: m.group(0).zfill(4), str(text or '').upper())
        return f'{pad(r[c_unit])}|{pad(r[c_at])}'

    def marks_of(r):
        # one entry per GLO: the task's alignment mark where that GLO is aligned
        # (shown in that GLO's colour, like the dashboard's coloured flags/dots)
        ids, mk = set(glo_ids(r)), marker_of(r)
        return [mk if g in ids else '' for g in range(1, 9)]

    # Week placements (Weeks sheet): guide weeks where the export has them,
    # otherwise the AT number (AT1 → Week 1 …), flagged as such on the card.
    week_rows = []                                      # (row id, year, tri, week key, card, …)
    for i, r in enumerate(rows):
        planned = re.findall(r'\d+', str(r.get('Planned Weeks', '') or ''))
        guide_weeks = re.findall(r'\d+', r.get('Guide Weeks', '') or '')
        if planned:                                     # the plan wins over the guide
            for wk in dict.fromkeys(planned):
                week_rows.append((i + 1, r[c_year], r[c_tri], wk, card_text(r, f'Planned: Wk {wk}'),
                                  marks_of(r), features_of(r)))
        elif guide_weeks:
            detail = r.get('Guide Week Detail', '')
            for wk in dict.fromkeys(guide_weeks):
                part = re.search(rf'Wk {wk} \(([^)]*)\)', detail)
                line = f'Guide: Wk {wk}' + (f' ({part.group(1)})' if part else '')
                week_rows.append((i + 1, r[c_year], r[c_tri], wk, card_text(r, line), marks_of(r), features_of(r)))
        else:
            m_at = re.search(r'\d+', r.get(c_at, '') or '')
            week_rows.append((i + 1, r[c_year], r[c_tri], m_at.group(0) if m_at else 'none',
                              card_text(r, FALLBACK_LINE if m_at else 'No week'), marks_of(r), features_of(r)))

    # ════════ Tasks (Excel Table) ════════
    n = len(rows)
    helper_cols = ([f'GLO{g} Mark' for g in range(1, 9)] + [f'GLO{g} Any' for g in range(1, 9)]
                   + ['In person', 'Online', 'Other mode', 'Weight %',
                      'Balance Line', 'All Line', 'Sort Key', 'Row ID', 'Card', 'Map Key', 'Visible'])
    t_cols = columns + flag_headers + helper_cols
    data = []
    for i, r in enumerate(rows):
        own = {p.strip() for p in str(r.get(c_paths, '') or '').split(';')}
        data.append([r[c] for c in columns] + ['●' if p in own else '' for p in pathways]
                    + marks_of(r) + any_marks_of(r) + features_of(r)
                    + [balance_line(r) + TYPE_DIVIDER + type_icon(r), all_line(r) + TYPE_DIVIDER + type_icon(r), sort_key(r),
                       i + 1, card_text(r), key_of(r), None])
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
              'Participation': 11, 'Student Output': 30, 'Tools': 16,
              'Assessment type': 18, 'Planned Weeks': 10}
    for i, h in enumerate(t_cols):
        if h in flag_headers:
            ws_tasks.set_column(i, i, 3.2)
        elif h in glo_cols:
            ws_tasks.set_column(i, i, 9)
        else:
            ws_tasks.set_column(i, i, widths.get(h, 12))
    hc = lambda name: len(columns) + len(flag_headers) + helper_cols.index(name)
    rowid_col, card_col, key_col, vis_col = hc('Row ID'), hc('Card'), hc('Map Key'), hc('Visible')
    mark_col0, ip_col, on_col, oth_col, wt_col = (hc('GLO1 Mark'), hc('In person'), hc('Online'),
                                                  hc('Other mode'), hc('Weight %'))
    line_col, all_line_col, any_col0 = hc('Balance Line'), hc('All Line'), hc('GLO1 Any')
    sort_rng = f"Tasks!${xl_col_to_name(hc('Sort Key'))}$2:${xl_col_to_name(hc('Sort Key'))}${n + 1}"
    tl_show_col = len(t_cols)                           # "Show (Timeline)", just right of the table
    # 1 while the row is shown, 0 once a filter/slicer hides it (plain A1 refs: robust everywhere)
    unit_letter = xl_col_to_name(columns.index(c_unit))
    for r in range(1, n + 1):
        ws_tasks.write_formula(r, vis_col, f'=SUBTOTAL(103,{unit_letter}{r + 1})', cell, 1)
    ws_tasks.write(0, tl_show_col, 'Show (Timeline)', hdr)
    ws_tasks.set_column(mark_col0, tl_show_col, 8, None, {'hidden': True})
    if has_guides:                                      # long quotes: keep the table compact
        for h in ('Guide Evidence', 'Guide Week Detail'):
            if h in t_cols:
                ws_tasks.set_column(t_cols.index(h), t_cols.index(h), 30 if h == 'Guide Evidence' else 16)
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
    # J Weight | K–R GLO1–8 code | S Mode | T Group | U Tools | V AoL task type | W Key | X–AE GLO1–8 rank
    out_hdr = (['Show', 'Type', 'Year', 'Trimester', 'Unit', 'Task', 'Assessment title', 'AoL',
                'Hurdle', 'Weight'] + [f'G{i}' for i in range(1, 9)]
               + ['Mode', 'Group', 'Tools', 'AoL task type', 'Key'] + [f'R{i}' for i in range(1, 9)])
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
            ws_out.write_formula(r, 0, f'=IF(COUNTIFS({key_rng},W{r + 1},{vis_rng},1)>0,1,0)', None, 1)
            ws_out.write_row(r, 1, ['Task', '', '', '', o['task'], o['title'], o['aol'], o['hurdle']])
            if o['weight'] is not None:
                ws_out.write_number(r, 9, o['weight'], pct)
            ws_out.write_row(r, 10, o['codes'])
            ws_out.write_row(r, 18, o['extra'])
            ws_out.write_string(r, 22, o['key'])
            ws_out.write_row(r, 23, o['ranks'])
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
            # roll-ups of the visible tasks: "Online 3 · Both 1", "Ind 3 · Grp 1", distinct tools
            for col_idx, labels in ((18, ('In person', 'Online', 'Both', 'Other')), (19, ('Ind', 'Grp', 'Both'))):
                cl = xl_col_to_name(col_idx)
                cnt = lambda lab: f'COUNTIFS({cl}{a}:{cl}{b},"{lab}",{show},1)'
                ws_out.write_formula(r, col_idx, '=TEXTJOIN(" · ",TRUE,' + ','.join(
                    f'IF({cnt(lab)}>0,"{lab} "&{cnt(lab)},"")' for lab in labels) + ')')
            ws_out.write_dynamic_array_formula(r, 20, r, 20,
                                               f'=IFERROR(TEXTJOIN(", ",TRUE,UNIQUE(FILTER(U{a}:U{b},'
                                               f'({show}=1)*(U{a}:U{b}<>"")))),"")')
        for g in range(8):
            rc = xl_col_to_name(23 + g)
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
    ws_sum.write_row(4, 17, ['Mode', 'Group', 'Tools', 'AoL task type'], sh)
    ws_sum.set_column('R:R', 13)      # Mode
    ws_sum.set_column('S:S', 11)      # Group
    ws_sum.set_column('T:T', 14)      # Tools
    ws_sum.set_column('U:U', 32)      # AoL task type
    ws_sum.set_row(4, 18)
    for g, name in enumerate(['Discipline-specific knowledge', 'Communication', 'Digital literacy',
                              'Critical thinking', 'Problem solving', 'Self-management', 'Teamwork',
                              'Global citizenship']):
        ws_sum.write_comment(4, 9 + g, f'GLO{g + 1} {name}', {'x_scale': 1.4, 'y_scale': 0.4})
    src = f'Outline!B2:V{last}'
    ws_sum.write_dynamic_array_formula(
        'A6', f'=FILTER(IF({src}="","",{src}),Outline!A2:A{last}=1,"No tasks match the filters on the Tasks sheet")')
    ws_sum.freeze_panes(5, 0)

    rng = (5, 0, 5 + m, 20)
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
    ws_sum.conditional_format(5, 17, 5 + m, 20, {'type': 'formula', 'criteria': '=$A6="Task"',
                                                 'format': wb.add_format({'font_color': '#374151'})})
    ws_sum.conditional_format(5, 5, 5 + m, 8, {'type': 'formula', 'criteria': '=$A6="Task"',
                                               'format': wb.add_format({'font_color': '#374151'})})
    # ════════ Timeline (AT groups) and Weeks ════════
    col_rng = lambda c_idx, last: f"{{}}!${xl_col_to_name(c_idx)}$2:${xl_col_to_name(c_idx)}${last}"
    tasks_rng = lambda name: col_rng(columns.index(name), n + 1).format('Tasks')
    title = ' · '.join(course_names) or in_path.stem
    trimesters = order_trimesters(r[c_tri] for r in rows)
    years = sorted({r[c_year] for r in rows}, key=year_key)
    if 'AT Group' in columns:
        counts = {}
        for r in rows:
            k = (r[c_year], r[c_tri], r['AT Group'] if r['AT Group'] in AT_GROUPS else 'Other')
            counts[k] = counts.get(k, 0) + 1
        build_grid(wb, ws_tl, F, title=title, years=years, trimesters=trimesters,
                   subtitle_prefix='Assessment timeline',
                   note='Tasks are grouped by AT number within each unit teaching trimester (not due dates). '
                        'Units taught in several trimesters appear in each.',
                   keys_for=lambda t: AT_GROUPS, key_label=lambda k: k, key_rule=lambda k: GROUP_RULE[k],
                   ranges={'year': tasks_rng(c_year), 'tri': tasks_rng(c_tri), 'key': tasks_rng('AT Group'),
                           'card': col_rng(card_col, n + 1).format('Tasks'),
                           'marks': [col_rng(mark_col0 + g, n + 1).format('Tasks') for g in range(8)]},
                   show={'ws': ws_tasks, 'sheet': 'Tasks', 'col': tl_show_col, 'n': n, 'vis': vis_col,
                         'marks': [mark_col0 + g for g in range(8)], 'ip': ip_col, 'on': on_col,
                         'oth': oth_col, 'wt': wt_col},
                   counts=counts)

    # WeekRows (hidden): one row per task per week; Visible follows the task's Tasks row
    nw = len(week_rows)
    ws_wr.write_row(0, 0, ['Row ID', 'Study Year', 'Trimester', 'Week', 'Card', 'Visible']
                    + [f'GLO{g} Mark' for g in range(1, 9)] + ['In person', 'Online', 'Other mode', 'Weight %',
                                                               'Show (Weeks)'])
    rowid_rng = col_rng(rowid_col, n + 1).format('Tasks')
    for i, (rid, y, t, wk, card, marks, feats) in enumerate(week_rows, start=1):
        ws_wr.write_row(i, 6, marks + feats)
        ws_wr.write_row(i, 0, [rid, y, t])
        ws_wr.write_string(i, 3, wk)
        ws_wr.write_string(i, 4, card)
        ws_wr.write_formula(i, 5, f'=IFERROR(INDEX({vis_rng},MATCH(A{i + 1},{rowid_rng},0)),0)', None, 1)
    ws_wr.hide()
    wcounts = {}
    for _, y, t, wk, *_ in week_rows:
        wcounts[(y, t, wk)] = wcounts.get((y, t, wk), 0) + 1
    sort_wk = lambda ks: sorted(ks, key=lambda k: (k == 'none', int(k) if k.isdigit() else 0))
    weeks_by_tri = {t: sort_wk({e[3] for e in week_rows if e[2] == t}) for t in trimesters}
    wr = lambda col_idx: col_rng(col_idx, nw + 1).format('WeekRows')
    n_guide = sum(1 for r in rows if re.search(r'\d', r.get('Guide Weeks', '') or ''))
    build_grid(wb, ws_wk, F, title=title, years=years, trimesters=trimesters,
               subtitle_prefix='Assessment weeks',
               note=(f'Weeks come from unit guides where the export has them ({n_guide} of {n} task rows). '
                     'Planned Weeks (filled in on Tasks) take priority over guide weeks. '
                     'Other tasks are placed by AT number (AT1 → Week 1, AT2 → Week 2 …), shown with a dashed '
                     'border; those are not due dates. A task with several guide weeks appears in each.'
                     if has_guides else
                     'This export has no unit-guide weeks (export again from dashboard v7.2+ with guides loaded). '
                     'Every task is placed by AT number (AT1 → Week 1 …); these are not due dates.'),
               keys_for=lambda t: weeks_by_tri.get(t, []),
               key_label=lambda k: 'No week' if k == 'none' else f'Week {k}',
               key_rule=lambda k: '#8A929C',
               ranges={'year': wr(1), 'tri': wr(2), 'key': wr(3), 'card': wr(4),
                       'marks': [wr(6 + g) for g in range(8)]},
               show={'ws': ws_wr, 'sheet': 'WeekRows', 'col': 18, 'n': nw, 'vis': 5,
                     'marks': list(range(6, 14)), 'ip': 14, 'on': 15, 'oth': 16, 'wt': 17},
               counts=wcounts, fallback_rule=True)

    build_balance(wb, ws_bal, F, title=title, years=years, rows=rows, c_year=c_year,
                  line_of=balance_line, glo_ids=glo_ids, marker_of=marker_of,
                  ranges={'year': tasks_rng(c_year), 'vis': vis_rng, 'sort': sort_rng,
                          'line': col_rng(line_col, n + 1).format('Tasks'),
                          'marks': [col_rng(mark_col0 + g, n + 1).format('Tasks') for g in range(8)]})
    n_guide_glo = sum(1 for r in rows if str(r.get('Guide GLOs', '') or '').strip())
    n_no_glo = sum(1 for r in rows if not any_glo_ids(r))
    build_balance(wb, ws_all, F, title=title, years=years, rows=rows, c_year=c_year,
                  line_of=all_line, glo_ids=any_glo_ids, marker_of=marker_of, mode='all',
                  source_note=(f'GLO alignment from unit guides for {n_guide_glo} of {n} task rows; the rest use CALM, '
                               f'which only aligns AoL tasks (lines marked "· CALM"). '
                               + (f'{n_no_glo} task rows have no GLO from either source and are not shown. '
                                  if n_no_glo else '')
                               + ('' if has_guide_glos else 'Export again from dashboard v7.4+ with unit guides '
                                                            'loaded to add guide alignment. ')),
                  ranges={'year': tasks_rng(c_year), 'vis': vis_rng, 'sort': sort_rng,
                          'line': col_rng(all_line_col, n + 1).format('Tasks'),
                          'marks': [col_rng(any_col0 + g, n + 1).format('Tasks') for g in range(8)]})
    build_legend(wb, ws_leg, F, title=title)
    ws_tl.activate()
    ws_tl.set_first_sheet()
    wb.close()
    return len(rows), sum(1 for o in outline if o['type'] == 'Task'), len(pathways)


ASSURANCE = [   # marker, label, fill — same four states as the dashboard timeline icons
    ('⚑', 'AoL + hurdle', '#F8D7D0'),
    ('●', 'AoL, non-hurdle', '#E4DCF3'),
    ('◆', 'Hurdle, non-AoL', '#F5E3C6'),
    ('○', 'Neither', '#F1F3F5'),
]
GLO_COLORS = ['#E76F51', '#F4A261', '#E9C46A', '#2A9D8F', '#607D8B', '#A06CD5', '#6D9DC5', '#C77DBA']  # dashboard
GLO_NAMES = ['Discipline-specific knowledge', 'Communication', 'Digital literacy', 'Critical thinking',
             'Problem solving', 'Self-management', 'Teamwork', 'Global citizenship']
FALLBACK_LINE = '≈ Week from AT number'
AT_GROUPS = ['AT1', 'AT2', 'AT3', 'Other']
GROUP_RULE = {'AT1': '#5277AD', 'AT2': '#A56D2B', 'AT3': '#56816D', 'Other': '#8A929C'}
BLOCK = 8                                               # columns per card: one per GLO


def order_trimesters(tris):
    tris = sorted(set(tris), key=tri_key)
    numbered = [t for t in tris if re.search(r'\d', t)] or tris
    return numbered + [t for t in tris if t not in numbered]   # "Not supplied" last


def build_grid(wb, ws, F, *, title, subtitle_prefix, note, years, trimesters, keys_for,
               key_label, key_rule, ranges, show, counts, fallback_rule=False):
    """Year bands × trimester × key columns (AT group or week), one live card per slot.

    Each card is 8 narrow columns wide: the card text (merged) and, on the row
    beneath, one cell per GLO showing the task's alignment mark (⚑ ● ◆ ○) in that
    GLO's dashboard colour. keys_for(tri) -> column keys under that trimester;
    ranges: year/tri/key/card/vis + marks[8] absolute source ranges;
    counts[(year, tri, key)] -> source rows (sets each band's depth)."""
    year_r, tri_r, key_r, card_r = (ranges[k] for k in ('year', 'tri', 'key', 'card'))
    marks_r = ranges['marks']
    slots = [(t, k) for t in trimesters for k in keys_for(t)]
    col0 = lambda j: 1 + j * BLOCK                      # first sheet column of slot j

    title_f = F(bold=True, font_size=14)
    sub_f = F(italic=True, font_color='#6B7280')
    tri_f = F(bold=True, bg_color='#EDF0F4', font_color='#303842', border=1, border_color='#D7DBE0')
    year_f = F(bold=True, font_size=12, valign='top', top=2, top_color='#6B7280',
               right=1, right_color='#D7DBE0')
    card_f = F(text_wrap=True, valign='top', font_size=9, left=1, left_color='#FFFFFF',
               right=1, right_color='#FFFFFF', top=1, top_color='#FFFFFF')
    mark_f = [F(align='center', valign='vcenter', font_size=10, bold=True, font_color=c,
                bottom=1, bottom_color='#FFFFFF', left=(1 if g == 0 else 0), left_color='#FFFFFF',
                right=(1 if g == BLOCK - 1 else 0), right_color='#FFFFFF')
              for g, c in enumerate(GLO_COLORS)]
    head_f = {}

    ws.hide_gridlines(2)
    ws.write('A1', title, title_f)
    # GLO filter: an Excel cell checkbox beside each GLO in the legend (row 4). Ticked
    # GLOs narrow every card, mark and count on this sheet to tasks aligned to ANY of
    # them; nothing ticked shows everything. Tasks-sheet filters still apply on top.
    boxes = [f'${xl_col_to_name(1 + BLOCK * g)}$4' for g in range(8)]
    ticked = '(' + '+'.join(boxes) + ')'
    # Mode (row 5): In person / Online / Other checkboxes, any-of; none ticked = all
    ip_box, on_box, oth_box = (f'${xl_col_to_name(1 + i * BLOCK)}$5' for i in range(3))
    # Weighting range (row 5): from/to %; unreadable weights only count while the range is 0–100
    w_min, w_max = f'${xl_col_to_name(1 + 4 * BLOCK)}$5', f'${xl_col_to_name(1 + 5 * BLOCK)}$5'

    # Every filter is worked out ONCE per source row, in a hidden "Show" column next to
    # the source data (1 = this card is shown on this sheet). Cards, marks and counts
    # then only test Show=1, which keeps formulas short and recalculation fast.
    src, here = show['ws'], f"'{ws.name}'!"
    for i in range(1, show['n'] + 1):
        ref = lambda c: f'{xl_col_to_name(c)}{i + 1}'
        glo_ok = (f'OR(({"+".join(here + b for b in boxes)})=0,('
                  + '+'.join(f'({ref(show["marks"][g])}<>"")*{here}{boxes[g]}' for g in range(8)) + ')>0)')
        mode_ok = (f'OR(({here}{ip_box}+{here}{on_box}+{here}{oth_box})=0,'
                   f'({ref(show["ip"])}*{here}{ip_box}+{ref(show["on"])}*{here}{on_box}'
                   f'+{ref(show["oth"])}*{here}{oth_box})>0)')
        wt = ref(show['wt'])
        weight_ok = (f'OR(AND({here}{w_min}<=0,{here}{w_max}>=100),'
                     f'AND({wt}<>"",{wt}>={here}{w_min},{wt}<={here}{w_max}))')
        src.write_formula(i, show['col'], f'=IF(AND({ref(show["vis"])}=1,{glo_ok},{mode_ok},{weight_ok}),1,0)',
                          None, 1)
    show_r = f"{show['sheet']}!${xl_col_to_name(show['col'])}$2:${xl_col_to_name(show['col'])}${show['n'] + 1}"
    glo_term = f'({show_r}=1)'
    shown = f'SUMPRODUCT(--({show_r}=1))'

    def header_formula(label, base, sep=' (', unit=')'):
        # "AT2 (6)  ⚑2  ●1  ◆1" — counts of the visible cards' alignment marks (first
        # character of the card text); an empty slot shows just "AT2 (0)"
        total = f'SUMPRODUCT({base})'
        marks = '&'.join(f'"  {mk}"&SUMPRODUCT({base}*(LEFT({card_r},1)="{mk}"))' for mk in ('⚑', '●', '◆'))
        formula = f'="{label}{sep}"&{total}&"{unit}"&IF({total}=0,"",{marks})'
        if len(formula) > 8000:                          # Excel's formula length limit is 8192
            raise ValueError(f'header formula too long ({len(formula)} chars)')
        return formula
    which = 'TEXTJOIN(", ",TRUE,' + ','.join(f'IF({b},"GLO{g + 1}","")' for g, b in enumerate(boxes)) + ')'
    ws.write_formula('A2', f'="{subtitle_prefix} · "&{shown}&" cards shown · "&IF({ticked}=0,'
                           f'"tick a GLO below to show only its tasks","GLO filter: "&{which}&" (untick to clear)")'
                           f'&IF({ip_box}+{on_box}+{oth_box}=0,""," · Mode: "&TEXTJOIN(" or ",TRUE,IF({ip_box},"in person",""),IF({on_box},"online",""),IF({oth_box},"other","")))'
                           f'&IF(AND({w_min}<=0,{w_max}>=100),""," · Weight "&{w_min}&"–"&{w_max}&"%")'
                           f'&" · Tasks-sheet filters also apply"', sub_f)
    # legend row 3: alignment fills and shapes; row 4: GLO colours
    c = 1
    for marker, label, fill in ASSURANCE:
        ws.merge_range(2, c, 2, c + 7, f'{marker} {label}',
                       F(bg_color=fill, font_size=9, border=1, border_color='#FFFFFF'))
        c += 8
    if fallback_rule:
        ws.merge_range(2, c, 2, c + 15, FALLBACK_LINE + ' (no guide week): dashed border',
                       F(font_size=9, border=3, border_color='#8A929C'))
    c = 1
    for g, (name, colour) in enumerate(zip(GLO_NAMES, GLO_COLORS)):
        ws.insert_checkbox(3, c, False, F(align='center'))
        ws.merge_range(3, c + 1, 3, c + 7, f'● GLO{g + 1} {name}', F(font_size=9, bold=True, font_color=colour))
        ws.conditional_format(3, c + 1, 3, c + 7, {
            'type': 'formula', 'criteria': f'={boxes[g]}=TRUE',
            'format': wb.add_format({'bg_color': '#EEF0F3', 'underline': 1})})
        c += 8
    ws.set_row(3, 18)
    # row 5: mode checkboxes and weighting range
    lab = F(font_size=9, bold=True, font_color='#303842')
    inp = F(font_size=10, bold=True, align='center', bg_color='#FFF8DC', border=1, border_color='#C8CDD3')
    ws.write(4, 0, 'Filter', F(font_size=9, bold=True, font_color='#6B7280'))
    mode_labels = ['In person (campus, site, field, placement)', 'Online', 'Other / unspecified mode']
    for i, label in enumerate(mode_labels):
        ws.insert_checkbox(4, 1 + i * BLOCK, False, F(align='center'))
        ws.merge_range(4, 2 + i * BLOCK, 4, (i + 1) * BLOCK, label, lab)
    ws.merge_range(4, 1 + 3 * BLOCK, 4, 4 * BLOCK, 'Weighting from (%)', F(font_size=9, bold=True, align='right',
                                                                           font_color='#303842'))
    ws.merge_range(4, 1 + 4 * BLOCK, 4, 3 + 4 * BLOCK, 0, inp)
    ws.merge_range(4, 4 + 4 * BLOCK, 4, 5 * BLOCK, 'to', F(font_size=9, bold=True, align='center'))
    ws.merge_range(4, 1 + 5 * BLOCK, 4, 3 + 5 * BLOCK, 100, inp)
    ws.merge_range(4, 4 + 5 * BLOCK, 4, 6 * BLOCK, '%  (unreadable weights show only at 0–100)', F(font_size=9,
                                                                                               font_color='#6B7280'))
    for cc in (1 + 4 * BLOCK, 1 + 5 * BLOCK):
        ws.data_validation(4, cc, 4, cc, {'validate': 'decimal', 'criteria': 'between', 'minimum': 0,
                                          'maximum': 100, 'error_title': 'Weighting',
                                          'error_message': 'Enter a percentage from 0 to 100.'})
    for i, box in enumerate((ip_box, on_box, oth_box)):
        ws.conditional_format(4, 2 + i * BLOCK, 4, (i + 1) * BLOCK, {'type': 'formula', 'criteria': f'={box}=TRUE',
                              'format': wb.add_format({'bg_color': '#EEF0F3', 'underline': 1})})
    ws.set_row(4, 18)
    ws.write(5, 1, note + ' Under each card, one cell per GLO (GLO1→GLO8, left to right) shows the '
                          'task\'s mark in that GLO\'s colour.', F(font_color='#6B7280', font_size=9))

    ws.set_column(0, 0, 7)
    ws.set_column(1, max(len(slots), 1) * BLOCK, 2.9)
    head = 7                                            # row 8: trimester headers
    col = 1
    for t in trimesters:
        width = len(keys_for(t)) * BLOCK
        if width:
            ws.merge_range(head, col, head, col + width - 1, '', tri_f)
            ws.write_formula(head, col, header_formula(t, f'({tri_r}="{t}")*{glo_term}',
                                                       sep=' · ', unit=' cards'), tri_f)
        col += width
    ws.write(head, 0, 'Year', tri_f)
    ws.freeze_panes(head + 1, 1)

    strip_rows = []
    row = head + 1
    for y in years:
        depth = max([counts.get((y, t, k), 0) for t, k in slots] + [1])
        ws.write(row, 0, y, year_f)
        for j, (t, k) in enumerate(slots):
            rule = key_rule(k)
            if rule not in head_f:
                head_f[rule] = F(bold=True, font_size=9, font_color='#59636F', bottom=2,
                                 bottom_color=rule, top=2, top_color='#6B7280')
            ws.merge_range(row, col0(j), row, col0(j) + BLOCK - 1, '', head_f[rule])
            base = f'({year_r}="{y}")*({tri_r}="{t}")*({key_r}="{k}")*{glo_term}'
            ws.write_formula(row, col0(j), header_formula(key_label(k), base), head_f[rule])
        ws.set_row(row, 16)
        r = row + 1
        for d in range(1, depth + 1):
            ws.set_row(r, 54)                           # card text
            ws.set_row(r + 1, 14)                       # GLO marks strip
            ws.write_blank(r, 0, None, F(right=1, right_color='#D7DBE0'))
            ws.write_blank(r + 1, 0, None, F(right=1, right_color='#D7DBE0'))
            strip_rows.append(r + 1)
            for j, (t, k) in enumerate(slots):
                crit = f'({year_r}="{y}")*({tri_r}="{t}")*({key_r}="{k}")*{glo_term}'
                ws.merge_range(r, col0(j), r, col0(j) + BLOCK - 1, '', card_f)
                ws.write_dynamic_array_formula(r, col0(j), r, col0(j),
                                               f'=IFERROR(INDEX(FILTER({card_r},{crit}),{d}),"")', card_f)
                for g in range(BLOCK):
                    ws.write_dynamic_array_formula(r + 1, col0(j) + g, r + 1, col0(j) + g,
                                                   f'=IFERROR(INDEX(FILTER({marks_r[g]},{crit}),{d})&"","")',   # &"": empty → blank, not 0
                                                   mark_f[g])
            r += 2
        row = r

    last_col = max(len(slots), 1) * BLOCK
    card_area = (head + 1, 1, row, last_col)
    # card fill by alignment (card cell text starts with the mark) …
    for marker, _, fill in ASSURANCE:
        ws.conditional_format(*card_area, {'type': 'text', 'criteria': 'begins with', 'value': marker,
                                           'format': wb.add_format({'bg_color': fill})})
    # … and the same fill on the GLO strip beneath, keyed off the card cell above it
    if strip_rows:
        for j in range(len(slots)):
            a, b = xl_col_to_name(col0(j)), xl_col_to_name(col0(j) + BLOCK - 1)
            multi = ' '.join(f'{a}{r + 1}:{b}{r + 1}' for r in strip_rows)
            first = strip_rows[0]
            for marker, _, fill in ASSURANCE:
                ws.conditional_format(first, col0(j), first, col0(j) + BLOCK - 1, {
                    'type': 'formula', 'criteria': f'=LEFT(${a}{first},1)="{marker}"',
                    'multi_range': multi, 'format': wb.add_format({'bg_color': fill})})
    if fallback_rule:
        ws.conditional_format(*card_area, {'type': 'text', 'criteria': 'containing', 'value': FALLBACK_LINE,
                                           'format': wb.add_format({'border': 3, 'border_color': '#8A929C',
                                                                    'font_color': '#4B5563'})})
    ws.set_landscape()
    ws.fit_to_pages(1, 0)
    ws.repeat_rows(head, head)


YEAR_LEVEL = {'Hons': 'Y4 (Hons)', 'PG': 'Y5 (PG)'}     # Balance column labels
# Task types as short codes. Add a type with one line: (code, label, regex). A task gets
# the type whose words appear FIRST in Assessment type (or CALM Student Output until that
# is filled in). Each line ends with ' §<code>', read back with TEXTAFTER(line,"§",-1);
# § never occurs in the data.
TASK_TYPES = [
    ('W', 'Written', r'written|report|essay|exam|test|quiz|portfolio|assignment|reflect|paper|proposal|journal'),
    ('IO', 'Interactive oral', r'interactive oral|viva|interview|\boral\b'),
    ('Pres', 'Presentation', r'presentation|present\b|poster|pitch|seminar|podcast|video'),
    ('QA', 'Q&A', r'q\s*&\s*a\b|questions? and answers?|questioning'),
    ('Prac', 'Practical', r'practical|\blab|skill|demonstrat|simulation|placement|field|experiment|clinical'),
    ('Perf', 'Performance', r'performance|perform\b|exhibition|showcase|recital'),
]
TYPE_OTHER = ('Oth', 'Other')
TYPE_DIVIDER = ' §'
TYPE_LABELS = [(code, label) for code, label, _ in TASK_TYPES] + [TYPE_OTHER]
AOL_TARGET = 2


def tint(hex_colour, amount=0.55):
    """Blend a #RRGGBB colour towards white (for the faded, not-counted marks)."""
    r, g, b = (int(hex_colour[i:i + 2], 16) for i in (1, 3, 5))
    mix = lambda v: round(v + (255 - v) * amount)
    return f'#{mix(r):02X}{mix(g):02X}{mix(b):02X}'


def build_balance(wb, ws, F, *, title, years, rows, c_year, line_of, glo_ids, marker_of, ranges,
                  mode='aol', source_note=''):
    """GLO × year level: every GLO should have at least 2 AoL tasks (⚑ or ●) at every level.

    Each cell: AoL count against the target, one mark per AoL task in the GLO's colour,
    the not-counted marks (◆ hurdle only, ○ neither) faded, and a line per AoL task
    (unit · AT · assessment type · weight). Live: follows filters on the Tasks sheet.
    Tasks repeated across trimesters are counted once (identical lines are de-duplicated)."""
    year_r, vis_r, line_r, marks_r = ranges['year'], ranges['vis'], ranges['line'], ranges['marks']
    first_char = f'LEFT({line_r},1)'
    is_aol = f'(({first_char}="⚑")+({first_char}="●"))'
    all_mode = mode == 'all'                            # All tasks sheet: every task, no AoL target
    keep = '' if all_mode else f'*{is_aol}'
    counted = (lambda r: True) if all_mode else (lambda r: marker_of(r) in ('⚑', '●'))

    # expected lines per GLO × year (unfiltered) to size the list rows
    max_lines = {}
    for g in range(1, 9):
        for y in years:
            lines = {line_of(r) for r in rows if r[c_year] == y and g in glo_ids(r) and counted(r)}
            max_lines[g] = max(max_lines.get(g, 0), len(lines))

    title_f = F(bold=True, font_size=14)
    sub_f = F(italic=True, font_color='#6B7280')
    head_f = F(bold=True, bg_color='#EDF0F4', font_color='#303842', border=1, border_color='#D7DBE0',
               align='center')
    status_f = F(bold=True, align='center', valign='top',
                 num_format='0" tasks"' if all_mode else f'"AoL "0" of {AOL_TARGET}"',
                 top=2, top_color='#6B7280')
    other_lab_f = F(font_size=8, font_color='#9AA0A6', align='center', valign='top')
    list_f = F(font_size=9, text_wrap=True, valign='top', bottom=1, bottom_color='#E5E7EB')

    ws.hide_gridlines(2)
    ws.write('A1', title, title_f)
    intro = (f'Every task under each GLO it addresses, by year level: ⚑ AoL + hurdle, ● AoL, ◆ hurdle only, '
             f'○ neither, in the GLO\'s colour. {source_note}' if all_mode else
             f'Target: every GLO assessed in at least {AOL_TARGET} AoL tasks (⚑ AoL + hurdle, ● AoL) at every '
             f'year level (CALM records GLO alignments on AoL tasks only). ')
    ws.write(3, 0, 'Task types: ' + ' · '.join(f'{c} {t}' for c, t in TYPE_LABELS) + '  — the first type mentioned '
                   'in Assessment type (or CALM Student Output until that is filled in). Codes under the marks are '
                   'in the same order as the tasks.', F(font_size=9, font_color='#4B5563'))
    ws.write(2, 0, intro + 'Type* = CALM Student Output, shown until the Assessment type column is filled in. '
                           'Follows filters on the Tasks sheet (e.g. one pathway\'s ● column = one major).',
             F(font_size=9, font_color='#4B5563'))
    ws.set_column(0, 0, 24)
    head = 5
    ws.write(head, 0, 'GLO', head_f)
    status_cells = []
    for j, y in enumerate(years):
        sc = 1 + 2 * j
        ws.set_column(sc, sc, 10)
        ws.set_column(sc + 1, sc + 1, 44)
        ws.merge_range(head, sc, head, sc + 1, YEAR_LEVEL.get(y, y), head_f)
    # whole-year type mix (AoL tasks on Balance, every task on All tasks), unique tasks
    mix_f = F(font_size=10, bold=True, align='center', valign='vcenter', bg_color='#F7F8FA',
              bottom=1, bottom_color='#D7DBE0')
    ws.write(head + 1, 0, 'Type mix (year)', F(bold=True, font_size=9, font_color='#4B5563', bg_color='#F7F8FA',
                                               bottom=1, bottom_color='#D7DBE0'))
    mix_cells = []                                      # filled in after the GLO bands (helper rows)
    for j, y in enumerate(years):
        sc = 1 + 2 * j
        ws.merge_range(head + 1, sc, head + 1, sc + 1, '', mix_f)
        mix_cells.append((sc, y))
    ws.set_row(head + 1, 22)
    ws.freeze_panes(head + 2, 1)

    row = head + 2
    for g in range(8):
        colour = GLO_COLORS[g]
        mG = marks_r[g]
        ws.merge_range(row, 0, row + 1, 0, f'● GLO{g + 1}\n{GLO_NAMES[g]}',
                       F(bold=True, font_color=colour, text_wrap=True, valign='top', top=2, top_color='#6B7280'))
        for j, y in enumerate(years):
            sc, dc = 1 + 2 * j, 2 + 2 * j
            base = f'({year_r}="{y}")*({vis_r}=1)*({mG}<>"")'
            # sorted by unit code then AT number; UNIQUE keeps that order (first occurrence)
            aol = f'SORTBY(FILTER({line_r},{base}{keep}),FILTER({ranges["sort"]},{base}{keep}))'
            ws.write_dynamic_array_formula(row, sc, row, sc, f'=IFERROR(ROWS(UNIQUE({aol})),0)', status_f)
            status_cells.append(xl_rowcol_to_cell(row, sc))
            # marks (⚑ ● ◆ ○ in the GLO colour) and, beneath, each task's type code — same order
            ws.write_dynamic_array_formula(row, dc, row, dc,
                                           f'=IFERROR(CONCAT(LEFT(UNIQUE({aol}),1))&CHAR(10)&'
                                           f'TEXTJOIN(" ",TRUE,TEXTAFTER(UNIQUE({aol}),"§",-1)),"")',
                                           F(bold=True, font_size=13, font_color=colour, valign='top', text_wrap=True,
                                             top=2, top_color='#6B7280'))
            # (CALM only aligns GLOs on AoL tasks, so ◆/○ never carry a GLO here; if planning
            #  adds such alignments they'd need their own row again — see tint())
            ws.merge_range(row + 1, sc, row + 1, dc, '', list_f)
            ws.write_dynamic_array_formula(row + 1, sc, row + 1, sc,
                                           f'=IFERROR(TEXTJOIN(CHAR(10),TRUE,UNIQUE({aol})),'
                                           f'"{"No task" if all_mode else "No AoL task"}")', list_f)
        ws.set_row(row, 38)
        ws.set_row(row + 1, max(14, 12.5 * max_lines[g + 1] + 4))
        row += 2

    # hidden helper rows: unique-task counts per year by type code, then the short mix text
    help_row = row + 2
    for k, (code, _) in enumerate(TYPE_LABELS + [('', 'total')]):
        ws.write(help_row + k, 0, code or 'total', F(font_color='#9AA0A6', font_size=8))
        ws.set_row(help_row + k, None, None, {'hidden': True})
    for sc, y in mix_cells:
        base = f'({year_r}="{y}")*({vis_r}=1){keep}'
        for k, (code, _) in enumerate(TYPE_LABELS):
            ws.write_dynamic_array_formula(help_row + k, sc, help_row + k, sc,
                                           f'=IFERROR(ROWS(UNIQUE(FILTER({line_r},{base}*'
                                           f'(TEXTAFTER({line_r},"§",-1)="{code}")))),0)')
        tot_row = help_row + len(TYPE_LABELS)
        ws.write_dynamic_array_formula(tot_row, sc, tot_row, sc,
                                       f'=IFERROR(ROWS(UNIQUE(FILTER({line_r},{base}))),0)')
        tot = xl_rowcol_to_cell(tot_row, sc)
        parts = '&'.join(f'IF({xl_rowcol_to_cell(help_row + k, sc)}>0," · {code} "&'
                         f'TEXT({xl_rowcol_to_cell(help_row + k, sc)}/{tot},"0%"),"")'
                         for k, (code, _) in enumerate(TYPE_LABELS))
        ws.write_formula(head + 1, sc, f'=IF({tot}=0,"—",MID({parts},4,999))', mix_f)

    if all_mode:                                        # only gaps are flagged: no task at all
        for cell in status_cells:
            r, c = cell_to_rowcol(cell)
            ws.conditional_format(r, c, r, c, {'type': 'cell', 'criteria': '==', 'value': 0,
                                               'format': wb.add_format({'bg_color': '#F8D7D0',
                                                                        'font_color': '#9B2C1A'})})
        gaps = '+'.join(f'({c}=0)' for c in status_cells)
        ws.write_formula('A2', f'="All tasks by GLO · "&({gaps})&" of {len(status_cells)} GLO × year cells have no '
                               f'task · follows filters on the Tasks sheet"', sub_f)
        ws.set_landscape()
        ws.fit_to_pages(1, 0)
        ws.repeat_rows(head, head)
        return
    # status colour: met / one short / none
    for cell in status_cells:
        r, c = cell_to_rowcol(cell)
        for crit, value, fill, font in (('>=', AOL_TARGET, '#D7EFD9', '#1E6B34'),
                                        ('between', None, '#FDE7B0', '#8A5A00'),
                                        ('==', 0, '#F8D7D0', '#9B2C1A')):
            rule = {'type': 'cell', 'criteria': crit,
                    'format': wb.add_format({'bg_color': fill, 'font_color': font})}
            if crit == 'between':                       # 1 … target-1: some AoL, not enough
                rule.update(minimum=1, maximum=AOL_TARGET - 1)
            else:
                rule['value'] = value
            ws.conditional_format(r, c, r, c, rule)
    met = '+'.join(f'({c}>={AOL_TARGET})' for c in status_cells)
    ws.write_formula('A2', f'="Assessment balance · "&({met})&" of {len(status_cells)} GLO × year cells have '
                           f'{AOL_TARGET}+ AoL tasks · follows filters on the Tasks sheet"', sub_f)
    ws.set_landscape()
    ws.fit_to_pages(1, 0)
    ws.repeat_rows(head, head)


def build_legend(wb, ws, F, *, title):
    """Quick guide and legend (last tab)."""
    h1 = F(bold=True, font_size=14)
    h2 = F(bold=True, font_size=11, font_color='#303842', top=2, top_color='#6B7280')
    txt = F(font_size=10, text_wrap=True, valign='top')
    key = F(font_size=10, bold=True, valign='top')
    ws.hide_gridlines(2)
    ws.set_column(0, 0, 4)
    ws.set_column(1, 1, 26)
    ws.set_column(2, 2, 90)
    ws.write('B1', 'Quick guide', h1)
    ws.write('B2', title, F(italic=True, font_color='#6B7280'))
    row = 3

    def section(name):
        nonlocal row
        row += 1
        ws.merge_range(row, 1, row, 2, name, h2)
        row += 1

    def item(k, v, kf=key):
        nonlocal row
        ws.write(row, 1, k, kf)
        ws.write(row, 2, v, txt)
        row += 1

    section('Sheets')
    item('Timeline', 'Every task as a card by study year × trimester × AT1/AT2/AT3/Other (not due dates).')
    item('Weeks', 'The same cards by week: Planned Weeks first, then unit-guide weeks, then the AT number '
                  '(AT1 → Week 1 …, dashed border, not a due date).')
    item('AOL type balance', 'GLO × year level. Target: every GLO assessed in at least 2 AoL tasks (⚑ or ●) at '
                             'every level. Lists each AoL task: unit · AT · type · weight.')
    item('All tasks', 'Same layout, every task under each GLO it addresses (unit-guide alignment where '
                      'available, otherwise CALM). Empty cells are flagged.')
    item('Summary', 'Year/trimester → unit → task list with GLO codes.')
    item('Tasks', 'The dashboard export, one row per task. Filter here and every view follows. '
                  'Planning columns are typed in here.')

    section('Alignment marks (cards, strips, balance)')
    for marker, label, fill in ASSURANCE:
        ws.write(row, 1, marker + '  ' + label, F(font_size=10, bold=True, bg_color=fill))
        ws.write(row, 2, {'⚑': 'AoL task that is also a hurdle. Counts towards the AoL target.',
                          '●': 'AoL task, not a hurdle. Counts towards the AoL target.',
                          '◆': 'Hurdle task, not AoL. Does not count towards the AoL target.',
                          '○': 'Neither AoL nor hurdle.'}[marker], txt)
        row += 1
    item('Card fill', 'Each card is tinted by its mark (colours above).')

    section('GLO colours')
    item('', 'Under each card, eight small cells (GLO1 → GLO8, left to right) show the task\'s mark in that '
             'GLO\'s colour. On the balance sheets the marks are in the row\'s GLO colour.')
    for g, (name, colour) in enumerate(zip(GLO_NAMES, GLO_COLORS)):
        item(f'● GLO{g + 1}', name, F(font_size=10, bold=True, font_color=colour))

    section('Task type codes')
    item('', 'Each task gets the type mentioned first in Assessment type, or CALM Student Output until that is '
             'filled in (marked *). Codes under the marks are in the same order as the task list; each task line '
             'ends with § and its code. The Type mix row shows each year\'s proportions.')
    for code, label, pattern in TASK_TYPES:
        item(code, f'{label}  (matches: {pattern.replace(chr(92) + "b", "").replace("|", ", ")})')
    item(TYPE_OTHER[0], TYPE_OTHER[1] + ' (none of the above)')

    section('Balance status')
    ws.write(row, 1, 'AoL 2 of 2', F(font_size=10, bold=True, bg_color='#D7EFD9', font_color='#1E6B34'))
    ws.write(row, 2, 'Target met (2 or more AoL tasks).', txt); row += 1
    ws.write(row, 1, 'AoL 1 of 2', F(font_size=10, bold=True, bg_color='#FDE7B0', font_color='#8A5A00'))
    ws.write(row, 2, 'Some AoL, below target.', txt); row += 1
    ws.write(row, 1, 'AoL 0 of 2', F(font_size=10, bold=True, bg_color='#F8D7D0', font_color='#9B2C1A'))
    ws.write(row, 2, 'No AoL task for this GLO at this level (on All tasks: no task at all).', txt); row += 1
    item('· CALM', 'On All tasks: GLOs come from CALM because no unit guide covered this task.')

    section('Card lines')
    item('Guide: Wk 6 (Quiz 2)', 'Week from the unit guide ("When and how to submit your work").')
    item('Planned: Wk 6', 'Week from the Planned Weeks column.')
    item('≈ Week from AT number', 'No guide or planned week; placed by AT number for layout only.')
    item('Type: … / Tools: …', 'From the Assessment type and Tools planning columns.')

    section('Filters')
    item('GLO checkboxes', 'Timeline and Weeks: tick one or more GLOs to show tasks aligned to any of them; '
                           'untick all to show everything.')
    item('Mode checkboxes', 'In person, Online, Other / unspecified (any of the ticked).')
    item('Weighting from–to', 'Type a % range; tasks with unreadable weights only show at 0–100.')
    item('Tasks filters', 'Filter any Tasks column (e.g. a pathway\'s ● column for one major, Study Year, '
                          'Alignment Mark). Every sheet follows.')

    section('Planning columns (Tasks sheet)')
    item('Tools', 'Tech used, e.g. FFIO, OT. Shown on cards.')
    item('Assessment type', 'The new task definitions. Used for type codes and balance lines instead of CALM.')
    item('Planned Weeks', 'Planned submission week(s), e.g. 6 or 6; 10. Takes priority on the Weeks sheet.')
    item('', 'These columns are read back into the dashboard (v7.3+) when the workbook is loaded as a '
             'planning workbook.')
    ws.set_landscape()
    ws.fit_to_pages(1, 0)


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
