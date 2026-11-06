# Curriculum Dashboard — Data Schema v2 (CALM-native)

Supersedes [old/SCHEMA-v1.md](old/SCHEMA-v1.md). Describes the **native CALM export format** read by dashboard v3.0 and later. Current as of **dashboard v5.0 / CALM export of 2026-10-27**.

For the step-by-step update process, see [HOW-TO-UPDATE-THE-DASHBOARD.md](HOW-TO-UPDATE-THE-DASHBOARD.md).

---

## Pipeline

```
CALM (Power BI)  →  CALM-dashboard-YYYYMMDD/*.xlsx   (3 exports, all faculties)
                 →  update-dashboard-data.py           (or "Extract SEBE.command";
                                                        syncs folders, runs extract-faculty.py SEBE)
                 →  CALM-dashboard-YYYYMMDD/SEBE/*.csv (same 3 files, SEBE only)
                    + CALM-dashboard-YYYYMMDD/SEBE-statistics.xlsx (make-statistics.py; same rules)
                 →  curriculum-dashboard-v5.html       (drag & drop)
```

The dashboard also accepts the full `.xlsx` exports directly. The extract step only reduces the volume.

---

## Source files

Every CALM export has one sheet, `Export` (matched case-insensitively), and ends with a footer row whose first cell is **`No filters applied`**. The dashboard and the extract script both skip that row.

Files are identified by their **headers**, not their names. The default CALM names are used throughout.

### File 1: `Student Pathways by Units`

One row per unit in each pathway, with the unit's GLO/CLO status.

```
Student Pathway | Unit Level of Study | Unit Code | GLO1 / CLO1 | … | GLO8 / CLO8
```

| Column | Notes |
|---|---|
| `Student Pathway` | `<course>_<pathway>`, for example `A250_Core Only`, `A310_Core + MJ-A310008 : English - Children's Literature`. The course code is everything before the first separator. Pathway types include Core Only, Core + MJ/ST/SP and MJ/ST/SP Only. |
| `Unit Level of Study` | Present but **not used**. See the year rule below. |
| `Unit Code` | For example `SLE111`. Join key to File 2. |
| `GLO1 / CLO1` … `GLO8 / CLO8` | One status per cell (see the vocabulary below). Note that CALM writes `GLO4  / CLO4` with two spaces, which the header matcher tolerates. A unit's status is the same in every pathway (verified on the Oct 2026 data), so the first row seen is used. |

Renamed from `Courses by Units` / `pathways by units` in earlier exports. The `2026 Teaching Periods` column that the April 2026 `Course by Units` export carried is **gone**.

### File 2: `Units by Assessment Tasks`

One row per assessment task.

```
Unit Code | Task Number | Assessment Title | AoL Task Type | Student Output | Mode | Participation | Hurdle | Weighting Information | GLO1 / CLO1 | … | GLO8 / CLO8
```

| Column | Notes |
|---|---|
| `Unit Code` | Join key to File 1. |
| `Task Number` | For example `SPB-AT-1`. Unique within a unit. (Older exports called this `AT Number`.) |
| `Assessment Title` | Text. (Older exports: `AT Title`.) |
| `AoL Task Type` | The **kind** of AoL task (Examination, Portfolio, Live oral presentation with Q&A, …). **Any value other than** `NA`, `N/A`, `Not applicable`, `TBC`, `No`, `Not AoL Task` or blank makes the task an AoL task. A whole-word `CLO` gives the CLO tier. |
| `Student Output`, `Mode`, `Participation` | Text. |
| `Hurdle` | Mostly `Yes` / `No`, but CALM contains free text (`Yes Part A`, `Yes report only`, `YNo`, `Not assured`, `10percent`, `NA`, `TBC`). A task is a hurdle if the value contains "yes" and not "non". |
| `Weighting Information` | Free text. Occasionally contains hurdle requirement wording, but there is **no dedicated field** for it. |
| `GLO1 / CLO1` … `GLO8 / CLO8` | Same vocabulary as File 1. |

### File 3: `Course Data` (optional)

One row per course. It supplies course names and faculty/school, and replaces the hand-maintained `course-titles.csv` (now in `old/`).

```
Course | Faculty | School | Course Name | Hurdle Count | Assessment Tasks
```

| Column | Notes |
|---|---|
| `Course` | Matches the course code prefix of `Student Pathway`. Codes can have 4 or 5 characters (`D3100`, `S737A`). |
| `Faculty` | One of the 4 faculties. The extract script filters on this. |
| `School` | 22 schools. Business & Law lists departments (Finance, Marketing, …). |
| `Course Name` | Shown as the pathway group heading in the dashboard. |
| `Hurdle Count`, `Assessment Tasks` | CALM's own counts. Not used by the dashboard. |

Student enrolments are **no longer exported** (the April 2026 Course Data had them).

### Other files in the export folder (not used)

| File | Contents |
|---|---|
| `CALM_AoL_data_model.xlsx` | Power BI semantic-model field list. Shows that `at_element_teaching_period` and `at_element_year` exist in CALM but are not exported. |
| `Unit Relationship Details.xlsx` | Unit × course × faculty (short) × school × pathway. Shows that units are shared across faculties: 544 of 1,726 units appear in courses from more than one faculty. |
| `Unit Relationship Summary.xlsx` | Count of related courses and pathways per unit. |
| `data.xlsx` | List of pathway names. |
| `data (1).xlsx` | List of CALM error types. |

### Legacy: `course-titles.csv`

Still accepted (`pathway_prefix, course_title, enrolment`) but superseded by Course Data.

---

## Outcome cell vocabulary

Each `GLOx / CLOx` cell holds **one** value, independent of the other seven.

| Value | Meaning | Visual (chip) |
|---|---|---|
| `Yes` | Assessed, hurdle | Filled |
| `Hurdle Yes` | Assessed, hurdle (explicit) | Filled |
| `Non-hurdle Yes` | Assessed, no hurdle | Outlined |
| `No` | Not assessed | Faded |
| `Not Assessed` | Explicitly not assessed | Faded |
| `Data Error` | CALM flagged broken data for this outcome | **Red** chip |
| `TBC` | Outcome not yet mapped | **Grey-blue** chip |
| `Form not received` | Legacy unit-level state (v1 data only) | ⚠ amber icon |

---

## Year derivation rule

The ring comes from the **first digit of the numeric part of `Unit Code`**. `Unit Level of Study` is ignored.

| Unit code | Ring | Label |
|---|---|---|
| `XXX0xx` | Non-major (mandatory but outside the major) | Non-major |
| `XXX1xx` / `2xx` / `3xx` | Year 1 / 2 / 3 | Y1 / Y2 / Y3 |
| `XXX4xx` | Honours | Hons |
| `XXX7xx` | Postgraduate (coursework) | PG |
| `XXX8xx` | Masters by Research | MRes |
| `XXX9xx` | PhD | PhD |
| anything else | Unclassified | ? |

Ring order, bottom to top: Non-major › Y1 › Y2 › Y3 › Hons › PG › MRes › PhD › ?. Empty rings are hidden.

---

## Teaching periods (stopgap)

CALM does not export teaching periods. Until it does, the dashboard uses a **built-in lookup** (`APR2026_TEACHING_PERIODS` in the HTML) taken from the `2026 Teaching Periods` column of the April 2026 export:

- 511 units. These cover 486 of 1,726 units overall, and 473 of the 532 SEBE units.
- These values are shown with a trailing `*` (for example `TRI-1, TRI-3*`).
- Units not in the lookup show `T0`.
- If a tasks export ever includes a teaching period column (`Teaching Period`, `Teaching Periods`, `2026 Teaching Periods` or `at_element_teaching_period`), it is used instead of the lookup.

Periods are displayed only. Units are spaced evenly around their year ring, not positioned by trimester.

---

## Importer behaviour

On upload, the dashboard checks each file's headers to decide what it is:

| If headers contain… | Treated as |
|---|---|
| `pathway_prefix` + `course_title` | Course titles (legacy CSV) |
| `Course` + `Course Name`, no GLO columns | Course titles (CALM Course Data) |
| Unit/AT-prefixed GLOs + `Student Pathway` + `AT Number` | Legacy v1 merged CSV |
| `GLOx / CLOx` + `Task Number` + `Assessment Title` | **File 2** (tasks) |
| `GLOx / CLOx` + `Student Pathway` | **File 1** (pathways) |

`.xlsx` files are read from the sheet named `export` (or the first sheet). `.csv` files must be UTF-8 **without** a byte-order mark. Units that appear only in the tasks file are still shown, with no pathways.

---

## Known gaps

1. **No teaching periods** in CALM. The stopgap is described above.
2. **No hurdle requirement wording.** Only the Yes/No flag (plus stray free text) is available.
3. **No unit owner.** Faculty and school belong to courses. A unit is included in a faculty extract if any course of that faculty uses it.
4. **No separate CLO values.** The paired `GLOx / CLOx` cells collapse to one status per outcome.
5. **Double degrees** belong to a single faculty in Course Data, so a faculty extract includes only the double degrees that faculty owns.
6. **No assessment due week, separate tool, or assessment location** in the current exports. The v6 timeline leaves due week unassigned and does not infer tools or locations from other text fields.

---

## Version

- **Schema v2** (this document): CALM-native. Revised 2026-10-29 for the Oct 2026 exports (new file names, Course Data, footer row, year levels 4/8/9, AoL rule, teaching period stopgap).
- **Dashboard v3.0+** reads schema v2. Dashboards v2.0–v2.7 read schema v1.
