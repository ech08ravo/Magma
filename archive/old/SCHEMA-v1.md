# Curriculum Dashboard v1 — Data Schema

## Overview
Combine **two CSV sheets** to create a unified curriculum model with assessment tasks, pathway filtering, and trimester-level detail.

---

## Sheet 1: Assessment Tasks
Maps individual assessment tasks to GLO/CLO outcomes.

### Columns:
```
unit_code, at_number, at_title, student_output, aol_task_type, mode, participation, hurdle, weighting, glo1, glo2, glo3, glo4, glo5, glo6, glo7, glo8
```

### Data Types:
- `unit_code`: Text (e.g., "ACG703", "SLE111")
- `at_number`: Integer (1, 2, 3... assessment task sequence)
- `at_title`: Text (e.g., "Design Project", "Lab Portfolio")
- `student_output`: Text (e.g., "Written, Practical, Design Artefact")
- `aol_task_type`: Text or "NA"
- `mode`: Text (e.g., "Online", "On-campus", "Online, On-campus")
- `participation`: Text (e.g., "Individual", "Group", "Individual, Group")
- `hurdle`: Boolean (Yes/No) — **triggers glow effect**
- `weighting`: Text (e.g., "30 percent of the total mark")
- `glo1–glo8`: One of:
  - `"Yes"` (assessed, no hurdle)
  - `"Hurdle Yes"` or `"Non-hurdle Yes"` (assessed, with/without hurdle)
  - `"No"` (not assessed)
  - `"Not Assessed"` (explicitly not assessed)

### Example Row:
```
ACG703, 1, Design Project, "Written, Practical, Other: Design Artefact", NA, "Online, On-campus", Individual, Yes, "30 percent of the total mark", Hurdle Yes, No, No, No, No, Not Assessed, Not Assessed, Not Assessed
```

---

## Sheet 2: Unit Coverage
Maps units to pathways, years, trimesters, and GLO/CLO coverage.

### Columns:
```
pathway_id, unit_level, unit_code, teaching_periods, glo1, glo2, glo3, glo4, glo5, glo6, glo7, glo8
```

### Data Types:
- `pathway_id`: Text (e.g., "D306_Core Only", "D307_Core Only", "D306_All")
- `unit_level`: Integer (1, 2, or 3 for Year 1–3)
- `unit_code`: Text (matches Sheet 1)
- `teaching_periods`: Text, comma-separated (e.g., "TRI-1", "TRI-2,TRI-3", "TRI-1,TRI-2,TRI-3")
- `glo1–glo8`: Same as Sheet 1 assessment mapping

### Example Rows:
```
D306_Core Only, 1, SRA143, TRI-1, "Non-hurdle Yes", "Non-hurdle Yes", No, No, No, No, "Non-hurdle Yes", "Not Assessed"
D307_Core Only, 1, MAE101, "TRI-1, TRI-3, TRI-2", "Non-hurdle Yes", "Not Assessed", No, "Non-hurdle Yes", "Not Assessed", "Not Assessed", "Not Assessed", No
```

---

## Merge Logic
1. **Group by Unit Code** (left join: Sheet 2 → Sheet 1)
2. For each unit in a pathway:
   - Inherit all assessment tasks from Sheet 1
   - Apply pathway-level GLO/CLO coverage from Sheet 2
   - Apply teaching period info from Sheet 2
   - Track hurdle status (from both sheets)

### Result Structure (internal):
```json
{
  "pathway_id": "D306_Core Only",
  "unit_code": "SRA143",
  "unit_level": 1,
  "teaching_periods": ["TRI-1"],
  "assessment_tasks": [
    {
      "at_number": 1,
      "at_title": "Design Project",
      "hurdle": true,
      "weighting": "30%",
      "glo_mappings": [1, 3, 4],
      "clo_mappings": [1, 6, 7]
    }
  ],
  "unit_glo_coverage": {
    "glo1": "Non-hurdle Yes",
    "glo2": "Non-hurdle Yes",
    ...
  }
}
```

---

## UI Interactions

### Pathway Selection:
- **Dropdown/Checkboxes** to select one or more pathways (e.g., D306_Core Only, D307_Core Only)
- **"Core Only" toggle** (Yes/No) — filters to only Core units within selected pathways

### 3D Model Behavior:
- **Year rings** broken into **3 sub-rings** (TRI-1, TRI-2, TRI-3)
- Units positioned in their trimester(s)
- **Zoom threshold = 15**: Below radius 15, assessment task nodes visible
  - Task nodes show `at_title`, `at_number`, `hurdle` status
  - Click task → detail panel shows `student_output`, `weighting`, `mode`, `participation`, GLO/CLO mappings

### Visual Indicators:
- **Hurdle nodes**: Glow effect + sparkle animation
- **Assessment badges**: Color-coded for hurdle (orange glow) vs non-hurdle (standard)
- **Unit nodes**: Color by strongest hurdle status (if any task is hurdle → bright)

---

## CSV Export Format

Use this format when exporting from Clippy/Gem:

### Sheet 1 — Assessment Tasks
```
unit_code,at_number,at_title,student_output,aol_task_type,mode,participation,hurdle,weighting,glo1,glo2,glo3,glo4,glo5,glo6,glo7,glo8
ACG703,1,Design Project,"Written, Practical, Design Artefact",NA,"Online, On-campus",Individual,Yes,30 percent,Hurdle Yes,No,No,No,No,Not Assessed,Not Assessed,Not Assessed
SLE111,1,Lab Portfolio,"Written, Practical",NA,On-campus,Individual,No,50 percent,Yes,No,Hurdle Yes,No,No,No,Yes,No
```

### Sheet 2 — Unit Coverage
```
pathway_id,unit_level,unit_code,teaching_periods,glo1,glo2,glo3,glo4,glo5,glo6,glo7,glo8
D306_Core Only,1,SRA143,TRI-1,Non-hurdle Yes,Non-hurdle Yes,No,No,No,No,Non-hurdle Yes,Not Assessed
D307_Core Only,1,MAE101,"TRI-1, TRI-3, TRI-2",Non-hurdle Yes,Not Assessed,No,Non-hurdle Yes,Not Assessed,Not Assessed,Not Assessed,No
```

---

## Implementation Steps (in Clippy/Gem)

1. **Extract** assessment task data → format as Sheet 1 CSV
2. **Extract** unit coverage data → format as Sheet 2 CSV
3. **Ensure Unit Code consistency** across both sheets
4. **Export** both CSVs to your Magma folder:
   - `curriculum-assessment-tasks-v1.csv`
   - `curriculum-unit-coverage-v1.csv`
5. **Dashboard will auto-load** and merge them on init

---

## Notes
- Hurdle status can be present at both task level (Sheet 1) and unit level (Sheet 2) — both trigger visual emphasis
- Teaching periods are comma-separated; dashboard parses and creates sub-rings per trimester
- If a unit appears in multiple pathways, it's listed separately for each (allows different trimester scheduling per pathway)
- "Not Assessed" vs "No" distinction is preserved in detail views
