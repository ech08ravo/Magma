# Test Data Setup

## What You Have

Two **example output CSV files** showing exactly what the VBA macro should produce:

1. **curriculum-assessment-tasks-v1-OUTPUT-EXAMPLE.csv**
2. **curriculum-unit-coverage-v1-OUTPUT-EXAMPLE.csv**

---

## How to Test

### Option A: Use These Files Directly in Dashboard
1. Copy both CSV files to your Magma folder (remove the "-OUTPUT-EXAMPLE" suffix)
2. Rename:
   - `curriculum-assessment-tasks-v1-OUTPUT-EXAMPLE.csv` → `curriculum-assessment-tasks-v1.csv`
   - `curriculum-unit-coverage-v1-OUTPUT-EXAMPLE.csv` → `curriculum-unit-coverage-v1.csv`
3. The new dashboard will auto-load them when you open it
4. Select pathways from dropdown to see the 3D model populate

### Option B: Run VBA Macro with Test Data
1. Create a new Excel file with two sheets:
   - Sheet 1: "Courses by Units"
   - Sheet 2: "Units by Assessment Tasks"
2. Enter the test data below into each sheet
3. Copy the VBA macro code from `VBA_MERGE_MACRO.txt`
4. Run the macro
5. Two CSVs will be generated in your Documents folder
6. Copy them to your Magma folder

---

## Test Data to Enter

### Sheet 1: "Courses by Units"

Headers:
```
Student Pathway | Unit Level of Study | Unit Code | 2026 Teaching Periods | GLO1/CLO1 | GLO2/CLO2 | GLO3/CLO3 | GLO4/CLO4 | GLO5/CLO5 | GLO6/CLO6 | GLO7/CLO7 | GLO8/CLO8
```

Test Data (copy into rows 2+):
```
D306_Core Only | 1 | SRA143 | TRI-1 | Non-hurdle Yes | Non-hurdle Yes | No | No | No | No | Non-hurdle Yes | Not Assessed
D306_Core Only | 1 | SRC163 | TRI-1 | No | No | No | Not Assessed | Not Assessed | No | Not Assessed | Not Assessed
D306_Core Only | 1 | SRC164 | TRI-2 | No | No | No | Not Assessed | Not Assessed | No | Not Assessed | Not Assessed
D306_Core Only | 2 | ACG703 | TRI-1 | Non-hurdle Yes | No | Hurdle Yes | No | No | Not Assessed | Not Assessed | No
D307_Core Only | 1 | MAE101 | TRI-1, TRI-3, TRI-2 | Non-hurdle Yes | Not Assessed | No | Non-hurdle Yes | Not Assessed | Not Assessed | Not Assessed | No
D307_Core Only | 1 | MAF101 | TRI-2, TRI-1, TRI-3 | Non-hurdle Yes | Not Assessed | Non-hurdle Yes | Non-hurdle Yes | Non-hurdle Yes | Not Assessed | Not Assessed | Not Assessed
```

### Sheet 2: "Units by Assessment Tasks"

Headers:
```
Unit Code | AT Number | AT Title | Student Output | AoL Task Type | Mode | Participation | Hurdle | Weighting Information | GLO1/CLO1 | GLO2/CLO2 | GLO3/CLO3 | GLO4/CLO4 | GLO5/CLO5 | GLO6/CLO6 | GLO7/CLO7 | GLO8/CLO8
```

Test Data (copy into rows 2+):
```
ACG703 | 1 | Design Project | Written, Practical, Other: Design Artefact | NA | Online, On-campus | Individual | No | 30 percent of the total mark of the unit. | No | No | No | No | No | Not Assessed | Not Assessed | Not Assessed
ACG703 | 2 | Group Critique | Oral, Written | NA | On-campus | Group | Yes | 20 percent of the total mark of the unit. | Yes | Hurdle Yes | No | No | No | Not Assessed | Not Assessed | No
SRA143 | 1 | Lab Report | Written, Practical | NA | On-campus | Individual | No | 50 percent of the total mark of the unit. | Hurdle Yes | Non-hurdle Yes | No | No | No | No | Non-hurdle Yes | Not Assessed
SRA143 | 2 | Final Exam | Written | NA | On-campus | Individual | Yes | 50 percent of the total mark of the unit. | Non-hurdle Yes | Hurdle Yes | No | No | No | No | Hurdle Yes | Not Assessed
SRC163 | 1 | Written Assignment | Written | NA | Online | Individual | No | 100 percent of the total mark of the unit. | No | No | No | Not Assessed | Not Assessed | No | Not Assessed | Not Assessed
MAE101 | 1 | Problem Sets | Written | NA | Online, On-campus | Individual | No | 40 percent of the total mark of the unit. | Non-hurdle Yes | Not Assessed | No | Non-hurdle Yes | Not Assessed | Not Assessed | Not Assessed | No
MAE101 | 2 | Midterm Exam | Written | NA | On-campus | Individual | Yes | 30 percent of the total mark of the unit. | Hurdle Yes | Not Assessed | No | Hurdle Yes | Not Assessed | Not Assessed | Not Assessed | No
MAE101 | 3 | Final Project | Written, Practical | NA | On-campus | Group | Yes | 30 percent of the total mark of the unit. | Non-hurdle Yes | Not Assessed | No | Non-hurdle Yes | Not Assessed | Not Assessed | Not Assessed | No
MAF101 | 1 | Written Exam | Written | NA | On-campus | Individual | Yes | 50 percent of the total mark of the unit. | Non-hurdle Yes | Not Assessed | Non-hurdle Yes | Non-hurdle Yes | Non-hurdle Yes | Not Assessed | Not Assessed | Not Assessed
```

---

## Expected Output

When you run the macro, you should get two CSVs matching the **-OUTPUT-EXAMPLE.csv** files.

### Key Features:
✅ Assessment tasks are **one per row** (units can have multiple rows)
✅ Unit coverage is **one per pathway per unit** (same unit in different pathways = different rows)
✅ Teaching periods are **comma-separated** if unit runs across multiple trimesters
✅ Hurdle status is **preserved** (Yes, No, Hurdle Yes, Non-hurdle Yes, Not Assessed)
✅ Comma-separated fields are **quoted** in CSV (e.g., "Written, Practical")

---

## Dashboard Integration

Once you have the CSVs:
1. Copy to Magma folder (or let the macro create them in Documents and move them)
2. I'll rebuild the dashboard to use them
3. Features enabled:
   - Pathway dropdown selector
   - Core Only toggle
   - Trimester sub-rings
   - Assessment task detail at zoom < 15
   - Hurdle glow effects

Ready to test! 🚀
