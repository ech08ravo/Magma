# Single CSV Merge - Instructions & Examples

## What Changed

Instead of **two separate CSVs**, the macro now outputs **one combined CSV** with all data in a single file:

- ✅ **curriculum-merged-v1.csv** — One file, fully combined
- ✅ Handles comma-separated values automatically (e.g., "Written, Practical", "TRI-1, TRI-2, TRI-3")
- ✅ Each row = one pathway + unit + assessment task combination
- ✅ Includes both unit-level and task-level GLO/CLO mappings

---

## Example Output Structure

Each row contains:
```
pathway_id | unit_level | unit_code | teaching_periods | at_number | at_title | student_output | aol_task_type | mode | participation | hurdle | weighting | unit_glo1-8 | task_glo1-8
```

### Example Row:
```
D306_Core Only, 1, SRA143, TRI-1, 1, Lab Report, "Written, Practical", NA, On-campus, Individual, No, 50 percent of the total mark of the unit., Hurdle Yes, Non-hurdle Yes, No, No, No, No, Non-hurdle Yes, Not Assessed, Hurdle Yes, Non-hurdle Yes, No, No, No, No, Non-hurdle Yes, Not Assessed
```

This single row tells you:
- 🎓 **Pathway:** D306_Core Only
- 📚 **Unit:** SRA143 (Year 1, TRI-1)
- 📋 **Assessment Task 1:** Lab Report
- 📊 **Task Details:** Written + Practical, On-campus, Individual, 50% weighting
- ✅ **Unit GLO Coverage:** GLO1,2,7 assessed; GLO3,4,5 not assessed; GLO6,8 not assessed
- ✅ **Task GLO Mapping:** Same as unit (inherits)

---

## Comma-Separated Value Handling

The macro **automatically handles commas in cells** by quoting fields:

### Fields That May Contain Commas:
- **Student Output:** "Written, Practical, Other: Design Artefact"
- **Mode:** "Online, On-campus"
- **Participation:** "Group, Individual"
- **Teaching Periods:** "TRI-1, TRI-2, TRI-3"
- **Weighting:** "30 percent of the total mark, includes hurdle requirement"

### CSV Quoting (Automatic):
```csv
D306_Core Only,1,ACG703,TRI-1,1,"Design Project","Written, Practical, Other: Design Artefact",NA,"Online, On-campus",Individual,No,"30 percent of the total mark",No,No,No,No,No,Not Assessed,Not Assessed,Not Assessed,No,No,No,No,No,Not Assessed,Not Assessed,Not Assessed
```

When the field has a comma, it's enclosed in quotes. No manual quoting needed!

---

## File Location

Files created in your **Documents folder**:
- `curriculum-merged-v1.csv` ← Single combined CSV

**To use with dashboard:**
1. Copy to Magma folder: `/Users/elizabeth.black/Library/Mobile Documents/com~apple~CloudDocs/Debusseycat/Magma/`
2. Dashboard will auto-load it

---

## Test Data (Quick Version)

### Sheet 1: "Courses by Units"
| Student Pathway | Unit Level of Study | Unit Code | 2026 Teaching Periods | GLO1/CLO1 | GLO2/CLO2 | GLO3/CLO3 | GLO4/CLO4 | GLO5/CLO5 | GLO6/CLO6 | GLO7/CLO7 | GLO8/CLO8 |
|---|---|---|---|---|---|---|---|---|---|---|---|
| D306_Core Only | 1 | SRA143 | TRI-1 | Non-hurdle Yes | Non-hurdle Yes | No | No | No | No | Non-hurdle Yes | Not Assessed |
| D306_Core Only | 1 | SRC163 | TRI-1 | No | No | No | Not Assessed | Not Assessed | No | Not Assessed | Not Assessed |
| D307_Core Only | 1 | MAE101 | TRI-1, TRI-3, TRI-2 | Non-hurdle Yes | Not Assessed | No | Non-hurdle Yes | Not Assessed | Not Assessed | Not Assessed | No |

### Sheet 2: "Units by Assessment Tasks"
| Unit Code | AT Number | AT Title | Student Output | AoL Task Type | Mode | Participation | Hurdle | Weighting Information | GLO1/CLO1 | GLO2/CLO2 | GLO3/CLO3 | GLO4/CLO4 | GLO5/CLO5 | GLO6/CLO6 | GLO7/CLO7 | GLO8/CLO8 |
|---|---|---|---|---|---|---|---|---|---|---|---|---|---|---|---|---|
| SRA143 | 1 | Lab Report | Written, Practical | NA | On-campus | Individual | No | 50 percent of the total mark | Hurdle Yes | Non-hurdle Yes | No | No | No | No | Non-hurdle Yes | Not Assessed |
| SRA143 | 2 | Final Exam | Written | NA | On-campus | Individual | Yes | 50 percent | Non-hurdle Yes | Hurdle Yes | No | No | No | No | Hurdle Yes | Not Assessed |
| MAE101 | 1 | Problem Sets | Written | NA | Online, On-campus | Individual | No | 40 percent | Non-hurdle Yes | Not Assessed | No | Non-hurdle Yes | Not Assessed | Not Assessed | Not Assessed | No |

---

## Expected Output

Run the macro and you'll get **curriculum-merged-v1.csv** with structure:

```
pathway_id,unit_level,unit_code,teaching_periods,at_number,at_title,student_output,aol_task_type,mode,participation,hurdle,weighting,unit_glo1,unit_glo2,unit_glo3,unit_glo4,unit_glo5,unit_glo6,unit_glo7,unit_glo8,task_glo1,task_glo2,task_glo3,task_glo4,task_glo5,task_glo6,task_glo7,task_glo8
D306_Core Only,1,SRA143,TRI-1,1,"Lab Report","Written, Practical",NA,On-campus,Individual,No,50 percent of the total mark of the unit.,Hurdle Yes,Non-hurdle Yes,No,No,No,No,Non-hurdle Yes,Not Assessed,Hurdle Yes,Non-hurdle Yes,No,No,No,No,Non-hurdle Yes,Not Assessed
D306_Core Only,1,SRA143,TRI-1,2,"Final Exam",Written,NA,On-campus,Individual,Yes,50 percent of the total mark of the unit.,Non-hurdle Yes,Hurdle Yes,No,No,No,No,Hurdle Yes,Not Assessed,Non-hurdle Yes,Hurdle Yes,No,No,No,No,Hurdle Yes,Not Assessed
D307_Core Only,1,MAE101,"TRI-1, TRI-3, TRI-2",1,"Problem Sets",Written,NA,"Online, On-campus",Individual,No,40 percent of the total mark of the unit.,Non-hurdle Yes,Not Assessed,No,Non-hurdle Yes,Not Assessed,Not Assessed,Not Assessed,No,Non-hurdle Yes,Not Assessed,No,Non-hurdle Yes,Not Assessed,Not Assessed,Not Assessed,No
```

---

## Dashboard Integration

Once you have the CSV:
1. Copy to Magma folder
2. I'll rebuild the dashboard to load **curriculum-merged-v1.csv**
3. Features enabled:
   - Pathway dropdown
   - Core Only toggle
   - Trimester sub-rings
   - Assessment task detail
   - Hurdle glow effects

Ready! 🚀
