# Data Setup Instructions for Dashboard v1

## What You Have
Three files in your Magma folder:

1. **SCHEMA-v1.md** — Full data schema explanation
2. **curriculum-assessment-tasks-v1-TEMPLATE.csv** — Assessment tasks template
3. **curriculum-unit-coverage-v1-TEMPLATE.csv** — Unit coverage template

---

## What You Need to Do

### Step 1: Extract Data in Clippy/Gem
Process your source data to match **two CSV outputs**:

#### CSV A: Assessment Tasks
- Extract per-unit assessment task data
- Include: unit code, task number, title, outputs, mode, participation, hurdle status, GLO/CLO mappings
- One row per assessment task (units can have multiple tasks)
- Save as: `curriculum-assessment-tasks-v1.csv`

#### CSV B: Unit Coverage  
- Extract unit pathway enrollment data
- Include: pathway ID, year level, unit code, teaching periods (trimester), GLO/CLO coverage status
- One row per unit per pathway (same unit can appear in multiple pathways)
- Save as: `curriculum-unit-coverage-v1.csv`

### Step 2: Validate Column Names
Ensure your CSVs match the templates **exactly** (case-sensitive):

**Assessment Tasks columns:**
```
unit_code,at_number,at_title,student_output,aol_task_type,mode,participation,hurdle,weighting,glo1,glo2,glo3,glo4,glo5,glo6,glo7,glo8
```

**Unit Coverage columns:**
```
pathway_id,unit_level,unit_code,teaching_periods,glo1,glo2,glo3,glo4,glo5,glo6,glo7,glo8
```

### Step 3: Standardize Values
Use **exact values** for these fields:

**hurdle column (Assessment Tasks):**
- `Yes` (hurdle = true)
- `No` (hurdle = false)

**GLO columns (both sheets):**
- `Yes` (assessed, no hurdle)
- `Hurdle Yes` or `Non-hurdle Yes` (explicitly noted)
- `No` (not assessed)
- `Not Assessed` (explicitly not assessed)

**teaching_periods (Unit Coverage):**
- Single: `TRI-1` or `TRI-2` or `TRI-3`
- Multiple: `TRI-1, TRI-2` (comma + space separated)

### Step 4: Save to Magma Folder
Export both CSVs to `/Users/elizabeth.black/Library/Mobile Documents/com~apple~CloudDocs/Debusseycat/Magma/`:
- `curriculum-assessment-tasks-v1.csv`
- `curriculum-unit-coverage-v1.csv`

### Step 5: Dashboard Auto-Load
Once files are in place, the next version of the dashboard will:
1. Auto-detect and load both CSVs
2. Merge by unit code
3. Build the 3D model with:
   - Pathway dropdown selector
   - "Core Only" toggle
   - Trimester sub-rings
   - Assessment task zoom-in detail
   - Hurdle glow effects

---

## Example Data Validation

### Assessment Tasks (Check):
```
✓ unit_code matches units in Coverage sheet
✓ at_number is sequential per unit (1, 2, 3...)
✓ hurdle is Yes or No (not "true"/"false")
✓ glo1–glo8 are one of: Yes, Hurdle Yes, Non-hurdle Yes, No, Not Assessed
✓ weighting is text (e.g., "30 percent", "50%")
```

### Unit Coverage (Check):
```
✓ pathway_id is consistent (e.g., all "D306_Core Only" or "D307_Core Only")
✓ unit_level is 1, 2, or 3
✓ unit_code matches Assessment Tasks
✓ teaching_periods use TRI-1/TRI-2/TRI-3 separated by ", "
✓ glo1–glo8 follow same values as Assessment Tasks
```

---

## Questions?
Refer to **SCHEMA-v1.md** for detailed column definitions and merge logic.

Once you have the CSVs ready, send them over and I'll build the new dashboard! 🚀
