# Balance sheet: spec (agreed 2026-11-09)

**Goal:** use the course's Excel data to support a balance of assessment types and assurance
modes for every GLO at every year level, across each course.

## Rule
- Every GLO must be assessed **at least twice in an AoL task at every year level**.
- AoL task = ⚑ (AoL + hurdle) or ● (AoL, non-hurdle).
- Many cells won't meet it yet; the sheet's job is to show where.

## Year levels
- Every level the export contains, each with the 2-AoL target.
- Labels: Y1, Y2, Y3, Hons → **Y4**, PG → **Y5** (from the unit code).

## What each GLO × year cell shows
- One mark per assessment of that GLO at that level, in the GLO's dashboard colour:
  ⚑ / ● count towards the target; ◆ (hurdle only) and ○ (neither) shown faded, not counted.
- AoL count against target, e.g. `AoL 1 of 2`, highlighted when below 2.
- For **every AoL task**: unit · AT, **assessment type**, **weighting**.

## Assessment type: two columns
- **Student Output**: the existing CALM field (coarse), kept as is.
- **Assessment type**: new planning column, **empty** until the new task definitions exist.
- Balance uses Assessment type where filled, otherwise falls back to Student Output (marked as fallback).
- Type mix: shown per task; no type-mix rule yet.

## Behaviour
- Live like the other views: follows Tasks filters (e.g. one pathway's ● column = one major)
  and the sheet's own checkboxes.
- Rebuilt by **Refresh views** in the VBA builder (planning edits flow through).
