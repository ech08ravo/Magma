# Curriculum map: later list

Things agreed but parked. Newest at the bottom.

## Done (2026-11-09)
- [x] **Timeline opens first.** Timeline is the first tab and the one shown on open.
- [x] **Tools in the dashboard.** Timeline cards show Tools and Type (Assessment type) when a planning workbook is loaded (dashboard v7.3).
- [x] **Read planning columns back into the dashboard** (v7.3): drop the Excel map (its Tasks sheet) in with the CALM files; Tools, Assessment type and Planned Weeks are matched by unit code + AT number. Exports carry them again.
- [x] **"Other / unspecified" mode checkbox** for tasks that are neither in person nor online.
- [x] **Balance sheet** (see BALANCE-SPEC.md).

## Done in the VBA builder (Curriculum Map Builder.xlsm, 2026-11-09)
- [x] **Refresh views** button: rebuilds every sheet from the current Tasks data (~19 s for 211 tasks).
- [x] **Import export…** button: loads a dashboard export (.csv or .xlsx) into Tasks, keeping planning and team-added columns (matched by unit + AT).
- [x] **Click a GLO or mode label** to toggle that filter on Timeline/Weeks.
- [x] Task type codes live on the editable **Lists** sheet.

## Done (2026-11-10)
- [x] Blank builder for colleagues (`Curriculum Map Builder (blank).xlsm`, made with `MakeBlankBuilder`).
- [x] Filter-restore fix: a cleared Timeline no longer hides every weighted card.
- [x] Dashboard v8.0 lite and its zip; repo tidied into v7/, v8/ and archive/; CALM data kept out of git.

## Open when the project resumes
- [ ] A dashboard change Elizabeth mentioned on 2026-11-09 ("need another change to the dashboard"), not yet specified.
- [ ] Decide on **Planned AoL** and **Planned GLOs** planning columns, so planned AoL changes show on the AOL type balance sheet.
- [ ] Remove the summary/description lines from the **All tasks** sheet too (as done on AOL type balance) and close up the empty rows.
- [ ] Start-sheet hint on the blank builder: "If Excel says the macro isn't available, close, reopen and click Enable Macros" (and the Windows Unblock step).
- [ ] Re-export S466 from dashboard v7.4/v8 (37 columns, with unit guides) and re-import to check guide weeks and GLOs.
- [ ] Older commits on GitHub/GitLab still contain CALM data files; rewrite history only if the repo needs to be clean (ask first).

## Optional, not started
- [ ] Year / Trimester checkboxes on the views (Tasks-sheet filters already do this).
- [ ] Tools tick-box filter, built from the Tools values present (once abbreviations settle; easiest in the VBA Refresh).
