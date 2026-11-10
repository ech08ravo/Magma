# How to use the Curriculum Dashboard (v8 lite)

The lite dashboard reads the files you download from CALM directly. You do not need Python, scripts or a prepared `SEBE` folder: download what CALM lets you see, open the dashboard and upload the files. It works the same on a Mac or a Windows PC.

**You need:** access to CALM (Power BI) and a current desktop version of Chrome or Edge.

A full version (v7.4) remains available from the dashboard maintainer as an alternative. It adds SEBE extraction scripts and a statistics workbook for faculty-wide reporting.

---

## Step 1: Set up the dashboard folder

Unzip the dashboard ZIP anywhere on your computer. Keep the `lib` folder beside the HTML file; the dashboard uses it to work offline. Do not open the HTML from inside the ZIP.

```text
Curriculum-Dashboard-v8/
├── curriculum-dashboard-v8.html
├── QUICK-START-v8.md
├── HOW-TO-USE-THE-DASHBOARD-v8.md
├── lib/
│   ├── three.min.js
│   ├── papaparse.min.js
│   ├── xlsx.full.min.js
│   ├── pdf.min.js
│   ├── pdf.worker.min.js
│   └── PDFJS-LICENSE.txt
├── CALM-downloads/               (empty: put your CALM files here)
└── Unit-Guides/                  (empty, optional: unit-guide PDFs named with unit codes)
```

---

## Step 2: Download your files from CALM

Export these reports from CALM, keep CALM's file names, and save them in the `CALM-downloads` folder:

- `Student Pathways by Units.xlsx` (required)
- `Units by Assessment Tasks.xlsx` (required)
- `Course Data.xlsx` (optional; adds course names, faculty and school)

You only see the courses CALM gives you access to, and the dashboard shows exactly those. When you download a newer set, replace the old files, or keep each set in a dated sub-folder such as `CALM-downloads/2026-11-10`.

---

## Step 3: Load the data into the dashboard

1. Open `curriculum-dashboard-v8.html` in Chrome or Edge. The title bar should show **v8.0 lite**.
2. Click **⬆ Upload CALM files** (top right). In the upload window, drag in the files from `CALM-downloads`, or click to browse and select them together. They can be in any order, and you can leave them as `.xlsx`.
3. Each file gets a green tick next to its name:
   - Student Pathways by Units
   - Units by Assessment Tasks
   - Course Data (optional)

   Other CALM downloads, such as `Unit Relationship Details.xlsx`, are skipped with a note; you do not need to remove them.
4. Click **Apply to Model**.

Your files are read in the browser and are not uploaded anywhere. After closing or reloading the page, select them again.

**Offline use:** keep the complete **`lib`** folder beside the dashboard. The dashboard, 3D view, spreadsheet handling and PDF reader use these local files. If a library is missing, some tools may try an online fallback; reading unit guides requires both local PDF.js files (`pdf.min.js` and `pdf.worker.min.js`). The optional display fonts are downloaded from Google Fonts when online; the dashboard uses system fonts if they cannot be reached.

Click **Quick tour** in the bottom-left corner for a four-step introduction to filtering, reading the model, opening unit details, and finding assessment timelines. Reopen it any time.

Use **Table** beside Rings and Matrix for a text-based list of units. It follows the current pathway, year, GLO and hurdle filters; activate a unit code to open its details. The table can be used with a keyboard and screen reader. Choose Rings or Matrix to return to the 3D view.

### Sharing the dashboard

Share the dashboard ZIP, not your CALM downloads. Each person downloads their own files from CALM. Do not send the HTML file by itself; it needs the `lib` folder. To give someone a new version, send a fresh copy of the complete folder and ask them to replace the old one.

### Optional: local unit guides

1. On the upload screen, leave **Also scan a local Unit-Guides folder** unchecked to use CALM data only. To add guide information, check it and choose the folder containing the PDFs. The browser asks you to select the folder; the dashboard does not search your computer automatically.
2. Guides must be text-readable PDFs named with the unit code, for example `SRT141.pdf`. Keep one guide per unit in the selected folder. The dashboard reads locally in your browser and does not upload guide files.
3. The scanner looks only in **When and how to submit your work**. Extracted weeks appear in an amber guide box in unit details and in the timeline's **Week** view. Guide data is labeled with its source filename and a text excerpt. Review the excerpt in the guide: a week mentioned in this section is not independently verified as a due date.
4. The Week view links guide assessment-task numbers to CALM task numbers. Sections without a matching task number can still appear in the unit detail box but not the timeline. Guide information is supplementary and does not change CALM-based timeline exports.
5. To return to CALM-only data, use **Clear guides** in the upload screen. If a guide cannot be read, the upload screen reports the file and error; other successfully read guides remain available.

### Assessment timeline

1. Click a unit in the dashboard to select it.
2. Open **⚙ Display** and choose **Open assessment timeline** after selecting a unit, or choose a course in the menu and click **Open course timeline** to view it without a related unit. The standalone course timeline respects the current pathway selection; with **All pathways** active, it includes all pathways in the chosen course. Use the expand button in the timeline's upper-right corner to make the panel fullscreen, and click it again to restore its normal size. The timeline shows a horizontal schedule with a separate row for each study year. Under each year label, a compact eight-GLO summary shows the strongest recorded task alignment for each GLO; hover an icon to see its counts and status. Each trimester is split into **AT1**, **AT2**, **AT3**, and **Other** columns, with tasks stacked vertically. The related-unit timeline includes tasks across courses that contain the selected unit; the current pathway selection can narrow the courses shown. The chart can still scroll horizontally or vertically when its content exceeds the screen.
3. Under the **Study year** and trimester headings, select one or more keyword buttons to highlight matching cards. **Quiz** is its own button and matches “quiz” or “quizzes” in either the assessment title or AoL task type. **Test / exam** also matches “End-of-Unit Assessment” titles. Keywords search both fields; selecting multiple keywords highlights cards matching any selection. Other cards remain visible but dimmed. Click a selected keyword again to clear it.
4. Each compact card shows the assessment title (the general assessment label from the export), weight, mode, and AoL task type, plus small color-coded dot, flag, diamond, or ring icons for mapped GLOs—without large GLO number labels. Long assessment titles are truncated on the card; hover to read them in full. `N/A` means the export marks AoL task type as not applicable; it does not mean the assessment title is missing. The GLO color identifies the outcome; the icon shape indicates AoL + hurdle, AoL without hurdle, hurdle without AoL, or neither (alignment). Hover over an icon for its GLO name and alignment, or over the card for other available details.
5. Choose **Trimester**, **List**, or **Summary** above the report. **Week** is enabled when the selected timeline has matching guide-sourced weeks. **Trimester** is the visual overview. **List** shows searchable assessment rows; selecting a count in **Summary** opens the list filtered to that study year and trimester. Summary counts represent assessment placements, so a unit offered in more than one trimester contributes in each.
6. Use **Export CSV** for a file that opens in Excel or **Export Excel** for an `.xlsx` workbook. Both exports are flat tables with one row per assessment in each unit teaching trimester, so multi-trimester units have a row for each placement. Columns include course and pathway, study year, trimester, AT group and number, unit code, assessment title (the general task type), AoL task type, AoL/hurdle status, weighting, mode, participation, student output, and GLO statuses. The export follows the timeline's current course/pathway scope and includes all tasks; keyword buttons only highlight cards and do not filter the exported rows. Units without a supplied teaching period are labeled **Not supplied**. Filenames include the course title (or unit code and related-course title) and `timeline_DDMMYY`.

**Important CALM data limits:** the current exports do not include assessment due week, a separate tool field, or location. The AT columns group by task number; they are not dates or estimates of when work is due. Trimester placement uses the unit's available teaching-period information; units with multiple listed trimesters appear in each. Units without a known teaching period are placed under **Teaching period not supplied**. Due-week, tool, and location filters can be added if CALM provides those fields.


---

## Step 4: Reading the dashboard

**Rings (bottom to top):** Y1 › Y2 › Y3 › Hons › PG › MRes › PhD. The first digit of the unit code sets the ring: 1xx = Y1, 4xx = Hons, 7xx = PG, 8xx = Masters by Research, 9xx = PhD, and 0xx = Non-major.

**Unit shapes:**

| Shape | Meaning |
|---|---|
| ◆ Spinning gem | AoL unit: at least one task has an AoL Task Type |
| ○ Small sphere | Developing: no AoL tasks |
| ◯ Pulsing coloured ring | AoL unit. Each arc is a GLO its AoL tasks assess, in that GLO's colour (grey = none). Selecting a GLO fades the other arcs. |
| ⚑ Red flag | At least one task is a hurdle |

Lines run from each unit to the GLOs it assesses. A unit with no lines assesses no GLO in the CALM data.

**Teaching period** (shown in the unit's tooltip and detail panel):
- `TRI-1`, `TRI-2, TRI-3` and so on: the trimesters the unit runs in.
- An **asterisk**, as in `TRI-1*`, means the period comes from the **April 2026** data, because CALM doesn't export teaching periods yet.
- `T0` means no teaching period is known.

**Click a unit** in Rings or Matrix, or activate its code in **Table**, to open its detail panel: year, AoL status, teaching period, the pathways it's in, its 8 GLOs as chips (filled = Yes, outlined = Non-hurdle Yes, faded = not assessed), and its assessment tasks, with a red **Hurdle** badge on hurdle tasks.

---

## Step 5: Choosing pathways

Open the **All pathways** menu (top left).

- **Search** by course name, course code or pathway name. Type a unit-code prefix, for example `SRT` or `SRT1`, to see matching units. Tick one or more units to show only those units; this temporarily overrides pathway and Core-only filters. Year, GLO and hurdle filters still apply. Clear the unit selection to return to pathway filtering. Matching pathway rows also appear below.
- **Tick a course heading** to show **all** that course's units: its core plus every major, specialisation and stream. Hover over a heading to see the course's faculty and school.
- **core only** (the small switch on a course heading) narrows that course to its core units while keeping other courses in the current selection. From **All pathways**, it preserves all other courses. Click it again to bring that course's majors and streams back. Courses with nothing but a core, such as Artificial Intelligence, have no switch.
- The top-bar **Core units only** filter keeps units that appear in a course's `Core Only` pathway. If every selected unit is already in a core pathway, the display will not change; the dashboard will tell you this.
- **Tick individual majors, specialisations or streams** under a course for finer choices. A half-tick (–) on a course heading means only some of its pathways are selected.
- **Clear** removes all ticks and shows every unit. **All cores** ticks every course's Core Only pathway. **All** ticks everything shown in the current search.
- If you untick everything, the view goes empty and the button says **No pathways selected**. Press **Clear** to see all units again.

Other filters in the top bar:
- **Core units only** (on/off) keeps your ticks but hides any unit that isn't in a course's core.
- **⚑ Hurdles** shows only units with a hurdle task.
- The **year buttons** (top right) show one year at a time.
- Click a **GLO marker** at the top of the 3D view to show only units assessing it. A **halo** means the GLO is assured (AoL + hurdle task) in every year in view, and a red **⚑** beside it means it's assessed but never by a hurdle task. Hover a marker for the detail.

---

## Step 6: Summarise selected pathways

With one or more pathways ticked, click **ⓘ Summary** (next to the pathway menu). The panel shows one table, **GLO coverage by year**. Hover the **ⓘ** next to its heading for the full key.

- **Rows** are years. **Columns** are 0, GLO1–GLO8 and All.
- **Year cells:** ⚑ means an AoL + hurdle task assesses that GLO that year, and ● means only AoL tasks that aren't hurdles do. Empty means no AoL task assesses it.
- **All:** ✓ when all 8 GLOs have ⚑ that year, ✗ when any don't. Hover the ✗ to see which.
- **Total row:** how many years each GLO has ⚑ and ●.
- **Column 0:** AoL or hurdle tasks with **no GLO mapped** in CALM, in grey.
- **Click a year** (or **Expand all**) to show its units underneath. Each unit line shows **AoL** / **Hurdle** badges and the teaching period, then one line per task with its marks under the GLO columns: ⚑ AoL + hurdle · ● AoL · ◆ hurdle, not AoL · ○ neither · faint = not assessed.
- Hover any cell or mark for the units and tasks behind it. Click a unit code to open its detail panel; **← Back to pathway summary** returns you.

Years start folded each time the Summary opens.

---

## Step 7 (optional): Edit and save changes

1. Click **✎ Edit Data** to open the editor. In it you can:
   - add or remove units in a pathway
   - change a unit's GLO values
   - change a task's weighting or hurdle
2. Changes show immediately, but **they are lost when the page reloads**.
3. To keep them, click **⬇ Export Modified XLSX**. Each exported workbook keeps the original uploaded document's filename and adds `MODIFIED_DDMMYY` before `.xlsx` (for example, `Student Pathways by Units MODIFIED_021126.xlsx`). Drop the exported workbooks into the dashboard next time instead of the originals.

Edits only change the dashboard's copy of the data. They **do not change CALM**. Report real corrections to the CALM owner.

---

## Known gaps in the CALM data

| Gap | Effect | Fix needed |
|---|---|---|
| No teaching period column | Periods come from the April 2026 snapshot (`*`) or show as T0 | Add `at_element_teaching_period` to the tasks export |
| No hurdle requirement wording | Only Yes/No is known, not *what* students must do (some wording appears in Weighting Information) | New CALM field |
| No unit owner (school/faculty) | Units are matched to a faculty through the courses they appear in | New CALM field |
| Free-text Hurdle values (`YNo`, `10percent`, …) | Treated as "not a hurdle" | Clean up in CALM |
| Units with no GLO mapping (e.g. SRE270) | No lines, all GLO chips faded | Check the mapping with the unit chair |

CALM data model owner: Andrew Frank.

