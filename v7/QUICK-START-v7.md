# Curriculum Dashboard v7 beta — quick start

After unzipping and moving the prepared `SEBE` folder, the files should be arranged like this:

```text
Curriculum-Dashboard-v7/
├── curriculum-dashboard-v7.html
├── lib/
│   ├── three.min.js
│   ├── papaparse.min.js
│   ├── xlsx.full.min.js
│   ├── pdf.min.js
│   └── pdf.worker.min.js
├── SEBE/
│   ├── Student Pathways by Units.csv
│   ├── Units by Assessment Tasks.csv
│   └── Course Data.csv
└── Unit-Guides/                  (optional; PDFs named with unit codes)
    ├── SRT141.pdf
    └── ...
```

1. **Unzip** the dashboard ZIP to a folder on your computer. Keep the `lib` folder beside `curriculum-dashboard-v7.html`.
2. **Move** the prepared `SEBE` folder into that same folder, beside the HTML file.
3. **Open** `curriculum-dashboard-v7.html` in Chrome or Edge. Keep the files together; the dashboard uses the local `lib` folder to work offline.
4. Click **⬆ Upload SEBE files** in the dashboard.
5. In the upload window, select all three CSV files from `SEBE` together: **Student Pathways by Units.csv**, **Units by Assessment Tasks.csv**, and **Course Data.csv**.
6. Confirm the three files appear in the upload window. Course Data is optional, but including it adds course names.
7. If you have unit-guide PDFs, check **Also scan a local Unit-Guides folder** and choose that folder. Skip this for CALM-only use.
8. Check that the selected files appear in the upload window, then click **Apply to Model**.
9. Wait for the dashboard to load. Units and filters should appear; click a unit to see its details. Guide-derived weeks appear separately in amber and in the **Week** timeline view.
10. To use the dashboard offline, leave the extracted files in place. Your selected CSVs and guides are read locally and are not uploaded; after closing or reloading, select them again.
