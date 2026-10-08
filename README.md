# Corneal Reports PDF Extractor — Windows App

A small Windows app that reads corneal topography PDFs (MYOPIA EXPERT 700
output) from a folder, extracts the clinical fields, and writes them out as
`.xlsx` and `.csv`.

The extraction logic is unchanged from the original `corneal_reports_fixed.py`
— it lives in `corneal_reports_core.py`. The Tkinter GUI wrapper
(`corneal_reports_app.py`) is the new entry point.

---

## What's in this zip

| File | Purpose |
| --- | --- |
| `corneal_reports_app.py` | GUI entry point (Tkinter, no extra deps) |
| `corneal_reports_core.py` | Original extraction logic (unchanged) |
| `requirements.txt` | Runtime + build dependencies |
| `corneal_reports_app.spec` | PyInstaller spec (single-file, no console) |
| `build_windows.bat` | **One-click** build → `dist\CornealReportsApp.exe` |
| `run_app.bat` | Run directly with Python (no .exe build) |
| `README.md` | This file |

---

## Option A — Build the .exe (recommended)

Requires **Python 3.9+** installed on Windows
(<https://www.python.org/downloads/>). During install, tick
**"Add python.exe to PATH"**.

1. Unzip this folder anywhere (e.g. `Desktop\corneal-reports-app`).
2. Double-click **`build_windows.bat`**.
3. Wait for the build to finish (≈1–3 minutes the first time).
4. Your executable is at:

   ```
   dist\CornealReportsApp.exe
   ```

Copy that single `.exe` anywhere. To use it:

1. Create a `reports\` folder **next to** the `.exe` and drop your PDFs in.
2. Double-click `CornealReportsApp.exe`.
3. Pick your `reports` folder and an `output` folder (or accept the defaults).
4. Click **▶ Process PDFs**.
5. Open the output folder (button in the app) to see:
   - `corneal_reports.xlsx` — main data (sheet `Extracted_Data`) + raw text (sheet `Raw_Text`) + errors (sheet `Errors`)
   - `corneal_reports.csv` — main data, CSV form
   - `corneal_reports_errors.csv` — only rows with warnings/errors

> Tip: the app remembers your folder picks until you change them. The default
> folders are `reports\` and `output\` next to the .exe, so for most users it's
> just: drop PDFs in `reports\`, double-click the .exe, click Process.

---

## Option B — Just run it (no .exe build)

If you don't need a standalone `.exe`:

1. Install Python 3.9+.
2. Double-click **`run_app.bat`** — it installs the deps and launches the GUI.

This is the fastest path while you're iterating.

---

## Why "build on Windows" instead of pre-built?

PyInstaller produces a binary for the OS it runs on. To get a real Windows
`.exe` (with proper DLLs and bootloader), PyInstaller must run on Windows.
That's why this zip contains a build script instead of the .exe itself —
the build is one click and the result is a clean, single-file executable.

---

## Output schema

The Excel file has three sheets:

- **`Extracted_Data`** — one row per PDF, columns include patient info,
  Sim-K (K1/K2/CYL with axes, D values), cornea decentration/diameter,
  photopic pupil X/Y/diameter, average pupillar power at 4.5 mm and 3 mm,
  AK (D, AGC, SI, Kpi), keratoconus assessment, keratorefractive indexes
  (SD, SAI, asphericity, Kc), axial length, device info.
- **`Raw_Text`** — extracted text per page, for debugging.
- **`Errors`** — rows whose `Status` is `WARNING` or `ERROR`, with the
  reason in `Error_Message`.

---

## Troubleshooting

- **`python` is not recognized** — install Python from python.org and tick
  "Add python.exe to PATH" in the installer.
- **Antivirus flags the .exe** — false positive is common with PyInstaller.
  Either add an exclusion, or sign the binary. The code is unchanged.
- **`No PDF files found`** — the input folder is empty or has no `*.pdf`.
- **`No selectable text found`** — the PDF is a scanned image. Run OCR
  (e.g. Adobe Acrobat → "Recognize Text") before processing.
- **Build is slow the first time** — PyInstaller is unpacking many wheels.
  Subsequent builds are fast.

---

## Project layout

```
corneal-reports-app/
├── corneal_reports_app.py        # GUI wrapper (entry point)
├── corneal_reports_core.py        # Original extraction logic
├── requirements.txt
├── corneal_reports_app.spec       # PyInstaller spec
├── build_windows.bat              # Build → .exe
├── run_app.bat                    # Run via Python
└── README.md
```
