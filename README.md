# PV String IV Curve Analyzer — v1.0

## Install on Windows
1. Install Python 3.11 or newer from python.org, and enable **Add Python to PATH**.
2. Extract the ZIP, open Command Prompt in this folder.
3. Run: `py -m pip install -r requirements.txt`
4. Run: `py -m streamlit run app.py`
5. Your browser opens the local app (usually http://localhost:8501).

## Using the app
- Enter project, block, inverter, expected string count, and cut string numbers in the sidebar.
- Upload one CSV per measured string, with filenames such as `b2-inv12-s4.csv` or `b2-inv12-s18(1).csv`. The complete number after `-s` identifies the string.
- CSV files need voltage and current numeric columns, e.g. `V(V),I(A)` (also supports `Voltage (V),Current (A)`).
- Review any import warnings and all FF results; export Word.
- The included `demo_csv` folder has synthetic sample measurements for testing only.

## Important limitations
- The generic importer supports simple voltage/current CSV tables; tester-specific formats with unusual headers/units need an adapter.
- FF = 100 × Pmax/(Voc × Isc). Pmax is the maximum measured V×I point, not a fitted maximum.
- Isc and Voc are linearly extrapolated from edge points when axis intercepts are absent; these are estimates, not tester-certified values.
- The graph connects measured points without smoothing or changing measurements.
- No irradiance, temperature, equipment ID, signature, or calibration metadata is imported yet.
- Word outputs are starter layouts, not an exact reproduction of any existing user template.
- Do not use as the sole basis for formal commissioning acceptance without validation against a calibrated instrument.

## Version 1.2: ABC Solar template reporting
- Uses the included original Word template for cover page, approval table, test information, six charts per page and string performance summary.
- Enter **Inverter Serial No.** in the sidebar; it is inserted into the Test Information table.
- The supplied template contains 18 string slots. Version 1.2 limits reports to 18 strings; a future template variant is needed for 21/28 strings.
- Existing STC readings in the source template are deliberately cleared rather than reused. No STC values are fabricated.
- PDF output is generated from the Word report using LibreOffice. On Windows, install LibreOffice and ensure `soffice`/`libreoffice` is available on PATH. If PDF conversion is unavailable, download Word.

## Version 1.3: Raw Data Files measured + STC
- Upload the tester's `Raw Data Files.csv` or `.xlsx` in the dedicated upload field, as well as the individual I-V curve CSVs.
- The importer recognizes a `File Name` column and measured columns such as `Pmax(W)`, `Voc(V)`, `Isc(A)`, `Vm(V)`, `Im(A)`, `Irradiance`, `PV Temp`, `En Temp`, `Incline`, `FF(%)`, plus corresponding columns prefixed or suffixed with `STC`.
- It also supports rows labeled `STC` in a `Type`, `Condition`, `Measurement Type`, or `Mode` column.
- Each string has a measured row followed immediately by its STC row, preserving the original ABC Solar summary table.
- The original user-provided Raw Data Files export has **not** been supplied in this turn. The importer has been tested with a representative fixture, but may need column mappings adjusted to the actual export. No STC values are invented.
- If no STC data is found, the report leaves STC cells blank and displays a warning in the app.

## Version 1.4: Word-only reporting
- Removed the PDF download button and all PDF conversion code.
- LibreOffice is **not required** on the server or on Windows.
- Reports remain editable DOCX files based on the original ABC Solar template.
- Measured and STC rows, inverter serial number and abnormal FF classification are retained.

### Version 1.6: Result cell colors
Normal results are shaded light green (#E2F0D9); Abnormal results are shaded light red (#F2DBDB), matching the ABC Solar Word template. Status shading is applied to the merged Result cells when generating the Word report.


## Version 1.7 — tester Item row correction
The Raw Data Files CSV uses `FileName`, `Item`, and the metric columns.
For each string, `Item=Value` supplies measured values and `Item=STC` supplies STC values.
Rows with PV module model names are ignored (they are nameplate specifications, not STC measurements).
The STC Data Preview displays imported STC fields before report download.
Tested with `Raw Data Files(20261121-101333).csv` containing 18 strings.
