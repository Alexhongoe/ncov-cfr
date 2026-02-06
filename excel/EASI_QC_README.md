# EASI QC Signals Macro

This module adds an Excel/VBA macro to calculate site-level QC signals for EASI data:

1. **Baseline SD unusually low** (per site):
   - Baseline visit = `AVISIT = 基线` or `AVISITN = 0`.
   - Compute SD of baseline `AVAL` per site.
   - Yellow flag: site SD <= 10th percentile of all site SDs.
   - Red flag: site SD <= 5th percentile **and** baseline sample size >= 10.

2. **Last-digit preference (0/5)**:
   - Uses integer part of `AVAL` and checks if last digit is 0 or 5.
   - Yellow: proportion > 45%.
   - Red: proportion > 60% **and** total `AVAL` count >= 20.

3. **Week-16 change outliers**:
   - Uses `AVISITN = 16` and `CHG` values.
   - Outlier defined as absolute deviation > 3 SD from global mean of week-16 `CHG`.
   - Yellow: outlier proportion > 2%.
   - Red: outlier proportion > 5%.

## How to use

1. Import `EASI_QC.bas` into your Excel workbook (VBA editor → File → Import File).
2. Put the data in a sheet named `DATA`, or make the data sheet active.
3. Ensure the header row is in row 1 and includes these columns:
   - `SITEID`, `SITENAM`, `AVISIT`, `AVISITN`, `AVAL`, `CHG`.
4. Run the macro `RunEasiQcSignals`.
5. Review the output in the `QC_SIGNALS` sheet.

## Output

The macro creates/overwrites a `QC_SIGNALS` sheet with per-site flags and summary metrics:
- `BaselineSD`, `BaselineFlag`
- `LastDigit0or5Pct`, `LastDigitFlag`
- `W16OutlierPct`, `W16Flag`

Baseline percentile thresholds are stored in the `Notes` column for reference.
