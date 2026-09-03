# NN Water Billing

## Monthly automation

Run from the workspace root:

```bash
python monthly_water_billing.py --month "<Month> <Year>" --pdf
```

By default, the command resolves paths as follows:

```text
images:  <year>/<Month>/input/
xlsx:    <year>/NN_WM_water.xlsx, with older naming patterns as fallbacks
output:  <year>/<Month>/output/
policy:  skills.md in the launch directory
```

The command updates the selected reading and cost columns, validates all meter identifiers and policy inputs, and creates Society-facing and NoBroker Hood Excel/PDF bills. Optional path arguments can override any default. If OCR cannot identify every reading, it stops before updating the workbook and writes an `*_input_review.json` report. Reviewed values can be supplied with `--readings-json`.

## Workspace layout

There is **one meter workbook per year**. All monthly readings and cost values for a year remain in that annual workbook as month columns; monthly bills are generated separately under the month output folder.

- `monthly_water_billing.py`: canonical monthly ingestion, validation, calculation, and output workflow.
- `calculate_water_bill.py`: reusable billing calculation core.
- `skills.md`: vendor rates, solar-meter policy, billing rules, and dependency requirements.
- `<year>/<Month>/input/`: source meter images by collection month.
- `<year>/<Month>/output/`: generated bills, PDFs, and monthly review reports.
- `<year>/NN_WM_water.xlsx`: the single active workbook for that year.
- `archive/ocr_experiments/`: superseded one-off OCR/parsing scripts retained for reference.
- `archive/legacy/`: older workbooks and intermediate artifacts retained for audit history.

## Required policy inputs

Keep vendor rates and faulty solar series current in `skills.md`. The workflow will stop if rates, the selected month columns, tanker counts, required worksheets, or complete meter readings are missing.