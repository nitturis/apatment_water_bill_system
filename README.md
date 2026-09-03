# NN Water Billing

## Environment setup

Create an isolated Python environment once per machine:

```bash
python3 -m venv .venv
source .venv/bin/activate
python -m pip install --upgrade pip
python -m pip install -r requirements.txt
```

The `requirements.txt` file installs these Python dependencies:

- `openpyxl`: read and update Excel workbooks.
- `Pillow`: load and enhance meter images.
- `pytesseract`: call Tesseract OCR from Python.

The OCR workflow also requires the system `tesseract` executable. On Debian/Ubuntu:

```bash
sudo apt-get install tesseract-ocr
```

For PDF generation, install LibreOffice as well. On Debian/Ubuntu:

```bash
sudo apt-get install libreoffice
```

On Windows, create and activate the environment with:

```powershell
python -m venv .venv
.venv\Scripts\Activate.ps1
python -m pip install --upgrade pip
python -m pip install -r requirements.txt
```

Activate the environment before each monthly run. Leave it with:

```bash
deactivate
```

The virtual environment is recommended but not mandatory if these dependencies are already installed in the selected Python interpreter.

## Monthly automation

Run from the workspace root:

```bash
python monthly_water_billing.py --month "<Month> <Year>" --config config/nn_water.json --pdf
```

By default, the command resolves paths as follows:

```text
images:  <year>/<Month>/input/
xlsx:    <year>/NN_WM_water.xlsx, with older naming patterns as fallbacks
output:  <year>/<Month>/output/
policy:  skills.md in the launch directory
```

The command updates the selected reading and cost columns, validates all meter identifiers and policy inputs, and creates Society-facing and NoBroker Hood Excel/PDF bills. Optional path arguments can override any default. If OCR cannot identify every reading, it stops before updating the workbook and writes an `*_input_review.json` report. Reviewed values can be supplied with `--readings-json`.

Use `--config` for a different apartment. Copy `config/nn_water.json`, then change the worksheet names, meter prefixes, vendor rates, tank capacity, faulty solar series, and output naming in the copy. The calculation code remains unchanged.

## Demo month

See [the August 2026 demo record](2026/August/README.md) for a real month folder containing input meter images, the annual workbook location, the command, and the generated Society-facing and NoBroker Hood outputs.

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