# August 2026 Demo Record

This folder is the complete example of one monthly billing run.

## Demo input

- Meter images: `input/*.jpeg`
- Annual workbook: `../NN_WM_water.xlsx`
- Policy: `../../skills.md`

The annual workbook contains one row per identifier in `water_meter_reading_quarter`:

| Column | Meaning |
|---|---|
| `Flat No` | Resident, solar, or common-meter identifier |
| `Jan_Reading` through the previous month | Historical meter readings |
| `Aug_Reading` | Current month reading to populate |

Identifiers include resident flats such as `G01`, `101`, and `317`, solar meters such as `SOLAR_G01`, and common meters such as `Common_G01`.

The `cost` sheet contains the monthly tanker quantities and cost formulas. Current vendor policy is KAVERI at `₹740` per tank and KIRAN at `₹950` per tank, with `700` litres per tank.

## Run the demo

From the workspace root:

```bash
python monthly_water_billing.py --month "August 2026" --pdf
```

The command automatically uses:

```text
Input images:  2026/August/input/
Annual XLSX:   2026/NN_WM_water.xlsx
Policy:        skills.md
Output:        2026/August/output/
```

If OCR cannot identify every row, the command stops and writes `August_2026_input_review.json` in the output folder. Corrected readings can be supplied with `--readings-json`.

## Demo output

The generated output folder contains:

- `August_2026_Final_Society_Bill.xlsx`: detailed bill and comparison analysis.
- `August_2026_Final_Society_Bill.pdf`: Society-facing PDF.
- `August_2026_Final_Society_Bill_NoBroker_Hood.xlsx`: resident-sharing workbook.
- `August_2026_Final_Society_Bill_NoBroker_Hood.pdf`: resident-sharing PDF without comparison columns.

This example demonstrates the year/month convention. For another year, use the same layout under `<year>/<Month>/` and keep one annual workbook at `<year>/NN_WM_water.xlsx`.