#!/usr/bin/env python3
"""Automate monthly water-meter ingestion, validation, and bill generation."""

from __future__ import annotations

import argparse
import json
import re
import shutil
import subprocess
import sys
from collections import defaultdict
from pathlib import Path

import openpyxl
from openpyxl import Workbook
from openpyxl.styles import Alignment, Font, PatternFill

from calculate_water_bill import MeterRow, build_records, normalize_series


MONTH_ALIASES = {
    "january": ("jan", "january"), "february": ("feb", "february"),
    "march": ("mar", "march"), "april": ("apr", "april"),
    "may": ("may",), "june": ("jun", "june"),
    "july": ("jul", "july"), "august": ("aug", "august"),
    "september": ("sep", "sept", "september"), "october": ("oct", "october"),
    "november": ("nov", "november"), "december": ("dec", "december"),
}


class DependencyError(RuntimeError):
    def __init__(self, missing: list[str], warnings: list[str] | None = None):
        self.missing = missing
        self.warnings = warnings or []
        super().__init__("Required billing information is missing")


def parse_args() -> argparse.Namespace:
    parser = argparse.ArgumentParser(description="Create monthly Society and NoBroker Hood water bills.")
    parser.add_argument("--image-folder", type=Path, help="Defaults to <year>/<Month>/input.")
    parser.add_argument("--input-xlsx", type=Path, help="Annual workbook; defaults to <year>/NN_WM_water.xlsx, with older naming patterns as fallbacks.")
    parser.add_argument("--output-folder", type=Path, help="Defaults to <year>/<Month>/output.")
    parser.add_argument("--skills-file", type=Path, help="Defaults to skills.md in the launch directory.")
    parser.add_argument("--month", required=True, help="Month to process, e.g. 'August 2026' or '2026-08'.")
    parser.add_argument("--config", type=Path, help="Apartment configuration JSON; defaults to config/nn_water.json.")
    parser.add_argument("--readings-json", type=Path, help="Optional reviewed OCR mapping: {flat_identifier: reading}.")
    parser.add_argument("--faulty-solar-fill-strategy", choices=["no-correction", "mean", "median"])
    parser.add_argument("--kaveri-tanks", type=int)
    parser.add_argument("--kiran-tanks", type=int)
    parser.add_argument("--pdf", action="store_true", help="Also export both generated workbooks to PDF.")
    return parser.parse_args()


def month_name(month: str) -> str:
    match = re.match(r"^\s*(\d{4})[-/ ]([01]?\d)\s*$", month)
    if match:
        names = list(MONTH_ALIASES)
        return f"{names[int(match.group(2)) - 1].title()} {match.group(1)}"
    match = re.match(r"^\s*([A-Za-z]+)\s+(\d{4})\s*$", month)
    if match and match.group(1).lower() in MONTH_ALIASES:
        return f"{match.group(1).title()} {match.group(2)}"
    raise ValueError("--month must look like 'August 2026' or '2026-08'")


def resolve_paths(args: argparse.Namespace, label: str) -> tuple[Path, Path, Path, Path]:
    year = label.split()[1]
    month = label.split()[0]
    launch_dir = Path.cwd()
    image_folder = args.image_folder or Path(year) / month / "input"
    output_folder = args.output_folder or Path(year) / month / "output"
    skills_file = args.skills_file or launch_dir / "skills.md"
    if args.input_xlsx:
        input_xlsx = args.input_xlsx
    else:
        month_number = list(MONTH_ALIASES).index(month.lower()) + 1
        candidates = [
            Path(year) / "NN_WM_water.xlsx",
            Path(year) / "NN_MM_water.xlsx",
            Path(year) / f"NN_{month_number:02d}_water.xlsx",
            Path(year) / f"NN_{month}_water.xlsx",
            Path(year) / f"NN_WM_{year}.xlsx",
            Path(year) / f"NN_Water_{year}.xlsx",
        ]
        input_xlsx = next((candidate for candidate in candidates if candidate.exists()), candidates[0])
    return image_folder, input_xlsx, output_folder, skills_file


def load_config(path: Path) -> dict[str, object]:
    try:
        config = json.loads(path.read_text(encoding="utf-8"))
    except (OSError, json.JSONDecodeError) as exc:
        raise DependencyError([f"valid configuration JSON: {path} ({exc})"]) from exc
    required = ("workbook", "vendors", "solar", "common", "outputs")
    missing = [key for key in required if key not in config]
    if missing:
        raise DependencyError([f"configuration keys: {', '.join(missing)}"])
    return config


def find_month_columns(ws, month: str) -> tuple[int, int, str]:
    label = month_name(month)
    month_word = label.split()[0].lower()
    aliases = MONTH_ALIASES[month_word]
    headers = {str(ws.cell(1, col).value).lower(): col for col in range(1, ws.max_column + 1)}
    current = next((col for header, col in headers.items() if any(alias in header for alias in aliases)), None)
    if current is None:
        raise ValueError(f"No reading column found for {label}; expected one of {aliases}")
    if current <= 1:
        raise ValueError(f"No previous reading column exists before {label}")
    return current - 1, current, label


def read_policy(path: Path) -> dict[str, object]:
    text = path.read_text(encoding="utf-8")
    faulty_match = re.search(r"Known non-working/faulty solar series[^:]*:\s*([^\n]+)", text, re.I)
    faulty = set(normalize_series(re.findall(r"\d+", faulty_match.group(1)))) if faulty_match else set()
    default_line = re.search(r"\*\*Default recommendation\*\*[^\n]*", text, re.I)
    strategy = "no-correction"
    if default_line:
        strategy_match = re.search(r"\b(no correction|mean correction|median correction)\b", default_line.group(0), re.I)
        if strategy_match:
            strategy = {"no correction": "no-correction", "mean correction": "mean", "median correction": "median"}[strategy_match.group(1).lower()]
    kiran = re.search(r"KIRAN[^\n]*?(?:₹|Rs\.?\s*)?(\d+(?:\.\d+)?)", text, re.I)
    kaveri = re.search(r"KAVERI[^\n]*?(?:₹|Rs\.?\s*)?(\d+(?:\.\d+)?)", text, re.I)
    return {
        "faulty_series": faulty,
        "fill_strategy": strategy,
        "kiran_rate": float(kiran.group(1)) if kiran else None,
        "kaveri_rate": float(kaveri.group(1)) if kaveri else None,
    }


def normalize_identifier(value: str) -> str:
    value = re.sub(r"[^A-Za-z0-9]", "", value).upper()
    if value.startswith("SOLAR"):
        value = "SOLAR_" + value[5:]
    if value.startswith("COMMON"):
        value = "COMMON_" + value[6:]
    return value


def load_identifiers(path: Path, sheet_name: str) -> list[str]:
    wb = openpyxl.load_workbook(path, data_only=True)
    if sheet_name not in wb.sheetnames:
        raise ValueError(f"Missing sheet {sheet_name!r}")
    ws = wb[sheet_name]
    return [str(ws.cell(row, 1).value) for row in range(2, ws.max_row + 1) if ws.cell(row, 1).value is not None]


def ocr_readings(folder: Path, identifiers: list[str]) -> tuple[dict[str, int], list[dict[str, object]]]:
    try:
        import pytesseract
        from PIL import Image, ImageEnhance
    except ImportError as exc:
        raise DependencyError([f"OCR dependency unavailable: {exc.name}; install pillow and pytesseract"]) from exc

    expected = {normalize_identifier(name): name for name in identifiers}
    found: dict[str, int] = {}
    review: list[dict[str, object]] = []
    image_paths = sorted(folder.glob("*.jpg")) + sorted(folder.glob("*.jpeg")) + sorted(folder.glob("*.png"))
    if not image_paths:
        raise DependencyError([f"No .jpg, .jpeg, or .png images found in {folder}"])

    for image_path in image_paths:
        image = Image.open(image_path)
        enhanced = ImageEnhance.Contrast(image.convert("L")).enhance(2.5)
        data = pytesseract.image_to_data(enhanced, config="--psm 6", output_type=pytesseract.Output.DICT)
        lines: dict[int, list[tuple[int, str]]] = defaultdict(list)
        for index, raw in enumerate(data["text"]):
            token = raw.strip()
            if token:
                lines[data["line_num"][index]].append((data["left"][index], token))
        for line_tokens in lines.values():
            line_tokens.sort()
            text = " ".join(token for _, token in line_tokens)
            normalized_text = normalize_identifier(text)
            matches = [original for normalized, original in expected.items() if normalized in normalized_text]
            numbers = [int(value) for value in re.findall(r"\d{2,}", text.replace(",", ""))]
            if len(matches) == 1 and numbers:
                # In the notebook layout, the right-most number is the current reading.
                candidate = numbers[-1]
                found[matches[0]] = candidate
            elif numbers:
                review.append({"image": str(image_path), "text": text, "numbers": numbers})
    return found, review


def parse_vendor_counts(formula: object, kaveri_rate: float, kiran_rate: float) -> tuple[int | None, int | None]:
    if not isinstance(formula, str):
        return None, None
    terms = re.findall(r"(\d+)\s*\*\s*(\d+(?:\.\d+)?)", formula)
    kaveri = kiran = None
    for count, rate in terms:
        if float(rate) == kaveri_rate:
            kaveri = int(count)
        elif float(rate) == kiran_rate:
            kiran = int(count)
    return kaveri, kiran


def load_meter_rows(path: Path, sheet: str, previous_col: int, current_col: int) -> tuple[list[MeterRow], list[str]]:
    wb = openpyxl.load_workbook(path, data_only=True)
    ws = wb[sheet]
    rows: list[MeterRow] = []
    issues: list[str] = []
    for row in range(2, ws.max_row + 1):
        name = ws.cell(row, 1).value
        if name is None:
            continue
        previous = ws.cell(row, previous_col).value
        current = ws.cell(row, current_col).value
        if current is None:
            issues.append(f"Missing current reading for {name}")
            continue
        rows.append(MeterRow(str(name), float(previous or 0), float(current)))
        if not str(name).startswith("SOLAR_") and float(current) < float(previous or 0):
            issues.append(f"Reading decreased for {name}: {previous} -> {current}")
    return rows, issues


def write_readings(path: Path, sheet: str, current_col: int, readings: dict[str, int]) -> None:
    wb = openpyxl.load_workbook(path)
    ws = wb[sheet]
    row_by_name = {str(ws.cell(row, 1).value): row for row in range(2, ws.max_row + 1)}
    for name, value in readings.items():
        if name in row_by_name:
            ws.cell(row_by_name[name], current_col).value = value
    wb.calculation.fullCalcOnLoad = True
    wb.calculation.forceFullCalc = True
    wb.calculation.calcMode = "auto"
    wb.save(path)


def style_table(ws, header_row: int, columns: int) -> None:
    fill = PatternFill("solid", fgColor="1F4E78")
    for col in range(1, columns + 1):
        cell = ws.cell(header_row, col)
        cell.fill = fill
        cell.font = Font(color="FFFFFF", bold=True)
        cell.alignment = Alignment(horizontal="center", wrap_text=True)
    ws.freeze_panes = f"A{header_row + 1}"


def update_cost_tab(path: Path, label: str, reading_column: int, meter_sheet: str, cost_sheet: str, litres_per_tank: int, kaveri_tanks: int, kiran_tanks: int, kaveri_rate: float, kiran_rate: float) -> None:
    wb = openpyxl.load_workbook(path)
    if cost_sheet not in wb.sheetnames:
        raise DependencyError([f"worksheet {cost_sheet!r}"])
    ws = wb[cost_sheet]
    month_word = label.split()[0].lower()
    aliases = MONTH_ALIASES[month_word]
    cost_column = next((col for col in range(2, ws.max_column + 1) if any(alias in str(ws.cell(1, col).value).lower() for alias in aliases)), None)
    if cost_column is None:
        raise DependencyError([f"cost-tab column for {label}"])
    reading_letter = openpyxl.utils.get_column_letter(reading_column)
    cost_letter = openpyxl.utils.get_column_letter(cost_column)
    last_row = wb[meter_sheet].max_row
    ws.cell(2, cost_column).value = f"=SUM('{meter_sheet}'!{reading_letter}2:{reading_letter}{last_row})"
    ws.cell(3, cost_column).value = f"=({kaveri_tanks}*{litres_per_tank})+({kiran_tanks}*{litres_per_tank})"
    ws.cell(4, cost_column).value = f"={cost_letter}2-{cost_letter}3"
    ws.cell(5, cost_column).value = f"=({kaveri_tanks}*{kaveri_rate})+({kiran_tanks}*{kiran_rate})"
    wb.calculation.fullCalcOnLoad = True
    wb.calculation.forceFullCalc = True
    wb.calculation.calcMode = "auto"
    wb.save(path)


def create_bills(output: Path, label: str, rows: list[MeterRow], previous_consumption: dict[str, float], tanker_cost: float, tanker_litres: int, faulty: set[str], strategy: str, solar_prefix: str, common_prefix: str) -> tuple[Path, Path, dict[str, object]]:
    summary, records, _ = build_records(rows, tanker_cost, faulty, strategy, solar_prefix, common_prefix)
    output.mkdir(parents=True, exist_ok=True)
    society_path = output / f"{label.replace(' ', '_')}_Final_Society_Bill.xlsx"
    nbh_path = output / f"{label.replace(' ', '_')}_Final_Society_Bill_NoBroker_Hood.xlsx"

    for path, include_analysis in ((society_path, True), (nbh_path, False)):
        wb = Workbook()
        ws = wb.active
        ws.title = f"{label} Bill"
        ws.append([f"{label} Water Bill"])
        ws.append([])
        ws.append(["Water tanker cost paid", tanker_cost])
        ws.append(["Common cost per flat", summary["Solar as common per flat"] + summary["Common facility per flat"]])
        ws.append(["Total flats billed", summary["Flat count"]])
        headers = ["Flat No", "Previous Reading", "Current Reading", "Consumption (x10 Litres)", "Individual Cost", "Common Cost", "Total Amount"]
        if include_analysis:
            headers += ["Last Month Consumption (x10 Litres)", "Change vs Last Month (x10 Litres)", "Change vs Last Month (%)", "Mean Consumption (x10 Litres)", "Consumption vs Mean (x10 Litres)"]
        ws.append(headers)
        style_table(ws, 6, len(headers))
        for record in records:
            name = str(record["Flat"])
            row = [name, record["Previous Reading"], record["Current Reading"], record["Consumption"], record["Individual Share"], record["Solar Share Common"] + record["Common Share"], record["Total Solar Common"]]
            if include_analysis:
                current_consumption = float(record["Consumption"])
                last_month = previous_consumption[name]
                row += [last_month, current_consumption - last_month, (current_consumption - last_month) / last_month if last_month else "", summary["Individual consumption"] / summary["Flat count"], current_consumption - summary["Individual consumption"] / summary["Flat count"]]
            ws.append(row)
        ws.append(["Total", None, None, f"=SUM(D7:D{6 + len(records)})", f"=SUM(E7:E{6 + len(records)})", f"=SUM(F7:F{6 + len(records)})", f"=SUM(G7:G{6 + len(records)})"])
        for col in range(1, len(headers) + 1):
            ws.column_dimensions[openpyxl.utils.get_column_letter(col)].width = 22
        ws.sheet_properties.pageSetUpPr.fitToPage = True
        ws.page_setup.fitToWidth = 1
        ws.page_setup.fitToHeight = 0
        wb.save(path)

    summary_path = output / f"{label.replace(' ', '_')}_Billing_Summary.xlsx"
    wb = Workbook()
    ws = wb.active
    ws.title = "Summary"
    ws.append(["Metric", "Value"])
    for key, value in summary.items(): ws.append([key, value])
    ws.append([])
    ws.append([f"{label} Water Source Summary", None])
    ws.append(["Source", "Quantity (litres)"])
    raw_total = sum(row.consumption for row in rows)
    borewell_litres = (raw_total - tanker_litres) * 10
    ws.append(["Tanker water", tanker_litres * 10])
    ws.append(["Borewell water", borewell_litres])
    ws.append(["Total water consumption", raw_total * 10])
    style_table(ws, 1, 2)
    style_table(ws, len(summary) + 4, 2)
    wb.save(summary_path)
    return society_path, nbh_path, summary


def export_pdf(workbook_path: Path, output_folder: Path) -> Path | None:
    if shutil.which("libreoffice") is None:
        return None
    subprocess.run(["libreoffice", "--headless", "--convert-to", "pdf", "--outdir", str(output_folder), str(workbook_path)], check=True, stdout=subprocess.PIPE, stderr=subprocess.PIPE)
    return workbook_path.with_suffix(".pdf")


def main() -> int:
    args = parse_args()
    try:
        label = month_name(args.month)
        args.image_folder, args.input_xlsx, args.output_folder, args.skills_file = resolve_paths(args, label)
        config_path = args.config or Path("config/nn_water.json")
        config = load_config(config_path)
        workbook_config = config["workbook"]
        vendor_config = config["vendors"]
        solar_config = config["solar"]
        common_config = config["common"]
        sheet = str(workbook_config["meter_sheet"])
        cost_sheet = str(workbook_config["cost_sheet"])
        policy = read_policy(args.skills_file)
        policy["faulty_series"] = set(normalize_series(solar_config["faulty_series"]))
        kaveri = vendor_config["KAVERI"]
        kiran = vendor_config["KIRAN"]
        kaveri_rate = float(kaveri["rate_per_tank"])
        kiran_rate = float(kiran["rate_per_tank"])
        litres_per_tank = int(kaveri["litres_per_tank"])
        missing: list[str] = []
        if kaveri_rate <= 0: missing.append("positive KAVERI rate in config")
        if kiran_rate <= 0: missing.append("positive KIRAN rate in config")
        wb = openpyxl.load_workbook(args.input_xlsx, data_only=False)
        if sheet not in wb.sheetnames: missing.append(f"worksheet {sheet!r}")
        if cost_sheet not in wb.sheetnames: missing.append(f"worksheet {cost_sheet!r}")
        if missing: raise DependencyError(missing)
        previous_col, current_col, _ = find_month_columns(wb[sheet], label)
        identifiers = load_identifiers(args.input_xlsx, sheet)
        ocr_found, review = ocr_readings(args.image_folder, identifiers)
        if args.readings_json:
            ocr_found.update({str(k): int(v) for k, v in json.loads(args.readings_json.read_text()).items()})
        expected = set(identifiers)
        missing_readings = sorted(expected - set(ocr_found))
        if missing_readings:
            report = args.output_folder / f"{label.replace(' ', '_')}_input_review.json"
            args.output_folder.mkdir(parents=True, exist_ok=True)
            report.write_text(json.dumps({"missing_readings": missing_readings, "ocr_review": review}, indent=2), encoding="utf-8")
            raise DependencyError([f"{len(missing_readings)} meter readings could not be identified; review {report}"])
        shutil.copy2(args.input_xlsx, args.input_xlsx.with_name(args.input_xlsx.stem + "_before_monthly_update.xlsx"))
        write_readings(args.input_xlsx, sheet, current_col, ocr_found)
        wb = openpyxl.load_workbook(args.input_xlsx, data_only=False)
        cost_ws = wb[cost_sheet] if cost_sheet in wb.sheetnames else None
        formula = None
        if cost_ws:
            month_word = label.split()[0].lower()
            formula = next((cost_ws.cell(5, col).value for col in range(2, cost_ws.max_column + 1) if any(alias in str(cost_ws.cell(1, col).value).lower() for alias in MONTH_ALIASES[month_word])), None)
        kaveri_tanks, kiran_tanks = parse_vendor_counts(formula, kaveri_rate, kiran_rate)
        kaveri_tanks = args.kaveri_tanks if args.kaveri_tanks is not None else kaveri_tanks
        kiran_tanks = args.kiran_tanks if args.kiran_tanks is not None else kiran_tanks
        if kaveri_tanks is None or kiran_tanks is None:
            raise DependencyError(["KAVERI/KIRAN tank counts in the cost formula or CLI arguments"])
        tanker_cost = kaveri_tanks * kaveri_rate + kiran_tanks * kiran_rate
        strategy = args.faulty_solar_fill_strategy or str(policy["fill_strategy"])
        rows, issues = load_meter_rows(args.input_xlsx, sheet, previous_col, current_col)
        if issues: raise DependencyError(issues)
        source_ws = openpyxl.load_workbook(args.input_xlsx, data_only=True)[sheet]
        previous_consumption = {}
        for row_number in range(2, source_ws.max_row + 1):
            name = source_ws.cell(row_number, 1).value
            if name is not None and not str(name).startswith(("SOLAR_", "Common_")):
                previous_consumption[str(name)] = float(source_ws.cell(row_number, previous_col).value or 0) - float(source_ws.cell(row_number, previous_col - 1).value or 0)
        tanker_litres = (kaveri_tanks + kiran_tanks) * 700
        update_cost_tab(args.input_xlsx, label, current_col, sheet, cost_sheet, litres_per_tank, kaveri_tanks, kiran_tanks, kaveri_rate, kiran_rate)
        society_path, nbh_path, summary = create_bills(args.output_folder, label, rows, previous_consumption, tanker_cost, tanker_litres, set(policy["faulty_series"]), strategy, str(solar_config["prefix"]), str(common_config["prefix"]))
        pdfs = [export_pdf(path, args.output_folder) for path in (society_path, nbh_path)] if args.pdf else []
        print(json.dumps({"society_bill": str(society_path), "nobroker_hood_bill": str(nbh_path), "pdfs": [str(p) for p in pdfs if p], "total_tanker_cost": tanker_cost, "summary": summary}, indent=2, default=str))
        return 0
    except DependencyError as exc:
        print("BILLING STOPPED: required information is missing or invalid.", file=sys.stderr)
        for item in exc.missing: print(f"- {item}", file=sys.stderr)
        for item in exc.warnings: print(f"Warning: {item}", file=sys.stderr)
        return 2
    except (KeyError, OSError, TypeError, ValueError, json.JSONDecodeError) as exc:
        print(f"BILLING STOPPED: {exc}", file=sys.stderr)
        return 1


if __name__ == "__main__":
    raise SystemExit(main())
