#!/usr/bin/env python3
"""Generate MPR YTD workbook from IR and MR source extracts."""

from __future__ import annotations

import argparse
import calendar
import shutil
from collections import Counter, defaultdict
from datetime import date, datetime
from pathlib import Path

from openpyxl import load_workbook

IR_COMPLETED_STATES = {"Closed", "Resolved"}
MR_COMPLETED_STATES = {"Closed"}
CANCELED_STATES = {"Canceled", "Cancelled"}
MR_OPEN_STATES = {"Open", "Assess", "Authorize"}

DEFAULT_APP = "TMS"

NATURE_COLS = {
    "Connect Direct": "B",
    "Data Related / User Error": "C",
    "Space Issue": "D",
    "Other Batch Related": "E",
    "Program Error": "F",
    "Environment": "G",
    "Not an Error / Invalid": "H",
}

PRIORITY_COLS = {
    "1 - Critical": "B",
    "2 - High": "C",
    "3 - Moderate": "D",
    "4 - Low": "E",
}

VALID_PRIORITY_COLS = {
    "1 - Critical": "I",
    "2 - High": "J",
    "3 - Moderate": "K",
    "4 - Low": "L",
}

MR_CATEGORY_COLS = {
    "Problem Resolution": "B",
    "Incident Resolution": "C",
    "Compliance": "D",
    "Audit": "E",
    "OS, DB Patches and Vendor Fixes": "F",
    "User Initiated Enhancements": "G",
    "ISG Initiated Enhancements": "H",
    "Project / Integration Support": "I",
    "Adhoc": "J",
}

ADHOC_TYPE_COLS = {
    "data / information extract": "B",
    "generate statistics": "C",
    "generate report": "D",
    "Compliance-related requests": "E",
    "others": "F",
    "Other": "F",
}

# Monthly TMS rows (Dec..Jan) in MPR YTD Totals
MPR_TMS_ROWS = {
    12: 44,
    11: 67,
    10: 90,
    9: 113,
    8: 136,
    7: 159,
    6: 182,
    5: 205,
    4: 228,
    3: 251,
    2: 274,
    1: 297,
}

NATURE_TMS_ROWS = {
    12: 44,
    11: 66,
    10: 88,
    9: 110,
    8: 132,
    7: 154,
    6: 176,
    5: 198,
    4: 220,
    3: 242,
    2: 264,
    1: 286,
}

PRIORITY_TMS_ROWS = {
    12: 44,
    11: 66,
    10: 88,
    9: 110,
    8: 132,
    7: 154,
    6: 176,
    5: 199,
    4: 221,
    3: 243,
    2: 266,
    1: 289,
}

# Completed MRs By Category - TMS row per month
MR_CATEGORY_TMS_ROWS = {
    12: 26,
    11: 49,
    10: 72,
    9: 95,
    8: 118,
    7: 141,
    6: 164,
    5: 187,
    4: 210,
    3: 233,
    2: 256,
    1: 279,
}

# Completed Adhoc MR - TMS row per month
ADHOC_TMS_ROWS = {
    12: 42,
    11: 64,
    10: 86,
    9: 108,
    8: 130,
    7: 152,
    6: 174,
    5: 196,
    4: 218,
    3: 240,
    2: 262,
    1: 284,
}


def normalize_text(value: str | None) -> str | None:
    if value is None:
        return None
    return " ".join(str(value).replace("\r", " ").split())


def as_date(value: datetime | date | None) -> date | None:
    if value is None:
        return None
    if isinstance(value, datetime):
        return value.date()
    if isinstance(value, date):
        return value
    return None


def in_month(value: datetime | date | None, year: int, month: int, as_of: date) -> bool:
    d = as_date(value)
    return bool(d and d.year == year and d.month == month and d <= as_of)


def month_end(year: int, month: int) -> date:
    return date(year, month, calendar.monthrange(year, month)[1])


def load_irs(path: Path) -> list[dict]:
    wb = load_workbook(path, data_only=True)
    ws = wb[wb.sheetnames[0]]
    rows: list[dict] = []
    for r in range(2, (ws.max_row or 1) + 1):
        opened = ws.cell(r, 3).value
        if opened is None:
            continue
        rows.append(
            {
                "app": ws.cell(r, 1).value or DEFAULT_APP,
                "number": ws.cell(r, 2).value,
                "opened": opened,
                "nature": normalize_text(ws.cell(r, 4).value),
                "priority": ws.cell(r, 5).value,
                "state": ws.cell(r, 6).value,
                "category": ws.cell(r, 7).value,
                "resolved": ws.cell(r, 8).value,
            }
        )
    return rows


def load_mrs(path: Path) -> list[dict]:
    wb = load_workbook(path, data_only=True)
    ws = wb[wb.sheetnames[0]]
    rows: list[dict] = []
    for r in range(2, (ws.max_row or 1) + 1):
        ch = ws.cell(r, 1).value
        short = ws.cell(r, 2).value
        ritm = ws.cell(r, 3).value
        state = ws.cell(r, 7).value
        opened = ws.cell(r, 9).value
        closed = ws.cell(r, 8).value
        if not any([ch, short, ritm, state, opened, closed]):
            continue
        rows.append(
            {
                "app": DEFAULT_APP,
                "number": ch or ritm,
                "ch": ch,
                "ritm": ritm,
                "short": short,
                "desc": ws.cell(r, 4).value,
                "category": normalize_text(ws.cell(r, 5).value),
                "adhoc_type": normalize_text(ws.cell(r, 6).value),
                "state": state,
                "opened": opened,
                "closed": closed,
            }
        )
    return rows


def ir_outstanding_as_of(rows: list[dict], as_of: date) -> list[dict]:
    out = []
    for row in rows:
        opened = as_date(row["opened"])
        if not opened or opened > as_of:
            continue
        if row["state"] in CANCELED_STATES:
            continue
        resolved = as_date(row["resolved"])
        if resolved and resolved <= as_of:
            continue
        out.append(row)
    return out


def mr_outstanding_as_of(rows: list[dict], as_of: date) -> list[dict]:
    out = []
    for row in rows:
        opened = as_date(row["opened"])
        closed = as_date(row["closed"])
        if opened and opened > as_of:
            continue
        if row["state"] in CANCELED_STATES:
            continue
        if closed and closed <= as_of:
            continue
        if row["state"] in MR_COMPLETED_STATES and closed is None:
            continue
        # Undated open-like tickets count as outstanding
        if opened is None and row["state"] not in MR_OPEN_STATES:
            continue
        if opened is None and row["state"] in MR_OPEN_STATES:
            out.append(row)
            continue
        if row["state"] in MR_OPEN_STATES or (
            row["state"] not in MR_COMPLETED_STATES | CANCELED_STATES
        ):
            out.append(row)
            continue
        # Closed/resolved after as_of still outstanding as of as_of
        if closed and closed > as_of:
            out.append(row)
    return out


def analyze_irs(rows: list[dict], year: int, as_of: date) -> dict:
    by_month: dict[int, dict] = {}
    bb = ir_outstanding_as_of(rows, date(year - 1, 12, 31))

    for month in range(1, 13):
        if date(year, month, 1) > as_of:
            continue
        received = [
            r for r in rows if in_month(r["opened"], year, month, as_of) and r["app"]
        ]
        completed = [
            r
            for r in rows
            if in_month(r["resolved"], year, month, as_of)
            and r["state"] in IR_COMPLETED_STATES
        ]
        canceled = [
            r
            for r in rows
            if r["state"] in CANCELED_STATES
            and in_month(r["opened"], year, month, as_of)
        ]
        by_app: dict = defaultdict(
            lambda: {
                "received": 0,
                "completed": 0,
                "canceled": 0,
                "nature": Counter(),
                "priority": Counter(),
                "valid_priority": Counter(),
            }
        )
        for r in received:
            app = r["app"]
            by_app[app]["received"] += 1
            if r["nature"]:
                by_app[app]["nature"][r["nature"]] += 1
            if r["priority"]:
                by_app[app]["priority"][r["priority"]] += 1
                if r["state"] not in CANCELED_STATES:
                    by_app[app]["valid_priority"][r["priority"]] += 1
        for r in completed:
            by_app[r["app"]]["completed"] += 1
        for r in canceled:
            by_app[r["app"]]["canceled"] += 1
        by_month[month] = {"by_app": dict(by_app)}

    return {
        "beginning_balance": Counter(r["app"] for r in bb),
        "by_month": by_month,
        "outstanding_as_of": Counter(
            r["app"] for r in ir_outstanding_as_of(rows, as_of)
        ),
    }


def analyze_mrs(rows: list[dict], year: int, as_of: date) -> dict:
    by_month: dict[int, dict] = {}

    for month in range(1, 13):
        if date(year, month, 1) > as_of:
            continue
        received = [r for r in rows if in_month(r["opened"], year, month, as_of)]
        completed = [
            r
            for r in rows
            if in_month(r["closed"], year, month, as_of)
            and r["state"] in MR_COMPLETED_STATES
        ]
        canceled = [
            r
            for r in rows
            if r["state"] in CANCELED_STATES
            and in_month(r["opened"], year, month, as_of)
        ]
        by_app: dict = defaultdict(
            lambda: {
                "received": 0,
                "completed": 0,
                "canceled": 0,
                "categories": Counter(),
                "adhoc_types": Counter(),
            }
        )
        for r in received:
            by_app[r["app"]]["received"] += 1
        for r in completed:
            app = r["app"]
            by_app[app]["completed"] += 1
            if r["category"]:
                by_app[app]["categories"][r["category"]] += 1
            if r["category"] == "Adhoc":
                adhoc_key = r["adhoc_type"] or "others"
                by_app[app]["adhoc_types"][adhoc_key] += 1
        for r in canceled:
            by_app[r["app"]]["canceled"] += 1
        by_month[month] = {"by_app": dict(by_app)}

    # Key accomplishments: completed MRs on/before as_of with a usable description
    accomplishments = []
    for r in rows:
        if r["state"] not in MR_COMPLETED_STATES:
            continue
        closed = as_date(r["closed"])
        if closed and closed > as_of:
            continue
        if closed is None:
            continue
        desc = r["desc"] or r["short"]
        if not desc:
            continue
        if desc.strip() == "ESR: Application - User maintenance request":
            desc = r["desc"] or desc
        accomplishments.append(
            {
                "app": r["app"],
                "description": normalize_text(desc),
                "category": r["category"],
                "closed": closed,
                "number": r["number"],
            }
        )
    accomplishments.sort(key=lambda x: x["closed"] or date.min, reverse=True)

    return {
        "by_month": by_month,
        "outstanding_as_of": Counter(
            r["app"] for r in mr_outstanding_as_of(rows, as_of)
        ),
        "accomplishments": accomplishments,
    }


def fill_workbook(
    template: Path,
    output: Path,
    ir_analysis: dict,
    mr_analysis: dict | None,
    year: int,
    as_of: date,
) -> None:
    shutil.copy2(template, output)
    wb = load_workbook(output)

    mpr = wb["MPR YTD Totals"]
    nature = wb["IRs Received by Nature of Error"]
    priority = wb["IRs Received By Priority - YTD "]
    mr_cat = wb["Completed MRs By Category - YTD"]
    adhoc = wb["Completed Adhoc MR YTD Totals"]
    key_acc = wb["MR - Key Acconplishment"]

    # Preserve template Jan MR beginning balance (B297) when extract has no prior-year BB.
    # IR beginning balance from extract.
    for app, count in ir_analysis["beginning_balance"].items():
        if app == "TMS":
            mpr["D297"] = count

    for month, data in ir_analysis["by_month"].items():
        if month not in MPR_TMS_ROWS or date(year, month, 1) > as_of:
            continue
        metrics = data["by_app"].get("TMS")
        if not metrics:
            continue
        mpr_row = MPR_TMS_ROWS[month]
        mpr[f"G{mpr_row}"] = metrics["received"]
        mpr[f"J{mpr_row}"] = metrics["completed"]
        mpr[f"M{mpr_row}"] = metrics["canceled"]

        nat_row = NATURE_TMS_ROWS[month]
        for nat_name, col in NATURE_COLS.items():
            nature[f"{col}{nat_row}"] = metrics["nature"].get(nat_name, 0)

        pri_row = PRIORITY_TMS_ROWS[month]
        for pri_name, col in PRIORITY_COLS.items():
            priority[f"{col}{pri_row}"] = metrics["priority"].get(pri_name, 0)
        for pri_name, col in VALID_PRIORITY_COLS.items():
            priority[f"{col}{pri_row}"] = metrics["valid_priority"].get(pri_name, 0)

    if mr_analysis:
        for month, data in mr_analysis["by_month"].items():
            if month not in MPR_TMS_ROWS or date(year, month, 1) > as_of:
                continue
            metrics = data["by_app"].get("TMS")
            if not metrics:
                # Still clear category leftovers for empty months
                metrics = {
                    "received": 0,
                    "completed": 0,
                    "canceled": 0,
                    "categories": Counter(),
                    "adhoc_types": Counter(),
                }

            mpr_row = MPR_TMS_ROWS[month]
            mpr[f"E{mpr_row}"] = metrics["received"]
            mpr[f"H{mpr_row}"] = metrics["completed"]
            mpr[f"K{mpr_row}"] = metrics["canceled"]

            cat_row = MR_CATEGORY_TMS_ROWS[month]
            for cat_name, col in MR_CATEGORY_COLS.items():
                mr_cat[f"{col}{cat_row}"] = metrics["categories"].get(cat_name, 0)

            adhoc_row = ADHOC_TMS_ROWS[month]
            for type_name, col in ADHOC_TYPE_COLS.items():
                # Avoid double-writing F for both "others" and "Other"
                if type_name == "others":
                    count = metrics["adhoc_types"].get("others", 0) + metrics[
                        "adhoc_types"
                    ].get("Other", 0)
                    adhoc[f"{col}{adhoc_row}"] = count
                elif type_name == "Other":
                    continue
                else:
                    adhoc[f"{col}{adhoc_row}"] = metrics["adhoc_types"].get(
                        type_name, 0
                    )

        # Key accomplishments (newest completed first)
        # Clear existing data rows under header
        for r in range(2, key_acc.max_row + 1):
            for c in range(1, 4):
                key_acc.cell(r, c).value = None
        for idx, item in enumerate(mr_analysis["accomplishments"], start=2):
            key_acc.cell(idx, 1).value = item["app"]
            key_acc.cell(idx, 2).value = item["description"]
            key_acc.cell(idx, 3).value = item["category"]

    # Clear months after as-of
    for month in range(as_of.month + 1, 13):
        if month not in MPR_TMS_ROWS:
            continue
        mpr_row = MPR_TMS_ROWS[month]
        for col in ("E", "G", "H", "J", "K", "M"):
            mpr[f"{col}{mpr_row}"] = None
        nat_row = NATURE_TMS_ROWS[month]
        for col in NATURE_COLS.values():
            nature[f"{col}{nat_row}"] = None
        pri_row = PRIORITY_TMS_ROWS[month]
        for col in list(PRIORITY_COLS.values()) + list(VALID_PRIORITY_COLS.values()):
            priority[f"{col}{pri_row}"] = None
        if month in MR_CATEGORY_TMS_ROWS:
            cat_row = MR_CATEGORY_TMS_ROWS[month]
            for col in MR_CATEGORY_COLS.values():
                mr_cat[f"{col}{cat_row}"] = None
        if month in ADHOC_TMS_ROWS:
            adhoc_row = ADHOC_TMS_ROWS[month]
            for col in ("B", "C", "D", "E", "F"):
                adhoc[f"{col}{adhoc_row}"] = None

    as_of_label = as_of.strftime("%B %d, %Y")
    mpr["A2"] = f"MPR Summary YTD Totals - {year} (as of {as_of_label})"
    nature["A2"] = (
        f"IRS Received By Nature of Error - YTD Totals {year} (as of {as_of_label})"
    )
    priority["A2"] = (
        f"IRS Received By Priority Level - YTD Totals {year} (as of {as_of_label})"
    )
    mr_cat["A3"] = f"Completed MRs By Category YTD Totals {year} (as of {as_of_label})"
    adhoc["A2"] = f"Completed Adhoc (MR) YTD Totals {year} (as of {as_of_label})"

    wb.save(output)


def print_summary(ir_analysis: dict, mr_analysis: dict | None, as_of: date) -> None:
    print(f"As of {as_of.isoformat()}")
    print(f"IR beginning balance: {dict(ir_analysis['beginning_balance'])}")
    print(f"IR outstanding: {dict(ir_analysis['outstanding_as_of'])}")
    for month in sorted(ir_analysis["by_month"]):
        tms = ir_analysis["by_month"][month]["by_app"].get("TMS", {})
        if not tms:
            continue
        print(
            f"  IR {month:02d}: recv={tms['received']} completed={tms['completed']} "
            f"canceled={tms['canceled']}"
        )
    if not mr_analysis:
        return
    print(f"MR outstanding: {dict(mr_analysis['outstanding_as_of'])}")
    for month in sorted(mr_analysis["by_month"]):
        tms = mr_analysis["by_month"][month]["by_app"].get("TMS", {})
        if not tms:
            continue
        print(
            f"  MR {month:02d}: recv={tms['received']} completed={tms['completed']} "
            f"canceled={tms['canceled']} categories={dict(tms['categories'])}"
        )
    print(f"Key accomplishments listed: {len(mr_analysis['accomplishments'])}")


def discover_source(source_dir: Path, prefix: str) -> Path | None:
    matches = sorted(source_dir.glob(f"{prefix}_*.xlsx"))
    return matches[-1] if matches else None


def main() -> None:
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument(
        "--template",
        type=Path,
        default=Path("MPR YTD {YYYY}_as of {Date Today}.xlsx"),
    )
    parser.add_argument("--source-dir", type=Path, default=Path("source"))
    parser.add_argument("--ir-source", type=Path, default=None)
    parser.add_argument("--mr-source", type=Path, default=None)
    parser.add_argument(
        "--output",
        type=Path,
        default=Path("output/MPR YTD 2026_as of November 25.xlsx"),
    )
    parser.add_argument("--year", type=int, default=2026)
    parser.add_argument("--as-of", type=date.fromisoformat, default=date(2026, 11, 25))
    args = parser.parse_args()

    ir_path = args.ir_source or discover_source(args.source_dir, "IR")
    mr_path = args.mr_source or discover_source(args.source_dir, "MR")
    if not ir_path or not ir_path.exists():
        raise SystemExit(f"IR source not found in {args.source_dir}")

    ir_rows = load_irs(ir_path)
    ir_analysis = analyze_irs(ir_rows, args.year, args.as_of)

    mr_analysis = None
    if mr_path and mr_path.exists():
        mr_rows = load_mrs(mr_path)
        mr_analysis = analyze_mrs(mr_rows, args.year, args.as_of)
        print(f"Loaded MR source: {mr_path} ({len(mr_rows)} rows)")
    else:
        print("No MR source found; IR-only fill")

    print(f"Loaded IR source: {ir_path} ({len(ir_rows)} rows)")
    print_summary(ir_analysis, mr_analysis, args.as_of)

    args.output.parent.mkdir(parents=True, exist_ok=True)
    fill_workbook(
        args.template, args.output, ir_analysis, mr_analysis, args.year, args.as_of
    )
    print(f"Wrote {args.output}")


if __name__ == "__main__":
    main()
