#!/usr/bin/env python3
"""Generate MPR YTD workbook from IR (and future MR/PR) source extracts."""

from __future__ import annotations

import argparse
import calendar
import shutil
from collections import Counter, defaultdict
from datetime import date, datetime
from pathlib import Path

from openpyxl import load_workbook

COMPLETED_STATES = {"Closed", "Resolved"}
CANCELED_STATES = {"Canceled", "Cancelled"}

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

# Monthly section TMS data rows in each sheet (Dec..Jan). Dec left blank when as-of is Nov.
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


def normalize_nature(value: str | None) -> str | None:
    if value is None:
        return None
    return " ".join(str(value).replace("\r", " ").split())


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
                "app": ws.cell(r, 1).value,
                "number": ws.cell(r, 2).value,
                "opened": opened,
                "nature": normalize_nature(ws.cell(r, 4).value),
                "priority": ws.cell(r, 5).value,
                "state": ws.cell(r, 6).value,
                "category": ws.cell(r, 7).value,
                "resolved": ws.cell(r, 8).value,
            }
        )
    return rows


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


def on_or_before(value: datetime | date | None, as_of: date) -> bool:
    d = as_date(value)
    return bool(d and d <= as_of)


def outstanding_as_of(rows: list[dict], as_of: date) -> list[dict]:
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


def month_end(year: int, month: int) -> date:
    return date(year, month, calendar.monthrange(year, month)[1])


def analyze_irs(rows: list[dict], year: int, as_of: date) -> dict:
    apps = sorted({r["app"] for r in rows if r["app"]})
    by_month: dict[int, dict] = {}

    bb = outstanding_as_of(rows, date(year - 1, 12, 31))

    for month in range(1, 13):
        if date(year, month, 1) > as_of:
            continue
        received = [
            r
            for r in rows
            if in_month(r["opened"], year, month, as_of) and r["app"]
        ]
        completed = [
            r
            for r in rows
            if in_month(r["resolved"], year, month, as_of)
            and r["state"] in COMPLETED_STATES
        ]
        # Cancelled timing not provided; attribute to open month.
        canceled = [
            r
            for r in rows
            if r["state"] in CANCELED_STATES
            and in_month(r["opened"], year, month, as_of)
        ]
        by_app = defaultdict(lambda: {
            "received": 0,
            "completed": 0,
            "canceled": 0,
            "nature": Counter(),
            "priority": Counter(),
            "valid_priority": Counter(),
        })
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

        prev_end = date(year - 1, 12, 31) if month == 1 else month_end(year, month - 1)
        if month == 11 and as_of.month == 11:
            curr_end = as_of
        elif month == as_of.month:
            curr_end = as_of
        else:
            curr_end = month_end(year, month)

        by_month[month] = {
            "by_app": dict(by_app),
            "outstanding_begin": Counter(
                r["app"] for r in outstanding_as_of(rows, prev_end)
            ),
            "outstanding_end": Counter(
                r["app"] for r in outstanding_as_of(rows, curr_end)
            ),
        }

    return {
        "apps": apps,
        "beginning_balance": Counter(r["app"] for r in bb),
        "by_month": by_month,
        "outstanding_as_of": Counter(r["app"] for r in outstanding_as_of(rows, as_of)),
    }


def fill_workbook(template: Path, output: Path, analysis: dict, year: int, as_of: date) -> None:
    shutil.copy2(template, output)
    wb = load_workbook(output)

    mpr = wb["MPR YTD Totals"]
    nature = wb["IRs Received by Nature of Error"]
    priority = wb["IRs Received By Priority - YTD "]

    # January IR beginning balance (Outstanding as of Dec prior year)
    for app, count in analysis["beginning_balance"].items():
        if app == "TMS":
            mpr["D297"] = count

    for month, data in analysis["by_month"].items():
        if month not in MPR_TMS_ROWS:
            continue
        # Only populate through as-of month; leave future months blank
        if date(year, month, 1) > as_of:
            continue

        for app, metrics in data["by_app"].items():
            if app != "TMS":
                continue

            mpr_row = MPR_TMS_ROWS[month]
            # Received / Completed / Cancelled IR columns
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
                priority[f"{col}{pri_row}"] = metrics["valid_priority"].get(
                    pri_name, 0
                )

    # Clear months after as-of (e.g. December when reporting November)
    for month in range(as_of.month + 1, 13):
        if month not in MPR_TMS_ROWS:
            continue
        mpr_row = MPR_TMS_ROWS[month]
        for col in ("G", "J", "M"):
            mpr[f"{col}{mpr_row}"] = None
        nat_row = NATURE_TMS_ROWS[month]
        for col in NATURE_COLS.values():
            nature[f"{col}{nat_row}"] = None
        pri_row = PRIORITY_TMS_ROWS[month]
        for col in list(PRIORITY_COLS.values()) + list(VALID_PRIORITY_COLS.values()):
            priority[f"{col}{pri_row}"] = None

    # Annotate report as-of on main title
    mpr["A2"] = f"MPR Summary YTD Totals - {year} (as of {as_of.strftime('%B %d, %Y')})"
    nature["A2"] = (
        f"IRS Received By Nature of Error - YTD Totals {year} "
        f"(as of {as_of.strftime('%B %d, %Y')})"
    )
    priority["A2"] = (
        f"IRS Received By Priority Level - YTD Totals {year} "
        f"(as of {as_of.strftime('%B %d, %Y')})"
    )

    wb.save(output)


def print_summary(analysis: dict, as_of: date) -> None:
    print(f"As of {as_of.isoformat()}")
    print(f"Beginning balance (IR): {dict(analysis['beginning_balance'])}")
    print(f"Outstanding as of report date: {dict(analysis['outstanding_as_of'])}")
    for month in sorted(analysis["by_month"]):
        data = analysis["by_month"][month]
        tms = data["by_app"].get("TMS", {})
        if not tms:
            continue
        print(
            f"  {month:02d}: recv={tms['received']} completed={tms['completed']} "
            f"canceled={tms['canceled']} nature={dict(tms['nature'])} "
            f"priority={dict(tms['priority'])}"
        )


def main() -> None:
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument(
        "--template",
        type=Path,
        default=Path("MPR YTD {YYYY}_as of {Date Today}.xlsx"),
    )
    parser.add_argument(
        "--source",
        type=Path,
        default=Path("source/IR_11262026.xlsx"),
    )
    parser.add_argument(
        "--output",
        type=Path,
        default=Path("output/MPR YTD 2026_as of November 25.xlsx"),
    )
    parser.add_argument("--year", type=int, default=2026)
    parser.add_argument("--as-of", type=date.fromisoformat, default=date(2026, 11, 25))
    args = parser.parse_args()

    rows = load_irs(args.source)
    analysis = analyze_irs(rows, args.year, args.as_of)
    print_summary(analysis, args.as_of)

    args.output.parent.mkdir(parents=True, exist_ok=True)
    fill_workbook(args.template, args.output, analysis, args.year, args.as_of)
    print(f"Wrote {args.output}")


if __name__ == "__main__":
    main()
