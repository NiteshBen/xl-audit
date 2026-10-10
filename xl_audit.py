#!/usr/bin/env python3
"""xl-audit: lint an Excel workbook (.xlsx) for formula landmines.

Catches the bugs that make inherited workbooks quietly wrong:
  - SUM ranges that stop short of the data (the "inserted row" bug)
  - broken references (#REF!, #NAME? ...)
  - cached error values in saved results (#DIV/0!, #N/A ...)
  - copy-paste-broken formulas that differ from their column's pattern
  - hardcoded constants inside formulas (=D2*1.18)

Only needs openpyxl. Exit codes: 0 clean, 1 problems found, 2 unreadable file.
"""
from __future__ import annotations

import argparse
import json
import re
import sys
from collections import Counter
from dataclasses import dataclass, field
from pathlib import Path

from openpyxl import load_workbook
from openpyxl.utils import column_index_from_string, get_column_letter

EXCEL_ERRORS = {"#REF!", "#DIV/0!", "#VALUE!", "#NAME?", "#N/A",
                "#NULL!", "#NUM!", "#GETTING_DATA"}
CELL_REF = r"\$?[A-Za-z]{1,3}\$?\d+(?:\s*:\s*\$?[A-Za-z]{1,3}\$?\d+)?"


@dataclass
class Finding:
    check: str
    severity: str          # "error" | "warn"
    title: str
    detail: str = ""
    examples: list = field(default_factory=list)

    def to_dict(self) -> dict:
        return {"check": self.check, "severity": self.severity, "title": self.title,
                "detail": self.detail, "examples": list(self.examples)}


def is_formula(value) -> bool:
    return isinstance(value, str) and value.startswith("=")


def analyze(path: Path, max_examples: int = 5) -> tuple[list[Finding], dict]:
    wb_f = load_workbook(path, data_only=False)   # formulas
    wb_v = load_workbook(path, data_only=True)    # last saved values
    findings: list[Finding] = []
    n_formulas = 0
    n_sheets = len(wb_f.worksheets)

    for ws_f, ws_v in zip(wb_f.worksheets, wb_v.worksheets):
        sheet = ws_f.title
        col_patterns: dict[str, list[tuple[str, str]]] = {}
        hardcoded_cols: dict[str, list[str]] = {}

        for row_f, row_v in zip(ws_f.iter_rows(), ws_v.iter_rows()):
            for cf, cv in zip(row_f, row_v):
                coord = f"{sheet}!{cf.coordinate}"
                formula = cf.value
                cached = cv.value

                # -- cached error values (fires on files saved by Excel/LibreOffice)
                if isinstance(cached, str) and cached in EXCEL_ERRORS:
                    findings.append(Finding(
                        "cached-errors", "error",
                        f"{coord} evaluates to {cached}",
                        "The last time this file was saved, the formula already "
                        "returned an error."))

                if not is_formula(formula):
                    continue
                n_formulas += 1

                # -- broken references in the formula text
                broken = [e for e in ("#REF!", "#NAME?") if e in formula]
                if broken:
                    findings.append(Finding(
                        "broken-refs", "error",
                        f"{coord} references {', '.join(broken)}",
                        "A referenced sheet, range or name was deleted.", [formula]))

                # -- hardcoded constants (=D2*1.18), grouped per column.
                # Only decimals and 3+ digit integers: small integer multipliers
                # (=A2*2) are common and rarely the bug.
                stripped = re.sub(r'"[^"]*"', '""', formula)          # text literals
                stripped = re.sub(r"[A-Za-z_]+\s*\(", "(", stripped)  # fn names
                stripped = re.sub(CELL_REF, "", stripped)             # refs & ranges
                if re.search(r"\d+\.\d+|\d{3,}", stripped):
                    hardcoded_cols.setdefault(cf.column_letter, []).append(coord)

                # -- collect patterns for the column-consistency check
                pattern = re.sub(r"(\$?)([A-Za-z]{1,3})(\$?)(\d+)",
                                 lambda m: m.group(2).upper() + "{}", formula)
                col = cf.column_letter
                col_patterns.setdefault(col, []).append((coord, pattern))

                # -- SUM scope check
                for m in re.finditer(
                        r"SUM\(\s*\$?([A-Za-z]{1,3})\$?(\d+)\s*:\s*\$?([A-Za-z]{1,3})\$?(\d+)\s*\)",
                        formula, re.IGNORECASE):
                    c1, r1, c2, r2 = m.groups()
                    if c1.upper() != c2.upper():
                        continue          # horizontal/multi-column range: skipped
                    col_idx = column_index_from_string(c1.upper())
                    r_end = int(r2)
                    # first cell below the range that is data, not another formula
                    for rr in range(r_end + 1, r_end + 4):
                        below_v = ws_v.cell(row=rr, column=col_idx).value
                        below_f = ws_f.cell(row=rr, column=col_idx).value
                        if below_v is None or (isinstance(below_v, str) and not below_v.strip()):
                            break         # gap: the range genuinely ends here
                        if is_formula(below_f):
                            continue      # a subtotal/other formula, not missed data
                        findings.append(Finding(
                            "sum-scope", "error",
                            f"{coord} sums {c1.upper()}{r1}:{c2.upper()}{r2} "
                            f"but data continues at {c1.upper()}{rr}",
                            "A row was probably inserted below the range: it is "
                            "silently excluded from the total.", [formula]))
                        break

        for col, cells in hardcoded_cols.items():
            findings.append(Finding(
                "hardcoded", "warn",
                f"{len(cells)} formula(s) in column {col} of {sheet!r} contain "
                f"hardcoded constants",
                "Rates and magic numbers belong in a labelled cell, not inside "
                "formulas.", cells[:max_examples]))

        # -- copy-paste drift: minority formulas in a column
        for col, cells in col_patterns.items():
            if len(cells) < 3:
                continue
            common, _ = Counter(p for _, p in cells).most_common(1)[0]
            odd = [coord for coord, p in cells if p != common]
            if odd and len(odd) < len(cells):        # ignore all-different columns
                findings.append(Finding(
                    "formula-drift", "warn",
                    f"{len(odd)} formula(s) in column {col} of {sheet!r} differ "
                    f"from the column's pattern",
                    "Classic copy-paste slip: one row computes something else "
                    "than its neighbours.", odd[:max_examples]))

    findings.sort(key=lambda f: (f.severity != "error", f.check))
    summary = {"file": path.name, "sheets": n_sheets, "formulas": n_formulas,
               "problems": sum(1 for f in findings)}
    return findings, summary


def render(findings: list[Finding], summary: dict) -> str:
    lines = [f"xl-audit report: {summary.get('file', '?')} "
             f"({summary.get('sheets', 0)} sheet(s), "
             f"{summary.get('formulas', 0)} formula(s))", ""]
    for f in findings:
        lines.append(f"{f.severity.upper():5}  {f.check:14}  {f.title}")
        if f.detail:
            lines.append(f"{'':7}{f.detail}")
        for ex in f.examples:
            lines.append(f"{'':11}- {ex}")
    if not findings:
        lines.append("No formula problems found.")
    lines.append("")
    errors = sum(1 for f in findings if f.severity == "error")
    warns = len(findings) - errors
    lines.append(f"verdict: {len(findings)} problem(s) ({errors} error, {warns} warning)"
                 if findings else "verdict: clean")
    return "\n".join(lines)


def main(argv=None) -> int:
    ap = argparse.ArgumentParser(
        prog="xl_audit",
        description="Lint an .xlsx workbook for formula landmines "
                    "(mis-scoped SUMs, broken refs, formula drift, hardcodes).")
    ap.add_argument("file", help="path to the .xlsx workbook")
    ap.add_argument("--json", action="store_true", help="machine-readable report")
    ap.add_argument("--max-examples", type=int, default=5)
    args = ap.parse_args(argv)

    path = Path(args.file)
    if not path.is_file():
        print(f"xl-audit: cannot read {args.file}", file=sys.stderr)
        return 2
    try:
        findings, summary = analyze(path, args.max_examples)
    except Exception as exc:                      # openpyxl raises various types
        print(f"xl-audit: failed to analyse {args.file}: {exc}", file=sys.stderr)
        return 2

    if args.json:
        print(json.dumps({"summary": summary,
                          "findings": [f.to_dict() for f in findings]}, indent=2))
    else:
        print(render(findings, summary))
    return 1 if findings else 0


if __name__ == "__main__":
    sys.exit(main())
