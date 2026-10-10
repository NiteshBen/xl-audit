# xl-audit

Lint an Excel workbook for the formula bugs that make it **quietly wrong** — before you trust a single number in it.

You inherit a spreadsheet that has been edited by four people since 2024. The totals look
plausible. Nobody notices that someone inserted a row last month and the `SUM` didn't grow
with it — so the quarterly report is off by exactly one invoice. That's the most expensive
class of spreadsheet bug: silent, plausible, wrong. xl-audit points a CLI at any `.xlsx`
and finds the landmines in seconds.

## What it checks

| Check | Severity | What it catches |
|---|---|---|
| sum-scope | error | `=SUM(D2:D12)` when the data actually continues at D13 — the "inserted row" bug that silently excludes rows from totals |
| broken-refs | error | formulas referencing `#REF!` / `#NAME?` (deleted sheets, ranges, names) |
| cached-errors | error | cells whose last saved value is an Excel error (`#DIV/0!`, `#N/A`, …) — fires on files saved by Excel/LibreOffice |
| formula-drift | warn | one formula in a column that differs from its neighbours' pattern — the classic copy-paste slip |
| hardcoded | warn | magic numbers inside formulas (`=D2*1.18`) — decimals and 3+ digit integers; benign multipliers like `*2` are left alone |

## How to run

```
pip install openpyxl        # the only dependency
py xl_audit.py examples/quarterly_budget.xlsx
```

```
xl-audit report: quarterly_budget.xlsx (2 sheet(s), 26 formula(s))

ERROR  broken-refs     Budget!G2 references #REF!
       A referenced sheet, range or name was deleted.
           - =#REF!*2
ERROR  sum-scope       Budget!D14 sums D2:D12 but data continues at D13
       A row was probably inserted below the range: it is silently excluded from the total.
           - =SUM(D2:D12)
WARN   formula-drift   1 formula(s) in column F of 'Budget' differ from the column's pattern
       Classic copy-paste slip: one row computes something else than its neighbours.
           - Budget!F8
WARN   hardcoded       12 formula(s) in column E of 'Budget' contain hardcoded constants
       Rates and magic numbers belong in a labelled cell, not inside formulas.
           - Budget!E2
           ...

verdict: 5 problem(s) (2 error, 3 warning)
```

`examples/quarterly_budget.xlsx` is a realistic inherited workbook with all four bug types
planted — regenerate it with `py scripts/make_example.py`.

### Machine-readable output and exit codes

```
py xl_audit.py report.xlsx --json
```

| Exit code | Meaning |
|---|---|
| 0 | no formula problems found |
| 1 | at least one error or warning |
| 2 | file could not be read |

## How the checks stay quiet

False positives are worse than no linter, so the checks are deliberately conservative:
a formula cell below a `SUM` range (subtotals) never counts as "missed data"; horizontal
`SUM` ranges are skipped; and hardcoded constants only fire for decimals or 3+ digit
integers, because `=A2*2` is almost always intentional.

## Limitations

- Single-column vertical `SUM` ranges only (where the inserted-row bug lives).
- Formula-pattern comparison normalises row numbers, not column shifts.
- `.xls` (pre-2007) files are not supported.

## Testing

```
py -m pytest
```

7 tests: each check on generated workbooks, including the false-positive guards
(subtotal below range, benign `*2` multiplier) and the CLI end-to-end with `--json`.

## License

MIT — see [LICENSE](LICENSE).
