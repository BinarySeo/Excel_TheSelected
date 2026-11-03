"""
verify_dashboard.py
-------------------
Independent check of SelectedProject.xlsx.

1. Data checks — the cleaned "Data" sheet still matches the Korean "Original Data"
   row for row: dates, quantities, prices, sales, the KRW → USD conversion, and that
   every Korean value was translated to exactly one English value.
2. KPI checks — recalculates every number shown on the dashboard from the raw rows
   with pandas and compares it with the pivot tables' saved values.

Usage:
    pip install -r requirements.txt
    python verify_dashboard.py                # exits 1 if any check fails
    python verify_dashboard.py --kpis kpis.csv  # also save the recalculated KPIs

The workbook is opened read-only and never modified (re-saving it from Python would
strip the dashboard's slicers).
"""

from __future__ import annotations

import argparse
import sys
import warnings
from dataclasses import dataclass
from pathlib import Path

import openpyxl
import pandas as pd

WORKBOOK = Path(__file__).resolve().parent / "SelectedProject.xlsx"
KRW_TO_USD = 0.00078          # rate used by the Price($) / Sales($) formulas
MONEY_TOL = 0.01              # cents
RATE_TOL = 0.0005             # 0.05 percentage points

# Original (Korean) column → cleaned (English) column, for row-by-row comparison
SAME_VALUE_COLUMNS = {"날짜": "Date", "수량": "Quantity", "만족도": "Rating"}
# Prices: the cleaned sheet keeps the ordered unit price (단가) even for cancelled orders,
# whose confirmed price (확정단가) is 0. Sales use the confirmed amount (확정매출).
MONEY_COLUMNS = {"단가": "Price(Won)", "확정매출": "Sales(Won)"}
# Korean category columns that were translated; each Korean value must map to one English value
TRANSLATED_COLUMNS = {
    "구분": "Gender",
    "판매채널": "Platform",
    "유입경로": "Ads",
    "재구매": "Return Customer",
    "주문취소": "Canceled?",
}
MONTHS = ["Jan", "Feb", "Mar", "Apr", "May", "Jun", "Jul", "Aug", "Sep", "Oct", "Nov", "Dec"]


# ---------------------------------------------------------------------------
# Loading
# ---------------------------------------------------------------------------

def read_sheet(wb, name: str) -> pd.DataFrame:
    """Read a sheet into a DataFrame, dropping the empty rows the Excel table extends over."""
    rows = list(wb[name].iter_rows(values_only=True))
    header = [h for h in rows[0] if h is not None]
    df = pd.DataFrame([r[: len(header)] for r in rows[1:]], columns=header)
    return df.dropna(how="all").reset_index(drop=True)


def load_workbook(path: Path = WORKBOOK) -> tuple[pd.DataFrame, pd.DataFrame, list[list]]:
    """Return (original data, cleaned data, pivot sheet grid). Uses Excel's saved values."""
    with warnings.catch_warnings():
        # openpyxl warns that it can't read slicers; we don't need them. Sheets are read
        # lazily, so the whole read stays inside this block.
        warnings.simplefilter("ignore")
        wb = openpyxl.load_workbook(path, read_only=True, data_only=True)
        original = read_sheet(wb, "Original Data")
        data = read_sheet(wb, "Data")
        pivot = [list(r) for r in wb["Pivot table"].iter_rows(values_only=True)]
        wb.close()
    return original, data, pivot


# ---------------------------------------------------------------------------
# Pivot sheet parsing
# ---------------------------------------------------------------------------

def find(grid: list[list], text: str) -> tuple[int, int]:
    """Row/column (0-based) of the first cell whose text starts with `text`."""
    for r, row in enumerate(grid):
        for c, value in enumerate(row):
            if isinstance(value, str) and value.strip().startswith(text):
                return r, c
    raise ValueError(f"'{text}' not found on the Pivot table sheet")


def pivot_block(grid: list[list], title: str, value_cols: int = 1) -> dict:
    """Read a pivot below `title`: label → value(s), from 'Row Labels' down to 'Grand Total'."""
    r, c = find(grid, title)
    while grid[r][c] != "Row Labels":
        r += 1
    out = {}
    for row in grid[r + 1:]:
        label = row[c]
        if label is None:
            break
        values = row[c + 1: c + 1 + value_cols]
        out[str(label).strip()] = values[0] if value_cols == 1 else tuple(values)
        if label == "Grand Total":
            break
    return out


def pivot_monthly(grid: list[list]) -> dict:
    """
    Monthly pivot → {(year, month): (quantity, sales $)} plus year and grand totals.
    Read row by row because month names repeat under each year.
    """
    r, c = find(grid, "1.Monthly Total sales")
    while grid[r][c] != "Row Labels":
        r += 1
    out, year = {}, None
    for row in grid[r + 1:]:
        label, values = row[c], tuple(row[c + 1: c + 3])
        if label is None:
            break
        label = str(label).strip()
        if label.isdigit():
            year = int(label)
            out[(year, "total")] = values
        elif label in MONTHS:
            out[(year, label)] = values
        elif label == "Grand Total":
            out[("all", "total")] = values
            break
    return out


def pivot_after(grid: list[list], title: str) -> float:
    """First number below a label (used for the average rating)."""
    r, c = find(grid, title)
    for row in grid[r + 1:]:
        if isinstance(row[c], (int, float)):
            return row[c]
    raise ValueError(f"No value under '{title}'")


# ---------------------------------------------------------------------------
# Checks
# ---------------------------------------------------------------------------

@dataclass
class Check:
    area: str
    name: str
    expected: object
    actual: object
    passed: bool
    blocking: bool = True   # non-blocking checks are reported as warnings only


def close(a, b, tol) -> bool:
    """True when both values exist and differ by at most tol."""
    return a is not None and b is not None and abs(float(a) - float(b)) <= tol


def data_checks(original: pd.DataFrame, data: pd.DataFrame) -> list[Check]:
    checks = [Check("data", "row count matches original", len(original), len(data),
                    len(original) == len(data))]
    if len(original) != len(data):
        return checks

    for ko, en in SAME_VALUE_COLUMNS.items():
        a, b = original[ko], data[en]
        if ko == "날짜":
            a, b = pd.to_datetime(a), pd.to_datetime(b)
        diff = int((a.fillna(-1) != b.fillna(-1)).sum())
        checks.append(Check("data", f"{en} matches original", 0, diff, diff == 0))

    for ko, en in MONEY_COLUMNS.items():
        diff = int(((original[ko] - data[en]).abs() > MONEY_TOL).sum())
        checks.append(Check("data", f"{en} matches original {ko}", 0, diff, diff == 0))

    for won, usd in [("Price(Won)", "Price($)"), ("Sales(Won)", "Sales($)")]:
        diff = int(((data[won] * KRW_TO_USD - data[usd]).abs() > MONEY_TOL).sum())
        checks.append(Check("data", f"{usd} = {won} × {KRW_TO_USD}", 0, diff, diff == 0))

    for ko, en in TRANSLATED_COLUMNS.items():
        per_value = pd.DataFrame({"ko": original[ko].fillna("(blank)"), "en": data[en]})
        ambiguous = int((per_value.groupby("ko")["en"].nunique() > 1).sum())
        checks.append(Check("data", f"{en}: each Korean value → one English value", 0, ambiguous, ambiguous == 0))

    canceled = data["Canceled?"] == "Yes"
    bad = int((data.loc[canceled, "Sales(Won)"] != 0).sum() + data.loc[canceled, "Rating"].notna().sum())
    checks.append(Check("data", "cancelled orders have no sales and no rating", 0, bad, bad == 0))

    kept = data[~canceled]
    bad = int(((kept["Quantity"] * kept["Price(Won)"] - kept["Sales(Won)"]).abs() > MONEY_TOL).sum())
    checks.append(Check("data", "sales = quantity × price for completed orders", 0, bad, bad == 0))

    bad = int((~data["Rating"].dropna().between(1, 5)).sum())
    checks.append(Check("data", "ratings between 1 and 5", 0, bad, bad == 0))

    # Excel pivots ignore case, so spelling variants are merged on the dashboard
    # but would be split by any case-sensitive tool (SQL, pandas, Power BI).
    for column in ["Category", "Sub-Category", "Platform", "Ads", "Province"]:
        values = data[column].dropna().unique()
        variants = len(values) - len({str(v).strip().lower() for v in values})
        checks.append(Check("data", f"{column}: consistent spelling/case", 0, variants,
                            variants == 0, blocking=False))
    return checks


def recalculate_kpis(data: pd.DataFrame) -> dict:
    """Every dashboard number, recalculated from the raw rows."""
    d = data.copy()
    d["Date"] = pd.to_datetime(d["Date"])
    n = len(d)
    monthly = d.groupby([d["Date"].dt.year, d["Date"].dt.month])[["Quantity", "Sales($)"]].sum()
    return {
        "total_sales": d["Sales($)"].sum(),
        "total_units": d["Quantity"].sum(),
        "units_excl_cancelled": d.loc[d["Canceled?"] == "No", "Quantity"].sum(),
        "monthly": {(y, MONTHS[m - 1]): (q, s) for (y, m), (q, s) in monthly.iterrows()},
        "yearly": d.groupby(d["Date"].dt.year)[["Quantity", "Sales($)"]].sum(),
        "platform_share": d.groupby("Platform")["Sales($)"].sum() / d["Sales($)"].sum(),
        "returning_rate": (d["Return Customer"] == "Yes").sum() / n,
        "cancel_rate": (d["Canceled?"] == "Yes").sum() / n,
        "avg_rating": d["Rating"].mean(),
        "sales_by_ads": d.groupby("Ads")["Sales($)"].sum(),
        # case-insensitive, matching how the Excel pivot groups "SweatPants" / "Sweatpants"
        "sales_by_subcategory": d.groupby(d["Sub-Category"].str.lower())["Sales($)"].sum(),
    }


def kpi_checks(k: dict, grid: list[list]) -> list[Check]:
    checks = []

    def add(name, expected, actual, tol):
        # a category missing from the data shows up as a failure, not a crash
        shown = None if actual is None else round(float(actual), 4)
        checks.append(Check("kpi", name, expected, shown, close(expected, actual, tol)))

    monthly = pivot_monthly(grid)
    add("total sales ($)", monthly[("all", "total")][1], k["total_sales"], MONEY_TOL)
    add("total units", monthly[("all", "total")][0], k["total_units"], 0)
    for year, row in k["yearly"].iterrows():
        add(f"{year} sales ($)", monthly[(year, "total")][1], row["Sales($)"], MONEY_TOL)
    bad = [key for key, (q, s) in k["monthly"].items()
           if key not in monthly or not close(monthly[key][0], q, 0) or not close(monthly[key][1], s, MONEY_TOL)]
    checks.append(Check("kpi", "all 17 monthly totals (units and $)", 0, len(bad), not bad))

    for platform, share in pivot_block(grid, "2.Sales per platform").items():
        if platform != "Grand Total":
            add(f"platform share: {platform}", share, k["platform_share"].get(platform), RATE_TOL)

    add("returning-customer rate", pivot_block(grid, "Return Rate")["Yes"], k["returning_rate"], RATE_TOL)
    add("cancellation rate", pivot_block(grid, "Cancel Rate")["Yes"], k["cancel_rate"], RATE_TOL)
    add("average rating", pivot_after(grid, "Average of Rating"), k["avg_rating"], 0.001)

    for source, sales in pivot_block(grid, "4.Ads").items():
        if source != "Grand Total":
            add(f"sales by ad source: {source}", sales, k["sales_by_ads"].get(source), MONEY_TOL)
    for sub, sales in pivot_block(grid, "5.Top Selling Product").items():
        if sub != "Grand Total":
            add(f"sales by sub-category: {sub}", sales, k["sales_by_subcategory"].get(sub.lower()), MONEY_TOL)
    return checks


def run(path: Path = WORKBOOK) -> tuple[list[Check], dict]:
    original, data, grid = load_workbook(path)
    kpis = recalculate_kpis(data)
    return data_checks(original, data) + kpi_checks(kpis, grid), kpis


# ---------------------------------------------------------------------------
# CLI
# ---------------------------------------------------------------------------

def main(argv=None) -> int:
    parser = argparse.ArgumentParser(description=__doc__.split("\n\n")[0])
    parser.add_argument("--workbook", type=Path, default=WORKBOOK)
    parser.add_argument("--kpis", type=Path, help="write recalculated headline KPIs to this CSV")
    args = parser.parse_args(argv)

    checks, kpis = run(args.workbook)
    width = max(len(c.name) for c in checks)
    for c in checks:
        status = "PASS" if c.passed else ("FAIL" if c.blocking else "WARN")
        print(f"{status}  [{c.area}] {c.name:<{width}}  expected {c.expected}, got {c.actual}")

    cancelled_units = kpis["total_units"] - kpis["units_excl_cancelled"]
    print(f"\nNote: the dashboard's {kpis['total_units']:,.0f} units include {cancelled_units:,.0f} units "
          f"from cancelled orders ({kpis['units_excl_cancelled']:,.0f} completed).")

    if args.kpis:
        headline = {
            "total_sales_usd": kpis["total_sales"],
            "total_units_incl_cancelled": kpis["total_units"],
            "units_completed_orders": kpis["units_excl_cancelled"],
            "returning_customer_rate": kpis["returning_rate"],
            "cancellation_rate": kpis["cancel_rate"],
            "average_rating": kpis["avg_rating"],
            **{f"platform_share_{p}": v for p, v in kpis["platform_share"].items()},
        }
        pd.Series(headline, name="value").round(4).to_csv(args.kpis, index_label="kpi")
        print(f"Saved {args.kpis}")

    failed = [c for c in checks if not c.passed and c.blocking]
    warned = [c for c in checks if not c.passed and not c.blocking]
    print(f"\n{len(checks) - len(failed) - len(warned)}/{len(checks)} checks passed, "
          f"{len(warned)} warning(s), {len(failed)} failure(s).")
    return 1 if failed else 0


if __name__ == "__main__":
    sys.exit(main())
