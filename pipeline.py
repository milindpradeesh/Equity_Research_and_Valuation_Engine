from __future__ import annotations

import logging
import os
import sys
import time
from typing import Any

import pandas as pd
import requests
import xlwings as xw

# ---------------------------------------------------------------------------
# Logging
# ---------------------------------------------------------------------------
logging.basicConfig(
    level=logging.INFO,
    format="%(asctime)s  %(levelname)-8s  %(message)s",
    datefmt="%H:%M:%S",
    stream=sys.stdout,
)
log = logging.getLogger(__name__)

# ---------------------------------------------------------------------------
# API key
# ---------------------------------------------------------------------------
try:
    from api_key_template import API_KEY  # type: ignore[import]
except ImportError as exc:
    raise SystemExit(
        "Missing api_key.py.  Create it with:  API_KEY = 'your_fmp_key'"
    ) from exc

# ---------------------------------------------------------------------------
# Constants
# ---------------------------------------------------------------------------
BASE_URL = "https://financialmodelingprep.com/stable"

# Metadata columns emitted by FMP that carry no numeric value.
_STATEMENT_METADATA: frozenset[str] = frozenset(
    {
        "symbol", "reportedcurrency", "cik",
        "fillingdate", "accepteddate",
        "calendaryear", "fiscalyear", "period",
        "link", "finallink",
    }
)

REQUIRED_METRICS: dict[str, list[str]] = {
    "income": [
        "revenue",
        "costOfRevenue",
        "grossProfit",
        "researchAndDevelopmentExpenses",
        "sellingGeneralAndAdministrativeExpenses",
        "operatingExpenses",
        "operatingIncome",
        "ebit",
        "interestExpense",
        "incomeBeforeTax",
        "incomeTaxExpense",
        "netIncome",
        "eps",
        "epsDiluted",
        "weightedAverageShsOut",
        "weightedAverageShsOutDil",
        "depreciationAndAmortization",
    ],
    "cashflow": [
        "netIncome",
        "depreciationAndAmortization",
        "changeInWorkingCapital",
        "accountsReceivables",
        "inventory",
        "accountsPayables",
        "netCashProvidedByOperatingActivities",
        "operatingCashFlow",
        "capitalExpenditure",
        "freeCashFlow",
    ],
    "balance": [
        "cashAndCashEquivalents",
        "shortTermInvestments",
        "cashAndShortTermInvestments",
        "netReceivables",
        "accountsReceivables",
        "inventory",
        "totalCurrentAssets",
        "propertyPlantEquipmentNet",
        "totalAssets",
        "accountPayables",
        "accountsPayables",
        "totalCurrentLiabilities",
        "shortTermDebt",
        "longTermDebt",
        "totalLiabilities",
        "retainedEarnings",
        "totalStockholdersEquity",
        "totalEquity",
        "totalDebt",
        "netDebt",
        "totalInvestments",
    ],
    "quote": [
        "symbol",
        "name",
        "price",
        "marketCap",
        "priceAvg50",
        "priceAvg200",
        "exchange",
        "sharesOutstanding",
        "timestamp",
    ],
}


# ---------------------------------------------------------------------------
# API layer
# ---------------------------------------------------------------------------
def fetch_endpoint(
    endpoint: str,
    ticker: str,
    *,
    retries: int = 3,
    backoff: float = 2.0,
) -> pd.DataFrame:
    """
    GET ``BASE_URL/<endpoint>?symbol=<ticker>&apikey=<key>``.

    Raises
    ------
    RuntimeError
        If all retry attempts fail or the response payload is not a list.
    """
    url = f"{BASE_URL}/{endpoint}"
    params = {"symbol": ticker, "apikey": API_KEY}

    for attempt in range(1, retries + 1):
        try:
            resp = requests.get(url, params=params, timeout=20)
            resp.raise_for_status()

            payload: Any = resp.json()

            # FMP returns an error dict on bad tickers / exhausted plans.
            if isinstance(payload, dict):
                raise ValueError(
                    f"Unexpected dict response from '{endpoint}': {payload}"
                )

            if not isinstance(payload, list):
                raise ValueError(
                    f"Expected list from '{endpoint}', got {type(payload).__name__}"
                )

            if not payload:
                log.warning("Empty payload returned for endpoint '%s'.", endpoint)
                return pd.DataFrame()

            return pd.DataFrame(payload)

        except Exception as exc:  # noqa: BLE001
            if attempt < retries:
                wait = backoff * attempt
                log.warning(
                    "Attempt %d/%d failed for '%s': %s  — retrying in %.0fs",
                    attempt, retries, endpoint, exc, wait,
                )
                time.sleep(wait)
            else:
                raise RuntimeError(
                    f"All {retries} attempts failed for '{endpoint}': {exc}"
                ) from exc

    # Unreachable, but satisfies type-checkers.
    raise RuntimeError("fetch_endpoint: unexpected exit")  # pragma: no cover


# ---------------------------------------------------------------------------
# Statement processing
# ---------------------------------------------------------------------------
def process_statement(raw: pd.DataFrame) -> pd.DataFrame:
    """
    Pivot a raw FMP statement DataFrame from row-per-year into
    column-per-year, strip metadata rows, and coerce values to numeric.

    The "date" column must exist; if it is absent the frame is returned
    unchanged with a warning.
    """
    if raw.empty:
        return raw

    if "date" not in raw.columns:
        log.warning(
            "process_statement: 'date' column not found — returning raw frame. "
            "Columns present: %s", list(raw.columns)
        )
        return raw

    df = raw.copy()
    df["date"] = pd.to_datetime(df["date"]).dt.year
    df = df.sort_values("date", ascending=False)
    df = df.set_index("date").T

    # Drop metadata rows (case-insensitive match).
    df = df[~df.index.str.lower().isin(_STATEMENT_METADATA)]

    # Coerce to numeric — keeps genuine zeros.
    df = df.apply(pd.to_numeric, errors="coerce")

    # Drop rows that are entirely NaN (i.e. truly absent data).
    # Do NOT drop rows that are all-zero — zeros are valid financial data.
    df = df.dropna(how="all")

    df = df.reset_index()
    df = df.rename(columns={"index": "Line Item"})

    return df


def filter_metrics(df: pd.DataFrame, required: list[str]) -> pd.DataFrame:
    """Keep only rows whose 'Line Item' appears in *required*; deduplicate."""
    if df.empty:
        return df
    mask = df["Line Item"].isin(required)
    return df[mask].drop_duplicates(subset=["Line Item"], keep="first").copy()


# ---------------------------------------------------------------------------
# Enrichment helpers
# ---------------------------------------------------------------------------
def add_income_fallbacks(df: pd.DataFrame) -> pd.DataFrame:
    """
    Provide alias rows for metrics that FMP may name differently:

    * ``epsDiluted``  ← ``eps``  (when diluted figure is absent)
    * ``operatingIncome`` ← ``ebit``  (when operatingIncome is absent)
    """
    items = set(df["Line Item"])

    aliases = [
        ("epsDiluted",      "eps"),
        ("operatingIncome", "ebit"),
    ]
    rows_to_add: list[pd.DataFrame] = []

    for target, source in aliases:
        if target not in items and source in items:
            alias = df[df["Line Item"] == source].copy()
            alias["Line Item"] = target
            rows_to_add.append(alias)
            log.info("Income fallback: '%s' aliased from '%s'.", target, source)

    if rows_to_add:
        df = pd.concat([df, *rows_to_add], ignore_index=True)

    return df


def resolve_operating_cash_flow(df: pd.DataFrame) -> pd.DataFrame:
    """
    FMP alternates between ``operatingCashFlow`` and
    ``netCashProvidedByOperatingActivities``.  Ensure both labels are
    present so downstream FCF computation is never silently blocked.
    """
    items = set(df["Line Item"])
    has_ocf  = "operatingCashFlow" in items
    has_ncpo = "netCashProvidedByOperatingActivities" in items

    if has_ncpo and not has_ocf:
        row = df[df["Line Item"] == "netCashProvidedByOperatingActivities"].copy()
        row["Line Item"] = "operatingCashFlow"
        df = pd.concat([df, row], ignore_index=True)
        log.info("CF alias: 'operatingCashFlow' ← 'netCashProvidedByOperatingActivities'.")

    elif has_ocf and not has_ncpo:
        row = df[df["Line Item"] == "operatingCashFlow"].copy()
        row["Line Item"] = "netCashProvidedByOperatingActivities"
        df = pd.concat([df, row], ignore_index=True)
        log.info("CF alias: 'netCashProvidedByOperatingActivities' ← 'operatingCashFlow'.")

    return df


def compute_derived_metrics(
    cashflow_df: pd.DataFrame,
    balance_df: pd.DataFrame,
) -> tuple[pd.DataFrame, pd.DataFrame]:
    """
    Compute ``freeCashFlow`` and ``netDebt`` when absent.

    freeCashFlow = operatingCashFlow − |capitalExpenditure|
    netDebt      = totalDebt − cashAndShortTermInvestments

    The absolute-value normalisation on capex handles FMP's inconsistent
    sign convention.  OCF sign is intentionally preserved — a negative OCF
    (cash-burning period) will correctly produce a negative FCF.
    """
    cf_items = set(cashflow_df["Line Item"])

    if "freeCashFlow" not in cf_items:
        if {"operatingCashFlow", "capitalExpenditure"}.issubset(cf_items):
            ocf   = cashflow_df[cashflow_df["Line Item"] == "operatingCashFlow"].iloc[:, 1:]
            capex = cashflow_df[cashflow_df["Line Item"] == "capitalExpenditure"].iloc[:, 1:]
            fcf_values = ocf.values - capex.abs().values
            fcf_row = pd.DataFrame(
                [["freeCashFlow", *fcf_values[0]]],
                columns=cashflow_df.columns,
            )
            cashflow_df = pd.concat([cashflow_df, fcf_row], ignore_index=True)
            log.info("Derived: 'freeCashFlow' computed from OCF − |capex|.")
        else:
            log.warning(
                "Cannot compute 'freeCashFlow': missing %s.",
                {"operatingCashFlow", "capitalExpenditure"} - cf_items,
            )

    b_items = set(balance_df["Line Item"])

    if "netDebt" not in b_items:
        if {"totalDebt", "cashAndShortTermInvestments"}.issubset(b_items):
            debt = balance_df[balance_df["Line Item"] == "totalDebt"].iloc[:, 1:]
            cash = balance_df[balance_df["Line Item"] == "cashAndShortTermInvestments"].iloc[:, 1:]
            nd_values = debt.values - cash.values
            nd_row = pd.DataFrame(
                [["netDebt", *nd_values[0]]],
                columns=balance_df.columns,
            )
            balance_df = pd.concat([balance_df, nd_row], ignore_index=True)
            log.info("Derived: 'netDebt' computed from totalDebt − cash.")
        else:
            log.warning(
                "Cannot compute 'netDebt': missing %s.",
                {"totalDebt", "cashAndShortTermInvestments"} - b_items,
            )

    return cashflow_df, balance_df


def add_dividends_if_present(
    target_df: pd.DataFrame,
    source_processed_df: pd.DataFrame,
) -> pd.DataFrame:
    """
    If any dividend-related line item exists in *source_processed_df*,
    append it to *target_df* normalised as ``dividendsPaid``.

    Parameters
    ----------
    target_df:
        The already-filtered cashflow DataFrame to append to.
    source_processed_df:
        The fully processed (but pre-filter) cashflow DataFrame used to
        search for dividend rows.
    """
    if target_df.empty or source_processed_df.empty:
        return target_df

    dividend_items = [
        item for item in source_processed_df["Line Item"]
        if "dividend" in str(item).lower()
    ]

    if not dividend_items:
        return target_df

    source_label = dividend_items[0]
    div_row = source_processed_df[source_processed_df["Line Item"] == source_label].copy()
    div_row["Line Item"] = "dividendsPaid"

    # Remove stale row before appending.
    target_df = target_df[target_df["Line Item"] != "dividendsPaid"].copy()
    target_df = pd.concat([target_df, div_row], ignore_index=True)
    log.info("Optional: 'dividendsPaid' appended from '%s'.", source_label)

    return target_df


# ---------------------------------------------------------------------------
# Quote processing
# ---------------------------------------------------------------------------
def process_quote(raw: pd.DataFrame) -> pd.DataFrame:
    """
    Return a two-column (Metric / Value) DataFrame for the first quote row,
    restricted to REQUIRED_METRICS["quote"].
    """
    if raw.empty:
        log.warning("process_quote: empty input.")
        return raw

    q = raw.iloc[0].copy()

    if "timestamp" in q.index:
        try:
            q["timestamp"] = pd.to_datetime(q["timestamp"], unit="s")
        except Exception:  # noqa: BLE001
            pass  # Leave as-is if conversion fails.

    q_df = q.to_frame().reset_index()
    q_df.columns = ["Metric", "Value"]
    q_df = q_df[q_df["Metric"].isin(REQUIRED_METRICS["quote"])].copy()

    return q_df


# ---------------------------------------------------------------------------
# Excel I/O
# ---------------------------------------------------------------------------
def write_sheet(wb: xw.Book, sheet_name: str, df: pd.DataFrame) -> None:
    """
    Write *df* to *sheet_name* in *wb*, clearing existing content first.
    Creates the sheet if it does not already exist.

    Raises
    ------
    RuntimeError
        Propagates any xlwings exception with added context.
    """
    try:
        try:
            sht = wb.sheets[sheet_name]
            sht.clear()
        except Exception:  # noqa: BLE001
            sht = wb.sheets.add(sheet_name)

        sht.range("A1").options(index=False).value = df
        log.info("Sheet '%s' written (%d rows).", sheet_name, len(df))

    except Exception as exc:
        raise RuntimeError(
            f"Failed to write sheet '{sheet_name}': {exc}"
        ) from exc


# ---------------------------------------------------------------------------
# Entry point
# ---------------------------------------------------------------------------
def main() -> None:
    if len(sys.argv) < 3:
        raise SystemExit(
            "Usage: python pipeline.py <TICKER> <EXCEL_PATH>"
        )

    ticker     = sys.argv[1].strip().upper()
    excel_path = sys.argv[2].strip()

    if not os.path.exists(excel_path):
        raise SystemExit(f"Excel file not found: {excel_path}")

    log.info("Pipeline started  |  ticker=%s  |  workbook=%s", ticker, excel_path)

    # -----------------------------------------------------------------------
    # Fetch
    # -----------------------------------------------------------------------
    log.info("Fetching financial statements from FMP…")
    income_raw   = fetch_endpoint("income-statement",      ticker)
    balance_raw  = fetch_endpoint("balance-sheet-statement", ticker)
    cashflow_raw = fetch_endpoint("cash-flow-statement",   ticker)
    quote_raw    = fetch_endpoint("quote",                 ticker)

    # -----------------------------------------------------------------------
    # Process
    # -----------------------------------------------------------------------
    log.info("Processing statements…")

    income_clean = filter_metrics(
        process_statement(income_raw),
        REQUIRED_METRICS["income"],
    )
    income_clean = add_income_fallbacks(income_clean)

    balance_clean = filter_metrics(
        process_statement(balance_raw),
        REQUIRED_METRICS["balance"],
    )

    # Process cashflow once; reuse the full processed frame for dividend
    # extraction so we never re-process the raw data.
    cashflow_processed = process_statement(cashflow_raw)
    cashflow_clean     = filter_metrics(cashflow_processed, REQUIRED_METRICS["cashflow"])
    cashflow_clean     = resolve_operating_cash_flow(cashflow_clean)
    cashflow_clean     = add_dividends_if_present(cashflow_clean, cashflow_processed)

    cashflow_clean, balance_clean = compute_derived_metrics(cashflow_clean, balance_clean)

    quote_clean = process_quote(quote_raw)

    # -----------------------------------------------------------------------
    # Write
    # -----------------------------------------------------------------------
    log.info("Writing to workbook…")
    wb = xw.Book(excel_path)

    try:
        write_sheet(wb, "Raw_Income",   income_clean)
        write_sheet(wb, "Raw_Balance",  balance_clean)
        write_sheet(wb, "Raw_CashFlow", cashflow_clean)
        write_sheet(wb, "Raw_Quote",    quote_clean)
        wb.save()
    except Exception:
        wb.close()
        raise

    log.info("=== SUCCESS: pipeline complete ===")


if __name__ == "__main__":
    main()