"""Clean marketplace files and build the static dashboard data payload."""

from __future__ import annotations

import json
import re
from pathlib import Path

import pandas as pd
from openpyxl import Workbook
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import get_column_letter

ROOT = Path(__file__).resolve().parent
DATA_JS = ROOT / "dashboard_data.js"
REPORT_XLSX = ROOT / "CartNest_Settlement_Report_Sep_2026.xlsx"


def clean_id(value: object) -> str:
    return re.sub(r"[^A-Z0-9]", "", str(value).strip().upper())


def parse_date(value: object) -> pd.Timestamp:
    text = str(value).strip()
    for fmt in ("%Y-%m-%d", "%d-%m-%Y", "%d/%m/%Y", "%m/%d/%Y"):
        parsed = pd.to_datetime(text, format=fmt, errors="coerce")
        if not pd.isna(parsed):
            return parsed
    return pd.NaT


def money(series: pd.Series) -> pd.Series:
    return pd.to_numeric(
        series.astype(str).str.replace("₹", "", regex=False).str.replace(",", "", regex=False),
        errors="coerce",
    ).fillna(0.0)


def load_data() -> tuple[pd.DataFrame, pd.DataFrame, pd.DataFrame, pd.DataFrame]:
    sales = pd.read_csv(ROOT / "sales_records.csv")
    settlements = pd.read_csv(ROOT / "settlement_report.csv")
    bank = pd.read_csv(ROOT / "bank_statement.csv")
    rates = pd.read_csv(ROOT / "fee_rate_card.csv")

    sales["order_item_id"] = sales["order_item_id"].map(clean_id)
    sales["order_id"] = sales["order_id"].map(clean_id)
    sales["dispatch_date"] = sales["dispatch_date"].map(parse_date)
    sales["selling_price"] = money(sales["selling_price"])
    sales["item_status"] = sales["item_status"].astype(str).str.strip().str.title()
    sales["category"] = sales["category"].astype(str).str.strip()

    settlements["order_item_ref"] = settlements["order_item_ref"].map(clean_id)
    settlements["txn_date"] = settlements["txn_date"].map(parse_date)
    settlements["settlement_date"] = settlements["settlement_date"].map(parse_date)
    for column in ["principal", "commission", "fixed_fee", "shipping_fee", "gst_on_fees",
                   "tcs", "tds", "net_amount"]:
        settlements[column] = money(settlements[column])

    bank["value_date"] = bank["value_date"].map(parse_date)
    bank["credit"] = money(bank["credit"])
    bank["debit"] = money(bank["debit"])

    rates["price_above"] = pd.to_numeric(rates["price_above"], errors="coerce").fillna(0)
    rates["price_up_to"] = pd.to_numeric(rates["price_up_to"], errors="coerce")
    rates["rate_pct"] = pd.to_numeric(rates["rate_pct"], errors="coerce")
    rates["rate_inr"] = pd.to_numeric(rates["rate_inr"], errors="coerce")
    return sales, settlements, bank, rates


def rate_for(rates: pd.DataFrame, category: str, price: float, fee_type: str) -> float:
    rows = rates[(rates["fee_type"] == fee_type) &
                 ((rates["category"] == category) | (rates["category"] == "ALL"))]
    rows = rows[(price > rows["price_above"]) & (rows["price_up_to"].isna() | (price <= rows["price_up_to"]))]
    if rows.empty:
        return 0.0
    return float(rows.iloc[0]["rate_pct"] if fee_type == "COMMISSION" else rows.iloc[0]["rate_inr"])


def iso(value: object) -> str:
    return value.strftime("%Y-%m-%d") if isinstance(value, pd.Timestamp) and not pd.isna(value) else ""


def build_payload() -> dict:
    sales, settlements, bank, rates = load_data()
    settlement_summary = (
        settlements.sort_values(["order_item_ref", "settlement_date"])
        .drop_duplicates("order_item_ref", keep="first")
        .rename(columns={"order_item_ref": "order_item_id"})
    )
    joined = sales.merge(
        settlement_summary[
            ["order_item_id", "principal", "commission", "fixed_fee", "shipping_fee",
             "gst_on_fees", "tcs", "tds", "net_amount", "settlement_date", "settlement_id"]
        ],
        on="order_item_id",
        how="left",
        indicator=True,
    )
    joined["missing"] = joined["_merge"].eq("left_only")
    joined["days_to_settle"] = (joined["settlement_date"] - joined["dispatch_date"]).dt.days
    joined["is_timing"] = joined["missing"] & (joined["dispatch_date"] > pd.Timestamp("2026-09-23"))

    commission_rates = joined.apply(
        lambda row: rate_for(rates, row["category"], row["selling_price"], "COMMISSION"), axis=1
    )
    fixed_rates = joined.apply(
        lambda row: rate_for(rates, "ALL", row["selling_price"], "FIXED_CLOSING_FEE"), axis=1
    )
    joined["expected_commission"] = joined["selling_price"] * commission_rates / 100
    joined["expected_fixed_fee"] = fixed_rates
    joined["expected_shipping_fee"] = rate_for(rates, "ALL", 0, "SHIPPING_FEE")
    joined["commission_overcharge"] = (
        joined["commission"].fillna(0) - joined["expected_commission"]
    ).clip(lower=0)
    joined["is_cancelled_exception"] = (
        joined["item_status"].eq("Cancelled") & joined["commission"].fillna(0).gt(0)
    )

    actual_cash = float(
        bank.loc[
            bank["narration"].str.contains("SHOPVERSE", case=False, na=False), "credit"
        ].sum()
    )
    gross_sales = float(sales["selling_price"].sum())
    returned_value = float(sales.loc[sales["item_status"].isin(["Returned", "Cancelled"]), "selling_price"].sum())
    expected_net = float(joined["selling_price"].sum() - joined["commission"].fillna(0).sum())
    expected_net -= float(joined["fixed_fee"].fillna(0).sum() + joined["shipping_fee"].fillna(0).sum())
    expected_net -= float(joined["gst_on_fees"].fillna(0).sum() + joined["tcs"].fillna(0).sum() + joined["tds"].fillna(0).sum())
    expected_net -= returned_value

    missing = joined[joined["missing"] & ~joined["is_timing"]].copy()
    timing = joined[joined["is_timing"]].copy()
    overcharges = joined[joined["commission_overcharge"] > 1].copy()
    cancelled = joined[joined["is_cancelled_exception"]].copy()

    def exception_rows(frame: pd.DataFrame, issue: str, amount_column: str) -> list[dict]:
        rows = []
        for _, row in frame.head(500).iterrows():
            rows.append({
                "id": row["order_item_id"],
                "date": iso(row["dispatch_date"]),
                "category": row["category"],
                "status": row["item_status"],
                "issue": issue,
                "amount": round(float(row[amount_column]), 2),
                "sellingPrice": round(float(row["selling_price"]), 2),
            })
        return rows

    category = (
        joined.groupby("category", dropna=False)
        .agg(sales=("selling_price", "sum"), items=("order_item_id", "count"),
             payout=("net_amount", "sum"), missing=("missing", "sum"))
        .reset_index()
    )
    daily = (
        sales.assign(date=sales["dispatch_date"].map(iso))
        .groupby("date")
        .agg(sales=("selling_price", "sum"), items=("order_item_id", "count"))
        .reset_index()
    )
    payout_by_date = (
        settlements.groupby(settlements["settlement_date"].map(iso))["net_amount"].sum()
        .round(2).to_dict()
    )

    payload = {
        "meta": {"period": "September 2026", "generatedAt": pd.Timestamp.now().isoformat(timespec="seconds")},
        "kpis": {
            "salesItems": int(len(sales)),
            "grossSales": round(gross_sales, 2),
            "settledItems": int(joined["net_amount"].notna().sum()),
            "settledPayout": round(float(joined["net_amount"].fillna(0).sum()), 2),
            "expectedNet": round(expected_net, 2),
            "actualCash": round(actual_cash, 2),
            "unreconciled": round(expected_net - actual_cash, 2),
            "timingItems": int(len(timing)),
            "timingAmount": round(float(timing["selling_price"].sum()), 2),
            "unpaidItems": int(len(missing)),
            "unpaidAmount": round(float(missing["selling_price"].sum()), 2),
            "overchargeItems": int(len(overcharges)),
            "overchargeAmount": round(float(overcharges["commission_overcharge"].sum()), 2),
            "cancelledItems": int(len(cancelled)),
            "cancelledAmount": round(float(cancelled["commission"].fillna(0).sum() + cancelled["fixed_fee"].fillna(0).sum()), 2),
            "returnsValue": round(returned_value, 2),
        },
        "bridge": [
            {"label": "Gross sales", "amount": round(gross_sales, 2), "kind": "positive"},
            {"label": "Commission & fixed fees", "amount": -round(float(joined["commission"].fillna(0).sum() + joined["fixed_fee"].fillna(0).sum()), 2), "kind": "negative"},
            {"label": "Shipping & GST", "amount": -round(float(joined["shipping_fee"].fillna(0).sum() + joined["gst_on_fees"].fillna(0).sum()), 2), "kind": "negative"},
            {"label": "TCS & TDS", "amount": -round(float(joined["tcs"].fillna(0).sum() + joined["tds"].fillna(0).sum()), 2), "kind": "negative"},
            {"label": "Returns / cancelled value", "amount": -round(returned_value, 2), "kind": "negative"},
            {"label": "Expected net payout", "amount": round(expected_net, 2), "kind": "total"},
            {"label": "Cash received in bank", "amount": -round(actual_cash, 2), "kind": "negative"},
            {"label": "Unreconciled difference", "amount": round(expected_net - actual_cash, 2), "kind": "total"},
        ],
        "categories": [
            {"name": str(row["category"]), "sales": round(float(row["sales"]), 2), "items": int(row["items"]),
             "payout": round(float(row["payout"]), 2), "missing": int(row["missing"])}
            for _, row in category.iterrows()
        ],
        "dailySales": [{"date": row["date"], "sales": round(float(row["sales"]), 2), "items": int(row["items"])} for _, row in daily.iterrows()],
        "dailyPayouts": payout_by_date,
        "exceptions": (
            exception_rows(timing, "T+7 due in October", "selling_price")
            + exception_rows(missing, "Unpaid shipped item", "selling_price")
            + exception_rows(overcharges, "Fee overcharge", "commission_overcharge")
            + exception_rows(cancelled, "Cancelled fee deduction", "commission")
        ),
    }
    return payload


def export_excel(payload: dict) -> None:
    """Create a concise client report with explanations beside the numbers."""
    wb = Workbook()
    summary = wb.active
    summary.title = "Client Summary"
    navy = "12243A"
    teal = "19A88E"
    pale = "EEF5F8"
    orange = "FFF2DF"
    red = "FDEDEC"
    white = "FFFFFF"
    thin = Side(style="thin", color="DDE5EC")

    def style_sheet(ws, widths: dict[str, int]) -> None:
        ws.freeze_panes = "A4"
        ws.sheet_view.showGridLines = False
        for col, width in widths.items():
            ws.column_dimensions[col].width = width
        for cell in ws[1]:
            cell.fill = PatternFill("solid", fgColor=navy)
            cell.font = Font(color=white, bold=True, size=14)
        ws.row_dimensions[1].height = 28

    def title(ws, text: str, subtitle: str, end_col: int) -> None:
        ws.merge_cells(start_row=1, start_column=1, end_row=1, end_column=end_col)
        ws["A1"] = text
        ws.merge_cells(start_row=2, start_column=1, end_row=2, end_column=end_col)
        ws["A2"] = subtitle
        ws["A2"].font = Font(color="68798C", italic=True, size=10)
        ws["A2"].alignment = Alignment(wrap_text=True)

    def header(ws, row: int, values: list[str]) -> None:
        for col, value in enumerate(values, 1):
            cell = ws.cell(row, col, value)
            cell.fill = PatternFill("solid", fgColor=teal)
            cell.font = Font(color=white, bold=True)
            cell.alignment = Alignment(wrap_text=True)
            cell.border = Border(bottom=thin)

    k = payload["kpis"]
    title(summary, "CartNest Marketplace Settlement Report", "September 2026 | Simple client view of sales, payout, cash received and follow-up items", 4)
    summary["A4"] = "Executive result"
    summary["A4"].font = Font(bold=True, size=12, color=navy)
    header(summary, 5, ["What we checked", "Result", "What it means", "Priority"])
    rows = [
        ["Gross sales recorded", k["grossSales"], f'{k["salesItems"]:,} items in CartNest sales records', "Reference"],
        ["Expected net payout", k["expectedNet"], "What should be received after fees, taxes and returns", "Reference"],
        ["Cash received from ShopVerse", k["actualCash"], "Credits identified in the bank statement", "Reference"],
        ["T+7 amount due in October", k["timingAmount"], f'{k["timingItems"]:,} recent dispatches are still within the normal payout window', "Monitor"],
        ["Unpaid shipped items", k["unpaidAmount"], f'{k["unpaidItems"]:,} shipped items have no matching settlement record', "Action required"],
        ["Fee overcharges", k["overchargeAmount"], f'{k["overchargeItems"]:,} items are above the rate-card commission calculation', "Action required"],
        ["Unreconciled difference", k["unreconciled"], "Difference between expected payout and bank cash; timing is shown separately", "Review"],
    ]
    for r, values in enumerate(rows, 6):
        for c, value in enumerate(values, 1):
            cell = summary.cell(r, c, value)
            cell.border = Border(bottom=thin)
            cell.alignment = Alignment(wrap_text=True, vertical="top")
            if c == 2:
                cell.number_format = '₹#,##0.00;[Red](₹#,##0.00)'
        summary.cell(r, 4).fill = PatternFill("solid", fgColor=orange if values[3] == "Monitor" else red if values[3] in ("Action required", "Review") else pale)
    summary["A15"] = "How the review was done"
    summary["A15"].font = Font(bold=True, size=12, color=navy)
    guide = [
        ["1", "Clean", "Standardized IDs, dates and currency values so records could be compared consistently."],
        ["2", "Match", "Compared every CartNest item with the ShopVerse settlement report."],
        ["3", "Separate timing", "Items dispatched from 24–30 September are treated as T+7 receivables due in October."],
        ["4", "Check fees", "Compared charged commission with the category and price-band rate card."],
        ["5", "Check cash", "Compared ShopVerse-labelled bank credits with the expected payout."],
    ]
    header(summary, 16, ["Step", "Review step", "Plain-English explanation"])
    for r, values in enumerate(guide, 17):
        for c, value in enumerate(values, 1):
            summary.cell(r, c, value).alignment = Alignment(wrap_text=True, vertical="top")
    style_sheet(summary, {"A": 29, "B": 19, "C": 72, "D": 18})
    summary.auto_filter.ref = "A5:D12"

    exceptions = wb.create_sheet("Exception Report")
    title(exceptions, "Exception Report", "Records that need monitoring, recovery or explanation", 6)
    header(exceptions, 4, ["Item ID", "Issue", "Dispatch date", "Category", "Amount", "Simple action"])
    actions = {
        "T+7 due in October": "Check October bank clearance",
        "Unpaid shipped item": "Raise settlement claim",
        "Fee overcharge": "Request fee correction",
        "Cancelled fee deduction": "Request reversal",
    }
    for r, item in enumerate(payload["exceptions"], 5):
        values = [item["id"], item["issue"], item["date"], item["category"], item["amount"], actions[item["issue"]]]
        for c, value in enumerate(values, 1):
            cell = exceptions.cell(r, c, value)
            cell.border = Border(bottom=thin)
            if c == 5:
                cell.number_format = '₹#,##0.00;[Red](₹#,##0.00)'
    style_sheet(exceptions, {"A": 17, "B": 28, "C": 15, "D": 18, "E": 16, "F": 28})
    exceptions.auto_filter.ref = f"A4:F{len(payload['exceptions']) + 4}"

    categories = wb.create_sheet("Category Summary")
    title(categories, "Category Summary", "Sales mix and settlement coverage by product category", 5)
    header(categories, 4, ["Category", "Items", "Gross sales", "Settled payout", "Missing item count"])
    for r, item in enumerate(payload["categories"], 5):
        values = [item["name"], item["items"], item["sales"], item["payout"], item["missing"]]
        for c, value in enumerate(values, 1):
            cell = categories.cell(r, c, value)
            cell.border = Border(bottom=thin)
            if c in (3, 4):
                cell.number_format = '₹#,##0.00;[Red](₹#,##0.00)'
    style_sheet(categories, {"A": 20, "B": 14, "C": 18, "D": 18, "E": 20})
    categories.auto_filter.ref = f"A4:E{len(payload['categories']) + 4}"

    log = wb.create_sheet("Cleaning Log")
    title(log, "Data Cleaning Log", "Transformations applied before reconciliation", 4)
    header(log, 4, ["Field", "Source", "Issue found", "Resolution"])
    cleaning_rows = [
        ["Item ID", "Sales + settlement files", "Mixed case, spaces and hyphens", "TRIM, UPPER and remove separators"],
        ["Date", "Sales, settlement and bank files", "Multiple date formats", "Converted to ISO date format"],
        ["Currency", "Settlement and bank files", "Blank values and currency symbols", "Converted to numeric rupee values"],
        ["Duplicate settlement references", "Settlement file", "Same item can occur more than once", "First dated record retained; duplicates excluded from summary"],
    ]
    for r, values in enumerate(cleaning_rows, 5):
        for c, value in enumerate(values, 1):
            log.cell(r, c, value).alignment = Alignment(wrap_text=True, vertical="top")
    style_sheet(log, {"A": 25, "B": 28, "C": 42, "D": 65})

    for ws in wb.worksheets:
        for row in ws.iter_rows():
            for cell in row:
                if cell.row > 2 and cell.alignment == Alignment():
                    cell.alignment = Alignment(vertical="top")
        ws.sheet_properties.pageSetUpPr.fitToPage = True
        ws.page_setup.fitToWidth = 1
        ws.page_setup.fitToHeight = 0
    wb.save(REPORT_XLSX)


if __name__ == "__main__":
    data = build_payload()
    DATA_JS.write_text(
        "window.RECONCILIATION_DATA = " + json.dumps(data, separators=(",", ":")) + ";\n",
        encoding="utf-8",
    )
    export_excel(data)
    print(f"Generated {DATA_JS.name} and {REPORT_XLSX.name} with {len(data['exceptions'])} exception rows.")
