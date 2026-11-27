"""
Deterministic Excel/CSV ingestion and spending analytics engine.

Handles messy real-world spreadsheets (title rows, varied column names,
separate Debit/Credit columns, currency symbols) and computes exact numbers
for the Ollama spending agent.
"""

from __future__ import annotations

import math
import re
from dataclasses import dataclass, field
from pathlib import Path
from typing import Any

import openpyxl
import pandas as pd


DATE_KEYWORDS = (
    "transaction date",
    "txn date",
    "posting date",
    "value date",
    "date",
    "dt",
    "timestamp",
    "time",
    "month",
)
AMOUNT_KEYWORDS = (
    "amount",
    "spent",
    "spend",
    "expense",
    "cost",
    "price",
    "total",
    "value",
    "amt",
)
DEBIT_KEYWORDS = ("debit", "withdrawal", "dr", "paid", "outflow", "expense")
CREDIT_KEYWORDS = ("credit", "deposit", "cr", "received", "inflow", "income")
CATEGORY_KEYWORDS = (
    "category",
    "cat",
    "expense type",
    "type",
    "group",
    "tag",
    "bucket",
    "head",
    "classification",
)
DESC_KEYWORDS = (
    "description",
    "narration",
    "particulars",
    "merchant",
    "payee",
    "vendor",
    "details",
    "item",
    "title",
    "memo",
    "notes",
    "remark",
    "remarks",
)
TYPE_KEYWORDS = ("type", "dr/cr", "cr/dr", "transaction type", "txn type", "kind")
PAYMENT_KEYWORDS = ("payment", "mode", "method", "account", "card", "upi", "wallet", "bank")

AUTO_CATEGORY_RULES: list[tuple[str, tuple[str, ...]]] = [
    ("Housing & Rent", ("rent", "mortgage", "lease", "maintenance", "hoa", "property", "society")),
    ("Utilities & Bills", ("electric", "electricity", "water", "gas", "internet", "wifi", "broadband", "phone", "mobile", "recharge", "bill", "utility")),
    ("Groceries", ("grocery", "groceries", "supermarket", "mart", "blinkit", "zepto", "instamart", "bigbasket", "walmart", "costco", "aldi", "trader", "whole foods", "veg", "milk")),
    ("Food & Dining", ("restaurant", "cafe", "coffee", "starbucks", "mcdonald", "kfc", "pizza", "burger", "swiggy", "zomato", "uber eats", "doordash", "dining", "lunch", "dinner", "breakfast", "bakery")),
    ("Transport & Fuel", ("uber", "ola", "lyft", "taxi", "cab", "fuel", "petrol", "diesel", "gas station", "metro", "train", "bus", "parking", "toll", "auto")),
    ("Travel & Vacation", ("flight", "airline", "hotel", "airbnb", "booking", "makemytrip", "trip", "vacation", "resort", "luggage")),
    ("Shopping & Electronics", ("amazon", "flipkart", "myntra", "target", "best buy", "apple", "electronics", "clothing", "shoes", "mall", "shopping", "gadget", "laptop", "phone")),
    ("Health & Medical", ("pharmacy", "medical", "hospital", "clinic", "doctor", "dental", "medicine", "health", "gym", "fitness")),
    ("Entertainment & Subscriptions", ("netflix", "spotify", "prime", "youtube", "disney", "hotstar", "movie", "cinema", "game", "steam", "subscription", "theatre")),
    ("Education", ("course", "udemy", "coursera", "book", "tuition", "school", "college", "exam")),
    ("Insurance & Finance", ("insurance", "emi", "loan", "interest", "tax", "fee", "bank charge", "sip", "mutual fund")),
]


def _normalize_col(name: Any) -> str:
    return re.sub(r"\s+", " ", str(name or "").strip().lower())


def _detect_currency_symbol(series_values: list[Any]) -> str:
    joined = " ".join(str(v) for v in series_values[:300] if v is not None)
    for sym in ("₹", "$", "€", "£", "¥", "AED", "SGD", "INR", "USD"):
        if sym in joined:
            if sym == "INR":
                return "₹"
            if sym == "USD":
                return "$"
            return sym
    return "$"


def _parse_number(val: Any) -> float | None:
    if val is None or (isinstance(val, float) and math.isnan(val)):
        return None
    if isinstance(val, (int, float)):
        return float(val)
    text = str(val).strip()
    if not text or text in {"-", "--", "N/A", "na", "null", "None"}:
        return None

    is_negative = False
    if text.startswith("(") and text.endswith(")"):
        is_negative = True
        text = text[1:-1].strip()
    if text.lower().endswith("dr"):
        text = text[:-2].strip()
    elif text.lower().endswith("cr"):
        text = text[:-2].strip()

    cleaned = re.sub(r"[^\d.\-+]", "", text.replace(",", ""))
    if not cleaned or cleaned in {"-", "+", "."}:
        return None
    try:
        num = float(cleaned)
        return -abs(num) if is_negative else num
    except ValueError:
        return None


def _auto_categorize(description: str) -> str:
    low = description.lower()
    for cat, keywords in AUTO_CATEGORY_RULES:
        if any(kw in low for kw in keywords):
            return cat
    return "Other"


def _find_best_header_row(raw_df: pd.DataFrame) -> int:
    """Scan top 15 rows to locate the most likely header row."""
    best_idx = 0
    best_score = -1
    all_keywords = DATE_KEYWORDS + AMOUNT_KEYWORDS + DEBIT_KEYWORDS + CREDIT_KEYWORDS + CATEGORY_KEYWORDS + DESC_KEYWORDS
    limit = min(15, len(raw_df))
    for i in range(limit):
        row_vals = [_normalize_col(v) for v in raw_df.iloc[i].tolist()]
        non_empty = [v for v in row_vals if v and v != "nan"]
        if len(non_empty) < 2:
            continue
        score = 0
        for cell in non_empty:
            if any(kw == cell or kw in cell for kw in all_keywords):
                score += 3
            elif not cell.replace(".", "").isdigit():
                score += 1
        if score > best_score:
            best_score = score
            best_idx = i
    return best_idx


def _match_column(columns: list[str], keywords: tuple[str, ...], exclude: set[str] | None = None) -> str | None:
    exclude = exclude or set()
    norm_map = {col: _normalize_col(col) for col in columns if col not in exclude}
    # Exact match first
    for kw in keywords:
        for col, norm in norm_map.items():
            if norm == kw:
                return col
    # Whole-word match next (prevents "cr" matching "description" or "dr" matching "address")
    for kw in keywords:
        pattern = re.compile(rf"\b{re.escape(kw)}\b")
        for col, norm in norm_map.items():
            if pattern.search(norm):
                return col
    # Substring fallback only for longer descriptive keywords (>= 4 chars)
    for kw in keywords:
        if len(kw) < 4:
            continue
        for col, norm in norm_map.items():
            if kw in norm:
                return col
    return None


@dataclass
class WorkbookMetadata:
    file_path: str
    sheet_name: str
    available_sheets: list[str]
    currency: str
    date_column: str
    amount_source: str
    category_column: str | None
    description_column: str | None
    payment_column: str | None
    rows_loaded: int
    expense_rows: int
    income_rows_excluded: int


@dataclass
class ExpenseAnalyzer:
    file_path: Path
    sheet_name: str | None = None
    currency_override: str | None = None
    df: pd.DataFrame = field(init=False)
    income_df: pd.DataFrame = field(init=False)
    meta: WorkbookMetadata = field(init=False)

    def __post_init__(self) -> None:
        self.file_path = Path(self.file_path)
        if not self.file_path.exists():
            raise FileNotFoundError(f"File not found: {self.file_path}")
        self._load_and_normalize()

    def _load_raw_sheet(self) -> tuple[pd.DataFrame, str, list[str]]:
        suffix = self.file_path.suffix.lower()
        if suffix == ".csv":
            raw = pd.read_csv(self.file_path, header=None, dtype=object)
            return raw, "CSV", ["CSV"]

        if suffix not in {".xlsx", ".xlsm", ".xls"}:
            raise ValueError(f"Unsupported file format: {suffix}. Expected .xlsx, .xlsm, .xls, or .csv")

        wb = openpyxl.load_workbook(self.file_path, read_only=True, data_only=True)
        sheets = wb.sheetnames
        wb.close()

        chosen_sheet = self.sheet_name
        if chosen_sheet and chosen_sheet not in sheets:
            raise ValueError(f"Sheet '{chosen_sheet}' not found. Available sheets: {sheets}")

        if not chosen_sheet:
            # Pick the sheet with the most data rows
            best_sheet = sheets[0]
            best_len = -1
            for s in sheets:
                candidate = pd.read_excel(self.file_path, sheet_name=s, header=None, dtype=object)
                non_empty_rows = int(candidate.dropna(how="all").shape[0])
                if non_empty_rows > best_len:
                    best_len = non_empty_rows
                    best_sheet = s
            chosen_sheet = best_sheet

        raw = pd.read_excel(self.file_path, sheet_name=chosen_sheet, header=None, dtype=object)
        return raw, chosen_sheet, sheets

    def _load_and_normalize(self) -> None:
        raw, chosen_sheet, sheets = self._load_raw_sheet()
        raw = raw.dropna(how="all").reset_index(drop=True)
        if raw.empty:
            raise ValueError("The selected spreadsheet is empty.")

        header_idx = _find_best_header_row(raw)
        headers = []
        seen: dict[str, int] = {}
        for i, h in enumerate(raw.iloc[header_idx].tolist()):
            name = str(h).strip() if h is not None and str(h).strip() != "nan" else f"Column_{i+1}"
            if name in seen:
                seen[name] += 1
                name = f"{name}_{seen[name]}"
            else:
                seen[name] = 1
            headers.append(name)

        data = raw.iloc[header_idx + 1 :].copy()
        data.columns = headers
        data = data.dropna(how="all").reset_index(drop=True)

        cols = list(data.columns)
        used: set[str] = set()

        date_col = _match_column(cols, DATE_KEYWORDS)
        if not date_col:
            # Fallback: find column with the most parseable dates
            best_col, best_valid = None, 0
            for c in cols:
                parsed = pd.to_datetime(data[c], errors="coerce")
                valid = int(parsed.notna().sum())
                if valid > best_valid:
                    best_valid = valid
                    best_col = c
            if best_col and best_valid >= max(2, int(len(data) * 0.2)):
                date_col = best_col
        if not date_col:
            raise ValueError(f"Could not detect a Date column. Columns found: {cols}")
        used.add(date_col)

        debit_col = _match_column(cols, DEBIT_KEYWORDS, exclude=used)
        credit_col = _match_column(cols, CREDIT_KEYWORDS, exclude=used | ({debit_col} if debit_col else set()))
        amount_col = _match_column(cols, AMOUNT_KEYWORDS, exclude=used)
        type_col = _match_column(cols, TYPE_KEYWORDS, exclude=used | ({amount_col} if amount_col else set()))

        category_col = _match_column(cols, CATEGORY_KEYWORDS, exclude=used | ({type_col} if type_col and _normalize_col(type_col) in ("dr/cr", "cr/dr", "transaction type", "txn type") else set()))
        if category_col == type_col and category_col is not None:
            # Check if values look like Dr/Cr or Income/Expense vs actual categories
            unique_vals = {_normalize_col(v) for v in data[category_col].dropna().unique()[:20]}
            if unique_vals.issubset({"debit", "credit", "dr", "cr", "expense", "income", "inflow", "outflow"}):
                type_col = category_col
                category_col = None

        desc_col = _match_column(cols, DESC_KEYWORDS, exclude=used | ({category_col} if category_col else set()))
        payment_col = _match_column(cols, PAYMENT_KEYWORDS, exclude=used | ({category_col, desc_col} - {None}))

        # Currency detection: check top metadata rows & Excel number formats first, then cell values
        sample_money_vals: list[Any] = []
        if header_idx > 0:
            for r_idx in range(header_idx):
                sample_money_vals.extend(raw.iloc[r_idx].dropna().tolist())
        if self.file_path.suffix.lower() in {".xlsx", ".xlsm"}:
            try:
                wb_fmt = openpyxl.load_workbook(self.file_path, read_only=True, data_only=False)
                ws_fmt = wb_fmt[chosen_sheet]
                for row_cells in ws_fmt.iter_rows(min_row=1, max_row=min(30, ws_fmt.max_row or 30)):
                    for cell in row_cells:
                        if cell.number_format and cell.number_format != "General":
                            sample_money_vals.append(cell.number_format)
                wb_fmt.close()
            except Exception:
                pass
        for c in (amount_col, debit_col, credit_col):
            if c and c in data.columns:
                sample_money_vals.extend(data[c].dropna().tolist()[:50])
        currency = self.currency_override or _detect_currency_symbol(sample_money_vals)

        # Build normalized records
        parsed_dates = pd.to_datetime(data[date_col], errors="coerce")
        records: list[dict[str, Any]] = []
        income_records: list[dict[str, Any]] = []

        amount_source = ""
        if debit_col and (not amount_col or debit_col != amount_col):
            amount_source = f"{debit_col}" + (f" / {credit_col}" if credit_col else "")
        elif amount_col:
            amount_source = amount_col
        else:
            raise ValueError(f"Could not detect an Amount or Debit column. Columns found: {cols}")

        for idx, row in data.iterrows():
            dt = parsed_dates.iloc[idx]
            if pd.isna(dt):
                continue

            desc = str(row[desc_col]).strip() if desc_col and pd.notna(row[desc_col]) else "Unspecified"
            if desc.lower() in {"nan", "none", ""}:
                desc = "Unspecified"

            # Skip obvious summary/total rows
            if desc.lower() in {"total", "grand total", "subtotal", "closing balance", "opening balance"}:
                continue

            raw_cat = str(row[category_col]).strip() if category_col and pd.notna(row[category_col]) else ""
            if raw_cat.lower() in {"nan", "none", ""}:
                raw_cat = ""

            pay_mode = str(row[payment_col]).strip() if payment_col and pd.notna(row[payment_col]) else ""
            if pay_mode.lower() in {"nan", "none"}:
                pay_mode = ""

            expense_amt: float | None = None
            income_amt: float | None = None

            if debit_col and credit_col:
                d_val = _parse_number(row[debit_col])
                c_val = _parse_number(row[credit_col])
                if d_val and abs(d_val) > 0:
                    expense_amt = abs(d_val)
                elif c_val and abs(c_val) > 0:
                    income_amt = abs(c_val)
            elif debit_col and not amount_col:
                d_val = _parse_number(row[debit_col])
                if d_val and abs(d_val) > 0:
                    expense_amt = abs(d_val)
            else:
                val = _parse_number(row[amount_col])
                if val is None or val == 0:
                    continue
                if type_col and pd.notna(row[type_col]):
                    t_norm = _normalize_col(row[type_col])
                    if any(k in t_norm for k in ("credit", "income", "cr", "salary", "refund", "deposit", "inflow")):
                        income_amt = abs(val)
                    else:
                        expense_amt = abs(val)
                else:
                    # Check if dataset mixes negative & positive numbers (bank statement style)
                    expense_amt = abs(val)
                    # We'll adjust sign semantics below if mixed signs exist
                    records.append({
                        "_raw_signed": val,
                        "date": dt,
                        "description": desc,
                        "category": raw_cat or _auto_categorize(desc),
                        "payment_mode": pay_mode,
                        "amount": abs(val),
                    })
                    continue

            if expense_amt and expense_amt > 0:
                records.append({
                    "_raw_signed": -expense_amt,
                    "date": dt,
                    "description": desc,
                    "category": raw_cat or _auto_categorize(desc),
                    "payment_mode": pay_mode,
                    "amount": round(expense_amt, 2),
                })
            elif income_amt and income_amt > 0:
                income_records.append({
                    "date": dt,
                    "description": desc,
                    "category": raw_cat or "Income",
                    "payment_mode": pay_mode,
                    "amount": round(income_amt, 2),
                })

        # If single amount column had both negative and positive values without type_col:
        # In bank statements, negatives are expenses and positives are credits/income.
        if records and not (debit_col and credit_col) and not type_col:
            neg_count = sum(1 for r in records if r["_raw_signed"] < 0)
            pos_count = sum(1 for r in records if r["_raw_signed"] > 0)
            if neg_count > 0 and pos_count > 0:
                # Negative rows are expenses, positive rows are income
                new_records = []
                for r in records:
                    if r["_raw_signed"] < 0:
                        r["amount"] = round(abs(r["_raw_signed"]), 2)
                        new_records.append(r)
                    else:
                        income_records.append({
                            "date": r["date"],
                            "description": r["description"],
                            "category": r["category"],
                            "payment_mode": r["payment_mode"],
                            "amount": round(abs(r["_raw_signed"]), 2),
                        })
                records = new_records

        if not records:
            raise ValueError("No valid expense rows found after parsing dates and amounts.")

        df = pd.DataFrame(records).drop(columns=["_raw_signed"])
        df["date"] = pd.to_datetime(df["date"])
        df = df.sort_values("date").reset_index(drop=True)
        df["month"] = df["date"].dt.to_period("M").astype(str)
        df["calendar_month"] = df["date"].dt.month_name()
        df["calendar_month_num"] = df["date"].dt.month
        df["day_of_week"] = df["date"].dt.day_name()
        df["is_weekend"] = df["date"].dt.dayofweek >= 5
        df["day_of_month"] = df["date"].dt.day

        self.df = df
        self.income_df = pd.DataFrame(income_records)
        self.meta = WorkbookMetadata(
            file_path=str(self.file_path),
            sheet_name=chosen_sheet,
            available_sheets=sheets,
            currency=currency,
            date_column=date_col,
            amount_source=amount_source,
            category_column=category_col,
            description_column=desc_col,
            payment_column=payment_col,
            rows_loaded=len(data),
            expense_rows=len(df),
            income_rows_excluded=len(income_records),
        )

    def get_overview(self) -> dict[str, Any]:
        monthly = self.df.groupby("month")["amount"].sum()
        total_spend = float(self.df["amount"].sum())
        total_income = float(self.income_df["amount"].sum()) if not self.income_df.empty else 0.0
        return {
            "file": self.meta.file_path,
            "sheet": self.meta.sheet_name,
            "currency": self.meta.currency,
            "date_range": {
                "start": self.df["date"].min().strftime("%Y-%m-%d"),
                "end": self.df["date"].max().strftime("%Y-%m-%d"),
                "num_months": int(monthly.shape[0]),
            },
            "total_spend": round(total_spend, 2),
            "total_income_detected": round(total_income, 2),
            "net_cashflow": round(total_income - total_spend, 2) if total_income > 0 else None,
            "transaction_count": int(len(self.df)),
            "average_transaction": round(float(self.df["amount"].mean()), 2),
            "median_transaction": round(float(self.df["amount"].median()), 2),
            "average_monthly_spend": round(float(monthly.mean()), 2),
            "median_monthly_spend": round(float(monthly.median()), 2),
            "largest_single_expense": {
                "date": self.df.loc[self.df["amount"].idxmax(), "date"].strftime("%Y-%m-%d"),
                "description": str(self.df.loc[self.df["amount"].idxmax(), "description"]),
                "category": str(self.df.loc[self.df["amount"].idxmax(), "category"]),
                "amount": round(float(self.df["amount"].max()), 2),
            },
            "detected_columns": {
                "date": self.meta.date_column,
                "amount": self.meta.amount_source,
                "category": self.meta.category_column or "Auto-categorized from description",
                "description": self.meta.description_column,
                "payment_mode": self.meta.payment_column,
            },
        }

    def get_where_you_spend_more(self, top_n: int = 10) -> dict[str, Any]:
        total = float(self.df["amount"].sum())
        cat_grp = (
            self.df.groupby("category")
            .agg(
                total_spend=("amount", "sum"),
                tx_count=("amount", "count"),
                avg_spend=("amount", "mean"),
                max_spend=("amount", "max"),
            )
            .sort_values("total_spend", ascending=False)
        )
        categories = []
        for cat, row in cat_grp.head(top_n).iterrows():
            sub = self.df[self.df["category"] == cat].sort_values("amount", ascending=False)
            top_items = [
                {
                    "date": r["date"].strftime("%Y-%m-%d"),
                    "description": r["description"],
                    "amount": round(float(r["amount"]), 2),
                }
                for _, r in sub.head(3).iterrows()
            ]
            categories.append({
                "category": str(cat),
                "total_spend": round(float(row["total_spend"]), 2),
                "share_percent": round(float(row["total_spend"]) / total * 100, 1) if total else 0.0,
                "transaction_count": int(row["tx_count"]),
                "average_transaction": round(float(row["avg_spend"]), 2),
                "max_transaction": round(float(row["max_spend"]), 2),
                "top_examples": top_items,
            })

        merch_grp = (
            self.df.groupby("description")
            .agg(
                total_spend=("amount", "sum"),
                tx_count=("amount", "count"),
                avg_spend=("amount", "mean"),
                category=("category", "first"),
            )
            .sort_values("total_spend", ascending=False)
        )
        merchants = [
            {
                "description": str(desc),
                "category": str(row["category"]),
                "total_spend": round(float(row["total_spend"]), 2),
                "share_percent": round(float(row["total_spend"]) / total * 100, 1) if total else 0.0,
                "transaction_count": int(row["tx_count"]),
                "average_transaction": round(float(row["avg_spend"]), 2),
            }
            for desc, row in merch_grp.head(top_n).iterrows()
        ]

        freq_grp = merch_grp.sort_values("tx_count", ascending=False)
        most_frequent = [
            {
                "description": str(desc),
                "transaction_count": int(row["tx_count"]),
                "total_spend": round(float(row["total_spend"]), 2),
            }
            for desc, row in freq_grp.head(5).iterrows()
        ]

        payment_modes = []
        if self.df["payment_mode"].str.len().gt(0).any():
            pm_grp = (
                self.df[self.df["payment_mode"].str.len() > 0]
                .groupby("payment_mode")
                .agg(total_spend=("amount", "sum"), tx_count=("amount", "count"))
                .sort_values("total_spend", ascending=False)
            )
            payment_modes = [
                {
                    "payment_mode": str(pm),
                    "total_spend": round(float(r["total_spend"]), 2),
                    "share_percent": round(float(r["total_spend"]) / total * 100, 1) if total else 0.0,
                    "transaction_count": int(r["tx_count"]),
                }
                for pm, r in pm_grp.iterrows()
            ]

        return {
            "top_category": categories[0] if categories else None,
            "categories": categories,
            "top_merchants_by_spend": merchants,
            "most_frequent_merchants": most_frequent,
            "payment_modes": payment_modes,
        }

    def get_monthly_spending(self) -> dict[str, Any]:
        m_grp = (
            self.df.groupby("month")
            .agg(
                total_spend=("amount", "sum"),
                tx_count=("amount", "count"),
                avg_tx=("amount", "mean"),
            )
            .sort_index()
        )
        avg_monthly = float(m_grp["total_spend"].mean())
        months_list = []
        prev_val: float | None = None
        for month_str, row in m_grp.iterrows():
            val = float(row["total_spend"])
            mom_pct = round((val - prev_val) / prev_val * 100, 1) if prev_val and prev_val > 0 else None
            vs_avg_pct = round((val - avg_monthly) / avg_monthly * 100, 1) if avg_monthly > 0 else 0.0
            top_cat = (
                self.df[self.df["month"] == month_str]
                .groupby("category")["amount"]
                .sum()
                .idxmax()
            )
            months_list.append({
                "month": str(month_str),
                "total_spend": round(val, 2),
                "transaction_count": int(row["tx_count"]),
                "average_transaction": round(float(row["avg_tx"]), 2),
                "mom_change_percent": mom_pct,
                "vs_monthly_avg_percent": vs_avg_pct,
                "top_category_in_month": str(top_cat),
            })
            prev_val = val

        peak_row = max(months_list, key=lambda x: x["total_spend"])
        low_row = min(months_list, key=lambda x: x["total_spend"])
        ranked = sorted(months_list, key=lambda x: x["total_spend"], reverse=True)

        return {
            "monthly_average": round(avg_monthly, 2),
            "peak_month": peak_row,
            "lowest_month": low_row,
            "top_3_highest_months": ranked[:3],
            "months_chronological": months_list,
        }

    def detect_recurring_expenses(self) -> list[dict[str, Any]]:
        """Detect merchants/descriptions appearing across multiple months with consistent amounts."""
        grouped = (
            self.df.groupby("description")
            .agg(
                months=("month", "nunique"),
                tx_count=("amount", "count"),
                avg_amount=("amount", "mean"),
                std_amount=("amount", "std"),
                total_spend=("amount", "sum"),
                category=("category", "first"),
            )
        )
        total_months = max(1, int(self.df["month"].nunique()))
        recurring = []
        for desc, row in grouped.iterrows():
            months_seen = int(row["months"])
            avg_amt = float(row["avg_amount"])
            std_amt = 0.0 if pd.isna(row["std_amount"]) else float(row["std_amount"])
            cv = std_amt / avg_amt if avg_amt > 0 else 1.0
            if months_seen >= min(3, total_months) and months_seen >= 2 and cv <= 0.25:
                recurring.append({
                    "description": str(desc),
                    "category": str(row["category"]),
                    "months_appeared": months_seen,
                    "typical_amount": round(avg_amt, 2),
                    "total_spend": round(float(row["total_spend"]), 2),
                })
        recurring.sort(key=lambda x: x["typical_amount"], reverse=True)
        return recurring[:15]

    def get_future_month_forecast(self) -> dict[str, Any]:
        """Forecast upcoming monthly spend & identify seasonal/calendar months with highest risk of high spend."""
        m_series = self.df.groupby("month")["amount"].sum().sort_index()
        months = list(m_series.index)
        values = [float(v) for v in m_series.values]

        last_period = pd.Period(months[-1], freq="M")
        next_months = [(last_period + i).strftime("%Y-%m") for i in range(1, 4)]

        # Simple trend + weighted recent average
        if len(values) >= 3:
            recent_3_avg = sum(values[-3:]) / 3.0
            # Linear slope across months
            n = len(values)
            x_mean = (n - 1) / 2.0
            y_mean = sum(values) / n
            denom = sum((i - x_mean) ** 2 for i in range(n)) or 1.0
            slope = sum((i - x_mean) * (values[i] - y_mean) for i in range(n)) / denom
            trend_direction = "increasing" if slope > (y_mean * 0.02) else ("decreasing" if slope < -(y_mean * 0.02) else "stable")
        else:
            recent_3_avg = sum(values) / max(1, len(values))
            slope = 0.0
            trend_direction = "insufficient_history"

        # Calendar month seasonality (if >= 6 months or multiple years)
        cal_grp = (
            self.df.groupby(["calendar_month_num", "calendar_month", "month"])["amount"]
            .sum()
            .reset_index()
            .groupby(["calendar_month_num", "calendar_month"])["amount"]
            .mean()
            .reset_index()
            .sort_values("amount", ascending=False)
        )
        seasonal_pattern = [
            {
                "calendar_month": str(r["calendar_month"]),
                "average_spend": round(float(r["amount"]), 2),
            }
            for _, r in cal_grp.iterrows()
        ]
        cal_lookup = {r["calendar_month"]: float(r["average_spend"]) for r in seasonal_pattern}

        overall_avg = sum(values) / max(1, len(values))
        forecasts = []
        for idx, nm in enumerate(next_months, start=1):
            p = pd.Period(nm, freq="M")
            cal_name = p.to_timestamp().strftime("%B")
            base_pred = max(0.0, recent_3_avg + slope * idx * 0.5)
            # Blend with seasonal calendar average if that calendar month was seen in data
            if cal_name in cal_lookup and len(values) >= 12:
                predicted = 0.5 * base_pred + 0.5 * cal_lookup[cal_name]
            else:
                predicted = base_pred
            forecasts.append({
                "month": nm,
                "calendar_month": cal_name,
                "projected_spend": round(predicted, 2),
                "vs_historical_avg_percent": round((predicted - overall_avg) / overall_avg * 100, 1) if overall_avg else 0.0,
            })

        recurring = self.detect_recurring_expenses()
        monthly_fixed_estimate = round(sum(r["typical_amount"] for r in recurring), 2)

        highest_forecast = max(forecasts, key=lambda x: x["projected_spend"]) if forecasts else None
        highest_seasonal_cal = seasonal_pattern[0] if seasonal_pattern else None

        return {
            "trend_direction": trend_direction,
            "monthly_slope_estimate": round(slope, 2),
            "recent_3_month_avg": round(recent_3_avg, 2),
            "estimated_recurring_monthly_commitments": monthly_fixed_estimate,
            "recurring_items": recurring[:10],
            "next_3_months_forecast": forecasts,
            "highest_projected_upcoming_month": highest_forecast,
            "highest_historical_calendar_month": highest_seasonal_cal,
            "seasonal_ranking": seasonal_pattern[:5],
            "note": (
                "Forecast blends recent 3-month spending, overall trend slope, and historical calendar-month pattern when >=12 months exist. "
                "If fewer than 12 months exist, seasonal spikes from unseen calendar months cannot be predicted purely from the sheet."
            ),
        }

    def explain_why_it_happens(self, target_month: str | None = None) -> dict[str, Any]:
        """Root-cause analysis for why spending spiked in the peak (or target) month and top categories."""
        monthly_totals = self.df.groupby("month")["amount"].sum()
        if not target_month:
            target_month = str(monthly_totals.idxmax())
        elif target_month not in monthly_totals.index:
            raise ValueError(f"Month '{target_month}' not found. Available months: {list(monthly_totals.index)}")

        target_df = self.df[self.df["month"] == target_month]
        other_df = self.df[self.df["month"] != target_month]

        target_total = float(target_df["amount"].sum())
        num_other_months = max(1, int(other_df["month"].nunique()))
        other_monthly_avg = float(other_df["amount"].sum()) / num_other_months if not other_df.empty else target_total
        excess_over_baseline = target_total - other_monthly_avg

        # Category comparison: target month vs baseline monthly avg per category
        target_cat = target_df.groupby("category")["amount"].sum()
        other_cat_avg = (other_df.groupby("category")["amount"].sum() / num_other_months) if not other_df.empty else pd.Series(dtype=float)

        all_cats = sorted(set(target_cat.index) | set(other_cat_avg.index))
        category_drivers = []
        for cat in all_cats:
            t_val = float(target_cat.get(cat, 0.0))
            b_val = float(other_cat_avg.get(cat, 0.0))
            diff = t_val - b_val
            pct_change = round((diff / b_val) * 100, 1) if b_val > 0 else (100.0 if t_val > 0 else 0.0)
            category_drivers.append({
                "category": str(cat),
                "target_month_spend": round(t_val, 2),
                "baseline_monthly_avg": round(b_val, 2),
                "difference": round(diff, 2),
                "percent_vs_baseline": pct_change,
            })
        category_drivers.sort(key=lambda x: x["difference"], reverse=True)

        # Top individual transactions driving the target month
        top_tx = [
            {
                "date": r["date"].strftime("%Y-%m-%d"),
                "description": str(r["description"]),
                "category": str(r["category"]),
                "amount": round(float(r["amount"]), 2),
                "share_of_month_percent": round(float(r["amount"]) / target_total * 100, 1) if target_total else 0.0,
            }
            for _, r in target_df.sort_values("amount", ascending=False).head(8).iterrows()
        ]

        # Anomaly / Outlier detection across entire dataset
        mean_tx = float(self.df["amount"].mean())
        std_tx = float(self.df["amount"].std()) if len(self.df) > 1 else 0.0
        outlier_threshold = max(mean_tx + 1.5 * std_tx, float(self.df["amount"].quantile(0.90)))
        month_outliers = target_df[target_df["amount"] >= outlier_threshold].sort_values("amount", ascending=False)
        outliers_list = [
            {
                "date": r["date"].strftime("%Y-%m-%d"),
                "description": str(r["description"]),
                "category": str(r["category"]),
                "amount": round(float(r["amount"]), 2),
            }
            for _, r in month_outliers.head(10).iterrows()
        ]

        # Behavioral habit signals across whole sheet
        weekend_spend = float(self.df[self.df["is_weekend"]]["amount"].sum())
        weekday_spend = float(self.df[~self.df["is_weekend"]]["amount"].sum())
        weekend_days = max(1, int(self.df[self.df["is_weekend"]]["date"].dt.date.nunique()))
        weekday_days = max(1, int(self.df[~self.df["is_weekend"]]["date"].dt.date.nunique()))

        # Small frequent purchases (< overall median or <= 25% of avg tx)
        small_threshold = float(self.df["amount"].median())
        small_tx = self.df[self.df["amount"] <= small_threshold]
        large_tx = self.df[self.df["amount"] > outlier_threshold]

        # Day-of-month pattern (Early 1-10, Mid 11-20, Late 21-31)
        early = float(self.df[self.df["day_of_month"] <= 10]["amount"].sum())
        mid = float(self.df[(self.df["day_of_month"] >= 11) & (self.df["day_of_month"] <= 20)]["amount"].sum())
        late = float(self.df[self.df["day_of_month"] >= 21]["amount"].sum())
        total_all = float(self.df["amount"].sum()) or 1.0

        # Frequency vs ticket size comparison in target month vs baseline
        target_tx_count = int(len(target_df))
        other_avg_tx_count = round(len(other_df) / num_other_months, 1) if not other_df.empty else float(target_tx_count)
        target_avg_ticket = round(float(target_df["amount"].mean()), 2)
        other_avg_ticket = round(float(other_df["amount"].mean()), 2) if not other_df.empty else target_avg_ticket

        return {
            "analyzed_month": target_month,
            "target_month_total": round(target_total, 2),
            "other_months_avg_total": round(other_monthly_avg, 2),
            "excess_spend_over_baseline": round(excess_over_baseline, 2),
            "frequency_vs_ticket_size": {
                "target_month_transactions": target_tx_count,
                "baseline_avg_transactions": other_avg_tx_count,
                "target_month_avg_ticket": target_avg_ticket,
                "baseline_avg_ticket": other_avg_ticket,
            },
            "top_category_surges": category_drivers[:6],
            "top_transactions_in_month": top_tx,
            "unusual_spike_transactions": outliers_list,
            "overall_spending_habits": {
                "weekend_spend_total": round(weekend_spend, 2),
                "weekday_spend_total": round(weekday_spend, 2),
                "avg_spend_per_active_weekend_day": round(weekend_spend / weekend_days, 2),
                "avg_spend_per_active_weekday": round(weekday_spend / weekday_days, 2),
                "small_frequent_purchases": {
                    "threshold_amount": round(small_threshold, 2),
                    "count": int(len(small_tx)),
                    "total_spend": round(float(small_tx["amount"].sum()), 2),
                    "share_percent": round(float(small_tx["amount"].sum()) / total_all * 100, 1),
                },
                "large_one_off_purchases": {
                    "threshold_amount": round(outlier_threshold, 2),
                    "count": int(len(large_tx)),
                    "total_spend": round(float(large_tx["amount"].sum()), 2),
                    "share_percent": round(float(large_tx["amount"].sum()) / total_all * 100, 1),
                },
                "month_timing_split_percent": {
                    "early_month_days_1_10": round(early / total_all * 100, 1),
                    "mid_month_days_11_20": round(mid / total_all * 100, 1),
                    "late_month_days_21_31": round(late / total_all * 100, 1),
                },
            },
        }

    def search_transactions(
        self,
        month: str | None = None,
        category: str | None = None,
        keyword: str | None = None,
        min_amount: float | None = None,
        limit: int = 20,
    ) -> dict[str, Any]:
        sub = self.df.copy()
        if month:
            sub = sub[sub["month"] == month]
        if category:
            sub = sub[sub["category"].str.lower().str.contains(category.lower(), na=False)]
        if keyword:
            mask = (
                sub["description"].str.lower().str.contains(keyword.lower(), na=False)
                | sub["category"].str.lower().str.contains(keyword.lower(), na=False)
            )
            sub = sub[mask]
        if min_amount is not None:
            sub = sub[sub["amount"] >= float(min_amount)]

        sub = sub.sort_values("amount", ascending=False)
        rows = [
            {
                "date": r["date"].strftime("%Y-%m-%d"),
                "month": r["month"],
                "category": str(r["category"]),
                "description": str(r["description"]),
                "amount": round(float(r["amount"]), 2),
            }
            for _, r in sub.head(limit).iterrows()
        ]
        return {
            "matching_count": int(len(sub)),
            "matching_total_spend": round(float(sub["amount"].sum()), 2),
            "showing_top": len(rows),
            "transactions": rows,
        }

    def build_full_analytics_bundle(self) -> dict[str, Any]:
        return {
            "overview": self.get_overview(),
            "where_you_spend_more": self.get_where_you_spend_more(top_n=10),
            "monthly_spending": self.get_monthly_spending(),
            "future_month_forecast": self.get_future_month_forecast(),
            "why_peak_month_happened": self.explain_why_it_happens(),
        }
