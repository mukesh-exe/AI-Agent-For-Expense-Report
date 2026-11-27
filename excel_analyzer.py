"""
Universal Excel/CSV ingestion and analytics engine.

Supports:
1. Single or multiple .xlsx, .xlsm, .xls, and .csv files
2. Multi-sheet workbooks (all tabs parsed and indexed; avoids summary double-counting)
3. Transaction logs (single Amount, Debit/Credit, Type, inline CR/DR, mixed signs)
4. Matrix / Pivot monthly reports (Category rows x Month columns -> unpivoted)
5. One-tab-per-month workbooks (e.g., Jan, Feb, Mar tabs)
6. General non-date tabular reports (Sales, Inventory, P&L, KPI tables)
"""

from __future__ import annotations

import calendar
import math
import re
import warnings
from dataclasses import dataclass, field
from datetime import datetime
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
    "department",
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
    "product",
    "name",
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

MONTH_NAME_MAP: dict[str, int] = {}
for _m_idx in range(1, 13):
    MONTH_NAME_MAP[calendar.month_name[_m_idx].lower()] = _m_idx
    MONTH_NAME_MAP[calendar.month_abbr[_m_idx].lower()] = _m_idx


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


def _parse_number_and_drcr(val: Any) -> tuple[float | None, str | None]:
    """Parse monetary value and optional inline DR/CR suffix."""
    if val is None or (isinstance(val, float) and math.isnan(val)):
        return None, None
    if isinstance(val, (int, float)):
        return float(val), None
    text = str(val).strip()
    if not text or text in {"-", "--", "N/A", "na", "null", "None"}:
        return None, None

    is_negative = False
    if text.startswith("(") and text.endswith(")"):
        is_negative = True
        text = text[1:-1].strip()

    drcr: str | None = None
    low = text.lower()
    if re.search(r"\bdr\.?$", low) or low.endswith("dr"):
        drcr = "dr"
        text = re.sub(r"(?i)\s*dr\.?$", "", text).strip()
    elif re.search(r"\bcr\.?$", low) or low.endswith("cr"):
        drcr = "cr"
        text = re.sub(r"(?i)\s*cr\.?$", "", text).strip()

    cleaned = re.sub(r"[^\d.\-+]", "", text.replace(",", ""))
    if not cleaned or cleaned in {"-", "+", "."}:
        return None, drcr
    try:
        num = float(cleaned)
        return (-abs(num) if is_negative else num), drcr
    except ValueError:
        return None, drcr


def _parse_number(val: Any) -> float | None:
    num, _ = _parse_number_and_drcr(val)
    return num


def _auto_categorize(description: str) -> str:
    low = description.lower()
    for cat, keywords in AUTO_CATEGORY_RULES:
        if any(kw in low for kw in keywords):
            return cat
    return "Other"


def _parse_month_label(label: Any, default_year: int | None = None) -> pd.Timestamp | None:
    """Detect whether a column header or sheet name represents a calendar month."""
    if isinstance(label, (datetime, pd.Timestamp)):
        return pd.Timestamp(year=label.year, month=label.month, day=1)

    text = str(label or "").strip()
    if not text:
        return None
    low = text.lower()
    if any(skip in low for skip in ("total", "avg", "average", "budget", "target", "diff", "variance")):
        return None

    # YYYY-MM or YYYY/MM
    m_ym = re.match(r"^(20\d{2}|19\d{2})[-/](\d{1,2})$", low)
    if m_ym:
        y, m = int(m_ym.group(1)), int(m_ym.group(2))
        if 1 <= m <= 12:
            return pd.Timestamp(year=y, month=m, day=1)

    # Month name optionally followed/preceded by year: "Jan", "January", "Jan-26", "Oct 2026", "2026 Oct"
    year = default_year or datetime.now().year
    y4 = re.search(r"\b(20\d{2}|19\d{2})\b", low)
    y2 = re.search(r"[-/\s'](\d{2})$", low)
    if y4:
        year = int(y4.group(1))
    elif y2:
        year = 2000 + int(y2.group(1))

    cleaned = re.sub(r"[\d\-/,'_\s]+", " ", low).strip()
    for token in cleaned.split():
        if token in MONTH_NAME_MAP:
            return pd.Timestamp(year=year, month=MONTH_NAME_MAP[token], day=1)

    # Also match "Jan26" or "Oct2026" without separator
    m_compact = re.match(r"^([a-z]{3,9})(\d{2,4})$", low)
    if m_compact and m_compact.group(1) in MONTH_NAME_MAP:
        yr_raw = int(m_compact.group(2))
        yr = yr_raw if yr_raw > 1900 else 2000 + yr_raw
        return pd.Timestamp(year=yr, month=MONTH_NAME_MAP[m_compact.group(1)], day=1)

    return None


def _smart_parse_dates(series: pd.Series) -> pd.Series:
    """Parse a date series cleanly, ignoring pure numeric/money values and supporting DD/MM formats."""
    if pd.api.types.is_datetime64_any_dtype(series):
        return pd.to_datetime(series, errors="coerce")

    non_null = series.dropna()
    if non_null.empty:
        return pd.Series([pd.NaT] * len(series), index=series.index)

    # Reject columns whose values are pure ints/floats (pandas would treat them as 1970 epoch ns)
    sample_vals = non_null.head(30).tolist()
    date_like_count = 0
    for v in sample_vals:
        if isinstance(v, (datetime, pd.Timestamp)):
            date_like_count += 1
        elif isinstance(v, str):
            s = v.strip()
            if re.search(r"[-/.,]", s) or any(m in s.lower() for m in MONTH_NAME_MAP) or re.match(r"^(19|20)\d{6}$", s):
                if not re.match(r"^[₹$€£¥\s()+-]*[\d,]+\.\d{1,2}\s*(cr|dr)?$", s.lower()):
                    date_like_count += 1
    if date_like_count == 0:
        return pd.Series([pd.NaT] * len(series), index=series.index)

    as_str = non_null.astype(str).head(50).tolist()
    if as_str and all(re.match(r"^\d{4}[-/]\d{1,2}[-/]\d{1,2}", s.strip()) for s in as_str[:10]):
        with warnings.catch_warnings():
            warnings.simplefilter("ignore", UserWarning)
            return pd.to_datetime(series, errors="coerce")

    with warnings.catch_warnings():
        warnings.simplefilter("ignore", UserWarning)
        parsed_default = pd.to_datetime(series, errors="coerce", dayfirst=False)
        parsed_dayfirst = pd.to_datetime(series, errors="coerce", dayfirst=True)
    if int(parsed_dayfirst.notna().sum()) > int(parsed_default.notna().sum()):
        return parsed_dayfirst
    return parsed_default


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
            if any(kw == cell or (len(kw) >= 4 and kw in cell) for kw in all_keywords):
                score += 3
            elif _parse_month_label(cell) is not None:
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
    for kw in keywords:
        for col, norm in norm_map.items():
            if norm == kw:
                return col
    for kw in keywords:
        pattern = re.compile(rf"\b{re.escape(kw)}\b")
        for col, norm in norm_map.items():
            if pattern.search(norm):
                return col
    for kw in keywords:
        if len(kw) < 4:
            continue
        for col, norm in norm_map.items():
            if kw in norm:
                return col
    return None


@dataclass
class SheetTableProfile:
    file_name: str
    sheet_name: str
    sheet_kind: str  # "transactions" | "matrix" | "general"
    columns: list[str]
    row_count: int
    currency: str
    dataframe: pd.DataFrame
    numeric_summary: list[dict[str, Any]]
    top_group_breakdown: list[dict[str, Any]]
    title_notes: list[str]


@dataclass
class WorkbookMetadata:
    file_path: str
    files_loaded: list[str]
    sheet_name: str
    available_sheets: list[str]
    parsed_sheets_summary: list[dict[str, Any]]
    report_mode: str  # "spending" | "general_report"
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
    file_path: Path | str | list[Path | str]
    sheet_name: str | None = None
    currency_override: str | None = None
    df: pd.DataFrame = field(init=False)
    income_df: pd.DataFrame = field(init=False)
    sheet_tables: list[SheetTableProfile] = field(init=False)
    meta: WorkbookMetadata = field(init=False)

    def __post_init__(self) -> None:
        if isinstance(self.file_path, (list, tuple)):
            self.file_paths = [Path(p) for p in self.file_path]
        else:
            self.file_paths = [Path(self.file_path)]
        for p in self.file_paths:
            if not p.exists():
                raise FileNotFoundError(f"File not found: {p}")
        self.file_path = self.file_paths[0]
        self._load_and_normalize()

    def _read_workbook_sheets(self, path: Path) -> list[tuple[str, pd.DataFrame, list[str]]]:
        """Return list of (sheet_name, raw_df, format_strings) for every target sheet."""
        suffix = path.suffix.lower()
        if suffix == ".csv":
            raw = pd.read_csv(path, header=None, dtype=object)
            return [("CSV", raw, [])]

        if suffix not in {".xlsx", ".xlsm", ".xls"}:
            raise ValueError(f"Unsupported file format: {suffix}. Expected .xlsx, .xlsm, .xls, or .csv")

        results: list[tuple[str, pd.DataFrame, list[str]]] = []
        with pd.ExcelFile(path) as xl:
            all_sheets = list(xl.sheet_names)
            if self.sheet_name:
                if self.sheet_name not in all_sheets:
                    raise ValueError(f"Sheet '{self.sheet_name}' not found in {path.name}. Available sheets: {all_sheets}")
                target_sheets = [self.sheet_name]
            else:
                target_sheets = all_sheets

            # Collect Excel cell number formats when available (.xlsx/.xlsm)
            sheet_formats: dict[str, list[str]] = {}
            if suffix in {".xlsx", ".xlsm"}:
                try:
                    wb_fmt = openpyxl.load_workbook(path, read_only=True, data_only=False)
                    for s in target_sheets:
                        fmts: list[str] = []
                        ws_fmt = wb_fmt[s]
                        for row_cells in ws_fmt.iter_rows(min_row=1, max_row=min(30, ws_fmt.max_row or 30)):
                            for cell in row_cells:
                                if cell.number_format and cell.number_format != "General":
                                    fmts.append(cell.number_format)
                        sheet_formats[s] = fmts
                    wb_fmt.close()
                except Exception:
                    pass

            for s in target_sheets:
                raw = pd.read_excel(xl, sheet_name=s, header=None, dtype=object)
                results.append((s, raw, sheet_formats.get(s, [])))
        return results

    def _profile_general_table(
        self,
        file_name: str,
        sheet_name: str,
        sheet_kind: str,
        data: pd.DataFrame,
        currency: str,
        title_notes: list[str],
    ) -> SheetTableProfile:
        """Build column/metric statistics for any arbitrary sheet (used for general reports + RAG/CAG)."""
        clean_df = data.copy()
        numeric_summary: list[dict[str, Any]] = []
        numeric_cols: list[str] = []

        for col in clean_df.columns:
            parsed_nums = clean_df[col].apply(_parse_number)
            valid_count = int(parsed_nums.notna().sum())
            if valid_count >= max(2, int(len(clean_df) * 0.4)) and valid_count > 0:
                clean_df[f"__num_{col}"] = parsed_nums
                numeric_cols.append(col)
                s = parsed_nums.dropna()
                numeric_summary.append({
                    "column": str(col),
                    "sum": round(float(s.sum()), 2),
                    "mean": round(float(s.mean()), 2),
                    "min": round(float(s.min()), 2),
                    "max": round(float(s.max()), 2),
                    "non_null_count": valid_count,
                })

        # Pick best categorical label column + primary numeric column for group breakdown
        top_group: list[dict[str, Any]] = []
        if numeric_cols:
            # Prefer amount/total/sales/revenue/spend column as primary metric
            metric_col = _match_column(numeric_cols, AMOUNT_KEYWORDS + ("sales", "revenue", "profit", "budget", "cost")) or numeric_cols[0]
            text_cols = [c for c in clean_df.columns if not str(c).startswith("__num_") and c not in numeric_cols]
            label_col = _match_column(text_cols, CATEGORY_KEYWORDS + DESC_KEYWORDS) or (text_cols[0] if text_cols else None)
            if label_col:
                grp = (
                    clean_df.groupby(clean_df[label_col].fillna("Unspecified").astype(str))[f"__num_{metric_col}"]
                    .agg(["sum", "count", "mean"])
                    .sort_values("sum", ascending=False)
                )
                total_m = float(grp["sum"].sum()) or 1.0
                for k, r in grp.head(10).iterrows():
                    if str(k).strip().lower() in {"total", "grand total", "subtotal"}:
                        continue
                    top_group.append({
                        "group": str(k),
                        "label_column": str(label_col),
                        "metric_column": str(metric_col),
                        "total": round(float(r["sum"]), 2),
                        "share_percent": round(float(r["sum"]) / total_m * 100, 1),
                        "count": int(r["count"]),
                    })

        display_df = data.copy()
        return SheetTableProfile(
            file_name=file_name,
            sheet_name=sheet_name,
            sheet_kind=sheet_kind,
            columns=list(display_df.columns),
            row_count=len(display_df),
            currency=currency,
            dataframe=display_df,
            numeric_summary=numeric_summary,
            top_group_breakdown=top_group,
            title_notes=title_notes,
        )

    def _parse_single_sheet(
        self,
        file_name: str,
        sheet_name: str,
        raw: pd.DataFrame,
        fmt_strings: list[str],
    ) -> tuple[str, list[dict[str, Any]], list[dict[str, Any]], SheetTableProfile | None, dict[str, Any]]:
        raw = raw.dropna(how="all").reset_index(drop=True)
        if raw.empty or len(raw) < 2:
            return "empty", [], [], None, {}

        header_idx = _find_best_header_row(raw)
        title_notes: list[str] = []
        if header_idx > 0:
            for r_idx in range(header_idx):
                vals = [str(v).strip() for v in raw.iloc[r_idx].dropna().tolist() if str(v).strip()]
                if vals:
                    title_notes.append(" | ".join(vals))

        headers: list[str] = []
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
        if data.empty:
            return "empty", [], [], None, {}

        cols = list(data.columns)
        sample_money_vals: list[Any] = list(title_notes) + list(fmt_strings)

        # Check if this sheet is a Monthly Matrix / Pivot table (Category rows x Month columns)
        default_year: int | None = None
        for note in title_notes + [sheet_name, file_name]:
            ym = re.search(r"\b(20\d{2}|19\d{2})\b", note)
            if ym:
                default_year = int(ym.group(1))
                break

        month_cols: list[tuple[str, pd.Timestamp]] = []
        for c in cols:
            m_ts = _parse_month_label(c, default_year=default_year)
            if m_ts is not None:
                month_cols.append((c, m_ts))

        date_col = _match_column(cols, DATE_KEYWORDS)
        used: set[str] = {date_col} if date_col else set()
        debit_col = _match_column(cols, DEBIT_KEYWORDS, exclude=used)
        credit_col = _match_column(cols, CREDIT_KEYWORDS, exclude=used | ({debit_col} if debit_col else set()))
        amount_col = _match_column(cols, AMOUNT_KEYWORDS, exclude=used)

        # Matrix detection: >= 2 month-named columns and no explicit transaction date column
        if len(month_cols) >= 2 and not date_col:
            non_month_cols = [c for c in cols if c not in {mc for mc, _ in month_cols}]
            label_col = _match_column(non_month_cols, CATEGORY_KEYWORDS + DESC_KEYWORDS) or (non_month_cols[0] if non_month_cols else None)
            if label_col:
                for mc, _ in month_cols[:4]:
                    sample_money_vals.extend(data[mc].dropna().tolist()[:15])
                cur = self.currency_override or _detect_currency_symbol(sample_money_vals)
                matrix_records: list[dict[str, Any]] = []
                for idx, row in data.iterrows():
                    cat = str(row[label_col]).strip() if pd.notna(row[label_col]) else ""
                    if not cat or cat.lower() in {"nan", "none", "total", "grand total", "subtotal", "sum"}:
                        continue
                    for mc, m_ts in month_cols:
                        val, drcr = _parse_number_and_drcr(row[mc])
                        if val is None or val == 0:
                            continue
                        if drcr == "cr":
                            continue
                        matrix_records.append({
                            "_raw_signed": -abs(val),
                            "date": m_ts,
                            "description": cat,
                            "category": cat,
                            "payment_mode": "",
                            "amount": round(abs(val), 2),
                            "source_file": file_name,
                            "source_sheet": sheet_name,
                            "row_num": header_idx + 2 + idx,
                        })
                if matrix_records:
                    profile = self._profile_general_table(file_name, sheet_name, "matrix", data, cur, title_notes)
                    meta_hint = {
                        "currency": cur,
                        "date_col": "Unpivoted Month Columns",
                        "amount_source": ", ".join(mc for mc, _ in month_cols[:3]) + ("..." if len(month_cols) > 3 else ""),
                        "category_col": label_col,
                        "desc_col": label_col,
                        "payment_col": None,
                        "rows_loaded": len(data),
                    }
                    return "matrix", matrix_records, [], profile, meta_hint

        # Transaction Log Detection
        tab_month_ts = _parse_month_label(sheet_name, default_year=default_year)
        if not date_col:
            best_col, best_valid = None, 0
            for c in cols:
                parsed = _smart_parse_dates(data[c])
                valid = int(parsed.notna().sum())
                if valid > best_valid:
                    best_valid = valid
                    best_col = c
            if best_col and best_valid >= max(2, int(len(data) * 0.25)):
                date_col = best_col
                used.add(date_col)

        # Re-detect amount columns excluding date_col
        used = {date_col} if date_col else set()
        debit_col = _match_column(cols, DEBIT_KEYWORDS, exclude=used)
        credit_col = _match_column(cols, CREDIT_KEYWORDS, exclude=used | ({debit_col} if debit_col else set()))
        amount_col = _match_column(cols, AMOUNT_KEYWORDS, exclude=used)
        type_col = _match_column(cols, TYPE_KEYWORDS, exclude=used | ({amount_col} if amount_col else set()))
        category_col = _match_column(
            cols,
            CATEGORY_KEYWORDS,
            exclude=used | ({type_col} if type_col and _normalize_col(type_col) in ("dr/cr", "cr/dr", "transaction type", "txn type") else set()),
        )
        if category_col == type_col and category_col is not None:
            unique_vals = {_normalize_col(v) for v in data[category_col].dropna().unique()[:20]}
            if unique_vals.issubset({"debit", "credit", "dr", "cr", "expense", "income", "inflow", "outflow"}):
                type_col = category_col
                category_col = None

        desc_col = _match_column(cols, DESC_KEYWORDS, exclude=used | ({category_col} if category_col else set()))
        payment_col = _match_column(cols, PAYMENT_KEYWORDS, exclude=used | ({category_col, desc_col} - {None}))

        for c in (amount_col, debit_col, credit_col):
            if c and c in data.columns:
                sample_money_vals.extend(data[c].dropna().tolist()[:50])
        cur = self.currency_override or _detect_currency_symbol(sample_money_vals)

        if (date_col or tab_month_ts is not None) and (amount_col or debit_col):
            parsed_dates = _smart_parse_dates(data[date_col]) if date_col else pd.Series([tab_month_ts] * len(data))
            tx_records: list[dict[str, Any]] = []
            inc_records: list[dict[str, Any]] = []

            for idx, row in data.iterrows():
                dt = parsed_dates.iloc[idx]
                if pd.isna(dt):
                    if tab_month_ts is not None:
                        dt = tab_month_ts
                    else:
                        continue

                desc = str(row[desc_col]).strip() if desc_col and pd.notna(row[desc_col]) else "Unspecified"
                if desc.lower() in {"nan", "none", ""}:
                    desc = "Unspecified"
                if desc.lower() in {"total", "grand total", "subtotal", "closing balance", "opening balance"}:
                    continue

                raw_cat = str(row[category_col]).strip() if category_col and pd.notna(row[category_col]) else ""
                if raw_cat.lower() in {"nan", "none", ""}:
                    raw_cat = ""
                if raw_cat.lower() in {"total", "grand total", "subtotal"}:
                    continue

                pay_mode = str(row[payment_col]).strip() if payment_col and pd.notna(row[payment_col]) else ""
                if pay_mode.lower() in {"nan", "none"}:
                    pay_mode = ""

                row_num = header_idx + 2 + idx
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
                    val, inline_drcr = _parse_number_and_drcr(row[amount_col])
                    if val is None or val == 0:
                        continue
                    if inline_drcr == "cr":
                        income_amt = abs(val)
                    elif inline_drcr == "dr":
                        expense_amt = abs(val)
                    elif type_col and pd.notna(row[type_col]):
                        t_norm = _normalize_col(row[type_col])
                        if any(k in t_norm for k in ("credit", "income", "cr", "salary", "refund", "deposit", "inflow", "revenue")):
                            income_amt = abs(val)
                        else:
                            expense_amt = abs(val)
                    else:
                        tx_records.append({
                            "_raw_signed": val,
                            "date": dt,
                            "description": desc,
                            "category": raw_cat or _auto_categorize(desc),
                            "payment_mode": pay_mode,
                            "amount": abs(val),
                            "source_file": file_name,
                            "source_sheet": sheet_name,
                            "row_num": row_num,
                        })
                        continue

                if expense_amt and expense_amt > 0:
                    tx_records.append({
                        "_raw_signed": -expense_amt,
                        "date": dt,
                        "description": desc,
                        "category": raw_cat or _auto_categorize(desc),
                        "payment_mode": pay_mode,
                        "amount": round(expense_amt, 2),
                        "source_file": file_name,
                        "source_sheet": sheet_name,
                        "row_num": row_num,
                    })
                elif income_amt and income_amt > 0:
                    inc_records.append({
                        "date": dt,
                        "description": desc,
                        "category": raw_cat or "Income",
                        "payment_mode": pay_mode,
                        "amount": round(income_amt, 2),
                        "source_file": file_name,
                        "source_sheet": sheet_name,
                        "row_num": row_num,
                    })

            if tx_records and not (debit_col and credit_col) and not type_col:
                neg_count = sum(1 for r in tx_records if r["_raw_signed"] < 0)
                pos_count = sum(1 for r in tx_records if r["_raw_signed"] > 0)
                if neg_count > 0 and pos_count > 0:
                    filtered_tx = []
                    for r in tx_records:
                        if r["_raw_signed"] < 0:
                            r["amount"] = round(abs(r["_raw_signed"]), 2)
                            filtered_tx.append(r)
                        else:
                            inc_records.append({
                                "date": r["date"],
                                "description": r["description"],
                                "category": r["category"],
                                "payment_mode": r["payment_mode"],
                                "amount": round(abs(r["_raw_signed"]), 2),
                                "source_file": r["source_file"],
                                "source_sheet": r["source_sheet"],
                                "row_num": r["row_num"],
                            })
                    tx_records = filtered_tx

            if tx_records:
                amount_source = f"{debit_col}" + (f" / {credit_col}" if credit_col else "") if (debit_col and (not amount_col or debit_col != amount_col)) else str(amount_col)
                profile = self._profile_general_table(file_name, sheet_name, "transactions", data, cur, title_notes)
                meta_hint = {
                    "currency": cur,
                    "date_col": date_col or f"Inferred from sheet '{sheet_name}'",
                    "amount_source": amount_source,
                    "category_col": category_col,
                    "desc_col": desc_col,
                    "payment_col": payment_col,
                    "rows_loaded": len(data),
                }
                return "transactions", tx_records, inc_records, profile, meta_hint

        # Fallback: General tabular sheet (no date or not an expense log)
        profile = self._profile_general_table(file_name, sheet_name, "general", data, cur, title_notes)
        return "general", [], [], profile, {"currency": cur, "rows_loaded": len(data)}

    def _load_and_normalize(self) -> None:
        all_tx_records: list[dict[str, Any]] = []
        all_matrix_records: list[dict[str, Any]] = []
        non_summary_matrix_records: list[dict[str, Any]] = []
        all_income_records: list[dict[str, Any]] = []
        self.sheet_tables = []
        available_sheets: list[str] = []
        parsed_sheets_summary: list[dict[str, Any]] = []

        first_hint: dict[str, Any] = {}
        total_rows_loaded = 0

        for path in self.file_paths:
            sheets = self._read_workbook_sheets(path)
            for s_name, raw_df, fmts in sheets:
                available_sheets.append(f"{path.name}:{s_name}" if len(self.file_paths) > 1 else s_name)
                kind, recs, incs, profile, hint = self._parse_single_sheet(path.name, s_name, raw_df, fmts)
                if kind == "empty":
                    continue
                total_rows_loaded += hint.get("rows_loaded", 0)
                if profile:
                    self.sheet_tables.append(profile)
                parsed_sheets_summary.append({
                    "file": path.name,
                    "sheet": s_name,
                    "kind": kind,
                    "rows": profile.row_count if profile else 0,
                    "columns": profile.columns if profile else [],
                })
                if not first_hint and kind in {"transactions", "matrix"}:
                    first_hint = hint
                elif not first_hint:
                    first_hint = hint

                if kind == "transactions":
                    all_tx_records.extend(recs)
                    all_income_records.extend(incs)
                elif kind == "matrix":
                    all_matrix_records.extend(recs)
                    if not any(kw in s_name.lower() for kw in ("summary", "overview", "dashboard", "pivot", "report")):
                        non_summary_matrix_records.extend(recs)

        if not self.sheet_tables:
            raise ValueError("No readable data tables found in the provided spreadsheet(s).")

        # Avoid double-counting a Summary matrix tab when detailed transaction tabs also exist
        if all_tx_records and non_summary_matrix_records:
            chosen_records = all_tx_records + non_summary_matrix_records
        elif all_tx_records:
            chosen_records = all_tx_records
        else:
            chosen_records = all_matrix_records

        currency = self.currency_override or first_hint.get("currency") or "$"

        if chosen_records:
            df = pd.DataFrame(chosen_records).drop(columns=["_raw_signed"])
            df["date"] = pd.to_datetime(df["date"])
            df = df.sort_values("date").reset_index(drop=True)
            df["month"] = df["date"].dt.to_period("M").astype(str)
            df["calendar_month"] = df["date"].dt.month_name()
            df["calendar_month_num"] = df["date"].dt.month
            df["day_of_week"] = df["date"].dt.day_name()
            df["is_weekend"] = df["date"].dt.dayofweek >= 5
            df["day_of_month"] = df["date"].dt.day
            report_mode = "spending"
            chosen_sheet_label = (
                parsed_sheets_summary[0]["sheet"]
                if len(parsed_sheets_summary) == 1
                else f"{len(parsed_sheets_summary)} sheets"
            )
        else:
            # General non-date tabular report mode
            df = pd.DataFrame(columns=["date", "description", "category", "payment_mode", "amount", "month", "calendar_month", "calendar_month_num", "day_of_week", "is_weekend", "day_of_month", "source_file", "source_sheet", "row_num"])
            report_mode = "general_report"
            chosen_sheet_label = (
                parsed_sheets_summary[0]["sheet"]
                if len(parsed_sheets_summary) == 1
                else f"{len(parsed_sheets_summary)} sheets"
            )

        self.df = df
        self.income_df = pd.DataFrame(all_income_records)
        self.meta = WorkbookMetadata(
            file_path=", ".join(str(p) for p in self.file_paths),
            files_loaded=[p.name for p in self.file_paths],
            sheet_name=chosen_sheet_label,
            available_sheets=available_sheets,
            parsed_sheets_summary=parsed_sheets_summary,
            report_mode=report_mode,
            currency=currency,
            date_column=first_hint.get("date_col", "N/A"),
            amount_source=first_hint.get("amount_source", "N/A"),
            category_column=first_hint.get("category_col"),
            description_column=first_hint.get("desc_col"),
            payment_column=first_hint.get("payment_col"),
            rows_loaded=total_rows_loaded,
            expense_rows=len(df),
            income_rows_excluded=len(all_income_records),
        )

    def get_sheet_profiles(self) -> list[dict[str, Any]]:
        """Return structured summaries of every sheet across all loaded files."""
        return [
            {
                "file": t.file_name,
                "sheet": t.sheet_name,
                "kind": t.sheet_kind,
                "row_count": t.row_count,
                "columns": t.columns,
                "title_notes": t.title_notes,
                "numeric_summary": t.numeric_summary,
                "top_group_breakdown": t.top_group_breakdown,
            }
            for t in self.sheet_tables
        ]

    def get_overview(self) -> dict[str, Any]:
        if self.meta.report_mode == "general_report":
            profiles = self.get_sheet_profiles()
            primary = profiles[0]
            top_num = primary["numeric_summary"][0] if primary["numeric_summary"] else None
            return {
                "file": self.meta.file_path,
                "files_loaded": self.meta.files_loaded,
                "sheet": self.meta.sheet_name,
                "report_mode": "general_report",
                "currency": self.meta.currency,
                "date_range": {"start": "N/A", "end": "N/A", "num_months": 0},
                "total_spend": top_num["sum"] if top_num else 0.0,
                "total_income_detected": 0.0,
                "net_cashflow": None,
                "transaction_count": sum(p["row_count"] for p in profiles),
                "average_transaction": top_num["mean"] if top_num else 0.0,
                "median_transaction": 0.0,
                "average_monthly_spend": 0.0,
                "median_monthly_spend": 0.0,
                "largest_single_expense": {"date": "N/A", "description": primary["sheet"], "category": top_num["column"] if top_num else "N/A", "amount": top_num["max"] if top_num else 0.0},
                "detected_columns": {"columns": primary["columns"]},
                "sheets_parsed": self.meta.parsed_sheets_summary,
            }

        monthly = self.df.groupby("month")["amount"].sum()
        total_spend = float(self.df["amount"].sum())
        total_income = float(self.income_df["amount"].sum()) if not self.income_df.empty else 0.0
        return {
            "file": self.meta.file_path,
            "files_loaded": self.meta.files_loaded,
            "sheet": self.meta.sheet_name,
            "report_mode": "spending",
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
            "sheets_parsed": self.meta.parsed_sheets_summary,
        }

    def get_where_you_spend_more(self, top_n: int = 10) -> dict[str, Any]:
        if self.meta.report_mode == "general_report":
            groups = []
            for t in self.sheet_tables:
                for g in t.top_group_breakdown:
                    groups.append({
                        "category": g["group"],
                        "total_spend": g["total"],
                        "share_percent": g["share_percent"],
                        "transaction_count": g["count"],
                        "average_transaction": round(g["total"] / max(1, g["count"]), 2),
                        "max_transaction": g["total"],
                        "top_examples": [],
                    })
            groups.sort(key=lambda x: x["total_spend"], reverse=True)
            groups = groups[:top_n]
            return {
                "top_category": groups[0] if groups else {"category": "N/A", "total_spend": 0.0, "share_percent": 0.0, "transaction_count": 0},
                "categories": groups,
                "top_merchants_by_spend": [],
                "most_frequent_merchants": [],
                "payment_modes": [],
            }

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
        if self.meta.report_mode == "general_report" or self.df.empty:
            return {
                "monthly_average": 0.0,
                "peak_month": {"month": "N/A", "total_spend": 0.0, "vs_monthly_avg_percent": 0.0, "top_category_in_month": "N/A"},
                "lowest_month": {"month": "N/A", "total_spend": 0.0, "vs_monthly_avg_percent": 0.0, "top_category_in_month": "N/A"},
                "top_3_highest_months": [],
                "months_chronological": [],
            }

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
        if self.df.empty:
            return []
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
        if self.meta.report_mode == "general_report" or self.df.empty:
            return {
                "trend_direction": "not_applicable",
                "monthly_slope_estimate": 0.0,
                "recent_3_month_avg": 0.0,
                "estimated_recurring_monthly_commitments": 0.0,
                "recurring_items": [],
                "next_3_months_forecast": [],
                "highest_projected_upcoming_month": None,
                "highest_historical_calendar_month": None,
                "seasonal_ranking": [],
                "note": "No date/month column detected in this general report.",
            }

        m_series = self.df.groupby("month")["amount"].sum().sort_index()
        months = list(m_series.index)
        values = [float(v) for v in m_series.values]

        last_period = pd.Period(months[-1], freq="M")
        next_months = [(last_period + i).strftime("%Y-%m") for i in range(1, 4)]

        if len(values) >= 3:
            recent_3_avg = sum(values[-3:]) / 3.0
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

        return {
            "trend_direction": trend_direction,
            "monthly_slope_estimate": round(slope, 2),
            "recent_3_month_avg": round(recent_3_avg, 2),
            "estimated_recurring_monthly_commitments": monthly_fixed_estimate,
            "recurring_items": recurring[:10],
            "next_3_months_forecast": forecasts,
            "highest_projected_upcoming_month": max(forecasts, key=lambda x: x["projected_spend"]) if forecasts else None,
            "highest_historical_calendar_month": seasonal_pattern[0] if seasonal_pattern else None,
            "seasonal_ranking": seasonal_pattern[:5],
            "note": (
                "Forecast blends recent 3-month spending, overall trend slope, and historical calendar-month pattern when >=12 months exist."
            ),
        }

    def explain_why_it_happens(self, target_month: str | None = None) -> dict[str, Any]:
        if self.meta.report_mode == "general_report" or self.df.empty:
            return {
                "analyzed_month": "N/A",
                "target_month_total": 0.0,
                "other_months_avg_total": 0.0,
                "excess_spend_over_baseline": 0.0,
                "frequency_vs_ticket_size": {
                    "target_month_transactions": 0,
                    "baseline_avg_transactions": 0,
                    "target_month_avg_ticket": 0.0,
                    "baseline_avg_ticket": 0.0,
                },
                "top_category_surges": [],
                "top_transactions_in_month": [],
                "unusual_spike_transactions": [],
                "overall_spending_habits": {
                    "weekend_spend_total": 0.0,
                    "weekday_spend_total": 0.0,
                    "avg_spend_per_active_weekend_day": 0.0,
                    "avg_spend_per_active_weekday": 0.0,
                    "small_frequent_purchases": {"threshold_amount": 0.0, "count": 0, "total_spend": 0.0, "share_percent": 0.0},
                    "large_one_off_purchases": {"threshold_amount": 0.0, "count": 0, "total_spend": 0.0, "share_percent": 0.0},
                    "month_timing_split_percent": {"early_month_days_1_10": 0.0, "mid_month_days_11_20": 0.0, "late_month_days_21_31": 0.0},
                },
            }

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

        top_tx = [
            {
                "date": r["date"].strftime("%Y-%m-%d"),
                "description": str(r["description"]),
                "category": str(r["category"]),
                "amount": round(float(r["amount"]), 2),
                "share_of_month_percent": round(float(r["amount"]) / target_total * 100, 1) if target_total else 0.0,
                "source": f"{r['source_file']} | {r['source_sheet']} | Row {r['row_num']}",
            }
            for _, r in target_df.sort_values("amount", ascending=False).head(8).iterrows()
        ]

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

        weekend_spend = float(self.df[self.df["is_weekend"]]["amount"].sum())
        weekday_spend = float(self.df[~self.df["is_weekend"]]["amount"].sum())
        weekend_days = max(1, int(self.df[self.df["is_weekend"]]["date"].dt.date.nunique()))
        weekday_days = max(1, int(self.df[~self.df["is_weekend"]]["date"].dt.date.nunique()))

        small_threshold = float(self.df["amount"].median())
        small_tx = self.df[self.df["amount"] <= small_threshold]
        large_tx = self.df[self.df["amount"] > outlier_threshold]

        early = float(self.df[self.df["day_of_month"] <= 10]["amount"].sum())
        mid = float(self.df[(self.df["day_of_month"] >= 11) & (self.df["day_of_month"] <= 20)]["amount"].sum())
        late = float(self.df[self.df["day_of_month"] >= 21]["amount"].sum())
        total_all = float(self.df["amount"].sum()) or 1.0

        return {
            "analyzed_month": target_month,
            "target_month_total": round(target_total, 2),
            "other_months_avg_total": round(other_monthly_avg, 2),
            "excess_spend_over_baseline": round(excess_over_baseline, 2),
            "frequency_vs_ticket_size": {
                "target_month_transactions": int(len(target_df)),
                "baseline_avg_transactions": round(len(other_df) / num_other_months, 1) if not other_df.empty else float(len(target_df)),
                "target_month_avg_ticket": round(float(target_df["amount"].mean()), 2),
                "baseline_avg_ticket": round(float(other_df["amount"].mean()), 2) if not other_df.empty else round(float(target_df["amount"].mean()), 2),
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
        if self.df.empty:
            return {"matching_count": 0, "matching_total_spend": 0.0, "showing_top": 0, "transactions": []}

        sub = self.df.copy()
        if month:
            sub = sub[sub["month"] == month]
        if category:
            sub = sub[sub["category"].str.lower().str.contains(category.lower(), na=False)]
        if keyword:
            mask = (
                sub["description"].str.lower().str.contains(keyword.lower(), na=False)
                | sub["category"].str.lower().str.contains(keyword.lower(), na=False)
                | sub["source_sheet"].str.lower().str.contains(keyword.lower(), na=False)
                | sub["source_file"].str.lower().str.contains(keyword.lower(), na=False)
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
                "source": f"{r['source_file']} | {r['source_sheet']} | Row {r['row_num']}",
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
            "sheet_profiles": self.get_sheet_profiles(),
            "where_you_spend_more": self.get_where_you_spend_more(top_n=10),
            "monthly_spending": self.get_monthly_spending(),
            "future_month_forecast": self.get_future_month_forecast(),
            "why_peak_month_happened": self.explain_why_it_happens(),
        }
