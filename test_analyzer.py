"""Unit tests for ExpenseAnalyzer and report generation."""

from __future__ import annotations

import tempfile
from pathlib import Path

import openpyxl

from create_sample_excel import generate_sample_excel
from excel_analyzer import ExpenseAnalyzer
from report_generator import save_reports


def test_sample_workbook_analytics() -> None:
    with tempfile.TemporaryDirectory() as tmp:
        xlsx_path = generate_sample_excel(Path(tmp) / "sample.xlsx")
        analyzer = ExpenseAnalyzer(xlsx_path)

        overview = analyzer.get_overview()
        assert overview["currency"] == "₹"
        assert overview["date_range"]["num_months"] == 10
        assert overview["total_income_detected"] == 950000.0

        where = analyzer.get_where_you_spend_more()
        assert len(where["categories"]) >= 5

        monthly = analyzer.get_monthly_spending()
        assert monthly["peak_month"]["month"] == "2026-10"

        why = analyzer.explain_why_it_happens()
        assert why["analyzed_month"] == "2026-10"
        top_surge_cats = [d["category"] for d in why["top_category_surges"][:3]]
        assert "Shopping & Electronics" in top_surge_cats
        assert "Travel & Vacation" in top_surge_cats

        forecast = analyzer.get_future_month_forecast()
        assert len(forecast["next_3_months_forecast"]) == 3
        assert len(forecast["recurring_items"]) >= 3

        bundle = analyzer.build_full_analytics_bundle()
        md_path, html_path = save_reports(bundle, "## Test Summary\n- Peak: October", "qwen2.5:latest", tmp, stem="test")
        assert md_path.exists() and html_path.exists()
        assert "2026-10" in html_path.read_text(encoding="utf-8")


def test_bank_statement_debit_credit_auto_category() -> None:
    with tempfile.TemporaryDirectory() as tmp:
        path = Path(tmp) / "bank.xlsx"
        wb = openpyxl.Workbook()
        ws = wb.active
        ws.append(["Txn Date", "Narration", "Withdrawal", "Deposit"])
        ws.append(["2026-08-01", "Salary Transfer", None, 5000])
        ws.append(["2026-08-03", "Starbucks Coffee", 6.50, None])
        ws.append(["2026-08-05", "Amazon Marketplace", 120.00, None])
        ws.append(["2026-09-10", "Rent Payment", 1500.00, None])
        ws.append(["2026-09-12", "Uber Trip", 25.00, None])
        wb.save(path)

        analyzer = ExpenseAnalyzer(path, currency_override="$")
        ov = analyzer.get_overview()
        assert ov["total_spend"] == 1651.50
        assert ov["total_income_detected"] == 5000.0
        monthly = analyzer.get_monthly_spending()
        assert monthly["peak_month"]["month"] == "2026-09"
        where = analyzer.get_where_you_spend_more()
        cats = {c["category"] for c in where["categories"]}
        assert "Housing & Rent" in cats
        assert "Food & Dining" in cats


if __name__ == "__main__":
    test_sample_workbook_analytics()
    test_bank_statement_debit_credit_auto_category()
    print("All analyzer & report tests passed!")
