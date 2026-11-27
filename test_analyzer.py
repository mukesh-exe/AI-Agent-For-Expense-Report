"""Comprehensive tests for Universal Excel parsing, Report generation, and Hybrid CAG+RAG."""

from __future__ import annotations

import tempfile
from pathlib import Path

import openpyxl

from create_sample_excel import generate_sample_excel
from excel_analyzer import ExpenseAnalyzer
from rag_cag_engine import ReportKnowledgeBase
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


def test_bank_statement_debit_credit_and_inline_drcr() -> None:
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

        # Inline CR/DR suffix test
        path2 = Path(tmp) / "bank_drcr.csv"
        path2.write_text(
            "Date,Description,Amount\n"
            "2026-07-01,Salary,4000 CR\n"
            "2026-07-04,Rent,1200 DR\n"
            "2026-07-08,Blinkit Groceries,150 DR\n",
            encoding="utf-8",
        )
        a2 = ExpenseAnalyzer(path2, currency_override="₹")
        ov2 = a2.get_overview()
        assert ov2["total_spend"] == 1350.0
        assert ov2["total_income_detected"] == 4000.0


def test_matrix_pivot_and_multi_tab_monthly() -> None:
    with tempfile.TemporaryDirectory() as tmp:
        # 1. Category x Month matrix budget
        matrix_path = Path(tmp) / "budget_matrix_2026.xlsx"
        wb = openpyxl.Workbook()
        ws = wb.active
        ws.title = "Annual Budget 2026"
        ws.append(["Category", "Jan", "Feb", "Mar", "Total"])
        ws.append(["Housing & Rent", 20000, 20000, 20000, 60000])
        ws.append(["Food & Dining", 4500, 5200, 9800, 19500])
        ws.append(["Travel", 1000, 500, 18000, 19500])
        ws.append(["Total", 25500, 25700, 47800, 99000])
        wb.save(matrix_path)

        a_mat = ExpenseAnalyzer(matrix_path, currency_override="₹")
        ov_m = a_mat.get_overview()
        assert ov_m["total_spend"] == 99000.0
        assert a_mat.get_monthly_spending()["peak_month"]["month"] == "2026-03"

        # 2. One-tab-per-month workbook (Jan, Feb)
        tabs_path = Path(tmp) / "monthly_tabs_2026.xlsx"
        wb2 = openpyxl.Workbook()
        ws_jan = wb2.active
        ws_jan.title = "Jan 2026"
        ws_jan.append(["Item", "Cost"])
        ws_jan.append(["Rent", 1500])
        ws_jan.append(["Coffee", 50])
        ws_feb = wb2.create_sheet("Feb 2026")
        ws_feb.append(["Item", "Cost"])
        ws_feb.append(["Rent", 1500])
        ws_feb.append(["Flight Ticket", 2200])
        wb2.save(tabs_path)

        a_tabs = ExpenseAnalyzer(tabs_path, currency_override="$")
        assert a_tabs.get_overview()["total_spend"] == 5250.0
        assert a_tabs.get_monthly_spending()["peak_month"]["month"] == "2026-02"


def test_general_non_date_report_and_hybrid_cag_rag() -> None:
    with tempfile.TemporaryDirectory() as tmp:
        sales_path = Path(tmp) / "regional_sales.xlsx"
        wb = openpyxl.Workbook()
        ws1 = wb.active
        ws1.title = "Regional Sales"
        ws1.append(["Region", "Product", "Units", "Revenue"])
        ws1.append(["North", "Solar Panel X", 120, 480000])
        ws1.append(["South", "Battery Pack Z", 85, 255000])
        ws1.append(["West", "Inverter Pro", 40, 120000])

        ws2 = wb.create_sheet("Notes & Risks")
        ws2.append(["Topic", "Status", "Notes"])
        ws2.append(["Supply Chain", "Delayed", "Battery Pack Z shipment delayed by 2 weeks due to port congestion"])
        wb.save(sales_path)

        analyzer = ExpenseAnalyzer(sales_path, currency_override="₹")
        assert analyzer.meta.report_mode == "general_report"
        bundle = analyzer.build_full_analytics_bundle()
        md_p, html_p = save_reports(bundle, "## Regional Summary\nNorth leads with ₹480,000.", "qwen2.5:latest", tmp, stem="sales")
        assert md_p.exists() and html_p.exists()

        kb = ReportKnowledgeBase(analyzer)
        kb.index_generated_report("## Executive Overview\nNorth leads revenue with ₹480,000 from Solar Panel X.")
        ret = kb.retrieve("Why was Battery Pack Z delayed?", top_k=5)
        assert "CAG" in ret["retrieval_mode"]
        assert "Regional Sales" in ret["cag_context"]
        joined_rag = "\n".join(c["text"] for c in ret["rag_chunks"])
        assert "port congestion" in joined_rag


if __name__ == "__main__":
    test_sample_workbook_analytics()
    test_bank_statement_debit_credit_and_inline_drcr()
    test_matrix_pivot_and_multi_tab_monthly()
    test_general_non_date_report_and_hybrid_cag_rag()
    print("All universal analyzer, report & CAG+RAG tests passed!")
