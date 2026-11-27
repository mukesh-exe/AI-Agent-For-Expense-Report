"""
Creates a realistic multi-month sample Excel workbook for testing and demoing
the Ollama Excel Spending Agent.
"""

from __future__ import annotations

import random
from datetime import date, timedelta
from pathlib import Path

import openpyxl
from openpyxl.styles import Alignment, Font, PatternFill


def generate_sample_excel(output_path: str | Path = "sample_expenses.xlsx") -> Path:
    out = Path(output_path)
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "Expenses_2026"

    # Title rows to test resilient header detection
    ws["A1"] = "Personal Expense & Spending Tracker (2026)"
    ws["A1"].font = Font(bold=True, size=14)
    ws["A2"] = "Currency: INR (₹) | Account: Primary Bank + Credit Card"
    ws["A2"].font = Font(italic=True, color="555555")

    headers = ["Date", "Category", "Description", "Payment Mode", "Type", "Amount"]
    header_row = 4
    fill = PatternFill(start_color="1F4E78", end_color="1F4E78", fill_type="solid")
    for col_idx, h in enumerate(headers, start=1):
        cell = ws.cell(row=header_row, column=col_idx, value=h)
        cell.font = Font(bold=True, color="FFFFFF")
        cell.fill = fill
        cell.alignment = Alignment(horizontal="center")

    rng = random.Random(42)
    rows: list[tuple[date, str, str, str, str, float]] = []

    # Generate Jan 2026 to Oct 2026 (10 months)
    for m in range(1, 11):
        # Salary (Income - should be excluded from spending!)
        rows.append((date(2026, m, 1), "Salary", "Monthly Salary Credit", "Bank Transfer", "Income", 95000.0))

        # Fixed monthly commitments
        rows.append((date(2026, m, 2), "Housing & Rent", "Apartment Monthly Rent", "Bank Transfer", "Expense", 22000.0))
        rows.append((date(2026, m, 5), "Utilities & Bills", "Fiber Internet & WiFi", "UPI", "Expense", 1199.0))
        rows.append((date(2026, m, 7), "Utilities & Bills", "Electricity Bill", "UPI", "Expense", round(rng.uniform(1800, 2900), 2)))
        rows.append((date(2026, m, 10), "Entertainment & Subscriptions", "Netflix & Spotify", "Credit Card", "Expense", 848.0))
        rows.append((date(2026, m, 15), "Health & Medical", "Gym Membership", "Credit Card", "Expense", 1500.0))

        # Regular groceries (4 times a month)
        for day in (4, 11, 18, 25):
            rows.append((
                date(2026, m, day),
                "Groceries",
                rng.choice(["Blinkit Groceries", "BigBasket Supermarket", "Local Fresh Mart"]),
                "UPI",
                "Expense",
                round(rng.uniform(1100, 2400), 2),
            ))

        # Frequent small food & coffee purchases (more on weekends)
        for d_off in range(1, 28, 3):
            dt = date(2026, m, min(28, d_off))
            if dt.weekday() >= 5:
                rows.append((dt, "Food & Dining", rng.choice(["Zomato Dinner", "Swiggy Order", "Cafe Coffee Day"]), "UPI", "Expense", round(rng.uniform(450, 1250), 2)))
            else:
                rows.append((dt, "Food & Dining", rng.choice(["Office Lunch", "Starbucks Coffee", "Bakery Snack"]), "UPI", "Expense", round(rng.uniform(180, 480), 2)))

        # Transport & Fuel
        for day in (6, 16, 24):
            rows.append((date(2026, m, day), "Transport & Fuel", rng.choice(["Shell Petrol Pump", "Uber Ride", "Metro Card Topup"]), "UPI", "Expense", round(rng.uniform(350, 1600), 2)))

        # Normal occasional shopping
        rows.append((date(2026, m, 14), "Shopping & Electronics", rng.choice(["Amazon Order", "Myntra Apparel"]), "Credit Card", "Expense", round(rng.uniform(1200, 3800), 2)))

    # Inject clear spike drivers in October 2026 (Festive & Tech Upgrade + Travel)
    rows.extend([
        (date(2026, 10, 12), "Shopping & Electronics", "Apple Store Laptop Upgrade", "Credit Card", "Expense", 54990.0),
        (date(2026, 10, 16), "Travel & Vacation", "IndiGo Return Flights", "Credit Card", "Expense", 16800.0),
        (date(2026, 10, 18), "Travel & Vacation", "Goa Resort Hotel Booking", "Credit Card", "Expense", 18500.0),
        (date(2026, 10, 20), "Shopping & Electronics", "Amazon Festive Gifts & Home Decor", "Credit Card", "Expense", 12450.0),
        (date(2026, 10, 22), "Food & Dining", "Festive Family Dinner Party", "Credit Card", "Expense", 6800.0),
    ])

    # Slight secondary bump in May 2026 (Summer AC & Weekend Trip)
    rows.extend([
        (date(2026, 5, 9), "Utilities & Bills", "Summer AC Servicing & Peak Electricity", "UPI", "Expense", 5400.0),
        (date(2026, 5, 23), "Travel & Vacation", "Weekend Hill Station Stay", "Credit Card", "Expense", 11200.0),
    ])

    rows.sort(key=lambda r: r[0])

    for idx, (dt, cat, desc, mode, tx_type, amt) in enumerate(rows, start=header_row + 1):
        ws.cell(row=idx, column=1, value=dt.strftime("%Y-%m-%d"))
        ws.cell(row=idx, column=2, value=cat)
        ws.cell(row=idx, column=3, value=desc)
        ws.cell(row=idx, column=4, value=mode)
        ws.cell(row=idx, column=5, value=tx_type)
        c = ws.cell(row=idx, column=6, value=amt)
        c.number_format = "₹#,##0.00"

    for col in ws.columns:
        max_len = max(len(str(cell.value or "")) for cell in col)
        ws.column_dimensions[col[0].column_letter].width = min(max_len + 4, 38)

    wb.save(out)
    return out


if __name__ == "__main__":
    created = generate_sample_excel()
    print(f"Sample Excel created at: {created.resolve()}")
