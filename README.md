# Universal Ollama Excel Analyzer & Hybrid CAG+RAG Agent

An AI agent powered by **Ollama** that analyzes **all Excel/CSV reports** across single or multiple files and tabs, generates Markdown & interactive HTML dashboards, and answers any follow-up question using **Hybrid CAG (Cache-Augmented Generation) + RAG (Retrieval-Augmented Generation)**.

## Supported Report Layouts
- **Transaction Logs:** Personal expenses, bank statements (`Debit`/`Credit`, `Withdrawal`/`Deposit`, `Type`, inline `CR`/`DR`, signed numbers)
- **Matrix / Pivot Budgets:** Category rows × Month columns (`Jan`, `Feb`, `Mar`, ...), automatically unpivoted
- **Multi-Tab Monthly Workbooks:** Separate tabs per month (`Jan 2026`, `Feb 2026`, ...) combined cleanly
- **General Tabular Reports:** Sales, Inventory, P&L, or multi-sheet operational reports without date columns
- **Formats:** `.xlsx`, `.xlsm`, `.xls`, and `.csv` (single file, multiple files, or folder)

## Web UI (Recommended)

Double-click `start_ui.bat` or run:

```powershell
.\.venv\Scripts\python.exe web_ui.py
```

Open `http://127.0.0.1:8501` to upload one or multiple Excel/CSV files, choose an Ollama model (`qwen2.5-coder:3b` for fastest speed or `qwen2.5:latest` for best writing & Q&A), watch reports stream live, and chat with CAG + RAG citations.

## CLI Usage

```powershell
.\.venv\Scripts\python.exe main.py "report.xlsx" --chat
.\.venv\Scripts\python.exe main.py jan.xlsx feb.xlsx --ask "Compare January vs February"
.\.venv\Scripts\python.exe main.py --demo --chat
```
