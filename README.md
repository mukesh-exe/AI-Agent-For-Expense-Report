# Ollama Excel Spending Analyzer Agent

An AI agent powered by **Ollama** that reads your Excel (`.xlsx`, `.xls`, `.csv`) expense sheets, computes exact analytics, generates Markdown & interactive HTML reports, and explains:
1. **Your full spending report**
2. **Where you spend more** (top categories, merchants, and frequent small expenses)
3. **Which month you spend / will spend more** (historical peak month + 3-month forecast & seasonal trend)
4. **Why it happens** (surging categories vs baseline, spike transactions, frequency vs ticket-size drivers, weekend/timing habits)

## Setup

Create/activate the virtual environment (already configured in `.venv`):

```powershell
.\.venv\Scripts\Activate.ps1
```

Make sure **Ollama** is running locally with at least one model (e.g., `qwen2.5:latest`).

## Web UI (Recommended)

Double-click `start_ui.bat` or run:

```powershell
.\.venv\Scripts\python.exe web_ui.py
```

This opens `http://127.0.0.1:8501` in your browser with drag-and-drop Excel upload, Ollama model selector, interactive charts, downloadable reports, and live AI chat.

## CLI Usage

Analyze your own Excel file:

```powershell
.\.venv\Scripts\python.exe main.py "C:\path\to\your_expenses.xlsx"
```

Try with the built-in demo Excel sheet + interactive Q&A:

```powershell
.\.venv\Scripts\python.exe main.py --demo --chat
```

### Useful Options
- `--model qwen2.5:latest` — Choose a specific Ollama model (`--list-models` lists installed models)
- `--sheet "Sheet1"` — Specify a sheet name if your workbook has multiple tabs
- `--currency "₹"` — Override the auto-detected currency symbol
- `--out reports` — Directory for saved `<name>_report.md` and `<name>_report.html`
- `--ask "Why did October spike?"` — Ask a one-off follow-up question
- `--chat` — Enter interactive follow-up Q&A mode after the report
