"""
CLI entrypoint for the Ollama Excel Spending Analyzer Agent.

Usage:
    python main.py path/to/expenses.xlsx
    python main.py --demo --chat
"""

from __future__ import annotations

import argparse
import sys
from pathlib import Path

from rich.console import Console
from rich.markdown import Markdown
from rich.panel import Panel
from rich.table import Table

from create_sample_excel import generate_sample_excel
from excel_analyzer import ExpenseAnalyzer
from ollama_agent import DEFAULT_OLLAMA_HOST, OllamaSpendingAgent, list_ollama_models
from report_generator import save_reports


console = Console()


def _print_quick_tables(bundle: dict) -> None:
    ov = bundle["overview"]
    cur = ov["currency"]
    where = bundle["where_you_spend_more"]
    monthly = bundle["monthly_spending"]
    forecast = bundle["future_month_forecast"]

    # Category table
    cat_tbl = Table(title="Top Spending Categories", show_lines=False)
    cat_tbl.add_column("Category", style="cyan")
    cat_tbl.add_column("Total Spend", justify="right", style="bold white")
    cat_tbl.add_column("Share", justify="right")
    cat_tbl.add_column("Txns", justify="right")
    for c in where["categories"][:7]:
        cat_tbl.add_row(c["category"], f"{cur}{c['total_spend']:,.2f}", f"{c['share_percent']}%", str(c["transaction_count"]))

    # Monthly table
    mon_tbl = Table(title="Monthly Spending & Forecast", show_lines=False)
    mon_tbl.add_column("Month", style="cyan")
    mon_tbl.add_column("Spend", justify="right", style="bold white")
    mon_tbl.add_column("vs Avg", justify="right")
    mon_tbl.add_column("Top Category")
    peak_m = monthly["peak_month"]["month"]
    for m in monthly["months_chronological"]:
        tag = " [red](PEAK)[/red]" if m["month"] == peak_m else ""
        mon_tbl.add_row(
            m["month"] + tag,
            f"{cur}{m['total_spend']:,.2f}",
            f"{m['vs_monthly_avg_percent']:+.1f}%",
            m["top_category_in_month"],
        )
    for f in forecast["next_3_months_forecast"]:
        mon_tbl.add_row(
            f"{f['month']} [magenta](Forecast)[/magenta]",
            f"{cur}{f['projected_spend']:,.2f}",
            f"{f['vs_historical_avg_percent']:+.1f}%",
            f"Trend: {forecast['trend_direction']}",
        )

    console.print(cat_tbl)
    console.print(mon_tbl)


def main(argv: list[str] | None = None) -> int:
    parser = argparse.ArgumentParser(
        description="Ollama AI Agent to analyze an Excel expense sheet, report where & which month you spend more, and explain why."
    )
    parser.add_argument("excel_file", nargs="?", help="Path to .xlsx / .xls / .csv expense file")
    parser.add_argument("--demo", action="store_true", help="Generate and analyze sample_expenses.xlsx")
    parser.add_argument("--sheet", help="Specific Excel sheet name to analyze")
    parser.add_argument("--currency", help="Override detected currency symbol (e.g. $, ₹, €)")
    parser.add_argument("--model", help="Ollama model name (defaults to best installed model, e.g. qwen2.5:latest)")
    parser.add_argument("--host", default=DEFAULT_OLLAMA_HOST, help="Ollama host URL (default: http://127.0.0.1:11434)")
    parser.add_argument("--out", default="reports", help="Directory to save Markdown & HTML reports")
    parser.add_argument("--ask", help="Ask a specific follow-up question after generating the report")
    parser.add_argument("--chat", action="store_true", help="Start interactive Q&A chat after the report")
    parser.add_argument("--deep", action="store_true", help="Use multi-turn autonomous tool calling for initial report (slower than default fast mode)")
    parser.add_argument("--list-models", action="store_true", help="List available Ollama models and exit")
    args = parser.parse_args(argv)

    if args.list_models:
        try:
            models = list_ollama_models(args.host)
        except Exception as exc:
            console.print(f"[red]Could not connect to Ollama at {args.host}: {exc}[/red]")
            return 1
        console.print("[bold green]Available Ollama models:[/bold green]")
        for m in models:
            console.print(f"  - {m}")
        return 0

    if args.demo:
        excel_path = generate_sample_excel("sample_expenses.xlsx")
        console.print(f"[green]Generated demo workbook:[/green] {excel_path.resolve()}")
    elif args.excel_file:
        excel_path = Path(args.excel_file)
    else:
        parser.print_help()
        console.print("\n[yellow]Tip:[/yellow] Pass your Excel path (`python main.py expenses.xlsx`) or try `--demo`.")
        return 1

    console.print(Panel.fit(f"[bold cyan]Ollama Excel Spending Agent[/bold cyan]\nFile: {excel_path}", border_style="cyan"))

    try:
        analyzer = ExpenseAnalyzer(excel_path, sheet_name=args.sheet, currency_override=args.currency)
    except Exception as exc:
        console.print(f"[red]Failed to parse Excel file:[/red] {exc}")
        return 1

    bundle = analyzer.build_full_analytics_bundle()
    _print_quick_tables(bundle)

    def _log_tool(name: str, tool_args: dict) -> None:
        arg_str = f"({tool_args})" if tool_args else "()"
        console.print(f"  [dim]-> Agent calling tool:[/dim] [yellow]{name}{arg_str}[/yellow]")

    try:
        agent = OllamaSpendingAgent(
            analyzer=analyzer,
            model=args.model,
            host=args.host,
            on_tool_call=_log_tool,
        )
    except Exception as exc:
        console.print(f"[red]Ollama initialization failed:[/red] {exc}")
        console.print("[yellow]Make sure Ollama is running (`ollama serve`) and at least one model is pulled.[/yellow]")
        return 1

    console.print(f"\n[bold green]Running Ollama Agent ({agent.model})...[/bold green]")
    with console.status("[bold cyan]Agent is inspecting your spreadsheet and writing the report...[/bold cyan]"):
        report_md, tools_used = agent.generate_full_report(fast_mode=not args.deep)

    console.print(Panel(Markdown(report_md), title=f"Spending Report ({agent.model})", border_style="green"))

    stem = excel_path.stem
    md_path, html_path = save_reports(bundle, report_md, agent.model, args.out, stem=stem)
    console.print(f"[bold cyan]Saved Markdown report:[/bold cyan] {md_path.resolve()}")
    console.print(f"[bold cyan]Saved HTML dashboard:[/bold cyan]  {html_path.resolve()}")
    console.print(f"[dim]Tools used by agent: {', '.join(tools_used)}[/dim]")

    if args.ask:
        console.print(f"\n[bold yellow]Question:[/bold yellow] {args.ask}")
        with console.status("[cyan]Thinking...[/cyan]"):
            ans = agent.ask(args.ask)
        console.print(Panel(Markdown(ans), title="Answer", border_style="blue"))

    if args.chat:
        console.print("\n[bold cyan]Interactive Q&A Mode[/bold cyan] (type 'exit' or 'quit' to stop)")
        while True:
            try:
                q = console.input("\n[bold green]Ask about your Excel sheet > [/bold green]").strip()
            except (EOFError, KeyboardInterrupt):
                break
            if not q or q.lower() in {"exit", "quit", "q"}:
                break
            with console.status("[cyan]Analyzing...[/cyan]"):
                ans = agent.ask(q)
            console.print(Markdown(ans))

    return 0


if __name__ == "__main__":
    raise SystemExit(main())
