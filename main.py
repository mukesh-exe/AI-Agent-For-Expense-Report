"""
CLI entrypoint for the Universal Ollama Excel Analyzer & Hybrid CAG+RAG Agent.

Usage:
    python main.py path/to/report.xlsx
    python main.py jan.xlsx feb.xlsx mar.xlsx --chat
    python main.py --demo --chat
"""

from __future__ import annotations

import argparse
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


def _expand_input_paths(inputs: list[str]) -> list[Path]:
    resolved: list[Path] = []
    for item in inputs:
        p = Path(item)
        if p.is_dir():
            for ext in ("*.xlsx", "*.xlsm", "*.xls", "*.csv"):
                resolved.extend(sorted(p.glob(ext)))
        else:
            resolved.append(p)
    return resolved


def _print_quick_tables(bundle: dict) -> None:
    ov = bundle["overview"]
    cur = ov["currency"]
    where = bundle["where_you_spend_more"]
    monthly = bundle["monthly_spending"]
    forecast = bundle["future_month_forecast"]
    profiles = bundle.get("sheet_profiles") or []

    if profiles:
        sh_tbl = Table(title="Parsed Sheets / Files", show_lines=False)
        sh_tbl.add_column("File", style="cyan")
        sh_tbl.add_column("Sheet")
        sh_tbl.add_column("Kind", style="magenta")
        sh_tbl.add_column("Rows", justify="right")
        for p in profiles:
            sh_tbl.add_row(p["file"], p["sheet"], p["kind"], str(p["row_count"]))
        console.print(sh_tbl)

    if where["categories"]:
        cat_tbl = Table(title="Top Categories / Groups", show_lines=False)
        cat_tbl.add_column("Category / Group", style="cyan")
        cat_tbl.add_column("Total", justify="right", style="bold white")
        cat_tbl.add_column("Share", justify="right")
        cat_tbl.add_column("Rows", justify="right")
        for c in where["categories"][:7]:
            cat_tbl.add_row(c["category"], f"{cur}{c['total_spend']:,.2f}", f"{c['share_percent']}%", str(c["transaction_count"]))
        console.print(cat_tbl)

    if monthly["months_chronological"]:
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
        console.print(mon_tbl)


def _print_qa_answer(result: dict) -> None:
    console.print(Panel(Markdown(result["answer"]), title=f"Answer ({result['retrieval_mode']})", border_style="blue"))
    if result.get("citations"):
        cites = " | ".join(f"{c['source']} ({c['score']})" for c in result["citations"][:4])
        console.print(f"[dim]Sources (CAG + RAG): {cites}[/dim]")


def main(argv: list[str] | None = None) -> int:
    parser = argparse.ArgumentParser(
        description="Universal Ollama Excel Report Analyzer & Hybrid CAG+RAG Q&A Agent."
    )
    parser.add_argument("excel_files", nargs="*", help="One or more .xlsx/.xlsm/.xls/.csv files or a folder")
    parser.add_argument("--demo", action="store_true", help="Generate and analyze sample_expenses.xlsx")
    parser.add_argument("--sheet", help="Analyze only a specific sheet name (default: all sheets)")
    parser.add_argument("--currency", help="Override detected currency symbol (e.g. $, ₹, €)")
    parser.add_argument("--model", help="Ollama model name (e.g. qwen2.5:latest, qwen2.5-coder:3b)")
    parser.add_argument("--host", default=DEFAULT_OLLAMA_HOST, help="Ollama host URL (default: http://127.0.0.1:11434)")
    parser.add_argument("--out", default="reports", help="Directory to save Markdown & HTML reports")
    parser.add_argument("--ask", help="Ask a follow-up question using CAG + RAG after generating the report")
    parser.add_argument("--chat", action="store_true", help="Start interactive CAG + RAG Q&A chat")
    parser.add_argument("--deep", action="store_true", help="Use multi-turn autonomous tool calling for initial report")
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
        excel_paths = [generate_sample_excel("sample_expenses.xlsx")]
        console.print(f"[green]Generated demo workbook:[/green] {excel_paths[0].resolve()}")
    elif args.excel_files:
        excel_paths = _expand_input_paths(args.excel_files)
    else:
        parser.print_help()
        console.print("\n[yellow]Tip:[/yellow] Pass Excel file(s) (`python main.py report.xlsx`) or try `--demo`.")
        return 1

    if not excel_paths:
        console.print("[red]No matching Excel/CSV files found.[/red]")
        return 1

    console.print(Panel.fit(f"[bold cyan]Ollama Universal Excel Agent (CAG + RAG)[/bold cyan]\nFiles: {', '.join(p.name for p in excel_paths)}", border_style="cyan"))

    try:
        analyzer = ExpenseAnalyzer(excel_paths, sheet_name=args.sheet, currency_override=args.currency)
    except Exception as exc:
        console.print(f"[red]Failed to parse Excel file(s):[/red] {exc}")
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
        return 1

    console.print(
        f"\n[bold green]Running Ollama Agent ({agent.model})...[/bold green] "
        f"[dim](RAG indexed {len(agent.kb.chunks)} chunks across {len(analyzer.sheet_tables)} sheet(s))[/dim]"
    )
    with console.status("[bold cyan]Generating report...[/bold cyan]"):
        report_md, tools_used = agent.generate_full_report(fast_mode=not args.deep)

    console.print(Panel(Markdown(report_md), title=f"Report ({agent.model})", border_style="green"))

    stem = excel_paths[0].stem if len(excel_paths) == 1 else "combined_reports"
    md_path, html_path = save_reports(bundle, report_md, agent.model, args.out, stem=stem)
    console.print(f"[bold cyan]Saved Markdown report:[/bold cyan] {md_path.resolve()}")
    console.print(f"[bold cyan]Saved HTML dashboard:[/bold cyan]  {html_path.resolve()}")
    console.print(f"[dim]Pipeline: {', '.join(tools_used)}[/dim]")

    if args.ask:
        console.print(f"\n[bold yellow]Question:[/bold yellow] {args.ask}")
        with console.status("[cyan]Retrieving CAG + RAG context and answering...[/cyan]"):
            res = agent.ask_with_sources(args.ask)
        _print_qa_answer(res)

    if args.chat:
        console.print("\n[bold cyan]Interactive CAG + RAG Q&A Mode[/bold cyan] (type 'exit' or 'quit' to stop)")
        while True:
            try:
                q = console.input("\n[bold green]Ask about any sheet/report > [/bold green]").strip()
            except (EOFError, KeyboardInterrupt):
                break
            if not q or q.lower() in {"exit", "quit", "q"}:
                break
            with console.status("[cyan]Retrieving & answering...[/cyan]"):
                res = agent.ask_with_sources(q)
            _print_qa_answer(res)

    return 0


if __name__ == "__main__":
    raise SystemExit(main())
