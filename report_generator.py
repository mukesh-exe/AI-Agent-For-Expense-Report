"""
Generates formatted Markdown and self-contained interactive HTML spending reports.
"""

from __future__ import annotations

import html
import re
from datetime import datetime
from pathlib import Path
from typing import Any


def _fmt_money(amount: float | int | None, currency: str) -> str:
    if amount is None:
        return "N/A"
    return f"{currency}{amount:,.2f}"


def _simple_markdown_to_html(md_text: str) -> str:
    """Lightweight Markdown-to-HTML converter for Ollama report prose & tables."""
    lines = md_text.splitlines()
    out: list[str] = []
    in_ul = False
    in_ol = False
    in_table = False
    in_code = False

    def close_lists() -> None:
        nonlocal in_ul, in_ol
        if in_ul:
            out.append("</ul>")
            in_ul = False
        if in_ol:
            out.append("</ol>")
            in_ol = False

    def close_table() -> None:
        nonlocal in_table
        if in_table:
            out.append("</tbody></table></div>")
            in_table = False

    def inline_fmt(text: str) -> str:
        escaped = html.escape(text)
        escaped = re.sub(r"`([^`]+)`", r"<code>\1</code>", escaped)
        escaped = re.sub(r"\*\*([^*]+)\*\*", r"<strong>\1</strong>", escaped)
        escaped = re.sub(r"\*([^*]+)\*", r"<em>\1</em>", escaped)
        return escaped

    i = 0
    while i < len(lines):
        line = lines[i]
        stripped = line.strip()

        if stripped.startswith("```"):
            close_lists()
            close_table()
            if not in_code:
                out.append("<pre><code>")
                in_code = True
            else:
                out.append("</code></pre>")
                in_code = False
            i += 1
            continue

        if in_code:
            out.append(html.escape(line))
            i += 1
            continue

        if not stripped:
            close_lists()
            close_table()
            i += 1
            continue

        # Markdown table
        if "|" in stripped and stripped.startswith("|"):
            close_lists()
            cells = [c.strip() for c in stripped.strip("|").split("|")]
            # Check if next line is separator
            if not in_table and i + 1 < len(lines) and re.match(r"^\|?[\s:\-|]+\|?$", lines[i + 1].strip()):
                out.append('<div class="table-wrap"><table><thead><tr>')
                for c in cells:
                    out.append(f"<th>{inline_fmt(c)}</th>")
                out.append("</tr></thead><tbody>")
                in_table = True
                i += 2
                continue
            elif in_table:
                out.append("<tr>")
                for c in cells:
                    out.append(f"<td>{inline_fmt(c)}</td>")
                out.append("</tr>")
                i += 1
                continue

        close_table()

        if stripped.startswith("### "):
            close_lists()
            out.append(f"<h4>{inline_fmt(stripped[4:])}</h4>")
        elif stripped.startswith("## "):
            close_lists()
            out.append(f"<h3>{inline_fmt(stripped[3:])}</h3>")
        elif stripped.startswith("# "):
            close_lists()
            out.append(f"<h2>{inline_fmt(stripped[2:])}</h2>")
        elif re.match(r"^[-*•]\s+", stripped):
            if in_ol:
                out.append("</ol>")
                in_ol = False
            if not in_ul:
                out.append("<ul>")
                in_ul = True
            out.append(f"<li>{inline_fmt(re.sub(r'^[-*•]\\s+', '', stripped))}</li>")
        elif re.match(r"^\d+\.\s+", stripped):
            if in_ul:
                out.append("</ul>")
                in_ul = False
            if not in_ol:
                out.append("<ol>")
                in_ol = True
            out.append(f"<li>{inline_fmt(re.sub(r'^\\d+\\.\\s+', '', stripped))}</li>")
        else:
            close_lists()
            out.append(f"<p>{inline_fmt(stripped)}</p>")

        i += 1

    close_lists()
    close_table()
    if in_code:
        out.append("</code></pre>")
    return "\n".join(out)


def save_reports(
    bundle: dict[str, Any],
    ai_narrative: str,
    model_name: str,
    output_dir: str | Path,
    stem: str = "spending",
) -> tuple[Path, Path]:
    out_dir = Path(output_dir)
    out_dir.mkdir(parents=True, exist_ok=True)

    overview = bundle["overview"]
    where = bundle["where_you_spend_more"]
    monthly = bundle["monthly_spending"]
    forecast = bundle["future_month_forecast"]
    why = bundle["why_peak_month_happened"]
    cur = overview["currency"]
    generated_at = datetime.now().strftime("%Y-%m-%d %H:%M")

    # 1. Save Markdown Report
    md_path = out_dir / f"{stem}_report.md"
    cat_lines = "\n".join(
        f"| {c['category']} | {_fmt_money(c['total_spend'], cur)} | {c['share_percent']}% | {c['transaction_count']} | {_fmt_money(c['average_transaction'], cur)} |"
        for c in where["categories"]
    )
    month_lines = "\n".join(
        f"| {m['month']} | {_fmt_money(m['total_spend'], cur)} | {m['transaction_count']} | "
        f"{(str(m['mom_change_percent']) + '%') if m['mom_change_percent'] is not None else '-'} | "
        f"{m['vs_monthly_avg_percent']:+}% | {m['top_category_in_month']} |"
        for m in monthly["months_chronological"]
    )
    forecast_lines = "\n".join(
        f"| {f['month']} ({f['calendar_month']}) | {_fmt_money(f['projected_spend'], cur)} | {f['vs_historical_avg_percent']:+}% |"
        for f in forecast["next_3_months_forecast"]
    )

    md_content = f"""# Excel Spending Analysis Report

- **Source File:** `{overview['file']}` (Sheet: `{overview['sheet']}`)
- **Date Range:** {overview['date_range']['start']} to {overview['date_range']['end']} ({overview['date_range']['num_months']} months)
- **Analyzed By Ollama Model:** `{model_name}`
- **Generated:** {generated_at}

---

## AI Agent Analysis

{ai_narrative}

---

## Exact Data Appendix

### Summary Metrics
- **Total Spending:** {_fmt_money(overview['total_spend'], cur)}
- **Average Monthly Spend:** {_fmt_money(overview['average_monthly_spend'], cur)}
- **Total Transactions:** {overview['transaction_count']}
- **Peak Spending Month:** {monthly['peak_month']['month']} ({_fmt_money(monthly['peak_month']['total_spend'], cur)})
- **Top Spending Category:** {where['top_category']['category']} ({_fmt_money(where['top_category']['total_spend'], cur)}, {where['top_category']['share_percent']}%)

### Where You Spend More (Categories)
| Category | Total Spend | Share | Transactions | Avg Transaction |
|---|---:|---:|---:|---:|
{cat_lines}

### Monthly Spending Breakdown
| Month | Total Spend | Transactions | MoM Change | vs Avg | Top Category |
|---|---:|---:|---:|---:|---|
{month_lines}

### Upcoming 3-Month Projection (Trend: {forecast['trend_direction']})
| Upcoming Month | Projected Spend | vs Historical Avg |
|---|---:|---:|
{forecast_lines}

### Why {why['analyzed_month']} Spiked (+{_fmt_money(why['excess_spend_over_baseline'], cur)} vs Baseline)
- **Peak Month Transactions vs Baseline:** {why['frequency_vs_ticket_size']['target_month_transactions']} vs {why['frequency_vs_ticket_size']['baseline_avg_transactions']} avg/month
- **Peak Month Avg Ticket vs Baseline:** {_fmt_money(why['frequency_vs_ticket_size']['target_month_avg_ticket'], cur)} vs {_fmt_money(why['frequency_vs_ticket_size']['baseline_avg_ticket'], cur)}
"""
    md_path.write_text(md_content, encoding="utf-8")

    # 2. Save Self-Contained HTML Dashboard Report
    html_path = out_dir / f"{stem}_report.html"

    max_cat_val = max((c["total_spend"] for c in where["categories"]), default=1.0) or 1.0
    cat_bars_html = "\n".join(
        f"""
        <div class="bar-row">
          <div class="bar-label"><span>{html.escape(c['category'])}</span><span>{_fmt_money(c['total_spend'], cur)} ({c['share_percent']}%)</span></div>
          <div class="bar-track"><div class="bar-fill" style="width:{min(100, c['total_spend'] / max_cat_val * 100):.1f}%"></div></div>
        </div>
        """
        for c in where["categories"][:8]
    )

    all_months_vals = [m["total_spend"] for m in monthly["months_chronological"]] + [
        f["projected_spend"] for f in forecast["next_3_months_forecast"]
    ]
    max_month_val = max(all_months_vals, default=1.0) or 1.0
    peak_month_str = monthly["peak_month"]["month"]

    month_bars_html = "\n".join(
        f"""
        <div class="col-item {'peak' if m['month'] == peak_month_str else ''}" title="{m['month']}: {_fmt_money(m['total_spend'], cur)}">
          <div class="col-val">{_fmt_money(m['total_spend'], cur)}</div>
          <div class="col-track"><div class="col-fill" style="height:{max(4, m['total_spend'] / max_month_val * 100):.1f}%"></div></div>
          <div class="col-lbl">{html.escape(m['month'])}</div>
        </div>
        """
        for m in monthly["months_chronological"]
    )
    forecast_bars_html = "\n".join(
        f"""
        <div class="col-item forecast" title="Projected {f['month']}: {_fmt_money(f['projected_spend'], cur)}">
          <div class="col-val">{_fmt_money(f['projected_spend'], cur)}</div>
          <div class="col-track"><div class="col-fill" style="height:{max(4, f['projected_spend'] / max_month_val * 100):.1f}%"></div></div>
          <div class="col-lbl">{html.escape(f['month'])}*</div>
        </div>
        """
        for f in forecast["next_3_months_forecast"]
    )

    drivers_rows = "\n".join(
        f"<tr><td>{html.escape(d['category'])}</td><td>{_fmt_money(d['target_month_spend'], cur)}</td>"
        f"<td>{_fmt_money(d['baseline_monthly_avg'], cur)}</td>"
        f"<td class='{'pos-bad' if d['difference'] > 0 else 'pos-good'}'>{d['difference']:+,.2f} ({d['percent_vs_baseline']:+.1f}%)</td></tr>"
        for d in why["top_category_surges"]
    )

    spike_tx_rows = "\n".join(
        f"<tr><td>{html.escape(t['date'])}</td><td>{html.escape(t['description'])}</td>"
        f"<td>{html.escape(t['category'])}</td><td>{_fmt_money(t['amount'], cur)}</td><td>{t['share_of_month_percent']}%</td></tr>"
        for t in why["top_transactions_in_month"][:6]
    )

    ai_html = _simple_markdown_to_html(ai_narrative)
    next_peak = forecast.get("highest_projected_upcoming_month") or {}

    html_content = f"""<!DOCTYPE html>
<html lang="en">
<head>
  <meta charset="UTF-8" />
  <meta name="viewport" content="width=device-width, initial-scale=1.0" />
  <title>Ollama Excel Spending Report</title>
  <style>
    :root {{
      --bg: #0f172a;
      --card: #1e293b;
      --border: #334155;
      --text: #e2e8f0;
      --muted: #94a3b8;
      --accent: #38bdf8;
      --peak: #f43f5e;
      --forecast: #a855f7;
      --good: #22c55e;
    }}
    * {{ box-sizing: border-box; }}
    body {{
      margin: 0; font-family: system-ui, -apple-system, Segoe UI, Roboto, sans-serif;
      background: var(--bg); color: var(--text); line-height: 1.6; padding: 28px 18px;
    }}
    .container {{ max-width: 1120px; margin: 0 auto; }}
    header {{ display: flex; flex-wrap: wrap; justify-content: space-between; align-items: flex-end; gap: 12px; margin-bottom: 24px; border-bottom: 1px solid var(--border); padding-bottom: 16px; }}
    h1 {{ margin: 0; font-size: 1.8rem; color: #fff; }}
    .meta {{ color: var(--muted); font-size: 0.9rem; }}
    .badge {{ display: inline-block; background: #0ea5e922; border: 1px solid var(--accent); color: var(--accent); padding: 2px 10px; border-radius: 999px; font-size: 0.82rem; }}
    .kpi-grid {{ display: grid; grid-template-columns: repeat(auto-fit, minmax(220px, 1fr)); gap: 14px; margin-bottom: 22px; }}
    .kpi {{ background: var(--card); border: 1px solid var(--border); border-radius: 12px; padding: 16px; }}
    .kpi-label {{ color: var(--muted); font-size: 0.82rem; text-transform: uppercase; letter-spacing: 0.04em; }}
    .kpi-val {{ font-size: 1.5rem; font-weight: 700; margin-top: 4px; color: #fff; }}
    .kpi-sub {{ font-size: 0.84rem; color: var(--muted); margin-top: 4px; }}
    .grid-2 {{ display: grid; grid-template-columns: repeat(auto-fit, minmax(360px, 1fr)); gap: 18px; margin-bottom: 22px; }}
    .card {{ background: var(--card); border: 1px solid var(--border); border-radius: 12px; padding: 20px; margin-bottom: 20px; }}
    .card h2, .card h3 {{ margin-top: 0; color: #fff; }}
    .bar-row {{ margin-bottom: 12px; }}
    .bar-label {{ display: flex; justify-content: space-between; font-size: 0.88rem; margin-bottom: 4px; }}
    .bar-track {{ height: 10px; background: #0f172a; border-radius: 999px; overflow: hidden; }}
    .bar-fill {{ height: 100%; background: linear-gradient(90deg, #38bdf8, #818cf8); border-radius: 999px; }}
    .col-chart {{ display: flex; align-items: flex-end; gap: 10px; overflow-x: auto; padding: 16px 4px 6px; min-height: 230px; }}
    .col-item {{ flex: 1 0 64px; display: flex; flex-direction: column; align-items: center; }}
    .col-val {{ font-size: 0.72rem; color: var(--muted); margin-bottom: 6px; white-space: nowrap; }}
    .col-track {{ width: 32px; height: 140px; background: #0f172a; border-radius: 6px; display: flex; align-items: flex-end; overflow: hidden; }}
    .col-fill {{ width: 100%; background: var(--accent); border-radius: 6px 6px 0 0; }}
    .col-item.peak .col-fill {{ background: var(--peak); }}
    .col-item.peak .col-lbl {{ color: #fda4af; font-weight: 700; }}
    .col-item.forecast .col-fill {{ background: var(--forecast); opacity: 0.85; }}
    .col-item.forecast .col-lbl {{ color: #d8b4fe; }}
    .col-lbl {{ font-size: 0.76rem; color: var(--muted); margin-top: 6px; }}
    .legend {{ display: flex; gap: 16px; font-size: 0.8rem; color: var(--muted); margin-top: 8px; }}
    .dot {{ display: inline-block; width: 10px; height: 10px; border-radius: 50%; margin-right: 5px; }}
    .table-wrap {{ overflow-x: auto; }}
    table {{ width: 100%; border-collapse: collapse; margin: 10px 0; font-size: 0.9rem; }}
    th, td {{ border-bottom: 1px solid var(--border); padding: 9px 10px; text-align: left; }}
    th {{ color: var(--muted); font-weight: 600; }}
    .pos-bad {{ color: #fda4af; font-weight: 600; }}
    .pos-good {{ color: #86efac; font-weight: 600; }}
    .ai-prose h2, .ai-prose h3, .ai-prose h4 {{ color: #7dd3fc; margin-top: 20px; }}
    .ai-prose code {{ background: #0f172a; padding: 2px 5px; border-radius: 4px; }}
  </style>
</head>
<body>
  <div class="container">
    <header>
      <div>
        <h1>Excel Spending Analysis Report</h1>
        <div class="meta">File: {html.escape(Path(overview['file']).name)} ({html.escape(overview['sheet'])}) &bull; {overview['date_range']['start']} to {overview['date_range']['end']} &bull; Generated {generated_at}</div>
      </div>
      <span class="badge">Ollama Model: {html.escape(model_name)}</span>
    </header>

    <div class="kpi-grid">
      <div class="kpi">
        <div class="kpi-label">Total Spending</div>
        <div class="kpi-val">{_fmt_money(overview['total_spend'], cur)}</div>
        <div class="kpi-sub">{overview['transaction_count']} transactions across {overview['date_range']['num_months']} months</div>
      </div>
      <div class="kpi">
        <div class="kpi-label">Where You Spend Most</div>
        <div class="kpi-val">{html.escape(where['top_category']['category'])}</div>
        <div class="kpi-sub">{_fmt_money(where['top_category']['total_spend'], cur)} ({where['top_category']['share_percent']}% of total)</div>
      </div>
      <div class="kpi">
        <div class="kpi-label">Peak Spending Month</div>
        <div class="kpi-val">{html.escape(monthly['peak_month']['month'])}</div>
        <div class="kpi-sub">{_fmt_money(monthly['peak_month']['total_spend'], cur)} ({monthly['peak_month']['vs_monthly_avg_percent']:+.1f}% vs avg)</div>
      </div>
      <div class="kpi">
        <div class="kpi-label">Next High-Spend Month Forecast</div>
        <div class="kpi-val">{html.escape(next_peak.get('month', 'N/A'))}</div>
        <div class="kpi-sub">Projected {_fmt_money(next_peak.get('projected_spend'), cur)} (Trend: {html.escape(forecast['trend_direction'])})</div>
      </div>
    </div>

    <div class="grid-2">
      <div class="card">
        <h3>Where You Spend More (Categories)</h3>
        {cat_bars_html}
      </div>
      <div class="card">
        <h3>Monthly Spending &amp; 3-Month Forecast</h3>
        <div class="col-chart">
          {month_bars_html}
          {forecast_bars_html}
        </div>
        <div class="legend">
          <span><span class="dot" style="background:var(--accent)"></span>Actual Month</span>
          <span><span class="dot" style="background:var(--peak)"></span>Peak Month</span>
          <span><span class="dot" style="background:var(--forecast)"></span>Projected (*Forecast)</span>
        </div>
      </div>
    </div>

    <div class="grid-2">
      <div class="card">
        <h3>Why {html.escape(why['analyzed_month'])} Spiked (vs Monthly Baseline)</h3>
        <div class="table-wrap">
          <table>
            <thead><tr><th>Category</th><th>{html.escape(why['analyzed_month'])}</th><th>Baseline Avg</th><th>Difference</th></tr></thead>
            <tbody>{drivers_rows}</tbody>
          </table>
        </div>
      </div>
      <div class="card">
        <h3>Top Transactions in {html.escape(why['analyzed_month'])}</h3>
        <div class="table-wrap">
          <table>
            <thead><tr><th>Date</th><th>Description</th><th>Category</th><th>Amount</th><th>Month %</th></tr></thead>
            <tbody>{spike_tx_rows}</tbody>
          </table>
        </div>
      </div>
    </div>

    <div class="card ai-prose">
      <h2>Ollama AI Agent Report &amp; Recommendations</h2>
      {ai_html}
    </div>
  </div>
</body>
</html>
"""
    html_path.write_text(html_content, encoding="utf-8")
    return md_path, html_path
