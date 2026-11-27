"""
Ollama-powered Excel Spending Analyst Agent.

Uses Ollama's /api/chat tool-calling loop so the local LLM inspects the Excel
analytics tools for exact numbers, explains where you spend more, which month
you spend / will spend more, and why it happens, and supports follow-up Q&A.
"""

from __future__ import annotations

import json
import os
from typing import Any, Callable

import requests

from excel_analyzer import ExpenseAnalyzer


DEFAULT_OLLAMA_HOST = os.environ.get("OLLAMA_HOST", "http://127.0.0.1:11434").rstrip("/")
PREFERRED_MODELS = (
    "qwen2.5:latest",
    "qwen3:4b",
    "qwen2.5-coder:3b",
    "llama3.1:latest",
    "llama3.2:latest",
    "mistral:latest",
    "phi3:latest",
)


def list_ollama_models(host: str = DEFAULT_OLLAMA_HOST, timeout: int = 5) -> list[str]:
    resp = requests.get(f"{host}/api/tags", timeout=timeout)
    resp.raise_for_status()
    data = resp.json()
    return [m["name"] for m in data.get("models", []) if m.get("name")]


def select_best_model(available: list[str], requested: str | None = None) -> str:
    if requested:
        if requested in available:
            return requested
        # Allow prefix match like "qwen2.5" -> "qwen2.5:latest"
        for m in available:
            if m.startswith(requested):
                return m
        raise ValueError(f"Requested Ollama model '{requested}' not found. Available: {available}")

    if not available:
        raise RuntimeError("No Ollama models found. Run `ollama pull qwen2.5` first.")

    for pref in PREFERRED_MODELS:
        if pref in available:
            return pref

    local_models = [m for m in available if "cloud" not in m.lower()]
    return local_models[0] if local_models else available[0]


def _build_tool_schemas() -> list[dict[str, Any]]:
    return [
        {
            "type": "function",
            "function": {
                "name": "get_overview",
                "description": "Get overall spending summary, date range, total spend, monthly average, largest expense, and detected columns.",
                "parameters": {"type": "object", "properties": {}, "required": []},
            },
        },
        {
            "type": "function",
            "function": {
                "name": "get_where_you_spend_more",
                "description": "Get top spending categories, top merchants/descriptions, most frequent merchants, and payment modes.",
                "parameters": {
                    "type": "object",
                    "properties": {
                        "top_n": {
                            "type": "integer",
                            "description": "Number of top categories and merchants to return (default 10).",
                        }
                    },
                    "required": [],
                },
            },
        },
        {
            "type": "function",
            "function": {
                "name": "get_monthly_spending",
                "description": "Get month-by-month spending totals, month-over-month changes, peak spending month, and lowest spending month.",
                "parameters": {"type": "object", "properties": {}, "required": []},
            },
        },
        {
            "type": "function",
            "function": {
                "name": "get_future_month_forecast",
                "description": "Get spending trend, recurring fixed commitments, seasonal calendar-month ranking, and projected spending for the next 3 months.",
                "parameters": {"type": "object", "properties": {}, "required": []},
            },
        },
        {
            "type": "function",
            "function": {
                "name": "explain_why_it_happens",
                "description": "Get root-cause drivers for why spending spiked in a target month (defaults to peak month): surging categories vs baseline, spike transactions, frequency vs ticket size, and weekday/weekend habits.",
                "parameters": {
                    "type": "object",
                    "properties": {
                        "target_month": {
                            "type": "string",
                            "description": "Optional month in YYYY-MM format. Omit to analyze the highest-spending peak month.",
                        }
                    },
                    "required": [],
                },
            },
        },
        {
            "type": "function",
            "function": {
                "name": "search_transactions",
                "description": "Search and filter individual transactions by month, category, keyword, or minimum amount.",
                "parameters": {
                    "type": "object",
                    "properties": {
                        "month": {"type": "string", "description": "Filter by YYYY-MM month."},
                        "category": {"type": "string", "description": "Filter by category substring."},
                        "keyword": {"type": "string", "description": "Filter by description/category keyword."},
                        "min_amount": {"type": "number", "description": "Minimum transaction amount."},
                        "limit": {"type": "integer", "description": "Max transactions to return (default 15)."},
                    },
                    "required": [],
                },
            },
        },
    ]


SYSTEM_PROMPT = """You are an expert Personal Finance & Excel Spending Analyst AI Agent running on Ollama.

Your job is to analyze the user's Excel expense sheet accurately using the provided analytical tools and write a clear, practical, insightful report.

Strict rules:
1. ALWAYS call the analytical tools before answering. Never invent, guess, or round numbers differently from tool outputs.
2. Always include the exact currency symbol reported by `get_overview` when mentioning amounts.
3. Explicitly answer ALL four parts of the user's request clearly with headings:
   - **Spending Overview Report** (total spend, date range, monthly average, transaction count, income/net if available)
   - **Where You Spend More** (top categories with % share, top merchants/payees, and frequent small drains)
   - **Which Month You Spend More (Peak Month & Upcoming Forecast)** (identify the highest historical month AND explain which upcoming/seasonal month is projected to have higher spending based on trend & recurring bills)
   - **Why It Happens (Root Causes & Habits)** (explain the exact surging categories, specific big-ticket spike transactions, frequency vs ticket-size shifts, and weekend/timing habits)
   - **Actionable Savings Plan** (3–5 concrete, data-backed steps tailored to the user's real numbers)
4. Keep the tone helpful, crisp, and easy for anyone to read. Use markdown tables or bullet lists where helpful.
"""


class OllamaSpendingAgent:
    def __init__(
        self,
        analyzer: ExpenseAnalyzer,
        model: str | None = None,
        host: str = DEFAULT_OLLAMA_HOST,
        on_tool_call: Callable[[str, dict[str, Any]], None] | None = None,
    ) -> None:
        self.analyzer = analyzer
        self.host = host.rstrip("/")
        self.on_tool_call = on_tool_call
        available = list_ollama_models(self.host)
        self.model = select_best_model(available, model)
        self.tools = _build_tool_schemas()
        self.messages: list[dict[str, Any]] = [{"role": "system", "content": SYSTEM_PROMPT}]
        self._tool_calls_supported = True

    def _execute_tool(self, name: str, args: dict[str, Any]) -> str:
        if self.on_tool_call:
            self.on_tool_call(name, args)

        try:
            if name == "get_overview":
                res = self.analyzer.get_overview()
            elif name == "get_where_you_spend_more":
                res = self.analyzer.get_where_you_spend_more(top_n=int(args.get("top_n", 10)))
            elif name == "get_monthly_spending":
                res = self.analyzer.get_monthly_spending()
            elif name == "get_future_month_forecast":
                res = self.analyzer.get_future_month_forecast()
            elif name == "explain_why_it_happens":
                res = self.analyzer.explain_why_it_happens(target_month=args.get("target_month"))
            elif name == "search_transactions":
                res = self.analyzer.search_transactions(
                    month=args.get("month"),
                    category=args.get("category"),
                    keyword=args.get("keyword"),
                    min_amount=args.get("min_amount"),
                    limit=int(args.get("limit", 15)),
                )
            else:
                res = {"error": f"Unknown tool: {name}"}
        except Exception as exc:
            res = {"error": str(exc)}

        return json.dumps(res, ensure_ascii=False)

    def _build_compact_report_context(self) -> str:
        """Build a token-efficient (~600 token) exact summary for fast single-pass reporting."""
        ov = self.analyzer.get_overview()
        cur = ov["currency"]
        where = self.analyzer.get_where_you_spend_more(top_n=7)
        mon = self.analyzer.get_monthly_spending()
        fc = self.analyzer.get_future_month_forecast()
        why = self.analyzer.explain_why_it_happens()

        lines = [
            f"FILE: {ov['file']} (Sheet: {ov['sheet']}) | Currency: {cur}",
            f"RANGE: {ov['date_range']['start']} to {ov['date_range']['end']} ({ov['date_range']['num_months']} months) | Txns: {ov['transaction_count']}",
            f"TOTAL SPEND: {cur}{ov['total_spend']:,.2f} | Avg/Month: {cur}{ov['average_monthly_spend']:,.2f} | Median/Month: {cur}{ov['median_monthly_spend']:,.2f}",
        ]
        if ov["total_income_detected"]:
            lines.append(f"TOTAL INCOME: {cur}{ov['total_income_detected']:,.2f} | Net Cashflow: {cur}{ov['net_cashflow']:,.2f}")
        lg = ov["largest_single_expense"]
        lines.append(f"LARGEST TXN: {lg['date']} {lg['description']} ({lg['category']}) = {cur}{lg['amount']:,.2f}")

        lines.append("\nTOP CATEGORIES (Where you spend more):")
        for c in where["categories"]:
            lines.append(f"- {c['category']}: {cur}{c['total_spend']:,.2f} ({c['share_percent']}%, {c['transaction_count']} txns, avg {cur}{c['average_transaction']:,.2f})")

        lines.append("\nTOP MERCHANTS:")
        for m in where["top_merchants_by_spend"][:5]:
            lines.append(f"- {m['description']} [{m['category']}]: {cur}{m['total_spend']:,.2f} ({m['transaction_count']}x)")

        lines.append("\nMOST FREQUENT:")
        for f_m in where["most_frequent_merchants"][:4]:
            lines.append(f"- {f_m['description']}: {f_m['transaction_count']}x, total {cur}{f_m['total_spend']:,.2f}")

        lines.append("\nMONTHLY SPENDING:")
        for m in mon["months_chronological"]:
            tag = " [PEAK]" if m["month"] == mon["peak_month"]["month"] else ""
            lines.append(f"- {m['month']}{tag}: {cur}{m['total_spend']:,.2f} ({m['vs_monthly_avg_percent']:+.1f}% vs avg, top: {m['top_category_in_month']})")

        lines.append(f"\nFORECAST (Trend: {fc['trend_direction']}, Fixed recurring ~{cur}{fc['estimated_recurring_monthly_commitments']:,.2f}/mo):")
        for nf in fc["next_3_months_forecast"]:
            lines.append(f"- {nf['month']} ({nf['calendar_month']}): projected {cur}{nf['projected_spend']:,.2f} ({nf['vs_historical_avg_percent']:+.1f}% vs avg)")
        if fc["recurring_items"]:
            rec_str = ", ".join(f"{r['description']} ({cur}{r['typical_amount']:,.0f})" for r in fc["recurring_items"][:5])
            lines.append(f"RECURRING BILLS: {rec_str}")

        freq = why["frequency_vs_ticket_size"]
        hab = why["overall_spending_habits"]
        lines.append(f"\nWHY PEAK MONTH ({why['analyzed_month']}) HAPPENED (+{cur}{why['excess_spend_over_baseline']:,.2f} over {cur}{why['other_months_avg_total']:,.2f} baseline):")
        lines.append(f"- Txns: {freq['target_month_transactions']} vs {freq['baseline_avg_transactions']} baseline | Avg ticket: {cur}{freq['target_month_avg_ticket']:,.2f} vs {cur}{freq['baseline_avg_ticket']:,.2f} baseline")
        lines.append("SURGING CATEGORIES vs BASELINE:")
        for d in why["top_category_surges"][:4]:
            lines.append(f"- {d['category']}: {cur}{d['target_month_spend']:,.2f} vs {cur}{d['baseline_monthly_avg']:,.2f} avg ({d['difference']:+,.2f}, {d['percent_vs_baseline']:+.1f}%)")
        lines.append("TOP SPIKE TRANSACTIONS IN PEAK MONTH:")
        for t in why["top_transactions_in_month"][:5]:
            lines.append(f"- {t['date']} {t['description']} ({t['category']}): {cur}{t['amount']:,.2f} ({t['share_of_month_percent']}% of month)")
        lines.append(
            f"HABITS: Weekend avg/day {cur}{hab['avg_spend_per_active_weekend_day']:,.2f} vs Weekday {cur}{hab['avg_spend_per_active_weekday']:,.2f} | "
            f"Small frequent (<={cur}{hab['small_frequent_purchases']['threshold_amount']:,.0f}): {hab['small_frequent_purchases']['count']} txns = {hab['small_frequent_purchases']['share_percent']}% | "
            f"Month split (Early/Mid/Late): {hab['month_timing_split_percent']['early_month_days_1_10']}% / {hab['month_timing_split_percent']['mid_month_days_11_20']}% / {hab['month_timing_split_percent']['late_month_days_21_31']}%"
        )
        return "\n".join(lines)

    def _call_ollama(
        self,
        messages: list[dict[str, Any]],
        use_tools: bool = True,
        num_ctx: int = 8192,
        num_predict: int | None = None,
        disable_thinking: bool = False,
        on_token: Callable[[str], None] | None = None,
    ) -> dict[str, Any]:
        options: dict[str, Any] = {"temperature": 0.1, "num_ctx": num_ctx}
        if num_predict is not None:
            options["num_predict"] = num_predict

        payload: dict[str, Any] = {
            "model": self.model,
            "messages": messages,
            "stream": bool(on_token),
            "keep_alive": "15m",
            "options": options,
        }
        if disable_thinking and any(k in self.model.lower() for k in ("qwen3", "deepseek-r1", "gpt-oss")):
            payload["think"] = False
        if use_tools and self._tool_calls_supported and not on_token:
            payload["tools"] = self.tools

        if on_token:
            resp = requests.post(f"{self.host}/api/chat", json=payload, stream=True, timeout=240)
            if resp.status_code == 400 and "think" in payload:
                payload.pop("think", None)
                resp = requests.post(f"{self.host}/api/chat", json=payload, stream=True, timeout=240)
            resp.raise_for_status()
            full_chunks: list[str] = []
            for raw_line in resp.iter_lines(decode_unicode=True):
                if not raw_line:
                    continue
                chunk = json.loads(raw_line)
                delta = (chunk.get("message") or {}).get("content") or ""
                if delta:
                    full_chunks.append(delta)
                    on_token(delta)
                if chunk.get("done"):
                    break
            return {"message": {"role": "assistant", "content": "".join(full_chunks)}}

        resp = requests.post(f"{self.host}/api/chat", json=payload, timeout=240)
        if resp.status_code == 400 and "think" in payload:
            payload.pop("think", None)
            resp = requests.post(f"{self.host}/api/chat", json=payload, timeout=240)
        if resp.status_code == 400 and use_tools and "tool" in resp.text.lower():
            # Model doesn't support native tool calling; disable and retry cleanly
            self._tool_calls_supported = False
            payload.pop("tools", None)
            resp = requests.post(f"{self.host}/api/chat", json=payload, timeout=240)

        resp.raise_for_status()
        return resp.json()

    def generate_full_report(
        self,
        fast_mode: bool = True,
        on_token: Callable[[str], None] | None = None,
    ) -> tuple[str, list[str]]:
        """Generate the full spending report (fast single-pass by default, or multi-step tool loop)."""
        if fast_mode:
            compact_ctx = self._build_compact_report_context()
            tools_used = [
                "fast_bundle:overview",
                "fast_bundle:where_you_spend_more",
                "fast_bundle:monthly_spending",
                "fast_bundle:future_forecast",
                "fast_bundle:why_it_happens",
            ]
            if self.on_tool_call:
                self.on_tool_call("build_compact_report_context", {"mode": "fast"})
            fast_prompt = (
                "Here are the exact verified numbers computed from the Excel sheet:\n\n"
                f"{compact_ctx}\n\n"
                "Using ONLY these exact numbers and currency symbol, write a crisp, well-structured Spending Report covering:\n"
                "1. **Spending Overview Report**\n"
                "2. **Where You Spend More** (top categories, merchants, frequent small purchases)\n"
                "3. **Which Month You Spend More (Peak Month & Upcoming Forecast)**\n"
                "4. **Why It Happens** (surging categories, spike transactions, habits)\n"
                "5. **Actionable Savings Plan** (3-5 concrete recommendations)"
            )
            self.messages.append({"role": "user", "content": fast_prompt})
            data = self._call_ollama(
                self.messages,
                use_tools=False,
                num_ctx=4096,
                num_predict=900,
                disable_thinking=True,
                on_token=on_token,
            )
            final_text = (data.get("message", {}).get("content") or "").strip()
            self.messages.append({"role": "assistant", "content": final_text})
            return final_text, tools_used

        tools_used: list[str] = []
        seen_calls: set[tuple[str, str]] = set()
        prompt = (
            "Analyze my Excel expense sheet thoroughly. Use your tools to gather the overview, "
            "where I spend more (categories and merchants), monthly spending breakdown, future month forecast, "
            "and why the peak spending happens. Call each needed tool once, then write the complete Spending Report."
        )
        self.messages.append({"role": "user", "content": prompt})

        for _ in range(6):
            data = self._call_ollama(self.messages, use_tools=self._tool_calls_supported)
            msg = data.get("message", {})
            tool_calls = msg.get("tool_calls") or []

            if tool_calls:
                self.messages.append(msg)
                all_duplicates = True
                for tc in tool_calls:
                    fn = tc.get("function", {})
                    name = fn.get("name", "")
                    raw_args = fn.get("arguments") or {}
                    if isinstance(raw_args, str):
                        try:
                            raw_args = json.loads(raw_args)
                        except json.JSONDecodeError:
                            raw_args = {}
                    call_key = (name, json.dumps(raw_args, sort_keys=True))
                    if call_key in seen_calls:
                        self.messages.append({
                            "role": "tool",
                            "content": json.dumps({"note": "Already called above. Use the earlier tool output to write your final report now."}),
                        })
                        continue
                    all_duplicates = False
                    seen_calls.add(call_key)
                    tools_used.append(name)
                    result_json = self._execute_tool(name, raw_args)
                    self.messages.append({
                        "role": "tool",
                        "content": result_json,
                    })
                if all_duplicates:
                    break
                continue

            content = (msg.get("content") or "").strip()
            # Fallback: if model didn't call any tools on turn 1 (or tools aren't supported),
            # inject exact precomputed analytics bundle so it never hallucinates.
            if not tools_used:
                bundle = self.analyzer.build_full_analytics_bundle()
                tools_used.extend(["bundle:overview", "bundle:where_you_spend_more", "bundle:monthly", "bundle:forecast", "bundle:why_peak_happened"])
                grounded_prompt = (
                    "Here are the exact verified analytics computed from the Excel sheet:\n\n"
                    f"```json\n{json.dumps(bundle, indent=2, ensure_ascii=False)}\n```\n\n"
                    "Using ONLY these exact numbers, write the complete report covering:\n"
                    "1. Spending Overview Report\n"
                    "2. Where You Spend More (Categories & Merchants)\n"
                    "3. Which Month You Spend More (Peak Month & Upcoming Forecast)\n"
                    "4. Why It Happens (Root Causes, Spike Transactions & Habits)\n"
                    "5. Actionable Savings Recommendations"
                )
                self.messages[-1] = {"role": "user", "content": grounded_prompt}
                fallback_data = self._call_ollama(self.messages, use_tools=False)
                fallback_msg = fallback_data.get("message", {})
                final_text = (fallback_msg.get("content") or "").strip()
                self.messages.append({"role": "assistant", "content": final_text})
                return final_text, tools_used

            self.messages.append({"role": "assistant", "content": content})
            return content, tools_used

        # If max iterations reached, ask for final synthesis without tools
        self.messages.append({
            "role": "user",
            "content": "You have gathered the data. Please write the final complete spending report now.",
        })
        final_data = self._call_ollama(self.messages, use_tools=False, num_ctx=8192, num_predict=600, disable_thinking=True)
        final_text = (final_data.get("message", {}).get("content") or "").strip()
        self.messages.append({"role": "assistant", "content": final_text})
        return final_text, tools_used

    def ask(self, question: str) -> str:
        """Answer a follow-up question about the Excel sheet using tool calls if needed."""
        self.messages.append({"role": "user", "content": question})

        if not self._tool_calls_supported:
            data = self._call_ollama(self.messages, use_tools=False, num_ctx=8192, num_predict=600, disable_thinking=True)
            ans = (data.get("message", {}).get("content") or "").strip()
            self.messages.append({"role": "assistant", "content": ans})
            return ans

        for _ in range(5):
            data = self._call_ollama(self.messages, use_tools=True, num_ctx=8192, num_predict=600, disable_thinking=True)
            msg = data.get("message", {})
            tool_calls = msg.get("tool_calls") or []
            if tool_calls:
                self.messages.append(msg)
                for tc in tool_calls:
                    fn = tc.get("function", {})
                    name = fn.get("name", "")
                    raw_args = fn.get("arguments") or {}
                    if isinstance(raw_args, str):
                        try:
                            raw_args = json.loads(raw_args)
                        except json.JSONDecodeError:
                            raw_args = {}
                    result_json = self._execute_tool(name, raw_args)
                    self.messages.append({"role": "tool", "content": result_json})
                continue

            ans = (msg.get("content") or "").strip()
            self.messages.append({"role": "assistant", "content": ans})
            return ans

        final_data = self._call_ollama(self.messages, use_tools=False, num_ctx=8192, num_predict=600, disable_thinking=True)
        ans = (final_data.get("message", {}).get("content") or "").strip()
        self.messages.append({"role": "assistant", "content": ans})
        return ans
