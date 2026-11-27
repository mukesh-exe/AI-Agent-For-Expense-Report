"""
Ollama-powered Universal Excel Report Analyst + Hybrid CAG & RAG Q&A Agent.

- Supports spending trackers, bank statements, multi-tab monthly workbooks,
  matrix/pivot budgets, and general non-date tabular reports across one or more files.
- Report generation uses fast single-pass grounded synthesis (or --deep multi-tool).
- Follow-up Q&A uses Hybrid CAG (Cache-Augmented Generation) + RAG
  (BM25/Metadata/Optional-Vector Retrieval) with source citations.
"""

from __future__ import annotations

import json
import os
from typing import Any, Callable

import requests

from excel_analyzer import ExpenseAnalyzer
from rag_cag_engine import ReportKnowledgeBase


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
                "description": "Get overall summary across all sheets/files, date range, total spend or key totals, and detected columns.",
                "parameters": {"type": "object", "properties": {}, "required": []},
            },
        },
        {
            "type": "function",
            "function": {
                "name": "get_where_you_spend_more",
                "description": "Get top spending categories/groups, top merchants, most frequent items, and payment modes.",
                "parameters": {
                    "type": "object",
                    "properties": {
                        "top_n": {"type": "integer", "description": "Number of top categories/merchants (default 10)."}
                    },
                    "required": [],
                },
            },
        },
        {
            "type": "function",
            "function": {
                "name": "get_monthly_spending",
                "description": "Get month-by-month spending totals, MoM changes, peak month, and lowest month.",
                "parameters": {"type": "object", "properties": {}, "required": []},
            },
        },
        {
            "type": "function",
            "function": {
                "name": "get_future_month_forecast",
                "description": "Get spending trend, recurring bills, seasonal ranking, and projected spending for the next 3 months.",
                "parameters": {"type": "object", "properties": {}, "required": []},
            },
        },
        {
            "type": "function",
            "function": {
                "name": "explain_why_it_happens",
                "description": "Get root-cause drivers for why spending spiked in a target month (defaults to peak month).",
                "parameters": {
                    "type": "object",
                    "properties": {
                        "target_month": {"type": "string", "description": "Optional YYYY-MM month."}
                    },
                    "required": [],
                },
            },
        },
        {
            "type": "function",
            "function": {
                "name": "search_transactions",
                "description": "Filter transactions by month, category, keyword, or minimum amount.",
                "parameters": {
                    "type": "object",
                    "properties": {
                        "month": {"type": "string", "description": "Filter by YYYY-MM month."},
                        "category": {"type": "string", "description": "Filter by category substring."},
                        "keyword": {"type": "string", "description": "Filter by description/sheet/file keyword."},
                        "min_amount": {"type": "number", "description": "Minimum transaction amount."},
                        "limit": {"type": "integer", "description": "Max transactions to return (default 15)."},
                    },
                    "required": [],
                },
            },
        },
        {
            "type": "function",
            "function": {
                "name": "search_report_knowledge",
                "description": "RAG search across ALL sheets, rows, summaries, and generated report sections.",
                "parameters": {
                    "type": "object",
                    "properties": {
                        "query": {"type": "string", "description": "Search query for the report knowledge base."},
                        "top_k": {"type": "integer", "description": "Number of top chunks to return (default 6)."},
                    },
                    "required": ["query"],
                },
            },
        },
    ]


SYSTEM_PROMPT = """You are an expert Financial & Excel Report Analyst AI Agent running on Ollama with Hybrid CAG + RAG capabilities.

Strict rules:
1. Ground every single number, date, sheet name, and entity ONLY in the verified CAG context, RAG chunks, or tool outputs provided. Never invent or guess figures.
2. Always include the exact currency symbol when quoting monetary amounts.
3. For spending/budget reports, clearly cover:
   - **Overview Report**
   - **Where You Spend More** (top categories, merchants, frequent items)
   - **Which Month You Spend More (Peak Month & Upcoming Forecast)**
   - **Why It Happens** (category surges, spike transactions, habits/drivers)
   - **Actionable Recommendations**
4. For general non-date tabular reports (e.g., Sales, Inventory, P&L), summarize all sheets, top performers/outliers, trends/drivers, and actionable insights.
5. In Q&A mode, cite the sheet/row or report section when referencing specific records, and say clearly if a fact is not present in the uploaded reports.
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
        self.kb = ReportKnowledgeBase(analyzer=analyzer, ollama_host=self.host, available_models=available)
        self.messages: list[dict[str, Any]] = [
            {"role": "system", "content": f"{SYSTEM_PROMPT}\n\n{self.kb.cag_context}"}
        ]
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
            elif name == "search_report_knowledge":
                ret = self.kb.retrieve(query=str(args.get("query", "")), top_k=int(args.get("top_k", 6)))
                res = {
                    "retrieval_mode": ret["retrieval_mode"],
                    "rag_chunks": ret["rag_chunks"],
                    "structured_matches": ret["structured_matches"],
                }
            else:
                res = {"error": f"Unknown tool: {name}"}
        except Exception as exc:
            res = {"error": str(exc)}

        return json.dumps(res, ensure_ascii=False)

    def _build_compact_report_context(self) -> str:
        """Build a token-efficient summary from the CAG context + spike drivers."""
        if self.analyzer.meta.report_mode == "general_report":
            lines = [self.kb.cag_context, "\nSAMPLE TOP ROWS PER SHEET:"]
            for t in self.analyzer.sheet_tables[:5]:
                lines.append(f"--- {t.file_name} :: {t.sheet_name} ---")
                for idx, row in t.dataframe.head(5).iterrows():
                    vals = [f"{c}={row[c]}" for c in t.columns if row[c] is not None and str(row[c]).strip() != ""]
                    lines.append(f"Row {idx + 1}: " + " | ".join(vals[:8]))
            return "\n".join(lines)

        ov = self.analyzer.get_overview()
        cur = ov["currency"]
        why = self.analyzer.explain_why_it_happens()
        lines = [self.kb.cag_context]
        if why["top_transactions_in_month"]:
            lines.append("\nTOP SPIKE TRANSACTIONS IN PEAK MONTH:")
            for t in why["top_transactions_in_month"][:6]:
                lines.append(f"- {t['date']} {t['description']} ({t['category']}): {cur}{t['amount']:,.2f} ({t['share_of_month_percent']}% of month)")
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
        """Generate the full report across all sheets/files, then index it into the RAG knowledge base."""
        if fast_mode:
            compact_ctx = self._build_compact_report_context()
            tools_used = ["cag:global_report_context", "rag:indexed_all_sheets"]
            if self.on_tool_call:
                self.on_tool_call("cag_fast_report", {"sheets": len(self.analyzer.sheet_tables)})

            if self.analyzer.meta.report_mode == "general_report":
                instructions = (
                    "Using ONLY the exact verified CAG summary and sample rows above, write a clear, comprehensive Report covering:\n"
                    "1. **Executive Overview** (files, sheets, row counts, key totals)\n"
                    "2. **Top Categories / Group Breakdown & Where Values Concentrate**\n"
                    "3. **Cross-Sheet Insights, Outliers & Key Drivers**\n"
                    "4. **Actionable Recommendations**"
                )
            else:
                instructions = (
                    "Using ONLY the exact verified CAG summary and peak transactions above, write a crisp, well-structured Spending Report covering:\n"
                    "1. **Spending Overview Report**\n"
                    "2. **Where You Spend More** (top categories, merchants, frequent small purchases)\n"
                    "3. **Which Month You Spend More (Peak Month & Upcoming Forecast)**\n"
                    "4. **Why It Happens** (surging categories, spike transactions, habits)\n"
                    "5. **Actionable Savings Plan** (3-5 concrete recommendations)"
                )

            fast_prompt = f"{compact_ctx}\n\n{instructions}"
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
            self.kb.index_generated_report(final_text)
            return final_text, tools_used

        tools_used = []
        seen_calls: set[tuple[str, str]] = set()
        prompt = (
            "Analyze ALL sheets of the uploaded Excel report(s) thoroughly. Call each needed analytical tool once, "
            "then write the complete Report."
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
                    self.messages.append({"role": "tool", "content": result_json})
                if all_duplicates:
                    break
                continue

            content = (msg.get("content") or "").strip()
            if not tools_used:
                # Fallback to compact CAG grounded generation
                self.messages.pop()
                return self.generate_full_report(fast_mode=True, on_token=on_token)

            self.messages.append({"role": "assistant", "content": content})
            self.kb.index_generated_report(content)
            return content, tools_used

        self.messages.append({"role": "user", "content": "Write the final complete report now using the tool results above."})
        final_data = self._call_ollama(self.messages, use_tools=False, disable_thinking=True)
        final_text = (final_data.get("message", {}).get("content") or "").strip()
        self.messages.append({"role": "assistant", "content": final_text})
        self.kb.index_generated_report(final_text)
        return final_text, tools_used

    def ask_with_sources(self, question: str) -> dict[str, Any]:
        """Answer any question using Hybrid CAG (cached system context) + RAG retrieval (+ tools if needed)."""
        retrieval = self.kb.retrieve(question, top_k=8)
        rag_block_lines = ["[RETRIEVED RAG EVIDENCE FROM REPORTS]"]
        for ch in retrieval["rag_chunks"]:
            rag_block_lines.append(f"- ({ch['source']}, score={ch['score']}): {ch['text']}")
        if retrieval["structured_matches"]:
            rag_block_lines.append("[EXACT STRUCTURED TABLE MATCHES]")
            rag_block_lines.append(json.dumps(retrieval["structured_matches"], ensure_ascii=False))

        augmented_user_msg = (
            f"{chr(10).join(rag_block_lines)}\n\n"
            f"USER QUESTION: {question}\n"
            "Answer concisely and accurately using ONLY the CAG cached summary above and the retrieved RAG evidence. "
            "Include exact amounts, currency, and cite the sheet/row or section where helpful."
        )

        # Keep conversation history compact (last 6 turns + system CAG prefix) so KV cache stays sharp
        convo_tail = self.messages[1:][-6:]
        qa_messages = [self.messages[0], *convo_tail, {"role": "user", "content": augmented_user_msg}]

        ans = ""
        if not self._tool_calls_supported:
            data = self._call_ollama(qa_messages, use_tools=False, num_ctx=8192, num_predict=600, disable_thinking=True)
            ans = (data.get("message", {}).get("content") or "").strip()
        else:
            for _ in range(3):
                data = self._call_ollama(qa_messages, use_tools=True, num_ctx=8192, num_predict=600, disable_thinking=True)
                msg = data.get("message", {})
                tool_calls = msg.get("tool_calls") or []
                if tool_calls:
                    qa_messages.append(msg)
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
                        qa_messages.append({"role": "tool", "content": result_json})
                    continue
                ans = (msg.get("content") or "").strip()
                break
            if not ans:
                final_data = self._call_ollama(qa_messages, use_tools=False, num_ctx=8192, num_predict=600, disable_thinking=True)
                ans = (final_data.get("message", {}).get("content") or "").strip()

        # Store clean user question + assistant answer in persistent conversation memory
        self.messages.append({"role": "user", "content": question})
        self.messages.append({"role": "assistant", "content": ans})

        return {
            "answer": ans,
            "retrieval_mode": retrieval["retrieval_mode"],
            "cag_used": True,
            "total_indexed_chunks": len(self.kb.chunks),
            "citations": [
                {"source": ch["source"], "kind": ch["kind"], "score": ch["score"]}
                for ch in retrieval["rag_chunks"][:5]
            ],
            "structured_matches": retrieval["structured_matches"],
        }

    def ask(self, question: str) -> str:
        """Backwards-compatible Q&A returning answer string."""
        return self.ask_with_sources(question)["answer"]
