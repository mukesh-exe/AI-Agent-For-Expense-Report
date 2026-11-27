"""
Hybrid CAG (Cache-Augmented Generation) + RAG (Retrieval-Augmented Generation)
knowledge base for answering any question across all Excel reports and sheets.

- CAG layer: Pre-computes and caches a compact, exact global knowledge summary
  (all sheet profiles, KPIs, category/group totals, monthly series, forecast,
  spike drivers, and generated report sections) in a stable prompt prefix so
  Ollama reuses its KV cache (`keep_alive`) and never misses global aggregates.
- RAG layer: Indexes every sheet overview, analytical section, generated report
  chunk, and raw row across all files/tabs with BM25 + month/number/entity
  matching (plus optional Ollama vector embeddings if an embedding model is
  installed) and retrieves exact cited chunks on every query.
"""

from __future__ import annotations

import calendar
import math
import re
from collections import Counter
from dataclasses import dataclass, field
from typing import Any

import requests

from excel_analyzer import ExpenseAnalyzer, _parse_number


STOPWORDS = {
    "a", "an", "and", "are", "as", "at", "be", "by", "did", "do", "does",
    "for", "from", "had", "has", "have", "how", "i", "in", "is", "it", "me",
    "my", "of", "on", "or", "show", "tell", "than", "that", "the", "their",
    "this", "to", "was", "were", "what", "when", "where", "which", "who",
    "why", "will", "with", "you", "your",
}

EMBEDDING_MODEL_HINTS = (
    "nomic-embed-text",
    "mxbai-embed-text",
    "all-minilm",
    "bge-m3",
    "bge-large",
    "snowflake-arctic-embed",
)

MONTH_NAME_TO_NUM = {}
for _m in range(1, 13):
    MONTH_NAME_TO_NUM[calendar.month_name[_m].lower()] = _m
    MONTH_NAME_TO_NUM[calendar.month_abbr[_m].lower()] = _m


def _tokenize(text: str) -> list[str]:
    raw_tokens = re.findall(r"[a-z0-9₹$€£]+(?:[-./][a-z0-9]+)*", text.lower())
    out: list[str] = []
    for tok in raw_tokens:
        if tok in STOPWORDS or len(tok) <= 1:
            continue
        out.append(tok)
        # Expand month names to numeric month tags so "October" matches "2026-10"
        if tok in MONTH_NAME_TO_NUM:
            out.append(f"m{MONTH_NAME_TO_NUM[tok]:02d}")
        m_ym = re.match(r"^(20\d{2}|19\d{2})[-/](\d{1,2})(?:[-/]\d{1,2})?$", tok)
        if m_ym:
            out.append(f"m{int(m_ym.group(2)):02d}")
            out.append(m_ym.group(1))
    return out


@dataclass
class KnowledgeChunk:
    chunk_id: str
    source: str
    kind: str  # "sheet_summary" | "report_section" | "row" | "analytics"
    text: str
    tokens: list[str] = field(default_factory=list)
    tf: Counter[str] = field(default_factory=Counter)
    embedding: list[float] | None = None


class ReportKnowledgeBase:
    """Builds and queries the Hybrid CAG + RAG knowledge base for all parsed reports."""

    def __init__(
        self,
        analyzer: ExpenseAnalyzer,
        ollama_host: str | None = None,
        available_models: list[str] | None = None,
    ) -> None:
        self.analyzer = analyzer
        self.ollama_host = (ollama_host or "").rstrip("/")
        self.available_models = available_models or []
        self.chunks: list[KnowledgeChunk] = []
        self.idf: dict[str, float] = {}
        self.avg_doc_len: float = 1.0
        self.embedding_model: str | None = self._detect_embedding_model()
        self.cag_context: str = ""
        self._report_indexed: bool = False
        self._build()

    def _detect_embedding_model(self) -> str | None:
        for m in self.available_models:
            low = m.lower()
            if any(hint in low for hint in EMBEDDING_MODEL_HINTS):
                return m
        return None

    def _embed_text(self, text: str) -> list[float] | None:
        if not self.embedding_model or not self.ollama_host:
            return None
        try:
            resp = requests.post(
                f"{self.ollama_host}/api/embeddings",
                json={"model": self.embedding_model, "prompt": text[:1500]},
                timeout=10,
            )
            if resp.status_code == 200:
                vec = resp.json().get("embedding")
                if isinstance(vec, list) and vec:
                    return vec
        except Exception:
            pass
        return None

    @staticmethod
    def _cosine(a: list[float], b: list[float]) -> float:
        dot = sum(x * y for x, y in zip(a, b))
        na = math.sqrt(sum(x * x for x in a))
        nb = math.sqrt(sum(y * y for y in b))
        return dot / (na * nb) if na > 0 and nb > 0 else 0.0

    def _add_chunk(self, chunk_id: str, source: str, kind: str, text: str) -> None:
        toks = _tokenize(text)
        if not toks:
            return
        self.chunks.append(
            KnowledgeChunk(
                chunk_id=chunk_id,
                source=source,
                kind=kind,
                text=text,
                tokens=toks,
                tf=Counter(toks),
            )
        )

    def _finalize_index(self) -> None:
        n_docs = max(1, len(self.chunks))
        df_counts: Counter[str] = Counter()
        total_len = 0
        for ch in self.chunks:
            total_len += len(ch.tokens)
            for term in set(ch.tokens):
                df_counts[term] += 1
        self.avg_doc_len = total_len / n_docs if n_docs else 1.0
        self.idf = {
            term: math.log((n_docs - freq + 0.5) / (freq + 0.5) + 1.0)
            for term, freq in df_counts.items()
        }

        # Optional dense embeddings when a dedicated embedding model is available and chunk count is moderate
        if self.embedding_model and len(self.chunks) <= 350:
            for ch in self.chunks:
                if ch.embedding is None:
                    ch.embedding = self._embed_text(ch.text)

    def _build_cag_context(self) -> str:
        """Build deterministic cached global context (CAG) covering ALL reports & sheets."""
        ov = self.analyzer.get_overview()
        cur = ov["currency"]
        profiles = self.analyzer.get_sheet_profiles()
        lines = [
            "=== CACHED GLOBAL REPORT KNOWLEDGE (CAG) ===",
            f"FILES: {', '.join(ov.get('files_loaded', [ov['file']]))} | Mode: {ov.get('report_mode', 'spending')} | Currency: {cur}",
            f"SHEETS PARSED ({len(profiles)}): "
            + "; ".join(f"{p['file']}:{p['sheet']} ({p['kind']}, {p['row_count']} rows)" for p in profiles),
        ]

        if ov.get("report_mode") == "spending" and not self.analyzer.df.empty:
            where = self.analyzer.get_where_you_spend_more(top_n=8)
            mon = self.analyzer.get_monthly_spending()
            fc = self.analyzer.get_future_month_forecast()
            why = self.analyzer.explain_why_it_happens()

            lines.append(
                f"SPENDING SUMMARY: {ov['date_range']['start']} to {ov['date_range']['end']} "
                f"({ov['date_range']['num_months']} months) | Txns: {ov['transaction_count']} | "
                f"Total Spend: {cur}{ov['total_spend']:,.2f} | Avg/Mo: {cur}{ov['average_monthly_spend']:,.2f}"
            )
            if ov.get("total_income_detected"):
                lines.append(
                    f"INCOME & NET: Income {cur}{ov['total_income_detected']:,.2f} | Net Cashflow {cur}{ov['net_cashflow']:,.2f}"
                )
            lg = ov["largest_single_expense"]
            lines.append(f"LARGEST EXPENSE: {lg['date']} {lg['description']} ({lg['category']}) = {cur}{lg['amount']:,.2f}")

            lines.append("TOP CATEGORIES: " + " | ".join(
                f"{c['category']}: {cur}{c['total_spend']:,.2f} ({c['share_percent']}%, {c['transaction_count']}x)"
                for c in where["categories"]
            ))
            if where["top_merchants_by_spend"]:
                lines.append("TOP MERCHANTS: " + " | ".join(
                    f"{m['description']}: {cur}{m['total_spend']:,.2f} ({m['transaction_count']}x)"
                    for m in where["top_merchants_by_spend"][:6]
                ))
            lines.append("MONTHLY TOTALS: " + " | ".join(
                f"{m['month']}{'*PEAK*' if m['month'] == mon['peak_month']['month'] else ''}: {cur}{m['total_spend']:,.2f} ({m['vs_monthly_avg_percent']:+.1f}%)"
                for m in mon["months_chronological"]
            ))
            if fc["next_3_months_forecast"]:
                lines.append(
                    f"FORECAST ({fc['trend_direction']}): "
                    + " | ".join(f"{f['month']}: {cur}{f['projected_spend']:,.2f}" for f in fc["next_3_months_forecast"])
                )
            if fc["recurring_items"]:
                lines.append(
                    "RECURRING COMMITMENTS: "
                    + ", ".join(f"{r['description']} ({cur}{r['typical_amount']:,.2f}/mo)" for r in fc["recurring_items"][:6])
                )
            lines.append(
                f"PEAK MONTH DRIVERS ({why['analyzed_month']}): +{cur}{why['excess_spend_over_baseline']:,.2f} above baseline "
                f"({cur}{why['other_months_avg_total']:,.2f}/mo). Surges: "
                + "; ".join(f"{d['category']} {d['difference']:+,.2f} ({d['percent_vs_baseline']:+.1f}%)" for d in why["top_category_surges"][:4])
            )
            hab = why["overall_spending_habits"]
            lines.append(
                f"HABITS: Weekend {cur}{hab['avg_spend_per_active_weekend_day']:,.2f}/day vs Weekday {cur}{hab['avg_spend_per_active_weekday']:,.2f}/day | "
                f"Early/Mid/Late month: {hab['month_timing_split_percent']['early_month_days_1_10']}% / "
                f"{hab['month_timing_split_percent']['mid_month_days_11_20']}% / {hab['month_timing_split_percent']['late_month_days_21_31']}%"
            )

        # Include EVERY sheet's schema and exact column totals so ANY non-expense / summary sheet is cached in CAG too!
        for p in profiles:
            num_str = "; ".join(
                f"{n['column']} sum={n['sum']:,.2f}, avg={n['mean']:,.2f}, max={n['max']:,.2f}"
                for n in p["numeric_summary"][:6]
            )
            grp_str = "; ".join(
                f"{g['group']}: {g['total']:,.2f} ({g['share_percent']}%)"
                for g in p["top_group_breakdown"][:6]
            )
            notes = f" | Notes: {' / '.join(p['title_notes'][:2])}" if p["title_notes"] else ""
            lines.append(
                f"SHEET [{p['file']} :: {p['sheet']}] ({p['kind']}, {p['row_count']} rows, cols={p['columns']}){notes}"
            )
            if num_str:
                lines.append(f"  Metrics: {num_str}")
            if grp_str:
                lines.append(f"  Top Groups: {grp_str}")

        return "\n".join(lines)

    def _build(self) -> None:
        self.cag_context = self._build_cag_context()
        self.chunks = []

        # 1. Index per-sheet summary chunks
        for p in self.analyzer.get_sheet_profiles():
            summary_lines = [
                f"Sheet {p['sheet']} from file {p['file']} (type: {p['kind']}, {p['row_count']} rows).",
                f"Columns: {', '.join(str(c) for c in p['columns'])}.",
            ]
            if p["title_notes"]:
                summary_lines.append("Notes: " + " | ".join(p["title_notes"]))
            if p["numeric_summary"]:
                summary_lines.append(
                    "Numeric totals: "
                    + "; ".join(f"{n['column']} sum={n['sum']:,.2f} mean={n['mean']:,.2f} min={n['min']:,.2f} max={n['max']:,.2f}" for n in p["numeric_summary"])
                )
            if p["top_group_breakdown"]:
                summary_lines.append(
                    "Top groups: "
                    + "; ".join(f"{g['group']}={g['total']:,.2f} ({g['share_percent']}%, {g['count']} rows)" for g in p["top_group_breakdown"])
                )
            self._add_chunk(
                chunk_id=f"sheet_summary:{p['file']}:{p['sheet']}",
                source=f"{p['file']} | {p['sheet']} | Summary",
                kind="sheet_summary",
                text="\n".join(summary_lines),
            )

        # 2. Index raw rows across EVERY sheet table for fine-grained RAG retrieval
        for table in self.analyzer.sheet_tables:
            cols = table.columns
            for idx, row in table.dataframe.iterrows():
                pairs = []
                for c in cols:
                    val = row[c]
                    if val is None or str(val).strip() in {"", "nan", "None"}:
                        continue
                    pairs.append(f"{c}: {val}")
                if not pairs:
                    continue
                row_text = f"[{table.file_name} | Sheet: {table.sheet_name} | Row {idx + 1}] " + " | ".join(pairs)
                self._add_chunk(
                    chunk_id=f"raw:{table.file_name}:{table.sheet_name}:{idx}",
                    source=f"{table.file_name} | {table.sheet_name} | Row {idx + 1}",
                    kind="row",
                    text=row_text,
                )

        # 3. Index normalized expense rows (adds auto-category, month YYYY-MM, weekday) if different/enriched
        if not self.analyzer.df.empty:
            cur = self.analyzer.meta.currency
            for idx, r in self.analyzer.df.iterrows():
                dt_str = r["date"].strftime("%Y-%m-%d")
                wk = "Weekend" if r["is_weekend"] else "Weekday"
                text = (
                    f"[Expense | {r['month']} | {dt_str} ({r['day_of_week']}, {wk})] "
                    f"Category: {r['category']} | Description: {r['description']} | "
                    f"Amount: {cur}{r['amount']:,.2f}"
                    + (f" | Mode: {r['payment_mode']}" if r["payment_mode"] else "")
                    + f" | Source: {r['source_file']} ({r['source_sheet']} Row {r['row_num']})"
                )
                self._add_chunk(
                    chunk_id=f"tx:{idx}",
                    source=f"{r['source_file']} | {r['source_sheet']} | Row {r['row_num']}",
                    kind="analytics",
                    text=text,
                )

        self._finalize_index()

    def index_generated_report(self, report_markdown: str) -> None:
        """Index the AI-generated narrative report sections into the RAG store + refresh CAG."""
        if not report_markdown or self._report_indexed:
            return
        sections = re.split(r"(?m)^(?=#{1,3}\s+)", report_markdown.strip())
        for idx, sec in enumerate(sections):
            sec = sec.strip()
            if not sec:
                continue
            first_line = sec.splitlines()[0].lstrip("#").strip()[:60]
            self._add_chunk(
                chunk_id=f"report_sec:{idx}",
                source=f"Generated Report | {first_line}",
                kind="report_section",
                text=sec,
            )
        self._report_indexed = True
        self._finalize_index()

    def _structured_query_fallback(self, query: str) -> list[dict[str, Any]]:
        """Detect explicit month/threshold/keyword filters in user question for exact Pandas aggregation."""
        matches: list[dict[str, Any]] = []
        q_low = query.lower()

        if self.analyzer.df.empty:
            return matches

        df = self.analyzer.df
        cur = self.analyzer.meta.currency

        # Detect month filter
        target_month: str | None = None
        ym_m = re.search(r"\b(20\d{2}-\d{2})\b", q_low)
        if ym_m:
            target_month = ym_m.group(1)
        else:
            for m_name, m_num in MONTH_NAME_TO_NUM.items():
                if re.search(rf"\b{re.escape(m_name)}\b", q_low):
                    sub_months = sorted(df[df["calendar_month_num"] == m_num]["month"].unique())
                    if sub_months:
                        target_month = sub_months[-1]
                    break

        # Detect minimum amount threshold like "above 5000", "over 10,000", "> 2000"
        min_amt: float | None = None
        m_amt = re.search(r"(?:above|over|greater than|more than|>|>=)\s*[₹$€£]?\s*([\d,]+(?:\.\d+)?)", q_low)
        if m_amt:
            min_amt = _parse_number(m_amt.group(1))

        # Detect category match
        matched_cat: str | None = None
        for cat in df["category"].unique():
            cat_low = str(cat).lower()
            if cat_low in q_low or any(part in q_low for part in re.split(r"[\s&/]+", cat_low) if len(part) >= 4):
                matched_cat = str(cat)
                break

        # Run exact filtered aggregation if at least one filter matched
        if target_month or min_amt is not None or matched_cat:
            res = self.analyzer.search_transactions(
                month=target_month,
                category=matched_cat,
                min_amount=min_amt,
                limit=8,
            )
            if res["matching_count"] > 0:
                matches.append({
                    "filter": {
                        "month": target_month,
                        "category": matched_cat,
                        "min_amount": min_amt,
                    },
                    "matching_count": res["matching_count"],
                    "matching_total": f"{cur}{res['matching_total_spend']:,.2f}",
                    "top_rows": res["transactions"],
                })

        return matches

    def retrieve(self, query: str, top_k: int = 8) -> dict[str, Any]:
        """Retrieve top-K RAG chunks + structured filters + cached CAG context for a user query."""
        q_tokens = _tokenize(query)
        k1 = 1.5
        b = 0.75

        q_vec = self._embed_text(query) if self.embedding_model else None

        scored: list[tuple[float, KnowledgeChunk]] = []
        for ch in self.chunks:
            bm25 = 0.0
            doc_len = len(ch.tokens) or 1
            for t in q_tokens:
                tf = ch.tf.get(t, 0)
                if tf == 0:
                    continue
                idf = self.idf.get(t, 0.0)
                denom = tf + k1 * (1.0 - b + b * (doc_len / self.avg_doc_len))
                bm25 += idf * (tf * (k1 + 1.0) / denom)

            # Slight boost to report summaries & sheet summaries when query asks broad overview questions
            if ch.kind in {"sheet_summary", "report_section"} and bm25 > 0:
                bm25 *= 1.15

            score = bm25
            if q_vec is not None and ch.embedding is not None:
                cos = max(0.0, self._cosine(q_vec, ch.embedding))
                score = bm25 + 4.0 * cos

            if score > 0:
                scored.append((score, ch))

        scored.sort(key=lambda x: x[0], reverse=True)
        selected = scored[:top_k]

        # Ensure at least top sheet summaries are included if lexical match was sparse
        if not selected:
            fallback = [ch for ch in self.chunks if ch.kind in {"sheet_summary", "report_section"}][:top_k]
            selected = [(0.1, ch) for ch in fallback]

        structured = self._structured_query_fallback(query)

        return {
            "retrieval_mode": "CAG + Hybrid BM25/Vector RAG" if self.embedding_model else "CAG + BM25 RAG",
            "cag_context": self.cag_context,
            "rag_chunks": [
                {
                    "chunk_id": ch.chunk_id,
                    "source": ch.source,
                    "kind": ch.kind,
                    "score": round(score, 3),
                    "text": ch.text,
                }
                for score, ch in selected
            ],
            "structured_matches": structured,
        }
