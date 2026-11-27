"""
Local Web UI for the Ollama Excel Spending Analyzer Agent.

Uses Python's built-in ThreadingHTTPServer (zero extra web dependencies) so it
starts instantaneously and serves a modern drag-and-drop dashboard + AI chat UI.

Run:
    python web_ui.py
"""

from __future__ import annotations

import argparse
import json
import tempfile
import threading
import uuid
import webbrowser
from email.parser import BytesParser
from email.policy import default as email_policy
from http import HTTPStatus
from http.server import BaseHTTPRequestHandler, ThreadingHTTPServer
from pathlib import Path
from typing import Any
from urllib.parse import urlparse

from create_sample_excel import generate_sample_excel
from excel_analyzer import ExpenseAnalyzer
from ollama_agent import DEFAULT_OLLAMA_HOST, OllamaSpendingAgent, list_ollama_models
from report_generator import save_reports


SESSIONS: dict[str, dict[str, Any]] = {}
SESSIONS_LOCK = threading.Lock()
REPORTS_DIR = Path("reports")


HTML_PAGE = r"""<!DOCTYPE html>
<html lang="en">
<head>
  <meta charset="UTF-8" />
  <meta name="viewport" content="width=device-width, initial-scale=1.0" />
  <title>Ollama Excel Spending Agent</title>
  <style>
    :root {
      --bg: #0b1120;
      --surface: #111827;
      --card: #1e293b;
      --border: #334155;
      --text: #f1f5f9;
      --muted: #94a3b8;
      --accent: #38bdf8;
      --accent-2: #818cf8;
      --peak: #f43f5e;
      --forecast: #a855f7;
      --good: #22c55e;
    }
    * { box-sizing: border-box; }
    body {
      margin: 0; font-family: system-ui, -apple-system, Segoe UI, Roboto, sans-serif;
      background: var(--bg); color: var(--text); line-height: 1.55;
    }
    .app { max-width: 1220px; margin: 0 auto; padding: 24px 18px 60px; }
    .topbar {
      display: flex; flex-wrap: wrap; justify-content: space-between; align-items: center;
      gap: 14px; margin-bottom: 22px; padding-bottom: 16px; border-bottom: 1px solid var(--border);
    }
    .brand h1 { margin: 0; font-size: 1.55rem; display: flex; align-items: center; gap: 10px; }
    .brand p { margin: 4px 0 0; color: var(--muted); font-size: 0.9rem; }
    .pill {
      background: #0ea5e91f; border: 1px solid var(--accent); color: var(--accent);
      padding: 4px 12px; border-radius: 999px; font-size: 0.82rem; font-weight: 600;
    }
    .panel {
      background: var(--card); border: 1px solid var(--border); border-radius: 14px;
      padding: 20px; margin-bottom: 20px;
    }
    .controls {
      display: grid; grid-template-columns: 1.5fr 1fr 0.7fr auto; gap: 14px; align-items: end;
    }
    @media (max-width: 900px) { .controls { grid-template-columns: 1fr; } }
    label { display: block; font-size: 0.82rem; color: var(--muted); margin-bottom: 6px; font-weight: 600; }
    .dropzone {
      border: 2px dashed #475569; border-radius: 10px; padding: 14px 16px;
      background: #0f172a; cursor: pointer; transition: border-color 0.2s, background 0.2s;
      display: flex; align-items: center; justify-content: space-between; gap: 10px;
    }
    .dropzone:hover, .dropzone.dragover { border-color: var(--accent); background: #17255444; }
    .dropzone span { font-size: 0.9rem; color: var(--text); overflow: hidden; text-overflow: ellipsis; white-space: nowrap; }
    select, input[type="text"] {
      width: 100%; padding: 11px 12px; border-radius: 10px; border: 1px solid var(--border);
      background: #0f172a; color: var(--text); font-size: 0.92rem;
    }
    .btn-row { display: flex; gap: 10px; flex-wrap: wrap; }
    button {
      border: none; border-radius: 10px; padding: 11px 18px; font-weight: 600;
      font-size: 0.92rem; cursor: pointer; transition: opacity 0.15s, transform 0.1s;
    }
    button:hover { opacity: 0.92; }
    button:active { transform: scale(0.99); }
    button:disabled { opacity: 0.5; cursor: not-allowed; }
    .btn-primary { background: linear-gradient(135deg, var(--accent), var(--accent-2)); color: #0b1120; }
    .btn-secondary { background: #334155; color: var(--text); }
    .status-bar {
      margin-top: 14px; padding: 10px 14px; border-radius: 10px; background: #0f172a;
      border: 1px solid var(--border); font-size: 0.88rem; color: var(--muted);
      display: none; align-items: center; gap: 10px;
    }
    .status-bar.visible { display: flex; }
    .spinner {
      width: 16px; height: 16px; border: 2px solid #334155; border-top-color: var(--accent);
      border-radius: 50%; animation: spin 0.8s linear infinite; flex-shrink: 0;
    }
    @keyframes spin { to { transform: rotate(360deg); } }
    .hidden { display: none !important; }
    .kpi-grid { display: grid; grid-template-columns: repeat(auto-fit, minmax(220px, 1fr)); gap: 14px; margin-bottom: 20px; }
    .kpi { background: var(--card); border: 1px solid var(--border); border-radius: 12px; padding: 16px; }
    .kpi-label { color: var(--muted); font-size: 0.78rem; text-transform: uppercase; letter-spacing: 0.05em; }
    .kpi-val { font-size: 1.45rem; font-weight: 700; margin-top: 4px; color: #fff; }
    .kpi-sub { font-size: 0.82rem; color: var(--muted); margin-top: 4px; }
    .grid-2 { display: grid; grid-template-columns: repeat(auto-fit, minmax(380px, 1fr)); gap: 18px; margin-bottom: 20px; }
    .bar-row { margin-bottom: 11px; }
    .bar-label { display: flex; justify-content: space-between; font-size: 0.87rem; margin-bottom: 4px; }
    .bar-track { height: 10px; background: #0f172a; border-radius: 999px; overflow: hidden; }
    .bar-fill { height: 100%; background: linear-gradient(90deg, var(--accent), var(--accent-2)); border-radius: 999px; }
    .col-chart { display: flex; align-items: flex-end; gap: 10px; overflow-x: auto; padding: 16px 4px 6px; min-height: 230px; }
    .col-item { flex: 1 0 60px; display: flex; flex-direction: column; align-items: center; }
    .col-val { font-size: 0.72rem; color: var(--muted); margin-bottom: 6px; white-space: nowrap; }
    .col-track { width: 32px; height: 140px; background: #0f172a; border-radius: 6px; display: flex; align-items: flex-end; overflow: hidden; }
    .col-fill { width: 100%; background: var(--accent); border-radius: 6px 6px 0 0; }
    .col-item.peak .col-fill { background: var(--peak); }
    .col-item.peak .col-lbl { color: #fda4af; font-weight: 700; }
    .col-item.forecast .col-fill { background: var(--forecast); opacity: 0.88; }
    .col-item.forecast .col-lbl { color: #d8b4fe; }
    .col-lbl { font-size: 0.75rem; color: var(--muted); margin-top: 6px; }
    .legend { display: flex; gap: 16px; font-size: 0.8rem; color: var(--muted); margin-top: 8px; }
    .dot { display: inline-block; width: 10px; height: 10px; border-radius: 50%; margin-right: 5px; }
    .table-wrap { overflow-x: auto; }
    table { width: 100%; border-collapse: collapse; font-size: 0.88rem; }
    th, td { border-bottom: 1px solid var(--border); padding: 8px 10px; text-align: left; }
    th { color: var(--muted); font-weight: 600; }
    .pos-bad { color: #fda4af; font-weight: 600; }
    .pos-good { color: #86efac; font-weight: 600; }
    .report-header { display: flex; justify-content: space-between; align-items: center; flex-wrap: wrap; gap: 10px; margin-bottom: 12px; }
    .ai-body h1, .ai-body h2, .ai-body h3, .ai-body h4 { color: #7dd3fc; margin-top: 18px; margin-bottom: 8px; }
    .ai-body p { margin: 8px 0; }
    .ai-body ul, .ai-body ol { padding-left: 22px; }
    .ai-body code { background: #0f172a; padding: 2px 6px; border-radius: 4px; }
    .ai-body table { margin: 12px 0; }
    .chat-box { display: flex; flex-direction: column; gap: 12px; }
    .chat-log {
      max-height: 360px; overflow-y: auto; background: #0f172a; border: 1px solid var(--border);
      border-radius: 10px; padding: 14px; display: flex; flex-direction: column; gap: 10px;
    }
    .msg { padding: 10px 14px; border-radius: 10px; max-width: 88%; font-size: 0.92rem; }
    .msg.user { align-self: flex-end; background: #1d4ed8; color: #fff; }
    .msg.bot { align-self: flex-start; background: #1e293b; border: 1px solid var(--border); }
    .chat-input-row { display: flex; gap: 10px; }
    .chat-input-row input { flex: 1; }
    a.dl-link {
      text-decoration: none; background: #334155; color: var(--text); padding: 8px 14px;
      border-radius: 8px; font-size: 0.85rem; font-weight: 600;
    }
    a.dl-link:hover { background: #475569; }
  </style>
</head>
<body>
  <div class="app">
    <div class="topbar">
      <div class="brand">
        <h1>Excel Spending AI Agent</h1>
        <p>Analyze any Excel/CSV expense sheet with local Ollama: full report, where you spend more, peak &amp; future month, and why it happens.</p>
      </div>
      <span class="pill" id="ollamaStatus">Connecting to Ollama...</span>
    </div>

    <div class="panel">
      <div class="controls">
        <div>
          <label>1. Upload Excel / CSV File</label>
          <div class="dropzone" id="dropzone">
            <span id="fileName">Click or drag &amp; drop .xlsx / .xls / .csv here</span>
            <button type="button" class="btn-secondary" style="padding:6px 12px;font-size:0.82rem;">Browse</button>
            <input type="file" id="fileInput" accept=".xlsx,.xlsm,.xls,.csv" class="hidden" />
          </div>
        </div>
        <div>
          <label for="modelSelect">2. Ollama Model</label>
          <select id="modelSelect"></select>
        </div>
        <div>
          <label for="currencyInput">3. Currency (Optional)</label>
          <input type="text" id="currencyInput" placeholder="Auto (₹, $, €)" />
        </div>
        <div class="btn-row">
          <button type="button" class="btn-primary" id="analyzeBtn">Analyze Sheet</button>
          <button type="button" class="btn-secondary" id="demoBtn">Run Demo</button>
        </div>
      </div>
      <div class="status-bar" id="statusBar">
        <div class="spinner" id="statusSpinner"></div>
        <span id="statusText">Working...</span>
      </div>
    </div>

    <div id="resultsSection" class="hidden">
      <div class="kpi-grid" id="kpiGrid"></div>

      <div class="grid-2">
        <div class="panel">
          <h3 style="margin-top:0">Where You Spend More (Categories)</h3>
          <div id="categoryBars"></div>
        </div>
        <div class="panel">
          <h3 style="margin-top:0">Monthly Spending &amp; 3-Month Forecast</h3>
          <div class="col-chart" id="monthlyChart"></div>
          <div class="legend">
            <span><span class="dot" style="background:var(--accent)"></span>Actual Month</span>
            <span><span class="dot" style="background:var(--peak)"></span>Peak Month</span>
            <span><span class="dot" style="background:var(--forecast)"></span>Projected (*Forecast)</span>
          </div>
        </div>
      </div>

      <div class="grid-2">
        <div class="panel">
          <h3 style="margin-top:0" id="whyTitle">Why Peak Month Spiked</h3>
          <div class="table-wrap" id="whyTable"></div>
        </div>
        <div class="panel">
          <h3 style="margin-top:0" id="spikeTitle">Top Spike Transactions</h3>
          <div class="table-wrap" id="spikeTable"></div>
        </div>
      </div>

      <div class="panel">
        <div class="report-header">
          <h2 style="margin:0">Ollama AI Agent Report</h2>
          <div class="btn-row">
            <a class="dl-link" id="dlHtml" href="#" target="_blank">Open Full HTML Report</a>
            <a class="dl-link" id="dlMd" href="#" download>Download Markdown</a>
          </div>
        </div>
        <div class="ai-body" id="aiReport"></div>
      </div>

      <div class="panel">
        <h3 style="margin-top:0">Ask Follow-up Questions (Ollama Agent Chat)</h3>
        <div class="chat-box">
          <div class="chat-log" id="chatLog">
            <div class="msg bot">Ask me anything about your spreadsheet — e.g., "Why did October cost more?", "How much did I spend on Food?", or "Which subscriptions repeat every month?"</div>
          </div>
          <div class="chat-input-row">
            <input type="text" id="chatInput" placeholder="Ask a question about your expenses..." />
            <button type="button" class="btn-primary" id="chatSendBtn">Ask</button>
          </div>
        </div>
      </div>
    </div>
  </div>

  <script>
    let selectedFile = null;
    let currentSessionId = null;

    const dropzone = document.getElementById('dropzone');
    const fileInput = document.getElementById('fileInput');
    const fileNameEl = document.getElementById('fileName');
    const modelSelect = document.getElementById('modelSelect');
    const currencyInput = document.getElementById('currencyInput');
    const analyzeBtn = document.getElementById('analyzeBtn');
    const demoBtn = document.getElementById('demoBtn');
    const statusBar = document.getElementById('statusBar');
    const statusSpinner = document.getElementById('statusSpinner');
    const statusText = document.getElementById('statusText');
    const resultsSection = document.getElementById('resultsSection');

    function setStatus(msg, loading = true, isError = false) {
      statusBar.classList.add('visible');
      statusText.textContent = msg;
      statusText.style.color = isError ? '#fda4af' : 'var(--muted)';
      statusSpinner.style.display = loading ? 'block' : 'none';
    }

    function escapeHtml(s) {
      return String(s ?? '').replace(/[&<>"']/g, c => ({'&':'&amp;','<':'&lt;','>':'&gt;','"':'&quot;',"'":'&#39;'}[c]));
    }

    function fmtMoney(val, cur) {
      if (val === null || val === undefined) return 'N/A';
      return cur + Number(val).toLocaleString(undefined, { minimumFractionDigits: 2, maximumFractionDigits: 2 });
    }

    // Minimal markdown renderer for AI prose & tables
    function renderMarkdown(md) {
      const lines = String(md || '').split(/\r?\n/);
      let out = [];
      let inUl = false, inOl = false, inTable = false, inCode = false;
      const closeLists = () => {
        if (inUl) { out.push('</ul>'); inUl = false; }
        if (inOl) { out.push('</ol>'); inOl = false; }
      };
      const closeTable = () => {
        if (inTable) { out.push('</tbody></table></div>'); inTable = false; }
      };
      const inline = s => escapeHtml(s)
        .replace(/`([^`]+)`/g, '<code>$1</code>')
        .replace(/\*\*([^*]+)\*\*/g, '<strong>$1</strong>')
        .replace(/\*([^*]+)\*/g, '<em>$1</em>');

      for (let i = 0; i < lines.length; i++) {
        const line = lines[i];
        const t = line.trim();
        if (t.startsWith('```')) {
          closeLists(); closeTable();
          out.push(inCode ? '</code></pre>' : '<pre><code>');
          inCode = !inCode;
          continue;
        }
        if (inCode) { out.push(escapeHtml(line)); continue; }
        if (!t) { closeLists(); closeTable(); continue; }

        if (t.startsWith('|') && t.includes('|')) {
          closeLists();
          const cells = t.replace(/^\||\|$/g, '').split('|').map(c => c.trim());
          if (!inTable && i + 1 < lines.length && /^\|?[\s:\-|]+\|?$/.test(lines[i + 1].trim())) {
            out.push('<div class="table-wrap"><table><thead><tr>' + cells.map(c => `<th>${inline(c)}</th>`).join('') + '</tr></thead><tbody>');
            inTable = true; i++; continue;
          } else if (inTable) {
            out.push('<tr>' + cells.map(c => `<td>${inline(c)}</td>`).join('') + '</tr>');
            continue;
          }
        }
        closeTable();

        if (t.startsWith('### ')) { closeLists(); out.push(`<h4>${inline(t.slice(4))}</h4>`); }
        else if (t.startsWith('## ')) { closeLists(); out.push(`<h3>${inline(t.slice(3))}</h3>`); }
        else if (t.startsWith('# ')) { closeLists(); out.push(`<h2>${inline(t.slice(2))}</h2>`); }
        else if (/^[-*•]\s+/.test(t)) {
          if (inOl) { out.push('</ol>'); inOl = false; }
          if (!inUl) { out.push('<ul>'); inUl = true; }
          out.push(`<li>${inline(t.replace(/^[-*•]\s+/, ''))}</li>`);
        } else if (/^\d+\.\s+/.test(t)) {
          if (inUl) { out.push('</ul>'); inUl = false; }
          if (!inOl) { out.push('<ol>'); inOl = true; }
          out.push(`<li>${inline(t.replace(/^\d+\.\s+/, ''))}</li>`);
        } else {
          closeLists();
          out.push(`<p>${inline(t)}</p>`);
        }
      }
      closeLists(); closeTable();
      if (inCode) out.push('</code></pre>');
      return out.join('\n');
    }

    async function loadModels() {
      const pill = document.getElementById('ollamaStatus');
      try {
        const res = await fetch('/api/models');
        const data = await res.json();
        if (!res.ok) throw new Error(data.error || 'Failed to load models');
        modelSelect.innerHTML = '';
        data.models.forEach(m => {
          const opt = document.createElement('option');
          opt.value = m; opt.textContent = m;
          if (m === data.recommended) opt.selected = true;
          modelSelect.appendChild(opt);
        });
        pill.textContent = `Ollama Ready (${data.models.length} models)`;
        pill.style.borderColor = 'var(--good)';
        pill.style.color = '#86efac';
      } catch (err) {
        pill.textContent = 'Ollama Offline';
        pill.style.borderColor = 'var(--peak)';
        pill.style.color = '#fda4af';
      }
    }

    dropzone.addEventListener('click', () => fileInput.click());
    fileInput.addEventListener('change', () => {
      if (fileInput.files.length) {
        selectedFile = fileInput.files[0];
        fileNameEl.textContent = selectedFile.name;
      }
    });
    dropzone.addEventListener('dragover', e => { e.preventDefault(); dropzone.classList.add('dragover'); });
    dropzone.addEventListener('dragleave', () => dropzone.classList.remove('dragover'));
    dropzone.addEventListener('drop', e => {
      e.preventDefault(); dropzone.classList.remove('dragover');
      if (e.dataTransfer.files.length) {
        selectedFile = e.dataTransfer.files[0];
        fileNameEl.textContent = selectedFile.name;
      }
    });

    function renderAnalytics(data) {
      currentSessionId = data.session_id;
      const b = data.bundle;
      const ov = b.overview;
      const where = b.where_you_spend_more;
      const mon = b.monthly_spending;
      const fc = b.future_month_forecast;
      const why = b.why_peak_month_happened;
      const cur = ov.currency;
      const nextPeak = fc.highest_projected_upcoming_month || {};

      document.getElementById('kpiGrid').innerHTML = `
        <div class="kpi">
          <div class="kpi-label">Total Spending</div>
          <div class="kpi-val">${fmtMoney(ov.total_spend, cur)}</div>
          <div class="kpi-sub">${ov.transaction_count} transactions (${ov.date_range.start} to ${ov.date_range.end})</div>
        </div>
        <div class="kpi">
          <div class="kpi-label">Where You Spend Most</div>
          <div class="kpi-val">${escapeHtml(where.top_category.category)}</div>
          <div class="kpi-sub">${fmtMoney(where.top_category.total_spend, cur)} (${where.top_category.share_percent}% of total)</div>
        </div>
        <div class="kpi">
          <div class="kpi-label">Peak Spending Month</div>
          <div class="kpi-val">${escapeHtml(mon.peak_month.month)}</div>
          <div class="kpi-sub">${fmtMoney(mon.peak_month.total_spend, cur)} (${mon.peak_month.vs_monthly_avg_percent >= 0 ? '+' : ''}${mon.peak_month.vs_monthly_avg_percent}% vs avg)</div>
        </div>
        <div class="kpi">
          <div class="kpi-label">Next High-Spend Month Forecast</div>
          <div class="kpi-val">${escapeHtml(nextPeak.month || 'N/A')}</div>
          <div class="kpi-sub">Projected ${fmtMoney(nextPeak.projected_spend, cur)} (Trend: ${escapeHtml(fc.trend_direction)})</div>
        </div>
      `;

      const maxCat = Math.max(...where.categories.map(c => c.total_spend), 1);
      document.getElementById('categoryBars').innerHTML = where.categories.slice(0, 8).map(c => `
        <div class="bar-row">
          <div class="bar-label"><span>${escapeHtml(c.category)}</span><span>${fmtMoney(c.total_spend, cur)} (${c.share_percent}%)</span></div>
          <div class="bar-track"><div class="bar-fill" style="width:${Math.min(100, c.total_spend / maxCat * 100).toFixed(1)}%"></div></div>
        </div>
      `).join('');

      const allVals = [
        ...mon.months_chronological.map(m => m.total_spend),
        ...fc.next_3_months_forecast.map(f => f.projected_spend)
      ];
      const maxMon = Math.max(...allVals, 1);
      const peakM = mon.peak_month.month;
      document.getElementById('monthlyChart').innerHTML =
        mon.months_chronological.map(m => `
          <div class="col-item ${m.month === peakM ? 'peak' : ''}">
            <div class="col-val">${fmtMoney(m.total_spend, cur)}</div>
            <div class="col-track"><div class="col-fill" style="height:${Math.max(5, m.total_spend / maxMon * 100).toFixed(1)}%"></div></div>
            <div class="col-lbl">${escapeHtml(m.month)}</div>
          </div>
        `).join('') +
        fc.next_3_months_forecast.map(f => `
          <div class="col-item forecast">
            <div class="col-val">${fmtMoney(f.projected_spend, cur)}</div>
            <div class="col-track"><div class="col-fill" style="height:${Math.max(5, f.projected_spend / maxMon * 100).toFixed(1)}%"></div></div>
            <div class="col-lbl">${escapeHtml(f.month)}*</div>
          </div>
        `).join('');

      document.getElementById('whyTitle').textContent = `Why ${why.analyzed_month} Spiked (vs Monthly Baseline)`;
      document.getElementById('whyTable').innerHTML = `
        <table>
          <thead><tr><th>Category</th><th>${escapeHtml(why.analyzed_month)}</th><th>Baseline Avg</th><th>Difference</th></tr></thead>
          <tbody>
            ${why.top_category_surges.map(d => `
              <tr>
                <td>${escapeHtml(d.category)}</td>
                <td>${fmtMoney(d.target_month_spend, cur)}</td>
                <td>${fmtMoney(d.baseline_monthly_avg, cur)}</td>
                <td class="${d.difference > 0 ? 'pos-bad' : 'pos-good'}">${d.difference >= 0 ? '+' : ''}${d.difference.toLocaleString(undefined,{minimumFractionDigits:2,maximumFractionDigits:2})} (${d.percent_vs_baseline >= 0 ? '+' : ''}${d.percent_vs_baseline}%)</td>
              </tr>
            `).join('')}
          </tbody>
        </table>
      `;

      document.getElementById('spikeTitle').textContent = `Top Transactions in ${why.analyzed_month}`;
      document.getElementById('spikeTable').innerHTML = `
        <table>
          <thead><tr><th>Date</th><th>Description</th><th>Category</th><th>Amount</th><th>Share</th></tr></thead>
          <tbody>
            ${why.top_transactions_in_month.slice(0, 6).map(t => `
              <tr>
                <td>${escapeHtml(t.date)}</td>
                <td>${escapeHtml(t.description)}</td>
                <td>${escapeHtml(t.category)}</td>
                <td>${fmtMoney(t.amount, cur)}</td>
                <td>${t.share_of_month_percent}%</td>
              </tr>
            `).join('')}
          </tbody>
        </table>
      `;

      document.getElementById('aiReport').innerHTML = '<p style="color:var(--muted)">Generating Ollama AI narrative report...</p>';
      document.getElementById('dlHtml').classList.add('hidden');
      document.getElementById('dlMd').classList.add('hidden');
      resultsSection.classList.remove('hidden');
    }

    async function runAnalysis(useDemo) {
      if (!useDemo && !selectedFile) {
        setStatus('Please select an Excel/CSV file first, or click "Run Demo".', false, true);
        return;
      }
      analyzeBtn.disabled = true;
      demoBtn.disabled = true;
      const chosenModel = modelSelect.value;
      setStatus('Parsing Excel sheet and computing spending analytics...', true);

      try {
        const form = new FormData();
        form.append('model', chosenModel);
        form.append('currency', currencyInput.value.trim());
        form.append('demo', useDemo ? '1' : '0');
        if (!useDemo && selectedFile) form.append('file', selectedFile);

        const parseRes = await fetch('/api/parse', { method: 'POST', body: form });
        const parseData = await parseRes.json();
        if (!parseRes.ok) throw new Error(parseData.error || 'Failed to parse spreadsheet');

        renderAnalytics(parseData);
        setStatus(`Charts ready! Streaming report from Ollama (${chosenModel})...`, true);

        const repRes = await fetch('/api/report_stream', {
          method: 'POST',
          headers: { 'Content-Type': 'application/json' },
          body: JSON.stringify({ session_id: parseData.session_id })
        });
        if (!repRes.ok) {
          const errJson = await repRes.json().catch(() => ({}));
          throw new Error(errJson.error || 'Ollama report generation failed');
        }

        const reader = repRes.body.getReader();
        const decoder = new TextDecoder();
        let buffer = '';
        let streamedMd = '';
        let doneMeta = null;
        const aiReportEl = document.getElementById('aiReport');

        while (true) {
          const { value, done } = await reader.read();
          if (done) break;
          buffer += decoder.decode(value, { stream: true });
          const lines = buffer.split('\n');
          buffer = lines.pop();
          for (const line of lines) {
            if (!line.trim()) continue;
            const ev = JSON.parse(line);
            if (ev.error) throw new Error(ev.error);
            if (ev.token) {
              streamedMd += ev.token;
              aiReportEl.innerHTML = renderMarkdown(streamedMd);
            }
            if (ev.done) {
              doneMeta = ev;
            }
          }
        }

        if (doneMeta) {
          aiReportEl.innerHTML = renderMarkdown(doneMeta.report_markdown || streamedMd);
          const dlHtml = document.getElementById('dlHtml');
          const dlMd = document.getElementById('dlMd');
          dlHtml.href = `/api/download/${encodeURIComponent(doneMeta.html_filename)}`;
          dlMd.href = `/api/download/${encodeURIComponent(doneMeta.md_filename)}`;
          dlHtml.classList.remove('hidden');
          dlMd.classList.remove('hidden');
          setStatus(`Done in ${doneMeta.elapsed_sec}s with ${doneMeta.model}!`, false);
        }
      } catch (err) {
        setStatus(`Error: ${err.message}`, false, true);
      } finally {
        analyzeBtn.disabled = false;
        demoBtn.disabled = false;
      }
    }

    analyzeBtn.addEventListener('click', () => runAnalysis(false));
    demoBtn.addEventListener('click', () => runAnalysis(true));

    async function sendChat() {
      const input = document.getElementById('chatInput');
      const q = input.value.trim();
      if (!q || !currentSessionId) return;
      input.value = '';
      const log = document.getElementById('chatLog');
      log.innerHTML += `<div class="msg user">${escapeHtml(q)}</div>`;
      const botDiv = document.createElement('div');
      botDiv.className = 'msg bot';
      botDiv.textContent = 'Thinking with Ollama...';
      log.appendChild(botDiv);
      log.scrollTop = log.scrollHeight;

      try {
        const res = await fetch('/api/ask', {
          method: 'POST',
          headers: { 'Content-Type': 'application/json' },
          body: JSON.stringify({ session_id: currentSessionId, question: q })
        });
        const data = await res.json();
        if (!res.ok) throw new Error(data.error || 'Chat failed');
        botDiv.innerHTML = renderMarkdown(data.answer);
      } catch (err) {
        botDiv.textContent = `Error: ${err.message}`;
      }
      log.scrollTop = log.scrollHeight;
    }

    document.getElementById('chatSendBtn').addEventListener('click', sendChat);
    document.getElementById('chatInput').addEventListener('keydown', e => {
      if (e.key === 'Enter') sendChat();
    });

    loadModels();
  </script>
</body>
</html>
"""


class SpendingUIHandler(BaseHTTPRequestHandler):
    ollama_host = DEFAULT_OLLAMA_HOST

    def log_message(self, format: str, *args: Any) -> None:
        # Keep console output clean
        return

    def _send_json(self, payload: dict[str, Any], status: int = HTTPStatus.OK) -> None:
        raw = json.dumps(payload, ensure_ascii=False).encode("utf-8")
        self.send_response(status)
        self.send_header("Content-Type", "application/json; charset=utf-8")
        self.send_header("Content-Length", str(len(raw)))
        self.end_headers()
        self.wfile.write(raw)

    def _send_html(self, body: str) -> None:
        raw = body.encode("utf-8")
        self.send_response(HTTPStatus.OK)
        self.send_header("Content-Type", "text/html; charset=utf-8")
        self.send_header("Content-Length", str(len(raw)))
        self.end_headers()
        self.wfile.write(raw)

    def do_GET(self) -> None:
        parsed = urlparse(self.path)
        if parsed.path == "/":
            self._send_html(HTML_PAGE)
            return

        if parsed.path == "/api/models":
            try:
                models = list_ollama_models(self.ollama_host)
                rec = "qwen2.5:latest" if "qwen2.5:latest" in models else (models[0] if models else "")
                self._send_json({"models": models, "recommended": rec})
            except Exception as exc:
                self._send_json({"error": str(exc)}, status=HTTPStatus.BAD_GATEWAY)
            return

        if parsed.path.startswith("/api/download/"):
            fname = Path(parsed.path.split("/api/download/", 1)[1]).name
            target = (REPORTS_DIR / fname).resolve()
            if not target.exists() or target.parent != REPORTS_DIR.resolve():
                self.send_error(HTTPStatus.NOT_FOUND, "Report not found")
                return
            content_type = "text/html; charset=utf-8" if target.suffix == ".html" else "text/markdown; charset=utf-8"
            raw = target.read_bytes()
            self.send_response(HTTPStatus.OK)
            self.send_header("Content-Type", content_type)
            self.send_header("Content-Length", str(len(raw)))
            self.end_headers()
            self.wfile.write(raw)
            return

        self.send_error(HTTPStatus.NOT_FOUND)

    def do_POST(self) -> None:
        parsed = urlparse(self.path)
        if parsed.path == "/api/parse":
            self._handle_parse()
            return
        if parsed.path == "/api/report":
            self._handle_report()
            return
        if parsed.path == "/api/report_stream":
            self._handle_report_stream()
            return
        if parsed.path == "/api/analyze":
            self._handle_analyze()
            return
        if parsed.path == "/api/ask":
            self._handle_ask()
            return
        self.send_error(HTTPStatus.NOT_FOUND)

    @staticmethod
    def _parse_multipart(content_type: str, body: bytes) -> tuple[dict[str, str], dict[str, tuple[str, bytes]]]:
        """Parse multipart/form-data without deprecated cgi module (Python 3.13+ safe)."""
        header_bytes = f"Content-Type: {content_type}\r\nMIME-Version: 1.0\r\n\r\n".encode("utf-8")
        msg = BytesParser(policy=email_policy).parsebytes(header_bytes + body)
        fields: dict[str, str] = {}
        files: dict[str, tuple[str, bytes]] = {}
        for part in msg.iter_parts():
            name = part.get_param("name", header="content-disposition")
            if not name:
                continue
            filename = part.get_filename()
            payload = part.get_payload(decode=True) or b""
            if filename:
                files[name] = (filename, payload)
            else:
                fields[name] = payload.decode("utf-8", errors="replace")
        return fields, files

    def _parse_upload_to_session(self) -> tuple[str, dict[str, Any]]:
        ctype = self.headers.get("Content-Type", "")
        if "multipart/form-data" not in ctype:
            raise ValueError("Expected multipart/form-data")

        length = int(self.headers.get("Content-Length", "0"))
        raw_body = self.rfile.read(length)
        fields, files = self._parse_multipart(ctype, raw_body)

        model = (fields.get("model") or "").strip() or None
        currency = (fields.get("currency") or "").strip() or None
        is_demo = fields.get("demo") == "1"

        if is_demo:
            excel_path = generate_sample_excel("sample_expenses.xlsx")
            stem = "sample_expenses"
        else:
            if "file" not in files:
                raise ValueError("No file uploaded")
            orig_name, file_bytes = files["file"]
            orig_name = Path(orig_name).name
            suffix = Path(orig_name).suffix.lower() or ".xlsx"
            stem = Path(orig_name).stem
            tmp_dir = Path(tempfile.mkdtemp(prefix="excel_agent_"))
            excel_path = tmp_dir / f"{stem}{suffix}"
            excel_path.write_bytes(file_bytes)

        analyzer = ExpenseAnalyzer(excel_path, currency_override=currency)
        bundle = analyzer.build_full_analytics_bundle()
        agent = OllamaSpendingAgent(analyzer=analyzer, model=model, host=self.ollama_host)

        session_id = uuid.uuid4().hex
        with SESSIONS_LOCK:
            SESSIONS[session_id] = {"agent": agent, "bundle": bundle, "stem": stem}
        return session_id, bundle

    def _handle_parse(self) -> None:
        try:
            session_id, bundle = self._parse_upload_to_session()
            self._send_json({"session_id": session_id, "bundle": bundle})
        except ValueError as exc:
            self._send_json({"error": str(exc)}, status=HTTPStatus.BAD_REQUEST)
        except Exception as exc:
            self._send_json({"error": str(exc)}, status=HTTPStatus.INTERNAL_SERVER_ERROR)

    def _handle_report(self) -> None:
        try:
            length = int(self.headers.get("Content-Length", "0"))
            body = json.loads(self.rfile.read(length).decode("utf-8"))
            session_id = body.get("session_id")
            with SESSIONS_LOCK:
                session = SESSIONS.get(session_id)
            if not session:
                self._send_json({"error": "Session expired. Please re-analyze the sheet."}, status=HTTPStatus.NOT_FOUND)
                return

            agent = session["agent"]
            bundle = session["bundle"]
            stem = session["stem"]
            report_md, tools_used = agent.generate_full_report(fast_mode=True)
            md_path, html_path = save_reports(bundle, report_md, agent.model, REPORTS_DIR, stem=stem)
            self._send_json({
                "session_id": session_id,
                "model": agent.model,
                "tools_used": tools_used,
                "report_markdown": report_md,
                "md_filename": md_path.name,
                "html_filename": html_path.name,
            })
        except Exception as exc:
            self._send_json({"error": str(exc)}, status=HTTPStatus.INTERNAL_SERVER_ERROR)

    def _handle_report_stream(self) -> None:
        import time as _time
        try:
            length = int(self.headers.get("Content-Length", "0"))
            body = json.loads(self.rfile.read(length).decode("utf-8"))
            session_id = body.get("session_id")
            with SESSIONS_LOCK:
                session = SESSIONS.get(session_id)
            if not session:
                self._send_json({"error": "Session expired. Please re-analyze the sheet."}, status=HTTPStatus.NOT_FOUND)
                return

            self.send_response(HTTPStatus.OK)
            self.send_header("Content-Type", "application/x-ndjson; charset=utf-8")
            self.send_header("Cache-Control", "no-cache")
            self.end_headers()

            def _emit(obj: dict[str, Any]) -> None:
                line = (json.dumps(obj, ensure_ascii=False) + "\n").encode("utf-8")
                self.wfile.write(line)
                self.wfile.flush()

            t0 = _time.perf_counter()
            agent = session["agent"]
            bundle = session["bundle"]
            stem = session["stem"]
            report_md, tools_used = agent.generate_full_report(
                fast_mode=True,
                on_token=lambda tok: _emit({"token": tok}),
            )
            md_path, html_path = save_reports(bundle, report_md, agent.model, REPORTS_DIR, stem=stem)
            elapsed = round(_time.perf_counter() - t0, 1)
            _emit({
                "done": True,
                "model": agent.model,
                "tools_used": tools_used,
                "report_markdown": report_md,
                "md_filename": md_path.name,
                "html_filename": html_path.name,
                "elapsed_sec": elapsed,
            })
        except Exception as exc:
            try:
                err_line = (json.dumps({"error": str(exc)}) + "\n").encode("utf-8")
                self.wfile.write(err_line)
                self.wfile.flush()
            except Exception:
                pass

    def _handle_analyze(self) -> None:
        try:
            session_id, bundle = self._parse_upload_to_session()
            with SESSIONS_LOCK:
                session = SESSIONS[session_id]
            agent = session["agent"]
            stem = session["stem"]
            report_md, tools_used = agent.generate_full_report(fast_mode=True)
            md_path, html_path = save_reports(bundle, report_md, agent.model, REPORTS_DIR, stem=stem)

            self._send_json({
                "session_id": session_id,
                "model": agent.model,
                "tools_used": tools_used,
                "bundle": bundle,
                "report_markdown": report_md,
                "md_filename": md_path.name,
                "html_filename": html_path.name,
            })
        except Exception as exc:
            self._send_json({"error": str(exc)}, status=HTTPStatus.INTERNAL_SERVER_ERROR)

    def _handle_ask(self) -> None:
        try:
            length = int(self.headers.get("Content-Length", "0"))
            body = json.loads(self.rfile.read(length).decode("utf-8"))
            session_id = body.get("session_id")
            question = (body.get("question") or "").strip()
            if not session_id or not question:
                self._send_json({"error": "Missing session_id or question"}, status=HTTPStatus.BAD_REQUEST)
                return

            with SESSIONS_LOCK:
                session = SESSIONS.get(session_id)
            if not session:
                self._send_json({"error": "Session expired. Please re-analyze the sheet."}, status=HTTPStatus.NOT_FOUND)
                return

            answer = session["agent"].ask(question)
            self._send_json({"answer": answer})
        except Exception as exc:
            self._send_json({"error": str(exc)}, status=HTTPStatus.INTERNAL_SERVER_ERROR)


def run_server(host: str = "127.0.0.1", port: int = 8501, open_browser: bool = True, ollama_host: str = DEFAULT_OLLAMA_HOST) -> None:
    SpendingUIHandler.ollama_host = ollama_host
    server = ThreadingHTTPServer((host, port), SpendingUIHandler)
    url = f"http://{host}:{port}"
    print(f"Ollama Excel Spending Agent UI running at: {url}")
    print("Press Ctrl+C to stop.")
    if open_browser:
        threading.Timer(0.8, lambda: webbrowser.open(url)).start()
    try:
        server.serve_forever()
    except KeyboardInterrupt:
        print("\nShutting down UI server.")
    finally:
        server.server_close()


if __name__ == "__main__":
    parser = argparse.ArgumentParser(description="Start the Web UI for Ollama Excel Spending Agent.")
    parser.add_argument("--host", default="127.0.0.1")
    parser.add_argument("--port", type=int, default=8501)
    parser.add_argument("--no-browser", action="store_true")
    parser.add_argument("--ollama-host", default=DEFAULT_OLLAMA_HOST)
    args = parser.parse_args()
    run_server(host=args.host, port=args.port, open_browser=not args.no_browser, ollama_host=args.ollama_host)
