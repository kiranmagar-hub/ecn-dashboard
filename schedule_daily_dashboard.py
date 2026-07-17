"""
Daily Scheduler for ECN Metrics Dashboard
Runs the ECN Metrics Generator automatically and optionally uploads to SharePoint
"""
import os
import sys
import subprocess
from datetime import datetime, timedelta
import json

# Configuration
# DB credentials: set ECN_DB_PASSWORD env var, or create a local config.local.py
# that sets CONFIG['db_password'] = 'your_password' (never commit that file)
import os as _os
CONFIG = {
    # Database connection
    'db_server': 'wilmatom1f',
    'db_database': 'WWECNRequests',
    'db_username': 'ECNRequestData',
    'db_password': _os.environ.get('ECN_DB_PASSWORD', ''),
    'db_table': 'Qry_EcnRequestGeneralInfoUpdate',

    # Date range (last 365 days)
    'days_back': 365,

    # Output folder
    'output_base': r'C:\Users\kmagar\OneDrive - Analog Devices, Inc\Documents\ECNMETRICS07022026',

    # SharePoint (using synced folder instead of API)
    'sharepoint_enabled': True,  # Set to True to enable SharePoint upload
    'sharepoint_synced_folder': r'C:\Users\kmagar\Analog Devices, Inc\Backend FNDY Promis Group Site - Backend Foundry ECN Metrics',
}

def _apply_chart_management(path):
    """Inject chart hide/show + drag-drop controls into the generated dashboard."""
    if not os.path.exists(path):
        print(f'Chart mgmt patch: file not found, skipping: {path}')
        return
    print('Applying chart management patch...')
    content = open(path, encoding='utf-8').read()

    sp_base = 'https://analog.sharepoint.com/sites/spmig_WWMFGbackendfndypromis/Shared%20Documents/Backend%20Foundry%20ECN%20Metrics'
    dl_style = 'font-size: 0.85em; text-decoration: none; font-weight: normal; margin-left: 10px;'

    css = """
        /* -- Chart Management Toolbar -- */
        .chart-mgmt-bar {
            background: white; border-radius: 12px;
            box-shadow: 0 4px 15px rgba(0,0,0,0.1);
            padding: 14px 20px; margin-bottom: 20px;
            display: flex; align-items: center; gap: 12px; flex-wrap: wrap;
        }
        .chart-mgmt-bar .mgmt-label { font-size: 0.85em; font-weight: 700; color: #555; text-transform: uppercase; letter-spacing: 0.5px; }
        .chart-toggle-btn {
            display: inline-flex; align-items: center; gap: 6px;
            padding: 6px 13px; border-radius: 20px;
            border: 2px solid #667eea; background: white; color: #667eea;
            font-size: 0.78em; font-weight: 600; cursor: pointer; transition: all 0.2s ease; white-space: nowrap;
        }
        .chart-toggle-btn:hover { background: #667eea; color: white; }
        .chart-toggle-btn.hidden-btn { border-color: #d1d5db; color: #9ca3af; background: #f9fafb; text-decoration: line-through; }
        .chart-toggle-btn.hidden-btn:hover { border-color: #667eea; color: #667eea; background: white; text-decoration: none; }
        .mgmt-divider { width: 1px; height: 24px; background: #e5e7eb; margin: 0 4px; }
        .mgmt-action-btn {
            display: inline-flex; align-items: center; gap: 5px;
            padding: 6px 13px; border-radius: 20px;
            border: 2px solid #e5e7eb; background: white; color: #555;
            font-size: 0.78em; font-weight: 600; cursor: pointer; transition: all 0.2s ease;
        }
        .mgmt-action-btn:hover { border-color: #667eea; color: #667eea; }
        .chart-container { position: relative; transition: opacity 0.3s, transform 0.2s; }
        .chart-container.chart-hidden { display: none !important; }
        .chart-container.drag-over { outline: 3px dashed #667eea; outline-offset: 4px; background: #f0f4ff; }
        .chart-container.dragging { opacity: 0.4; transform: scale(0.98); }
        .chart-title-row { display: flex; align-items: flex-start; justify-content: space-between; margin-bottom: 20px; gap: 10px; }
        .chart-title-row .chart-title { margin-bottom: 0; flex: 1; }
        .chart-controls { display: flex; align-items: center; gap: 6px; opacity: 0; transition: opacity 0.2s; flex-shrink: 0; }
        .chart-container:hover .chart-controls { opacity: 1; }
        .chart-ctrl-btn {
            width: 28px; height: 28px; border-radius: 6px;
            border: 1px solid #e5e7eb; background: white; cursor: pointer;
            display: flex; align-items: center; justify-content: center;
            font-size: 14px; transition: all 0.15s; color: #6b7280;
        }
        .chart-ctrl-btn:hover { background: #f3f4f6; border-color: #667eea; color: #667eea; }
        .chart-ctrl-btn.drag-handle { cursor: grab; }
        .chart-ctrl-btn.drag-handle:active { cursor: grabbing; }
        @media (max-width: 768px) { .chart-mgmt-bar { padding: 10px 14px; gap: 8px; } .chart-toggle-btn, .mgmt-action-btn { font-size: 0.72em; padding: 5px 10px; } }
        @media (min-width: 1920px) { .chart-toggle-btn, .mgmt-action-btn { font-size: 0.88em; padding: 8px 16px; } }
"""

    def wrap_chart(canvas_id, title_html, fw=False):
        fw_class = ' full-width' if fw else ''
        return (f'<div class="chart-container{fw_class}" id="wrap-{canvas_id}" draggable="true">\n'
                f'                            <div class="chart-title-row">\n'
                f'                                <h3 class="chart-title">{title_html}</h3>\n'
                f'                                <div class="chart-controls">\n'
                f'                                    <button class="chart-ctrl-btn drag-handle" title="Drag to move">&#8597;</button>\n'
                f'                                    <button class="chart-ctrl-btn" title="Toggle full width" onclick="toggleFullWidth(this)">&#8596;</button>\n'
                f'                                    <button class="chart-ctrl-btn" title="Hide chart" onclick="hideChartFromHandle(\'{canvas_id}\')">&#10005;</button>\n'
                f'                                </div>\n'
                f'                            </div>\n'
                f'                            <canvas id="{canvas_id}"></canvas>\n'
                f'                        </div>')

    def tbtn(canvas_id, label):
        return f'<button class="chart-toggle-btn" data-chart="{canvas_id}" onclick="toggleChart(\'{canvas_id}\',this)">{label}</button>'

    overview_charts = [
        wrap_chart('monthlyTrendsChart', f'Monthly Cycle Time Trends <a href="{sp_base}/ECN_Quarterly_Trends.xlsx" download="ECN_Quarterly_Trends.xlsx" style="{dl_style} color:#667eea;">&#128229; Download Excel</a>', fw=True),
        wrap_chart('quarterlyTopicTrendsChart', 'Quarterly Processing Cycle Time Trends by ECN Type (Top 5 Types)', fw=True),
        wrap_chart('topicChart', 'Top 10 ECN Topics by Volume'),
        wrap_chart('topicCTChart', 'Cycle Time by Topic'),
        wrap_chart('coordinatorChart', 'Top 10 Coordinators by Volume'),
        wrap_chart('siteChart', 'Site Distribution'),
        wrap_chart('ecnsOver100DaysChart', f'&#9888;&#65039; ECNs Open for More Than 100 Days (By Type &amp; State) <a href="{sp_base}/ECN_Over_100_Days.xlsx" download="ECN_Over_100_Days.xlsx" style="{dl_style} color:#dc2626;">&#128229; Download Excel</a>', fw=True),
        wrap_chart('ecnTypeStateChart', f'&#128203; Open ECNs by Type &amp; State (As of Today) <a href="{sp_base}/ECN_Open_ECNs.xlsx" download="ECN_Open_ECNs.xlsx" style="{dl_style} color:#667eea;">&#128229; Download Open ECNs Excel</a>', fw=True),
        wrap_chart('stateChart', 'State Distribution', fw=True),
        wrap_chart('rushChart', 'Rush vs Regular Requests'),
        wrap_chart('voidTrendChart', 'Void Rate Trend'),
        wrap_chart('percentile90thChart', f'90th Percentile ECNs by Type (Slowest 10% - Avg Processing CT) <a href="{sp_base}/ECN_90th_Percentile.xlsx" download="ECN_90th_Percentile.xlsx" style="{dl_style} color:#667eea;">&#128229; Download Excel</a>', fw=True),
    ]
    toolbar_overview = '\n                        '.join([
        tbtn('monthlyTrendsChart', 'Monthly Trends'), tbtn('quarterlyTopicTrendsChart', 'Quarterly by Type'),
        tbtn('topicChart', 'Topics by Volume'), tbtn('topicCTChart', 'Cycle Time by Topic'),
        tbtn('coordinatorChart', 'Coordinators'), tbtn('siteChart', 'Site Distribution'),
        tbtn('ecnsOver100DaysChart', 'Over 100 Days'), tbtn('ecnTypeStateChart', 'Open by Type/State'),
        tbtn('stateChart', 'State Distribution'), tbtn('rushChart', 'Rush vs Regular'),
        tbtn('voidTrendChart', 'Void Rate Trend'), tbtn('percentile90thChart', '90th Percentile'),
    ])
    new_overview_block = (
        '<!-- Chart Management Toolbar: Overview -->\n'
        '                    <div class="chart-mgmt-bar" id="overviewMgmtBar">\n'
        '                        <span class="mgmt-label">&#128065; Charts:</span>\n'
        f'                        {toolbar_overview}\n'
        '                        <div class="mgmt-divider"></div>\n'
        '                        <button class="mgmt-action-btn" onclick="showAllCharts(\'overview-grid\')">&#10003; Show All</button>\n'
        '                        <button class="mgmt-action-btn" onclick="resetLayout(\'overview-grid\')">&#8635; Reset Layout</button>\n'
        '                    </div>\n\n'
        '                    <div class="chart-grid" id="overview-grid">\n'
        '                        ' + '\n'.join(overview_charts) + '\n'
        '                    </div>\n'
        '                </div>'
    )

    kpi_charts = [
        wrap_chart('ftrTrendChart', 'FTR Rate Trend', fw=True),
        wrap_chart('holdTopicChart', 'Hold Rate by Topic'),
        wrap_chart('coordWorkloadChart', 'Coordinator Workload Distribution'),
        wrap_chart('rushTrendChart', 'Rush Rate Trend'),
        wrap_chart('mfgSiteChart', 'Top Manufacturing Sites'),
    ]
    toolbar_kpi = '\n                        '.join([
        tbtn('ftrTrendChart', 'FTR Rate Trend'), tbtn('holdTopicChart', 'Hold Rate by Topic'),
        tbtn('coordWorkloadChart', 'Coordinator Workload'), tbtn('rushTrendChart', 'Rush Rate Trend'),
        tbtn('mfgSiteChart', 'Mfg Sites'),
    ])
    new_kpi_block = (
        '<!-- Additional KPI Charts -->\n'
        '                    <div class="chart-mgmt-bar" id="kpisMgmtBar">\n'
        '                        <span class="mgmt-label">&#128065; Charts:</span>\n'
        f'                        {toolbar_kpi}\n'
        '                        <div class="mgmt-divider"></div>\n'
        '                        <button class="mgmt-action-btn" onclick="showAllCharts(\'kpis-grid\')">&#10003; Show All</button>\n'
        '                        <button class="mgmt-action-btn" onclick="resetLayout(\'kpis-grid\')">&#8635; Reset Layout</button>\n'
        '                    </div>\n\n'
        '                    <div class="chart-grid" id="kpis-grid">\n'
        '                        ' + '\n'.join(kpi_charts) + '\n'
        '                    </div>\n'
        '                </div>'
    )

    js = """
    <script>
    // -- Chart Management: Hide/Show + Drag-and-Drop --
    const toolbarBtns = {};
    document.querySelectorAll('.chart-toggle-btn').forEach(btn => {
        toolbarBtns[btn.dataset.chart] = btn;
    });
    const originalOrders = {};
    document.querySelectorAll('.chart-grid').forEach(grid => {
        originalOrders[grid.id] = Array.from(grid.children).map(c => c.id);
    });
    function toggleChart(canvasId, btn) {
        const wrap = document.getElementById('wrap-' + canvasId);
        if (!wrap) return;
        const isHidden = wrap.classList.contains('chart-hidden');
        if (isHidden) { wrap.classList.remove('chart-hidden'); btn.classList.remove('hidden-btn'); }
        else          { wrap.classList.add('chart-hidden');    btn.classList.add('hidden-btn'); }
    }
    function hideChartFromHandle(canvasId) {
        const wrap = document.getElementById('wrap-' + canvasId);
        if (!wrap) return;
        wrap.classList.add('chart-hidden');
        const btn = toolbarBtns[canvasId];
        if (btn) btn.classList.add('hidden-btn');
    }
    function showAllCharts(gridId) {
        const grid = document.getElementById(gridId);
        if (!grid) return;
        grid.querySelectorAll('.chart-container.chart-hidden').forEach(w => w.classList.remove('chart-hidden'));
        Object.values(toolbarBtns).forEach(btn => btn.classList.remove('hidden-btn'));
    }
    function toggleFullWidth(btn) {
        const wrap = btn.closest('.chart-container');
        if (wrap) wrap.classList.toggle('full-width');
    }
    function resetLayout(gridId) {
        const grid = document.getElementById(gridId);
        const order = originalOrders[gridId];
        if (!grid || !order) return;
        order.forEach(id => { const el = document.getElementById(id); if (el) grid.appendChild(el); });
    }
    let dragSrc = null;
    function initDragDrop(grid) {
        grid.querySelectorAll('.chart-container[draggable]').forEach(card => {
            card.addEventListener('dragstart', e => { dragSrc = card; card.classList.add('dragging'); e.dataTransfer.effectAllowed = 'move'; });
            card.addEventListener('dragend', () => { card.classList.remove('dragging'); grid.querySelectorAll('.chart-container').forEach(c => c.classList.remove('drag-over')); dragSrc = null; });
            card.addEventListener('dragover', e => { e.preventDefault(); if (dragSrc && dragSrc !== card) { grid.querySelectorAll('.chart-container').forEach(c => c.classList.remove('drag-over')); card.classList.add('drag-over'); } });
            card.addEventListener('dragleave', () => card.classList.remove('drag-over'));
            card.addEventListener('drop', e => { e.preventDefault(); if (dragSrc && dragSrc !== card) { const cards = Array.from(grid.querySelectorAll('.chart-container')); if (cards.indexOf(dragSrc) < cards.indexOf(card)) grid.insertBefore(dragSrc, card.nextSibling); else grid.insertBefore(dragSrc, card); } card.classList.remove('drag-over'); });
        });
    }
    document.querySelectorAll('.chart-container[draggable]').forEach(card => {
        card.addEventListener('mousedown', e => { card.draggable = e.target.classList.contains('drag-handle'); });
        card.addEventListener('dragend', () => { card.draggable = true; });
    });
    document.querySelectorAll('.chart-grid').forEach(grid => initDragDrop(grid));
    </script>
"""

    # Skip if already patched
    if 'chart-mgmt-bar' in content:
        print('Chart mgmt patch: already applied, skipping.')
        return

    # 1. Inject CSS
    content = content.replace('    </style>\n</head>', css + '    </style>\n</head>', 1)

    # 2. Replace overview grid
    overview_start = content.find('<div class="chart-grid">')
    overview_end   = content.find('</div>\n                </div>', overview_start)
    if overview_start != -1 and overview_end != -1:
        old_block = content[overview_start:overview_end + len('</div>\n                </div>')]
        content = content.replace(old_block, new_overview_block, 1)

    # 3. Replace KPI grid
    kpi_marker = '                    <!-- Additional KPI Charts -->'
    kpi_start = content.find(kpi_marker)
    kpi_end   = content.find('</div>\n                </div>', kpi_start)
    if kpi_start != -1 and kpi_end != -1:
        old_kpi = content[kpi_start:kpi_end + len('</div>\n                </div>')]
        content = content.replace(old_kpi, new_kpi_block, 1)

    # 4. Inject JS before the LAST </body> (real HTML one, not jsPDF template strings)
    last_body_pos = content.rfind('</body>')
    content = content[:last_body_pos] + js + '</body>' + content[last_body_pos + len('</body>'):]

    open(path, 'w', encoding='utf-8').write(content)
    print(f'Chart mgmt patch applied: {path}')


def _apply_tile_features(path):
    """Add hover tooltips and click-to-download to stat tiles."""
    if not os.path.exists(path):
        return
    content = open(path, encoding='utf-8').read()
    if 'has-tooltip' in content:
        print('Tile features: already applied, skipping.')
        return

    SP_BASE = 'https://analog.sharepoint.com/sites/spmig_WWMFGbackendfndypromis/Shared%20Documents/Backend%20Foundry%20ECN%20Metrics'

    tile_css = """
        /* -- Tile tooltip & download enhancements -- */
        .stat-card { position: relative; cursor: default; overflow: visible; }
        .stat-card.has-tooltip { cursor: help; }
        .stat-card.has-download { cursor: pointer; }
        .stat-card.has-download:hover { transform: translateY(-3px); box-shadow: 0 8px 25px rgba(0,0,0,0.18); }
        .stat-card.has-download::after {
            content: '\\21E9  Download Excel';
            display: block; margin-top: 8px;
            font-size: 0.72em; font-weight: 600; color: #667eea;
            letter-spacing: 0.3px; opacity: 0; transition: opacity 0.2s;
        }
        .stat-card.has-download:hover::after { opacity: 1; }
        .tile-tooltip {
            display: none; position: absolute; bottom: calc(100% + 10px); left: 50%;
            transform: translateX(-50%); background: #1e293b; color: #f1f5f9;
            border-radius: 10px; padding: 14px 16px; min-width: 240px; max-width: 320px;
            box-shadow: 0 8px 30px rgba(0,0,0,0.35); z-index: 9999;
            font-size: 0.82em; line-height: 1.55; pointer-events: none;
        }
        .tile-tooltip::before {
            content: ''; position: absolute; top: 100%; left: 50%;
            transform: translateX(-50%); border: 7px solid transparent;
            border-top-color: #1e293b;
        }
        .stat-card:hover .tile-tooltip { display: block; }
        .tile-tooltip .tt-row { display: flex; justify-content: space-between; gap: 12px; padding: 3px 0; border-bottom: 1px solid rgba(255,255,255,0.08); }
        .tile-tooltip .tt-row:last-child { border-bottom: none; }
        .tile-tooltip .tt-label { color: #94a3b8; }
        .tile-tooltip .tt-val { font-weight: 700; color: #e2e8f0; white-space: nowrap; }
        .tile-tooltip .tt-head { font-weight: 700; color: #a5b4fc; margin-bottom: 8px; font-size: 0.9em; text-transform: uppercase; letter-spacing: 0.5px; }
        .tile-tooltip .tt-divider { border-top: 1px solid rgba(255,255,255,0.12); margin: 6px 0; }
"""

    old_stats = """                    <div class="stats-grid">
                        <div class="stat-card">
                            <div class="stat-label">Total Requests</div>
                            <div class="stat-value" id="totalRequests">-</div>
                        </div>
                        <div class="stat-card green">
                            <div class="stat-label">Avg Processing CT</div>
                            <div class="stat-value" id="avgProcCT">-</div>
                            <div class="stat-label">days</div>
                        </div>
                        <div class="stat-card orange">
                            <div class="stat-label">Avg Total CT</div>
                            <div class="stat-value" id="avgTotalCT">-</div>
                            <div class="stat-label">days</div>
                        </div>
                        <div class="stat-card purple">
                            <div class="stat-label">Median Total CT</div>
                            <div class="stat-value" id="medianTotalCT">-</div>
                            <div class="stat-label">days</div>
                        </div>
                        <div class="stat-card red">
                            <div class="stat-label">Void Rate</div>
                            <div class="stat-value" id="voidRate">-</div>
                            <div class="stat-label">%</div>
                        </div>
                        <div class="stat-card green">
                            <div class="stat-label">FTR Rate</div>
                            <div class="stat-value" id="ftrRate">-</div>
                            <div class="stat-label">%</div>
                        </div>
                        <div class="stat-card blue">
                            <div class="stat-label">50th Percentile (Proc CT)</div>
                            <div class="stat-value" id="percentile50">-</div>
                            <div class="stat-label">days</div>
                        </div>
                        <div class="stat-card orange">
                            <div class="stat-label">75th Percentile (Proc CT)</div>
                            <div class="stat-value" id="percentile75">-</div>
                            <div class="stat-label">days</div>
                        </div>
                        <div class="stat-card red">
                            <div class="stat-label">90th Percentile (Proc CT)</div>
                            <div class="stat-value" id="percentile90">-</div>
                            <div class="stat-label">days</div>
                        </div>
                    </div>"""

    new_stats = (
        '                    <div class="stats-grid">\n\n'
        '                        <div class="stat-card has-tooltip">\n'
        '                            <div class="stat-label">Total Requests</div>\n'
        '                            <div class="stat-value" id="totalRequests">-</div>\n'
        '                            <div class="tile-tooltip" id="tt-total"><div class="tt-head">Request Breakdown</div>'
        '<div class="tt-row"><span class="tt-label">Closed</span><span class="tt-val" id="tt-closed">-</span></div>'
        '<div class="tt-row"><span class="tt-label">Void</span><span class="tt-val" id="tt-void">-</span></div>'
        '<div class="tt-row"><span class="tt-label">Rush</span><span class="tt-val" id="tt-rush">-</span></div>'
        '<div class="tt-row"><span class="tt-label">On Hold</span><span class="tt-val" id="tt-hold">-</span></div>'
        '<div class="tt-divider"></div>'
        '<div class="tt-row"><span class="tt-label">Date Range</span><span class="tt-val" id="tt-daterange">-</span></div>'
        '</div></div>\n\n'
        '                        <div class="stat-card green has-tooltip">\n'
        '                            <div class="stat-label">Avg Processing CT</div>\n'
        '                            <div class="stat-value" id="avgProcCT">-</div>\n'
        '                            <div class="stat-label">days</div>\n'
        '                            <div class="tile-tooltip"><div class="tt-head">Processing CT Detail</div>'
        '<div class="tt-row"><span class="tt-label">Avg Total CT</span><span class="tt-val" id="tt-avgtotalct">-</span></div>'
        '<div class="tt-row"><span class="tt-label">Median Proc CT</span><span class="tt-val" id="tt-medianprocct">-</span></div>'
        '<div class="tt-row"><span class="tt-label">Max Proc CT</span><span class="tt-val" id="tt-maxprocct">-</span></div>'
        '<div class="tt-divider"></div>'
        '<div class="tt-row"><span class="tt-label">Rush Avg CT</span><span class="tt-val" id="tt-rushct">-</span></div>'
        '<div class="tt-row"><span class="tt-label">With Holds Avg CT</span><span class="tt-val" id="tt-holdct">-</span></div>'
        '</div></div>\n\n'
        '                        <div class="stat-card orange has-tooltip">\n'
        '                            <div class="stat-label">Avg Total CT</div>\n'
        '                            <div class="stat-value" id="avgTotalCT">-</div>\n'
        '                            <div class="stat-label">days</div>\n'
        '                            <div class="tile-tooltip"><div class="tt-head">Slowest ECN Types (Avg Total CT)</div>'
        '<div id="tt-slowtopics">-</div></div></div>\n\n'
        '                        <div class="stat-card purple has-tooltip">\n'
        '                            <div class="stat-label">Median Total CT</div>\n'
        '                            <div class="stat-value" id="medianTotalCT">-</div>\n'
        '                            <div class="stat-label">days</div>\n'
        '                            <div class="tile-tooltip"><div class="tt-head">CT Distribution</div>'
        '<div class="tt-row"><span class="tt-label">50th pct (Proc)</span><span class="tt-val" id="tt-p50">-</span></div>'
        '<div class="tt-row"><span class="tt-label">75th pct (Proc)</span><span class="tt-val" id="tt-p75">-</span></div>'
        '<div class="tt-row"><span class="tt-label">90th pct (Proc)</span><span class="tt-val" id="tt-p90">-</span></div>'
        '<div class="tt-divider"></div>'
        '<div class="tt-row"><span class="tt-label">Median Total CT</span><span class="tt-val" id="tt-mediantotalct">-</span></div>'
        '</div></div>\n\n'
        '                        <div class="stat-card red has-tooltip">\n'
        '                            <div class="stat-label">Void Rate</div>\n'
        '                            <div class="stat-value" id="voidRate">-</div>\n'
        '                            <div class="stat-label">%</div>\n'
        '                            <div class="tile-tooltip"><div class="tt-head">Top Void Reasons</div>'
        '<div id="tt-voidreasons">-</div>'
        '<div class="tt-divider"></div>'
        '<div class="tt-row"><span class="tt-label">Total Voided</span><span class="tt-val" id="tt-voidcount">-</span></div>'
        '</div></div>\n\n'
        '                        <div class="stat-card green has-tooltip">\n'
        '                            <div class="stat-label">FTR Rate</div>\n'
        '                            <div class="stat-value" id="ftrRate">-</div>\n'
        '                            <div class="stat-label">%</div>\n'
        '                            <div class="tile-tooltip"><div class="tt-head">First-Time-Right Detail</div>'
        '<div class="tt-row"><span class="tt-label">FTR Count</span><span class="tt-val" id="tt-ftrcount">-</span></div>'
        '<div class="tt-row"><span class="tt-label">Hold Rate</span><span class="tt-val" id="tt-holdrate">-</span></div>'
        '<div class="tt-row"><span class="tt-label">Hold Count</span><span class="tt-val" id="tt-holdcount">-</span></div>'
        '<div class="tt-divider"></div>'
        '<div class="tt-head" style="margin-top:4px">Top Hold Reason</div>'
        '<div id="tt-holdreasons">-</div>'
        '</div></div>\n\n'
        f'                        <div class="stat-card blue has-download has-tooltip" onclick="window.open(\'{SP_BASE}/ECN_90th_Percentile.xlsx\')" title="Click to download percentile breakdown Excel">\n'
        '                            <div class="stat-label">50th Percentile (Proc CT)</div>\n'
        '                            <div class="stat-value" id="percentile50">-</div>\n'
        '                            <div class="stat-label">days</div>\n'
        '                            <div class="tile-tooltip"><div class="tt-head">Top ECN Types at 50th pct</div>'
        '<div id="tt-p50types">-</div></div></div>\n\n'
        f'                        <div class="stat-card orange has-download has-tooltip" onclick="window.open(\'{SP_BASE}/ECN_90th_Percentile.xlsx\')" title="Click to download percentile breakdown Excel">\n'
        '                            <div class="stat-label">75th Percentile (Proc CT)</div>\n'
        '                            <div class="stat-value" id="percentile75">-</div>\n'
        '                            <div class="stat-label">days</div>\n'
        '                            <div class="tile-tooltip"><div class="tt-head">Top ECN Types at 75th pct</div>'
        '<div id="tt-p75types">-</div></div></div>\n\n'
        f'                        <div class="stat-card red has-download has-tooltip" onclick="window.open(\'{SP_BASE}/ECN_90th_Percentile.xlsx\')" title="Click to download percentile breakdown Excel">\n'
        '                            <div class="stat-label">90th Percentile (Proc CT)</div>\n'
        '                            <div class="stat-value" id="percentile90">-</div>\n'
        '                            <div class="stat-label">days</div>\n'
        '                            <div class="tile-tooltip"><div class="tt-head">Top ECN Types at 90th pct</div>'
        '<div id="tt-p90types">-</div></div></div>\n\n'
        '                    </div>'
    )

    tooltip_js = """
            // -- Populate tile tooltips --
            (function() {
                const kpis   = data.advanced_kpis;
                const rush   = data.rush_comparison;
                const holds  = kpis.hold_analysis;
                const voids  = kpis.void_analysis;
                const ftr    = kpis.ftr_analysis;
                const topics = data.topic_comparison;
                function fmt(v, d=1) { return v != null ? Number(v).toFixed(d) : '-'; }
                function fmtN(v) { return v != null ? Number(v).toLocaleString() : '-'; }
                function shortTopic(t) { return t ? t.replace(/^\\(\\d+[A-Z]?\\)\\s+/, '').replace(/^\\{\\d+\\}\\s+/, '') : t; }
                function rows(items, lFn, vFn, limit=3) {
                    return (items||[]).slice(0,limit).map(r =>
                        '<div class="tt-row"><span class="tt-label">'+ lFn(r) +'</span><span class="tt-val">'+ vFn(r) +'</span></div>'
                    ).join('') || '-';
                }
                document.getElementById('tt-closed').textContent    = fmtN(stats.total_closed);
                document.getElementById('tt-void').textContent      = fmtN(stats.total_void);
                document.getElementById('tt-hold').textContent      = fmtN(holds.hold_count);
                document.getElementById('tt-rush').textContent      = rush.length ? fmtN(rush[0].RequestNum) : '-';
                document.getElementById('tt-avgtotalct').textContent   = fmt(stats.avg_total_ct) + ' days';
                document.getElementById('tt-medianprocct').textContent = fmt(stats.median_proc_ct) + ' days';
                document.getElementById('tt-maxprocct').textContent    = fmt(stats.max_proc_ct,0) + ' days';
                document.getElementById('tt-rushct').textContent       = rush.length ? fmt(rush[0]['ProcCT(days)']) + ' days' : '-';
                const withHold = (holds.ct_comparison||[]).find(r => r.Category === 'With Holds');
                document.getElementById('tt-holdct').textContent = withHold ? fmt(withHold['ProcCT(days)']) + ' days' : '-';
                const slowTopics = [...(topics||[])].sort((a,b)=> b['TotalCT(Days)'] - a['TotalCT(Days)']).slice(0,3);
                document.getElementById('tt-slowtopics').innerHTML = rows(slowTopics, r => shortTopic(r['ECN Topic']), r => fmt(r['TotalCT(Days)']) + 'd');
                document.getElementById('tt-p50').textContent         = fmt(stats.percentile_50_proc_ct) + ' days';
                document.getElementById('tt-p75').textContent         = fmt(stats.percentile_75_proc_ct) + ' days';
                document.getElementById('tt-p90').textContent         = fmt(stats.percentile_90_proc_ct) + ' days';
                document.getElementById('tt-mediantotalct').textContent = fmt(stats.median_total_ct) + ' days';
                document.getElementById('tt-voidcount').textContent = fmtN(voids.void_count);
                document.getElementById('tt-voidreasons').innerHTML = rows(voids.by_reason, r => r.VoidReason.substring(0,30)+(r.VoidReason.length>30?'...':''), r => fmtN(r.RequestNum));
                document.getElementById('tt-ftrcount').textContent  = fmtN(ftr.ftr_count);
                document.getElementById('tt-holdrate').textContent  = fmt(holds.hold_rate) + '%';
                document.getElementById('tt-holdcount').textContent = fmtN(holds.hold_count);
                document.getElementById('tt-holdreasons').innerHTML = rows((holds.by_reason||[]), r => (r.HoldReason||'').substring(0,30)+((r.HoldReason||'').length>30?'...':''), r => fmtN(r.RequestNum), 2);
                document.getElementById('tt-p50types').innerHTML = rows((stats.top_categories_50th||[]), r => shortTopic(r['ECN Topic']), r => fmt(r['ProcCT(days)']) + 'd avg');
                document.getElementById('tt-p75types').innerHTML = rows((stats.top_categories_75th||[]), r => shortTopic(r['ECN Topic']), r => fmt(r['ProcCT(days)']) + 'd avg');
                document.getElementById('tt-p90types').innerHTML = rows((stats.top_categories_90th||[]), r => shortTopic(r['ECN Topic']), r => fmt(r['ProcCT(days)']) + 'd avg');
                const dr = stats.date_range;
                if (dr) {
                    const fd = d => new Date(d).toLocaleDateString('en-US',{month:'short',day:'numeric',year:'numeric'});
                    document.getElementById('tt-daterange').textContent = fd(dr.start) + ' - ' + fd(dr.end);
                }
                document.querySelectorAll('.stat-card').forEach(card => {
                    card.addEventListener('mouseenter', () => {
                        const tt = card.querySelector('.tile-tooltip');
                        if (!tt) return;
                        const rect = card.getBoundingClientRect();
                        if (window.innerHeight - rect.bottom < 200) { tt.style.top = 'auto'; tt.style.bottom = 'calc(100% + 10px)'; }
                        else { tt.style.top = ''; tt.style.bottom = ''; }
                    });
                });
            })();
"""

    anchor = "document.getElementById('percentile90').textContent = stats.percentile_90_proc_ct.toFixed(2);"

    content = content.replace('    </style>\n</head>', tile_css + '    </style>\n</head>', 1)
    if old_stats in content:
        content = content.replace(old_stats, new_stats, 1)
    if anchor in content:
        content = content.replace(anchor, anchor + tooltip_js, 1)

    open(path, 'w', encoding='utf-8').write(content)
    print(f'Tile features applied: {path}')


def generate_dashboard():
    """Run the ECN Metrics Generator to create dashboard"""
    print('='*80)
    print(f'ECN METRICS DAILY GENERATION - {datetime.now().strftime("%Y-%m-%d %H:%M:%S")}')
    print('='*80)

    # Calculate date range
    end_date = datetime.now()
    start_date = end_date - timedelta(days=CONFIG['days_back'])

    print(f'\nDate Range: {start_date.strftime("%Y-%m-%d")} to {end_date.strftime("%Y-%m-%d")}')
    print(f'Database: {CONFIG["db_server"]} / {CONFIG["db_database"]}')
    print(f'Table: {CONFIG["db_table"]}')

    # Import and run the metrics generator
    try:
        import pyodbc
        import pandas as pd

        # Connect to database
        print('\nConnecting to database...')
        conn_str = (
            f"DRIVER={{SQL Server}};"
            f"SERVER={CONFIG['db_server']};"
            f"DATABASE={CONFIG['db_database']};"
            f"UID={CONFIG['db_username']};"
            f"PWD={CONFIG['db_password']}"
        )
        conn = pyodbc.connect(conn_str)

        # Query data
        print('Querying data...')
        query = f"""
        SELECT * FROM {CONFIG['db_table']}
        WHERE SubmitDate >= '{start_date.strftime("%Y-%m-%d")}'
        AND SubmitDate <= '{end_date.strftime("%Y-%m-%d")}'
        """

        df = pd.read_sql(query, conn)
        conn.close()

        print(f'Retrieved {len(df):,} records')

        # Clean up old folders first (keep only current date folder)
        import shutil
        import glob
        import time

        date_stamp = datetime.now().strftime("%Y-%m-%d")
        current_folder_name = f'ECN_Metrics_{date_stamp}'

        print('\nCleaning up old dashboard folders...')
        base_folder = CONFIG['output_base']
        old_folders = glob.glob(os.path.join(base_folder, 'ECN_Metrics_*'))

        deleted_count = 0
        skipped_count = 0

        for old_folder in old_folders:
            folder_name = os.path.basename(old_folder)

            # Skip if this is today's folder (we'll recreate it)
            if folder_name == current_folder_name:
                print(f'  Skipping current date folder: {folder_name}')
                skipped_count += 1
                continue

            try:
                shutil.rmtree(old_folder)
                print(f'  [OK] Deleted: {folder_name}')
                deleted_count += 1
            except PermissionError:
                print(f'  Warning: Skipped (file in use): {folder_name}')
                skipped_count += 1
            except Exception as e:
                print(f'  ✗ Error deleting {folder_name}: {e}')
                skipped_count += 1

        print(f'\nCleanup summary: Deleted {deleted_count}, Skipped {skipped_count}')

        # Create output folder (date only, no time)
        output_folder = os.path.join(CONFIG['output_base'], current_folder_name)

        # If folder exists, remove it first to ensure fresh data
        if os.path.exists(output_folder):
            try:
                shutil.rmtree(output_folder)
                print(f'Removed existing folder for fresh generation')
            except Exception as e:
                print(f'Warning: Could not remove existing folder: {e}')

        os.makedirs(output_folder, exist_ok=True)

        print(f'\nOutput folder: {output_folder}')

        # Save raw data to Excel
        excel_file = os.path.join(output_folder, 'source_data.xlsx')
        print(f'Saving source data to Excel...')
        with pd.ExcelWriter(excel_file, engine='openpyxl') as writer:
            df.to_excel(writer, sheet_name='Document_TB11', index=False)

        # Copy coordinator mapping file
        coordinator_file = r'C:\Users\kmagar\Claude\projects\BEF ECN CT 0624 update\ECN Coordinators.xlsx'
        if os.path.exists(coordinator_file):
            dest_file = os.path.join(output_folder, 'ECN Coordinators.xlsx')
            shutil.copy2(coordinator_file, dest_file)
            print('Copied ECN Coordinators.xlsx')
        else:
            print('Warning: Warning: ECN Coordinators.xlsx not found - will use user IDs instead of names')

        # Copy logo file
        logo_file = r'C:\Users\kmagar\Claude\projects\BEF ECN CT 0624 update\ADI-Logo-RGB-FullColor.png'
        if os.path.exists(logo_file):
            dest_logo = os.path.join(output_folder, 'ADI-Logo-RGB-FullColor.png')
            shutil.copy2(logo_file, dest_logo)
            print('Copied ADI logo')

        # Copy JavaScript libraries for embedding
        project_dir = r'C:\Users\kmagar\Claude\projects\BEF ECN CT 0624 update'
        for lib_file in ['chart.umd.min.js', 'jspdf.umd.min.js']:
            src_lib = os.path.join(project_dir, lib_file)
            if os.path.exists(src_lib):
                dest_lib = os.path.join(output_folder, lib_file)
                shutil.copy2(src_lib, dest_lib)
                print(f'Copied {lib_file}')

        # Run process_data.py to generate metrics
        print('\nGenerating data.json...')

        # Change to output folder so process_data.py outputs there
        original_dir = os.getcwd()
        os.chdir(output_folder)

        try:
            # Import and run process_data
            import importlib.util
            spec = importlib.util.spec_from_file_location("process_data", os.path.join(original_dir, "process_data.py"))
            process_data = importlib.util.module_from_spec(spec)
            spec.loader.exec_module(process_data)

            print('[OK] data.json created')

            # Run create_enhanced_dashboard.py to generate HTML
            print('\nGenerating HTML dashboard...')
            spec2 = importlib.util.spec_from_file_location("create_dashboard", os.path.join(original_dir, "create_enhanced_dashboard.py"))
            create_dashboard = importlib.util.module_from_spec(spec2)
            spec2.loader.exec_module(create_dashboard)

            print('[OK] Dashboard HTML created')

            # Generate Executive Summary PDF
            print('\nGenerating Executive Summary PDF...')
            spec3 = importlib.util.spec_from_file_location("generate_pdf", os.path.join(original_dir, "generate_executive_summary_pdf.py"))
            generate_pdf = importlib.util.module_from_spec(spec3)
            spec3.loader.exec_module(generate_pdf)

            # Read data.json and generate PDF
            with open('data.json', 'r') as f:
                data_for_pdf = json.load(f)

            generate_pdf.generate_executive_summary_pdf(data_for_pdf, 'ECN_Executive_Summary.pdf')
            print('[OK] Executive Summary PDF created')

            # Run outlier analysis
            print('\nRunning Outlier Analysis...')
            spec4 = importlib.util.spec_from_file_location("analyze_outliers", os.path.join(original_dir, "analyze_outliers_for_optimization.py"))
            analyze_outliers = importlib.util.module_from_spec(spec4)
            spec4.loader.exec_module(analyze_outliers)
            print('[OK] Outlier Analysis completed')

        finally:
            # Change back to original directory
            os.chdir(original_dir)

        print(f'\nDashboard generated successfully!')
        print(f'Location: {output_folder}')

        # Apply chart management patch (hide/show + drag-drop)
        dashboard_file = os.path.join(output_folder, 'ECN_Metrics_Dashboard.html')
        _apply_chart_management(dashboard_file)

        # Apply tile tooltip + percentile download enhancements
        _apply_tile_features(dashboard_file)

        # Find the dashboard HTML file (re-confirm path)

        print(f'\nChecking for dashboard file: {dashboard_file}')
        print(f'File exists: {os.path.exists(dashboard_file)}')
        print(f'SharePoint enabled: {CONFIG["sharepoint_enabled"]}')

        if os.path.exists(dashboard_file):
            print(f'\nDashboard: {dashboard_file}')

            # Upload to SharePoint if enabled
            if CONFIG['sharepoint_enabled']:
                print('Initiating SharePoint upload...')
                upload_to_sharepoint(dashboard_file, output_folder)
            else:
                print('SharePoint upload is disabled in CONFIG')
        else:
            print('WARNING: Dashboard HTML file not found!')

        return output_folder

    except Exception as e:
        print(f'\nERROR: {e}')
        import traceback
        traceback.print_exc()
        return None

def upload_to_sharepoint(dashboard_file, output_folder):
    """Copy dashboard to synced SharePoint folder (OneDrive will auto-upload)"""
    print('\n' + '='*80)
    print('SHAREPOINT SYNC')
    print('='*80)

    try:
        import shutil
        import glob

        sharepoint_folder = CONFIG['sharepoint_synced_folder']

        # Check if SharePoint folder exists
        if not os.path.exists(sharepoint_folder):
            print(f'\nWarning: SharePoint folder not found: {sharepoint_folder}')
            print('Please make sure the SharePoint folder is synced to your PC.')
            return

        print(f'\nSharePoint folder: {sharepoint_folder}')

        # Copy all files from output folder to SharePoint
        files_to_copy = [
            'ECN_Metrics_Dashboard.html',
            'ECN_Executive_Summary.pdf',
            'ADI-Logo-RGB-FullColor.png',
            'source_data.xlsx',
            'data.json',
            'ECN_90th_Percentile.xlsx',
            'ECN_Void_by_Reason.xlsx',
            'ECN_Quarterly_Trends.xlsx',
            'ECN_Open_ECNs.xlsx',
            'ECN_Over_100_Days.xlsx'
        ]

        # Add outlier analysis file with date pattern matching
        import glob
        outlier_files = glob.glob(os.path.join(output_folder, 'ECN_Outlier_Analysis_*.xlsx'))
        if outlier_files:
            # Get the most recent outlier file
            latest_outlier = max(outlier_files, key=os.path.getctime)
            # Copy it with a standard name
            standard_outlier_name = 'ECN_Outlier_Analysis.xlsx'
            shutil.copy2(latest_outlier, os.path.join(output_folder, standard_outlier_name))
            files_to_copy.append(standard_outlier_name)

        copied_count = 0
        for filename in files_to_copy:
            source_file = os.path.join(output_folder, filename)
            if os.path.exists(source_file):
                dest_file = os.path.join(sharepoint_folder, filename)
                shutil.copy2(source_file, dest_file)
                print(f'  [OK] Copied: {filename}')
                copied_count += 1
            else:
                print(f'  Warning: Not found: {filename}')

        print(f'\n[OK] SharePoint sync complete! ({copied_count} files copied)')
        print(f'OneDrive will automatically upload to SharePoint.')
        print(f'\nSharePoint folder: {sharepoint_folder}')

    except Exception as e:
        print(f'\nSharePoint sync failed: {e}')
        print(f'Error type: {type(e).__name__}')
        import traceback
        traceback.print_exc()

def setup_windows_task_scheduler():
    """Generate instructions for setting up Windows Task Scheduler"""
    print('\n' + '='*80)
    print('WINDOWS TASK SCHEDULER SETUP INSTRUCTIONS')
    print('='*80)

    script_path = os.path.abspath(__file__)
    python_path = sys.executable

    print(f'''
To run this script daily automatically:

1. Open Windows Task Scheduler
   - Press Win+R, type "taskschd.msc", press Enter

2. Click "Create Basic Task..."
   - Name: "ECN Metrics Daily Dashboard"
   - Description: "Generate ECN metrics dashboard daily"

3. Trigger: Daily
   - Start time: 6:00 AM (or your preferred time)
   - Recur every: 1 days

4. Action: Start a program
   - Program/script: {python_path}
   - Arguments: "{script_path}"
   - Start in: {os.path.dirname(script_path)}

5. Click Finish

The dashboard will be generated automatically every day at 6 AM.
Output location: {CONFIG['output_base']}

To enable SharePoint upload:
1. Install: pip install Office365-REST-Python-Client
2. Edit CONFIG in this file:
   - Set sharepoint_enabled = True
   - Fill in your SharePoint site URL
   - Fill in your credentials
''')

if __name__ == '__main__':
    import argparse

    parser = argparse.ArgumentParser(description='ECN Metrics Daily Dashboard Generator')
    parser.add_argument('--setup', action='store_true', help='Show Windows Task Scheduler setup instructions')
    parser.add_argument('--run', action='store_true', help='Run dashboard generation now')

    args = parser.parse_args()

    if args.setup:
        setup_windows_task_scheduler()
    elif args.run or len(sys.argv) == 1:
        # Run by default
        output_folder = generate_dashboard()
        if output_folder:
            print(f'\n{"="*80}')
            print('SUCCESS!')
            print('='*80)
            print(f'\nDashboard created in: {output_folder}')
    else:
        parser.print_help()
