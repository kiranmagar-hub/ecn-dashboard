import sys, re

path = 'C:/Users/kmagar/Claude/projects/BEF ECN CT 0624 update/ECN_Metrics_Dashboard.html'
content = open(path, encoding='utf-8').read()

if 'chart-mgmt-bar' in content:
    print('Already patched - skipping to avoid duplicates.')
    sys.exit(0)

# ── 1. Inject CSS before </style> ──
css = """
        /* ── Chart Management Toolbar ── */
        .chart-mgmt-bar {
            background: white;
            border-radius: 12px;
            box-shadow: 0 4px 15px rgba(0,0,0,0.1);
            padding: 14px 20px;
            margin-bottom: 20px;
            display: flex;
            align-items: center;
            gap: 12px;
            flex-wrap: wrap;
        }
        .chart-mgmt-bar .mgmt-label {
            font-size: 0.85em; font-weight: 700; color: #555;
            text-transform: uppercase; letter-spacing: 0.5px;
        }
        .chart-toggle-btn {
            display: inline-flex; align-items: center; gap: 6px;
            padding: 6px 13px; border-radius: 20px;
            border: 2px solid #667eea; background: white; color: #667eea;
            font-size: 0.78em; font-weight: 600; cursor: pointer;
            transition: all 0.2s ease; white-space: nowrap;
        }
        .chart-toggle-btn:hover { background: #667eea; color: white; }
        .chart-toggle-btn.hidden-btn {
            border-color: #d1d5db; color: #9ca3af;
            background: #f9fafb; text-decoration: line-through;
        }
        .chart-toggle-btn.hidden-btn:hover {
            border-color: #667eea; color: #667eea;
            background: white; text-decoration: none;
        }
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
        .chart-title-row {
            display: flex; align-items: flex-start;
            justify-content: space-between; margin-bottom: 20px; gap: 10px;
        }
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
        @media (max-width: 768px) {
            .chart-mgmt-bar { padding: 10px 14px; gap: 8px; }
            .chart-toggle-btn, .mgmt-action-btn { font-size: 0.72em; padding: 5px 10px; }
        }
        @media (min-width: 1920px) {
            .chart-toggle-btn, .mgmt-action-btn { font-size: 0.88em; padding: 8px 16px; }
        }
"""
content = content.replace('    </style>\n</head>', css + '    </style>\n</head>', 1)
print('CSS injected:', 'chart-mgmt-bar' in content)

# ── Helper: build a wrapped chart div ──
def wrap_chart(canvas_id, title_html, fw=False):
    fw_class = ' full-width' if fw else ''
    return f'''<div class="chart-container{fw_class}" id="wrap-{canvas_id}" draggable="true">
                            <div class="chart-title-row">
                                <h3 class="chart-title">{title_html}</h3>
                                <div class="chart-controls">
                                    <button class="chart-ctrl-btn drag-handle" title="Drag to move">&#8597;</button>
                                    <button class="chart-ctrl-btn" title="Toggle full width" onclick="toggleFullWidth(this)">&#8596;</button>
                                    <button class="chart-ctrl-btn" title="Hide chart" onclick="hideChartFromHandle('{canvas_id}')">&#10005;</button>
                                </div>
                            </div>
                            <canvas id="{canvas_id}"></canvas>
                        </div>'''

# ── Helper: toolbar button ──
def tbtn(canvas_id, label):
    return f'<button class="chart-toggle-btn" data-chart="{canvas_id}" onclick="toggleChart(\'{canvas_id}\',this)">{label}</button>'

# ── 2. Overview grid ──
sp_base = 'https://analog.sharepoint.com/sites/spmig_WWMFGbackendfndypromis/Shared%20Documents/Backend%20Foundry%20ECN%20Metrics'
dl_style = 'font-size: 0.85em; text-decoration: none; font-weight: normal; margin-left: 10px;'

overview_charts = [
    wrap_chart('monthlyTrendsChart',
        f'Monthly Cycle Time Trends <a href="{sp_base}/ECN_Quarterly_Trends.xlsx" download="ECN_Quarterly_Trends.xlsx" style="{dl_style} color:#667eea;">&#128229; Download Excel</a>',
        fw=True),
    wrap_chart('quarterlyTopicTrendsChart',
        'Quarterly Processing Cycle Time Trends by ECN Type (Top 5 Types)', fw=True),
    wrap_chart('topicChart', 'Top 10 ECN Topics by Volume'),
    wrap_chart('topicCTChart', 'Cycle Time by Topic'),
    wrap_chart('ecnsOver100DaysChart',
        f'&#9888;&#65039; ECNs Open for More Than 100 Days (By Type &amp; State) <a href="{sp_base}/ECN_Over_100_Days.xlsx" download="ECN_Over_100_Days.xlsx" style="{dl_style} color:#dc2626;">&#128229; Download Excel</a>',
        fw=True),
    wrap_chart('ecnTypeStateChart',
        f'&#128203; Open ECNs by Type &amp; State (As of Today) <a href="{sp_base}/ECN_Open_ECNs.xlsx" download="ECN_Open_ECNs.xlsx" style="{dl_style} color:#667eea;">&#128229; Download Open ECNs Excel</a>',
        fw=True),
    wrap_chart('stateChart', 'State Distribution', fw=True),
    wrap_chart('rushChart', 'Rush vs Regular Requests'),
    wrap_chart('voidTrendChart', 'Void Rate Trend'),
    wrap_chart('percentile90thChart',
        f'90th Percentile ECNs by Type (Slowest 10% - Avg Processing CT) <a href="{sp_base}/ECN_90th_Percentile.xlsx" download="ECN_90th_Percentile.xlsx" style="{dl_style} color:#667eea;">&#128229; Download Excel</a>',
        fw=True),
]

toolbar_overview = '\n                        '.join([
    tbtn('monthlyTrendsChart', 'Monthly Trends'),
    tbtn('quarterlyTopicTrendsChart', 'Quarterly by Type'),
    tbtn('topicChart', 'Topics by Volume'),
    tbtn('topicCTChart', 'Cycle Time by Topic'),
    tbtn('ecnsOver100DaysChart', 'Over 100 Days'),
    tbtn('ecnTypeStateChart', 'Open by Type/State'),
    tbtn('stateChart', 'State Distribution'),
    tbtn('rushChart', 'Rush vs Regular'),
    tbtn('voidTrendChart', 'Void Rate Trend'),
    tbtn('percentile90thChart', '90th Percentile'),
])

new_overview_block = f'''<!-- Chart Management Toolbar: Overview -->
                    <div class="chart-mgmt-bar" id="overviewMgmtBar">
                        <span class="mgmt-label">&#128065; Charts:</span>
                        {toolbar_overview}
                        <div class="mgmt-divider"></div>
                        <button class="mgmt-action-btn" onclick="showAllCharts('overview-grid')">&#10003; Show All</button>
                        <button class="mgmt-action-btn" onclick="resetLayout('overview-grid')">&#8635; Reset Layout</button>
                    </div>

                    <div class="chart-grid" id="overview-grid">
                        {chr(10).join(overview_charts)}
                    </div>
                </div>'''

# Find and replace the overview chart-grid block
overview_start = content.find('<div class="chart-grid">')
overview_end   = content.find('</div>\n                </div>', overview_start)
old_overview   = content[overview_start:overview_end + len('</div>\n                </div>')]
content = content.replace(old_overview, new_overview_block, 1)
print('Overview grid replaced')

# ── 3. KPI grid ──
kpi_charts = [
    wrap_chart('ftrTrendChart', 'FTR Rate Trend', fw=True),
    wrap_chart('holdTopicChart', 'Hold Rate by Topic'),
    wrap_chart('coordWorkloadChart', 'Coordinator Workload Distribution'),
    wrap_chart('rushTrendChart', 'Rush Rate Trend'),
    wrap_chart('mfgSiteChart', 'Top Manufacturing Sites'),
]

toolbar_kpi = '\n                        '.join([
    tbtn('ftrTrendChart', 'FTR Rate Trend'),
    tbtn('holdTopicChart', 'Hold Rate by Topic'),
    tbtn('coordWorkloadChart', 'Coordinator Workload'),
    tbtn('rushTrendChart', 'Rush Rate Trend'),
    tbtn('mfgSiteChart', 'Mfg Sites'),
])

new_kpi_block = f'''<!-- Additional KPI Charts -->
                    <div class="chart-mgmt-bar" id="kpisMgmtBar">
                        <span class="mgmt-label">&#128065; Charts:</span>
                        {toolbar_kpi}
                        <div class="mgmt-divider"></div>
                        <button class="mgmt-action-btn" onclick="showAllCharts('kpis-grid')">&#10003; Show All</button>
                        <button class="mgmt-action-btn" onclick="resetLayout('kpis-grid')">&#8635; Reset Layout</button>
                    </div>

                    <div class="chart-grid" id="kpis-grid">
                        {chr(10).join(kpi_charts)}
                    </div>
                </div>'''

old_kpi_marker = '                    <!-- Additional KPI Charts -->'
kpi_start = content.find(old_kpi_marker)
kpi_end   = content.find('</div>\n                </div>', kpi_start)
old_kpi   = content[kpi_start:kpi_end + len('</div>\n                </div>')]
content   = content.replace(old_kpi, new_kpi_block, 1)
print('KPI grid replaced')

# ── 4. Inject JS before </body> ──
js = """
    <script>
    // ── Chart Management: Hide/Show + Drag-and-Drop ──
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
            card.addEventListener('dragstart', e => {
                dragSrc = card; card.classList.add('dragging');
                e.dataTransfer.effectAllowed = 'move';
            });
            card.addEventListener('dragend', () => {
                card.classList.remove('dragging');
                grid.querySelectorAll('.chart-container').forEach(c => c.classList.remove('drag-over'));
                dragSrc = null;
            });
            card.addEventListener('dragover', e => {
                e.preventDefault();
                if (dragSrc && dragSrc !== card) {
                    grid.querySelectorAll('.chart-container').forEach(c => c.classList.remove('drag-over'));
                    card.classList.add('drag-over');
                }
            });
            card.addEventListener('dragleave', () => card.classList.remove('drag-over'));
            card.addEventListener('drop', e => {
                e.preventDefault();
                if (dragSrc && dragSrc !== card) {
                    const cards = Array.from(grid.querySelectorAll('.chart-container'));
                    if (cards.indexOf(dragSrc) < cards.indexOf(card)) grid.insertBefore(dragSrc, card.nextSibling);
                    else grid.insertBefore(dragSrc, card);
                }
                card.classList.remove('drag-over');
            });
        });
    }

    document.querySelectorAll('.chart-container[draggable]').forEach(card => {
        card.addEventListener('mousedown', e => {
            card.draggable = e.target.classList.contains('drag-handle');
        });
        card.addEventListener('dragend', () => { card.draggable = true; });
    });

    document.querySelectorAll('.chart-grid').forEach(grid => initDragDrop(grid));
    </script>
"""
# Use rfind to replace the LAST </body> (real HTML one, not the jsPDF template string ones)
last_body_pos = content.rfind('</body>')
content = content[:last_body_pos] + js + '</body>' + content[last_body_pos + len('</body>'):]
print('JS injected:', 'toggleChart' in content)

open(path, 'w', encoding='utf-8').write(content)
print('Saved.')
print('chart-mgmt-bar count:', content.count('chart-mgmt-bar'))
print('wrap- divs count:', content.count('id="wrap-'))
