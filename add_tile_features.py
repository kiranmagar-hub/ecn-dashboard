"""
Adds to the ECN dashboard stat tiles:
  1. Hover tooltips with drill-down detail on the 6 KPI tiles
  2. Click-to-download Excel on the 3 percentile tiles
Then copies the result to the SharePoint sync folder.
"""
import sys, shutil
sys.stdout.reconfigure(encoding='utf-8', errors='replace')

PROJ  = r'C:\Users\kmagar\Claude\projects\BEF ECN CT 0624 update\ECN_Metrics_Dashboard.html'
SP    = r'C:\Users\kmagar\Analog Devices, Inc\Backend FNDY Promis Group Site - Backend Foundry ECN Metrics\ECN_Metrics_Dashboard.html'
SP_BASE = 'https://analog.sharepoint.com/sites/spmig_WWMFGbackendfndypromis/Shared%20Documents/Backend%20Foundry%20ECN%20Metrics'

c = open(PROJ, encoding='utf-8').read()

# ── 1. CSS ───────────────────────────────────────────────────────────────────
TILE_CSS = """
        /* ── Tile tooltip & download enhancements ── */
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

        /* tooltip bubble — default: above tile, arrow points down */
        .tile-tooltip {
            display: none;
            position: absolute; bottom: calc(100% + 10px); left: 50%;
            transform: translateX(-50%);
            background: #1e293b; color: #f1f5f9;
            border-radius: 10px; padding: 14px 16px;
            min-width: 240px; max-width: 320px;
            box-shadow: 0 8px 30px rgba(0,0,0,0.35);
            z-index: 9999; font-size: 0.82em; line-height: 1.55;
            pointer-events: none;
        }
        .tile-tooltip::before {
            content: '';
            position: absolute; top: 100%; left: 50%;
            transform: translateX(-50%);
            border: 7px solid transparent;
            border-top-color: #1e293b;
        }
        /* flip: below tile, arrow points up */
        .tile-tooltip.tt-below {
            bottom: auto; top: calc(100% + 10px);
        }
        .tile-tooltip.tt-below::before {
            top: auto; bottom: 100%;
            border-top-color: transparent;
            border-bottom-color: #1e293b;
        }
        .stat-card:hover .tile-tooltip { display: block; }
        .tile-tooltip .tt-row {
            display: flex; justify-content: space-between; gap: 12px;
            padding: 3px 0; border-bottom: 1px solid rgba(255,255,255,0.08);
        }
        .tile-tooltip .tt-row:last-child { border-bottom: none; }
        .tile-tooltip .tt-label { color: #94a3b8; }
        .tile-tooltip .tt-val { font-weight: 700; color: #e2e8f0; white-space: nowrap; }
        .tile-tooltip .tt-head {
            font-weight: 700; color: #a5b4fc; margin-bottom: 8px;
            font-size: 0.9em; text-transform: uppercase; letter-spacing: 0.5px;
        }
        .tile-tooltip .tt-divider { border-top: 1px solid rgba(255,255,255,0.12); margin: 6px 0; }
"""

c = c.replace('    </style>\n</head>', TILE_CSS + '    </style>\n</head>', 1)
print('CSS injected')

# ── 2. Replace stat-card HTML blocks ─────────────────────────────────────────
OLD_STATS = """                    <div class="stats-grid">
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

NEW_STATS = f"""                    <div class="stats-grid">

                        <!-- Total Requests -->
                        <div class="stat-card has-tooltip">
                            <div class="stat-label">Total Requests</div>
                            <div class="stat-value" id="totalRequests">-</div>
                            <div class="tile-tooltip" id="tt-total">
                                <div class="tt-head">Request Breakdown</div>
                                <div class="tt-row"><span class="tt-label">Closed</span><span class="tt-val" id="tt-closed">-</span></div>
                                <div class="tt-row"><span class="tt-label">Void</span><span class="tt-val" id="tt-void">-</span></div>
                                <div class="tt-row"><span class="tt-label">Rush</span><span class="tt-val" id="tt-rush">-</span></div>
                                <div class="tt-row"><span class="tt-label">On Hold</span><span class="tt-val" id="tt-hold">-</span></div>
                                <div class="tt-divider"></div>
                                <div class="tt-row"><span class="tt-label">Date Range</span><span class="tt-val" id="tt-daterange">-</span></div>
                            </div>
                        </div>

                        <!-- Avg Processing CT -->
                        <div class="stat-card green has-tooltip">
                            <div class="stat-label">Avg Processing CT</div>
                            <div class="stat-value" id="avgProcCT">-</div>
                            <div class="stat-label">days</div>
                            <div class="tile-tooltip">
                                <div class="tt-head">Processing CT Detail</div>
                                <div class="tt-row"><span class="tt-label">Avg Total CT</span><span class="tt-val" id="tt-avgtotalct">-</span></div>
                                <div class="tt-row"><span class="tt-label">Median Proc CT</span><span class="tt-val" id="tt-medianprocct">-</span></div>
                                <div class="tt-row"><span class="tt-label">Max Proc CT</span><span class="tt-val" id="tt-maxprocct">-</span></div>
                                <div class="tt-divider"></div>
                                <div class="tt-row"><span class="tt-label">Rush Avg CT</span><span class="tt-val" id="tt-rushct">-</span></div>
                                <div class="tt-row"><span class="tt-label">With Holds Avg CT</span><span class="tt-val" id="tt-holdct">-</span></div>
                            </div>
                        </div>

                        <!-- Avg Total CT -->
                        <div class="stat-card orange has-tooltip">
                            <div class="stat-label">Avg Total CT</div>
                            <div class="stat-value" id="avgTotalCT">-</div>
                            <div class="stat-label">days</div>
                            <div class="tile-tooltip">
                                <div class="tt-head">Slowest ECN Types (Avg Total CT)</div>
                                <div id="tt-slowtopics">-</div>
                            </div>
                        </div>

                        <!-- Median Total CT -->
                        <div class="stat-card purple has-tooltip">
                            <div class="stat-label">Median Total CT</div>
                            <div class="stat-value" id="medianTotalCT">-</div>
                            <div class="stat-label">days</div>
                            <div class="tile-tooltip">
                                <div class="tt-head">CT Distribution</div>
                                <div class="tt-row"><span class="tt-label">50th pct (Proc)</span><span class="tt-val" id="tt-p50">-</span></div>
                                <div class="tt-row"><span class="tt-label">75th pct (Proc)</span><span class="tt-val" id="tt-p75">-</span></div>
                                <div class="tt-row"><span class="tt-label">90th pct (Proc)</span><span class="tt-val" id="tt-p90">-</span></div>
                                <div class="tt-divider"></div>
                                <div class="tt-row"><span class="tt-label">Median Total CT</span><span class="tt-val" id="tt-mediantotalct">-</span></div>
                            </div>
                        </div>

                        <!-- Void Rate -->
                        <div class="stat-card red has-tooltip">
                            <div class="stat-label">Void Rate</div>
                            <div class="stat-value" id="voidRate">-</div>
                            <div class="stat-label">%</div>
                            <div class="tile-tooltip">
                                <div class="tt-head">Top Void Reasons</div>
                                <div id="tt-voidreasons">-</div>
                                <div class="tt-divider"></div>
                                <div class="tt-row"><span class="tt-label">Total Voided</span><span class="tt-val" id="tt-voidcount">-</span></div>
                            </div>
                        </div>

                        <!-- FTR Rate -->
                        <div class="stat-card green has-tooltip">
                            <div class="stat-label">FTR Rate</div>
                            <div class="stat-value" id="ftrRate">-</div>
                            <div class="stat-label">%</div>
                            <div class="tile-tooltip">
                                <div class="tt-head">First-Time-Right Detail</div>
                                <div class="tt-row"><span class="tt-label">FTR Count</span><span class="tt-val" id="tt-ftrcount">-</span></div>
                                <div class="tt-row"><span class="tt-label">Hold Rate</span><span class="tt-val" id="tt-holdrate">-</span></div>
                                <div class="tt-row"><span class="tt-label">Hold Count</span><span class="tt-val" id="tt-holdcount">-</span></div>
                                <div class="tt-divider"></div>
                                <div class="tt-head" style="margin-top:4px">Top Hold Reason</div>
                                <div id="tt-holdreasons">-</div>
                            </div>
                        </div>

                        <!-- 50th Percentile — click downloads Excel -->
                        <div class="stat-card blue has-download" onclick="window.open('{SP_BASE}/ECN_90th_Percentile.xlsx')" title="Click to download percentile breakdown Excel">
                            <div class="stat-label">50th Percentile (Proc CT)</div>
                            <div class="stat-value" id="percentile50">-</div>
                            <div class="stat-label">days</div>
                            <div class="tile-tooltip">
                                <div class="tt-head">Top ECN Types at 50th pct</div>
                                <div id="tt-p50types">-</div>
                            </div>
                        </div>

                        <!-- 75th Percentile — click downloads Excel -->
                        <div class="stat-card orange has-download has-tooltip" onclick="window.open('{SP_BASE}/ECN_90th_Percentile.xlsx')" title="Click to download percentile breakdown Excel">
                            <div class="stat-label">75th Percentile (Proc CT)</div>
                            <div class="stat-value" id="percentile75">-</div>
                            <div class="stat-label">days</div>
                            <div class="tile-tooltip">
                                <div class="tt-head">Top ECN Types at 75th pct</div>
                                <div id="tt-p75types">-</div>
                            </div>
                        </div>

                        <!-- 90th Percentile — click downloads Excel -->
                        <div class="stat-card red has-download has-tooltip" onclick="window.open('{SP_BASE}/ECN_90th_Percentile.xlsx')" title="Click to download percentile breakdown Excel">
                            <div class="stat-label">90th Percentile (Proc CT)</div>
                            <div class="stat-value" id="percentile90">-</div>
                            <div class="stat-label">days</div>
                            <div class="tile-tooltip">
                                <div class="tt-head">Top ECN Types at 90th pct</div>
                                <div id="tt-p90types">-</div>
                            </div>
                        </div>

                    </div>"""

if OLD_STATS in c:
    c = c.replace(OLD_STATS, NEW_STATS, 1)
    print('Stat cards replaced')
else:
    print('ERROR: stat card block not found — check whitespace')

# ── 3. JS — inject tooltip population after the existing stats assignments ───
# Find the anchor line we'll insert after
ANCHOR = "document.getElementById('percentile90').textContent = stats.percentile_90_proc_ct.toFixed(2);"

TOOLTIP_JS = """
            // ── Populate tile tooltips ──────────────────────────────────────
            (function() {
                const kpis   = data.advanced_kpis;
                const rush   = data.rush_comparison;
                const holds  = kpis.hold_analysis;
                const voids  = kpis.void_analysis;
                const ftr    = kpis.ftr_analysis;
                const topics = data.topic_comparison;

                function fmt(v, decimals=1) { return v != null ? Number(v).toFixed(decimals) : '-'; }
                function fmtN(v) { return v != null ? Number(v).toLocaleString() : '-'; }
                function shortTopic(t) {
                    // strip leading "(N)  " code
                    return t ? t.replace(/^\\(\\d+[A-Z]?\\)\\s+/, '').replace(/^\\{\\d+\\}\\s+/, '') : t;
                }
                function rows(items, labelFn, valFn, limit=3) {
                    return (items||[]).slice(0,limit).map(r =>
                        '<div class="tt-row"><span class="tt-label">'+ labelFn(r) +'</span><span class="tt-val">'+ valFn(r) +'</span></div>'
                    ).join('') || '-';
                }

                // Total Requests tile
                document.getElementById('tt-closed').textContent    = fmtN(stats.total_closed);
                document.getElementById('tt-void').textContent      = fmtN(stats.total_void);
                document.getElementById('tt-hold').textContent      = fmtN(holds.hold_count);
                document.getElementById('tt-rush').textContent      = rush.length ? fmtN(rush[0].RequestNum) : '-';

                // Avg Proc CT tile
                document.getElementById('tt-avgtotalct').textContent   = fmt(stats.avg_total_ct) + ' days';
                document.getElementById('tt-medianprocct').textContent = fmt(stats.median_proc_ct) + ' days';
                document.getElementById('tt-maxprocct').textContent    = fmt(stats.max_proc_ct,0) + ' days';
                document.getElementById('tt-rushct').textContent       = rush.length ? fmt(rush[0]['ProcCT(days)']) + ' days' : '-';
                const holdComp = holds.ct_comparison || [];
                const withHold = holdComp.find(r => r.Category === 'With Holds');
                document.getElementById('tt-holdct').textContent = withHold ? fmt(withHold['ProcCT(days)']) + ' days' : '-';

                // Avg Total CT tile — slowest topics by avg total ct
                const slowTopics = [...(topics||[])].sort((a,b)=> b['TotalCT(Days)'] - a['TotalCT(Days)']).slice(0,3);
                document.getElementById('tt-slowtopics').innerHTML = rows(slowTopics,
                    r => shortTopic(r['ECN Topic']),
                    r => fmt(r['TotalCT(Days)']) + 'd');

                // Median Total CT tile
                document.getElementById('tt-p50').textContent         = fmt(stats.percentile_50_proc_ct) + ' days';
                document.getElementById('tt-p75').textContent         = fmt(stats.percentile_75_proc_ct) + ' days';
                document.getElementById('tt-p90').textContent         = fmt(stats.percentile_90_proc_ct) + ' days';
                document.getElementById('tt-mediantotalct').textContent = fmt(stats.median_total_ct) + ' days';

                // Void Rate tile
                document.getElementById('tt-voidcount').textContent = fmtN(voids.void_count);
                document.getElementById('tt-voidreasons').innerHTML = rows(voids.by_reason,
                    r => r.VoidReason.substring(0,30) + (r.VoidReason.length>30?'…':''),
                    r => fmtN(r.RequestNum));

                // FTR Rate tile
                document.getElementById('tt-ftrcount').textContent  = fmtN(ftr.ftr_count);
                document.getElementById('tt-holdrate').textContent  = fmt(holds.hold_rate) + '%';
                document.getElementById('tt-holdcount').textContent = fmtN(holds.hold_count);
                document.getElementById('tt-holdreasons').innerHTML = rows((holds.by_reason||[]),
                    r => (r.HoldReason||'').substring(0,30) + ((r.HoldReason||'').length>30?'…':''),
                    r => fmtN(r.RequestNum), 2);

                // Percentile tiles — top ECN types per bucket
                const p50cats = (stats.top_categories_50th||[]).slice(0,3);
                const p75cats = (stats.top_categories_75th||[]).slice(0,3);
                const p90cats = (stats.top_categories_90th||[]).slice(0,3);
                document.getElementById('tt-p50types').innerHTML = rows(p50cats,
                    r => shortTopic(r['ECN Topic']),
                    r => fmt(r['ProcCT(days)']) + 'd avg');
                document.getElementById('tt-p75types').innerHTML = rows(p75cats,
                    r => shortTopic(r['ECN Topic']),
                    r => fmt(r['ProcCT(days)']) + 'd avg');
                document.getElementById('tt-p90types').innerHTML = rows(p90cats,
                    r => shortTopic(r['ECN Topic']),
                    r => fmt(r['ProcCT(days)']) + 'd avg');

                // Date range on Total Requests tooltip
                const dr = stats.date_range;
                if (dr) {
                    const fmt2 = d => new Date(d).toLocaleDateString('en-US',{month:'short',day:'numeric',year:'numeric'});
                    document.getElementById('tt-daterange').textContent = fmt2(dr.start) + ' – ' + fmt2(dr.end);
                }

                // Keep tooltip on the correct side when near bottom of viewport
                document.querySelectorAll('.stat-card').forEach(card => {
                    card.addEventListener('mouseenter', () => {
                        const tt = card.querySelector('.tile-tooltip');
                        if (!tt) return;
                        const rect = card.getBoundingClientRect();
                        if (rect.top < 220) { tt.classList.add('tt-below'); }
                        else { tt.classList.remove('tt-below'); }
                    });
                });
            })();
"""

if ANCHOR in c:
    c = c.replace(ANCHOR, ANCHOR + TOOLTIP_JS, 1)
    print('Tooltip JS injected')
else:
    print('ERROR: JS anchor not found')

# ── 4. Save & deploy ─────────────────────────────────────────────────────────
open(PROJ, 'w', encoding='utf-8').write(c)
shutil.copy2(PROJ, SP)
print()
print('Saved to project folder.')
print('Copied to SharePoint sync folder.')
print()
print('Verification:')
print('  has-tooltip count :', c.count('has-tooltip'))
print('  has-download count:', c.count('has-download'))
print('  tt- element count :', c.count('id="tt-'))
