"""
Create enhanced dashboard with tabs: Overview, Executive Summary, and Advanced KPIs
"""
import json
import base64
import os
from datetime import datetime

# Read the data
with open('data.json', 'r') as f:
    data = json.load(f)

# Convert data to JavaScript
data_js = json.dumps(data, indent=2)

# Get current timestamp for "Last Updated"
last_updated = datetime.now().strftime('%B %d, %Y at %I:%M %p')

# SharePoint base URL for downloads
sharepoint_base_url = 'https://analog.sharepoint.com/sites/spmig_WWMFGbackendfndypromis/Shared%20Documents/Backend%20Foundry%20ECN%20Metrics'

# Embed logo as base64 for SharePoint compatibility
logo_base64 = ""
logo_file = 'ADI-Logo-RGB-FullColor.png'
if os.path.exists(logo_file):
    with open(logo_file, 'rb') as f:
        logo_base64 = base64.b64encode(f.read()).decode('utf-8')

# Embed Chart.js and jsPDF libraries for SharePoint compatibility
chartjs_code = ""
jspdf_code = ""

if os.path.exists('chart.umd.min.js'):
    with open('chart.umd.min.js', 'r', encoding='utf-8') as f:
        chartjs_code = f.read()
    print("[OK] Chart.js library loaded for embedding")

if os.path.exists('jspdf.umd.min.js'):
    with open('jspdf.umd.min.js', 'r', encoding='utf-8') as f:
        jspdf_code = f.read()
    print("[OK] jsPDF library loaded for embedding")

html_content = f"""<!DOCTYPE html>
<html lang="en">
<head>
    <meta charset="UTF-8">
    <meta name="viewport" content="width=device-width, initial-scale=1.0">
    <title>BEF ECN Cycle Time Dashboard - FY27</title>
    <script>
    // Embedded Chart.js library for SharePoint compatibility
    {chartjs_code}
    </script>
    <script>
    // Embedded jsPDF library for SharePoint compatibility
    {jspdf_code}
    </script>
    <style>
        * {{
            margin: 0;
            padding: 0;
            box-sizing: border-box;
        }}

        body {{
            font-family: 'Segoe UI', Tahoma, Geneva, Verdana, sans-serif;
            background: linear-gradient(135deg, #2563eb 0%, #1e3a8a 100%);
            padding: 20px;
            min-height: 100vh;
        }}

        .container {{
            max-width: 1400px;
            margin: 0 auto;
            width: 100%;
            padding: 0 10px;
        }}

        header {{
            background: white;
            padding: 30px;
            border-radius: 15px;
            box-shadow: 0 10px 30px rgba(0,0,0,0.2);
            margin-bottom: 30px;
            text-align: center;
            position: relative;
        }}

        .download-section {{
            margin-top: 20px;
            padding-top: 20px;
            border-top: 2px solid #e5e7eb;
            display: flex;
            justify-content: center;
            gap: 15px;
            flex-wrap: wrap;
        }}

        .download-btn {{
            background: linear-gradient(135deg, #2563eb 0%, #1e40af 100%);
            color: white;
            padding: 12px 24px;
            border-radius: 8px;
            text-decoration: none;
            font-weight: 600;
            display: inline-flex;
            align-items: center;
            gap: 8px;
            transition: all 0.3s ease;
            box-shadow: 0 4px 6px rgba(37, 99, 235, 0.3);
        }}

        .download-btn:hover {{
            transform: translateY(-2px);
            box-shadow: 0 6px 12px rgba(37, 99, 235, 0.4);
            background: linear-gradient(135deg, #1e40af 0%, #1e3a8a 100%);
        }}

        .download-btn.excel {{
            background: linear-gradient(135deg, #059669 0%, #047857 100%);
            box-shadow: 0 4px 6px rgba(5, 150, 105, 0.3);
        }}

        .download-btn.excel:hover {{
            background: linear-gradient(135deg, #047857 0%, #065f46 100%);
            box-shadow: 0 6px 12px rgba(5, 150, 105, 0.4);
        }}

        .download-btn.summary {{
            background: linear-gradient(135deg, #7c3aed 0%, #6d28d9 100%);
            box-shadow: 0 4px 6px rgba(124, 58, 237, 0.3);
        }}

        .download-btn.summary:hover {{
            background: linear-gradient(135deg, #6d28d9 0%, #5b21b6 100%);
            box-shadow: 0 6px 12px rgba(124, 58, 237, 0.4);
        }}

        .download-icon {{
            font-size: 18px;
        }}

        /* System Userid highlighting */
        .system-userid {{
            color: #dc2626;
            font-weight: 700;
            background: #fee2e2;
            padding: 2px 8px;
            border-radius: 4px;
            display: inline-block;
        }}

        .header-logo {{
            position: absolute;
            top: 20px;
            right: 30px;
            height: 100px;
            width: auto;
        }}

        h1 {{
            color: #333;
            font-size: 2.5em;
            margin-bottom: 10px;
        }}

        .subtitle {{
            color: #666;
            font-size: 1.1em;
        }}

        /* Tab Navigation */
        .tabs {{
            background: white;
            border-radius: 15px;
            box-shadow: 0 5px 15px rgba(0,0,0,0.1);
            margin-bottom: 30px;
            overflow: hidden;
        }}

        .tab-buttons {{
            display: flex;
            background: #f8f9fa;
            border-bottom: 2px solid #e0e0e0;
        }}

        .tab-button {{
            flex: 1;
            padding: 20px;
            background: transparent;
            border: none;
            cursor: pointer;
            font-size: 1.1em;
            font-weight: 600;
            color: #666;
            transition: all 0.3s ease;
            border-bottom: 3px solid transparent;
        }}

        .tab-button:hover {{
            background: #e8eeff;
            color: #667eea;
        }}

        .tab-button.active {{
            background: white;
            color: #667eea;
            border-bottom: 3px solid #667eea;
        }}

        .tab-content {{
            display: none;
            padding: 30px;
            animation: fadeIn 0.5s;
        }}

        .tab-content.active {{
            display: block;
        }}

        @keyframes fadeIn {{
            from {{ opacity: 0; transform: translateY(10px); }}
            to {{ opacity: 1; transform: translateY(0); }}
        }}

        .stats-grid {{
            display: grid;
            grid-template-columns: repeat(auto-fit, minmax(200px, 1fr));
            gap: 20px;
            margin-bottom: 30px;
        }}

        .stat-card {{
            background: linear-gradient(135deg, #667eea 0%, #764ba2 100%);
            padding: 25px;
            border-radius: 10px;
            box-shadow: 0 5px 15px rgba(0,0,0,0.1);
            text-align: center;
            transition: transform 0.3s ease;
            color: white;
        }}

        .stat-card:hover {{
            transform: translateY(-5px);
            box-shadow: 0 8px 25px rgba(0,0,0,0.2);
        }}

        .stat-card.green {{
            background: linear-gradient(135deg, #48bb78 0%, #38a169 100%);
        }}

        .stat-card.orange {{
            background: linear-gradient(135deg, #f59e0b 0%, #d97706 100%);
        }}

        .stat-card.red {{
            background: linear-gradient(135deg, #ef4444 0%, #dc2626 100%);
        }}

        .stat-card.purple {{
            background: linear-gradient(135deg, #8b5cf6 0%, #7c3aed 100%);
        }}

        .stat-value {{
            font-size: 2.5em;
            font-weight: bold;
            color: white;
            margin: 10px 0;
        }}

        .stat-label {{
            color: rgba(255,255,255,0.9);
            font-size: 0.9em;
            text-transform: uppercase;
            letter-spacing: 1px;
        }}

        .chart-grid {{
            display: grid;
            grid-template-columns: repeat(auto-fit, minmax(min(100%, 500px), 1fr));
            gap: 30px;
            margin-bottom: 30px;
        }}

        .chart-container {{
            background: white;
            padding: 25px;
            border-radius: 15px;
            box-shadow: 0 10px 30px rgba(0,0,0,0.1);
        }}

        .chart-container.full-width {{
            grid-column: 1 / -1;
        }}

        .chart-title {{
            font-size: 1.3em;
            color: #333;
            margin-bottom: 20px;
            font-weight: 600;
        }}

        canvas {{
            max-height: 400px;
        }}

        .date-range {{
            background: #f0f9ff;
            color: #1e40af;
            padding: 12px 24px;
            border-radius: 8px;
            display: inline-block;
            margin-top: 15px;
            font-size: 1.05em;
            font-weight: 600;
            border: 2px solid #3b82f6;
        }}

        .loading {{
            text-align: center;
            padding: 100px;
            color: white;
            font-size: 1.5em;
        }}

        /* Executive Summary Styles */
        .summary-section {{
            background: white;
            padding: 30px;
            border-radius: 15px;
            box-shadow: 0 5px 15px rgba(0,0,0,0.1);
            margin-bottom: 30px;
        }}

        .summary-title {{
            font-size: 2em;
            color: #667eea;
            margin-bottom: 20px;
            border-bottom: 3px solid #667eea;
            padding-bottom: 10px;
        }}

        .summary-subtitle {{
            font-size: 1.4em;
            color: #333;
            margin-top: 25px;
            margin-bottom: 15px;
            font-weight: 600;
        }}

        .summary-text {{
            color: #555;
            line-height: 1.8;
            font-size: 1.05em;
            margin-bottom: 15px;
        }}

        .insight-box {{
            background: #f0f7ff;
            border-left: 4px solid #667eea;
            padding: 20px;
            margin: 20px 0;
            border-radius: 5px;
        }}

        .insight-box.warning {{
            background: #fff8e1;
            border-left-color: #f59e0b;
        }}

        .insight-box.success {{
            background: #f0fff4;
            border-left-color: #48bb78;
        }}

        .insight-box.danger {{
            background: #fef2f2;
            border-left-color: #ef4444;
        }}

        .metric-highlight {{
            font-weight: bold;
            color: #667eea;
            font-size: 1.1em;
        }}

        .recommendations {{
            list-style: none;
            padding: 0;
        }}

        .recommendations li {{
            padding: 12px 15px;
            margin: 10px 0;
            background: #f8f9fa;
            border-radius: 8px;
            border-left: 3px solid #667eea;
        }}

        .recommendations li:before {{
            content: "→ ";
            color: #667eea;
            font-weight: bold;
            margin-right: 10px;
        }}

        /* KPI Table */
        .kpi-table {{
            width: 100%;
            border-collapse: collapse;
            margin: 20px 0;
        }}

        .kpi-table th {{
            background: #667eea;
            color: white;
            padding: 15px;
            text-align: left;
            font-weight: 600;
            border-bottom: 2px solid #5568d3;
        }}

        .kpi-table td {{
            padding: 12px 15px;
            border-bottom: 1px solid #e0e0e0;
        }}

        .kpi-table tr:hover {{
            background: #f8f9fa;
        }}

        .kpi-table tr:last-child td {{
            border-bottom: none;
        }}

        .badge {{
            display: inline-block;
            padding: 4px 12px;
            border-radius: 12px;
            font-size: 0.85em;
            font-weight: 600;
        }}

        .badge.good {{
            background: #d4edda;
            color: #155724;
        }}

        .badge.warning {{
            background: #fff3cd;
            color: #856404;
        }}

        .badge.danger {{
            background: #f8d7da;
            color: #721c24;
        }}

        /* Mobile-specific styles */
        @media (max-width: 768px) {{
            body {{
                padding: 10px;
            }}

            header {{
                padding: 15px;
            }}

            .header-logo {{
                position: static;
                height: 60px;
                margin: 10px auto;
                display: block;
            }}

            h1 {{
                font-size: 1.5em !important;
            }}

            .download-section {{
                flex-direction: column;
                gap: 10px;
            }}

            .download-btn {{
                width: 100%;
                justify-content: center;
                padding: 10px 16px;
                font-size: 14px;
            }}

            /* Hide download section on mobile */
            .download-section {{
                display: none !important;
            }}

            /* Hide complex charts and tables on mobile */
            .chart-container {{
                display: none !important;
            }}

            .kpi-table {{
                font-size: 13px;
            }}

            .kpi-table th,
            .kpi-table td {{
                padding: 8px 6px;
            }}

            /* Show simplified view notice */
            .mobile-notice {{
                display: block !important;
                background: #fef3c7;
                color: #92400e;
                padding: 12px;
                border-radius: 8px;
                margin-bottom: 15px;
                text-align: center;
                font-size: 14px;
                border: 2px solid #fbbf24;
            }}

            /* Simplify tabs */
            .tab-nav {{
                flex-direction: column;
                gap: 5px;
            }}

            .tab-button {{
                width: 100%;
                padding: 10px;
                font-size: 14px;
            }}

            /* Hide Advanced KPIs tab on mobile - too complex */
            .tab-button[onclick*="kpis"] {{
                display: none;
            }}

            #kpis-tab {{
                display: none !important;
            }}

            .stat-card {{
                min-width: auto;
                flex: 1 1 calc(50% - 10px);
            }}

            .stat-value {{
                font-size: 1.5em;
            }}
        }}

        .mobile-notice {{
            display: none;
        }}
    </style>
</head>
<body>
    <div class="container">
        <header>
            <img src="data:image/png;base64,{logo_base64}" alt="ADI Logo" class="header-logo">
            <h1>BEF ECN Cycle Time Dashboard</h1>
            <p class="subtitle">Engineering Change Notice Metrics - Last 12 Months</p>
            <div class="date-range" id="dateRange"></div>
            <p style="margin-top: 10px; color: #666; font-size: 0.9em;">
                <strong>Last Updated:</strong> {last_updated}
            </p>

            <div class="download-section">
                <a href="{sharepoint_base_url}/source_data.xlsx" download="ECN_Source_Data.xlsx" class="download-btn excel" title="Download complete Excel file with all ECN records and details">
                    <span class="download-icon">📥</span>
                    Download Full ECN Data (Excel)
                </a>
                <a href="{sharepoint_base_url}/ECN_Quarterly_Trends.xlsx" download="ECN_Quarterly_Trends.xlsx" class="download-btn excel" title="Download quarterly and monthly cycle time trends">
                    <span class="download-icon">📈</span>
                    Download Quarterly Trends (Excel)
                </a>
                <a href="{sharepoint_base_url}/ECN_Executive_Summary.pdf" download="ECN_Executive_Summary.pdf" class="download-btn summary" title="Download executive summary as PDF file">
                    <span class="download-icon">📄</span>
                    Download Executive Summary (PDF)
                </a>
                <a href="{sharepoint_base_url}/data.json" download="ECN_Metrics_Data.json" class="download-btn" title="Download all calculated metrics in JSON format">
                    <span class="download-icon">📊</span>
                    Download Metrics (JSON)
                </a>
            </div>
        </header>

        <div class="loading" id="loading">Loading dashboard...</div>

        <div id="dashboard" style="display: none;">
            <!-- Mobile Notice -->
            <div class="mobile-notice">
                📱 <strong>Mobile View:</strong> Showing simplified KPIs only. Charts and Advanced KPIs are hidden for better mobile performance. View on desktop for full dashboard.
            </div>

            <!-- Tab Navigation -->
            <div class="tabs">
                <div class="tab-buttons">
                    <button class="tab-button active" onclick="switchTab('overview')">Overview</button>
                    <button class="tab-button" onclick="switchTab('kpis')">Advanced KPIs</button>
                    <button class="tab-button" onclick="switchTab('executive')">Executive Summary</button>
                </div>

                <!-- Overview Tab -->
                <div id="overview-tab" class="tab-content active">
                    <!-- Summary Stats -->
                    <div class="stats-grid">
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
                    </div>

                    <!-- Charts -->
                    <div class="chart-grid">
                        <div class="chart-container full-width">
                            <h3 class="chart-title">
                                Monthly Cycle Time Trends
                                <a href="{sharepoint_base_url}/ECN_Quarterly_Trends.xlsx" download="ECN_Quarterly_Trends.xlsx" style="float: right; font-size: 0.85em; color: #667eea; text-decoration: none; font-weight: normal;" title="Download Excel with monthly and quarterly trends">
                                    📥 Download Excel
                                </a>
                            </h3>
                            <canvas id="monthlyTrendsChart"></canvas>
                        </div>

                        <div class="chart-container full-width">
                            <h3 class="chart-title">Quarterly Processing Cycle Time Trends by ECN Type (Top 5 Types)</h3>
                            <canvas id="quarterlyTopicTrendsChart"></canvas>
                        </div>

                        <div class="chart-container">
                            <h3 class="chart-title">Top 10 ECN Topics by Volume</h3>
                            <canvas id="topicChart"></canvas>
                        </div>

                        <div class="chart-container">
                            <h3 class="chart-title">Cycle Time by Topic</h3>
                            <canvas id="topicCTChart"></canvas>
                        </div>

                        <div class="chart-container">
                            <h3 class="chart-title">Top 10 Coordinators by Volume</h3>
                            <canvas id="coordinatorChart"></canvas>
                        </div>

                        <div class="chart-container">
                            <h3 class="chart-title">Site Distribution</h3>
                            <canvas id="siteChart"></canvas>
                        </div>

                        <div class="chart-container full-width">
                            <h3 class="chart-title">
                                ⚠️ ECNs Open for More Than 100 Days (By Type & State)
                                <a href="{sharepoint_base_url}/ECN_Over_100_Days.xlsx" download="ECN_Over_100_Days.xlsx" style="float: right; font-size: 0.85em; color: #dc2626; text-decoration: none; font-weight: normal;" title="Download ECNs open more than 100 days">
                                    📥 Download Excel
                                </a>
                            </h3>
                            <canvas id="ecnsOver100DaysChart"></canvas>
                        </div>

                        <div class="chart-container full-width">
                            <h3 class="chart-title">
                                📋 Open ECNs by Type & State (As of Today)
                                <a href="{sharepoint_base_url}/ECN_Open_ECNs.xlsx" download="ECN_Open_ECNs.xlsx" style="float: right; font-size: 0.85em; color: #667eea; text-decoration: none; font-weight: normal;" title="Download detailed list of all open ECNs">
                                    📥 Download Open ECNs Excel
                                </a>
                            </h3>
                            <canvas id="ecnTypeStateChart"></canvas>
                        </div>

                        <div class="chart-container full-width">
                            <h3 class="chart-title">State Distribution</h3>
                            <canvas id="stateChart"></canvas>
                        </div>

                        <div class="chart-container">
                            <h3 class="chart-title">Rush vs Regular Requests</h3>
                            <canvas id="rushChart"></canvas>
                        </div>

                        <div class="chart-container">
                            <h3 class="chart-title">Void Rate Trend</h3>
                            <canvas id="voidTrendChart"></canvas>
                        </div>

                        <div class="chart-container full-width">
                            <h3 class="chart-title">
                                90th Percentile ECNs by Type (Slowest 10% - Avg Processing CT)
                                <a href="{sharepoint_base_url}/ECN_90th_Percentile.xlsx" download="ECN_90th_Percentile.xlsx" style="float: right; font-size: 0.85em; color: #667eea; text-decoration: none; font-weight: normal;" title="Download detailed list of slowest ECNs">
                                    📥 Download Excel
                                </a>
                            </h3>
                            <canvas id="percentile90thChart"></canvas>
                        </div>
                    </div>
                </div>

                <!-- Executive Summary Tab -->
                <div id="executive-tab" class="tab-content">
                    <div class="summary-section">
                        <h2 class="summary-title">Executive Summary</h2>

                        <div id="execSummaryContent">
                            <!-- Content will be dynamically generated -->
                        </div>
                    </div>
                </div>

                <!-- Advanced KPIs Tab -->
                <div id="kpis-tab" class="tab-content">
                    <div class="summary-section">
                        <h2 class="summary-title">Advanced Key Performance Indicators</h2>

                        <div id="kpisContent">
                            <!-- Content will be dynamically generated -->
                        </div>
                    </div>

                    <!-- Additional KPI Charts -->
                    <div class="chart-grid">
                        <div class="chart-container full-width">
                            <h3 class="chart-title">FTR Rate Trend</h3>
                            <canvas id="ftrTrendChart"></canvas>
                        </div>

                        <div class="chart-container">
                            <h3 class="chart-title">Hold Rate by Topic</h3>
                            <canvas id="holdTopicChart"></canvas>
                        </div>

                        <div class="chart-container">
                            <h3 class="chart-title">Coordinator Workload Distribution</h3>
                            <canvas id="coordWorkloadChart"></canvas>
                        </div>

                        <div class="chart-container">
                            <h3 class="chart-title">Rush Rate Trend</h3>
                            <canvas id="rushTrendChart"></canvas>
                        </div>

                        <div class="chart-container">
                            <h3 class="chart-title">Top Manufacturing Sites</h3>
                            <canvas id="mfgSiteChart"></canvas>
                        </div>
                    </div>
                </div>
            </div>
        </div>
    </div>

    <script>
        const embeddedData = {data_js};
        let globalData = null;

        // Tab switching
        function switchTab(tabName) {{
            // Hide all tabs
            document.querySelectorAll('.tab-content').forEach(tab => {{
                tab.classList.remove('active');
            }});
            document.querySelectorAll('.tab-button').forEach(btn => {{
                btn.classList.remove('active');
            }});

            // Show selected tab
            document.getElementById(tabName + '-tab').classList.add('active');
            event.target.classList.add('active');
        }}

        /*
  BEF ECN METRICS - EXECUTIVE SUMMARY
  Generated: ${{new Date().toLocaleString()}}
  Period: ${{dateRange}}
================================================================================

OVERVIEW
================================================================================
Total ECN Requests:        ${{stats.total_requests.toLocaleString()}}
Closed ECNs:              ${{stats.total_closed.toLocaleString()}}
Void ECNs:                ${{stats.total_void.toLocaleString()}}

Average Processing Time:   ${{stats.avg_proc_ct.toFixed(2)}} days
Average Total Cycle Time:  ${{stats.avg_total_ct.toFixed(2)}} days

PROCESSING TIME PERCENTILES (Closed ECNs Only)
================================================================================
Based on ${{stats.total_closed.toLocaleString()}} Closed ECNs
(excluding ${{stats.total_void.toLocaleString()}} Void ECNs)

50th Percentile:  ${{stats.percentile_50_proc_ct.toFixed(2)}} days
  → 50% of ECNs processed in ${{stats.percentile_50_proc_ct.toFixed(2)}} days or less

75th Percentile:  ${{stats.percentile_75_proc_ct.toFixed(2)}} days
  → 75% of ECNs processed in ${{stats.percentile_75_proc_ct.toFixed(2)}} days or less

90th Percentile:  ${{stats.percentile_90_proc_ct.toFixed(2)}} days
  → 90% of ECNs processed in ${{stats.percentile_90_proc_ct.toFixed(2)}} days or less

TOP ECN CATEGORIES BY PERCENTILE RANGE
================================================================================

50th Percentile Range (Fastest - up to ${{stats.percentile_50_proc_ct.toFixed(2)}} days):
`;

            // Add top 3 categories for 50th percentile
            if (stats.top_categories_50th) {{
                stats.top_categories_50th.forEach((cat, idx) => {{
                    summary += `  ${{idx + 1}}. ${{cat.category}}: ${{cat.count.toLocaleString()}} ECNs (avg ${{cat.avg_proc_ct.toFixed(2)}} days)\\n`;
                }});
            }}

            summary += `
75th Percentile Range (${{stats.percentile_50_proc_ct.toFixed(2)}} - ${{stats.percentile_75_proc_ct.toFixed(2)}} days):
`;

            // Add top 3 categories for 75th percentile
            if (stats.top_categories_75th) {{
                stats.top_categories_75th.forEach((cat, idx) => {{
                    summary += `  ${{idx + 1}}. ${{cat.category}}: ${{cat.count.toLocaleString()}} ECNs (avg ${{cat.avg_proc_ct.toFixed(2)}} days)\\n`;
                }});
            }}

            summary += `
90th Percentile Range (${{stats.percentile_75_proc_ct.toFixed(2)}} - ${{stats.percentile_90_proc_ct.toFixed(2)}} days):
`;

            // Add top 3 categories for 90th percentile
            if (stats.top_categories_90th) {{
                stats.top_categories_90th.forEach((cat, idx) => {{
                    summary += `  ${{idx + 1}}. ${{cat.category}}: ${{cat.count.toLocaleString()}} ECNs (avg ${{cat.avg_proc_ct.toFixed(2)}} days)\\n`;
                }});
            }}

            summary += `
KEY INSIGHTS
================================================================================

Processing Efficiency:
  • Median processing time: ${{stats.percentile_50_proc_ct.toFixed(2)}} days
  • Most ECNs (90%) completed within: ${{stats.percentile_90_proc_ct.toFixed(2)}} days
  • Average across all ECNs: ${{stats.avg_proc_ct.toFixed(2)}} days

Volume Analysis:
  • Total requests processed: ${{stats.total_requests.toLocaleString()}}
  • Successfully closed: ${{stats.total_closed.toLocaleString()}} (${{(stats.total_closed/stats.total_requests*100).toFixed(1)}}%)
  • Voided: ${{stats.total_void.toLocaleString()}} (${{(stats.total_void/stats.total_requests*100).toFixed(1)}}%)

RECOMMENDATIONS
================================================================================

1. Focus on Top Categories
   The top ECN categories by volume show where most effort is concentrated.
   Consider streamlining processes for these high-volume categories.

2. Monitor Percentile Trends
   Track 50th and 90th percentiles month-over-month to identify
   process improvements or degradations.

3. Address Outliers
   ECNs beyond the 90th percentile (${{stats.percentile_90_proc_ct.toFixed(2)}}+ days) should be
   reviewed for process bottlenecks or special circumstances.

4. Reduce Void Rate
   Current void rate of ${{(stats.total_void/stats.total_requests*100).toFixed(1)}}% represents potential waste.
   Investigate root causes and implement preventive measures.

DATA NOTES
================================================================================

Percentile Calculations:
  • Based on Processing Cycle Time (ProcCT) in days
  • Includes only Closed ECNs (Void ECNs excluded)
  • Measured from submission to processing completion

Period Coverage:
  • Date Range: ${{dateRange}}
  • Total Records: ${{stats.total_requests.toLocaleString()}}

For detailed metrics, charts, and analysis:
  • Open dashboard.html in your web browser
  • Review BEF_ECN_Metrics.pptx PowerPoint presentation
  • Explore data.json for raw metrics data

================================================================================
  END OF EXECUTIVE SUMMARY
================================================================================
*/

        // Load and process data
        try {{
            globalData = embeddedData;
            renderDashboard(embeddedData);
        }} catch(error) {{
            console.error('Error loading data:', error);
            document.getElementById('loading').textContent = 'Error loading data: ' + error.message;
        }}

        function renderDashboard(data) {{
            // Hide loading, show dashboard
            document.getElementById('loading').style.display = 'none';
            document.getElementById('dashboard').style.display = 'block';

            // Update summary stats
            const stats = data.overall_stats;
            document.getElementById('totalRequests').textContent = stats.total_requests.toLocaleString();
            document.getElementById('avgProcCT').textContent = stats.avg_proc_ct.toFixed(2);
            document.getElementById('avgTotalCT').textContent = stats.avg_total_ct.toFixed(2);
            document.getElementById('medianTotalCT').textContent = stats.median_total_ct.toFixed(2);
            document.getElementById('voidRate').textContent = data.advanced_kpis.void_analysis.void_rate.toFixed(2);
            document.getElementById('ftrRate').textContent = data.advanced_kpis.ftr_analysis.ftr_rate.toFixed(2);
            document.getElementById('percentile50').textContent = stats.percentile_50_proc_ct.toFixed(2);
            document.getElementById('percentile75').textContent = stats.percentile_75_proc_ct.toFixed(2);
            document.getElementById('percentile90').textContent = stats.percentile_90_proc_ct.toFixed(2);
            // Format date range nicely
            const startDate = new Date(stats.date_range.start);
            const endDate = new Date(stats.date_range.end);
            const options = {{ year: 'numeric', month: 'short', day: 'numeric' }};
            const formattedStart = startDate.toLocaleDateString('en-US', options);
            const formattedEnd = endDate.toLocaleDateString('en-US', options);
            document.getElementById('dateRange').textContent = `Data Period: ${{formattedStart}} to ${{formattedEnd}} (Last 12 Months)`;

            // Create charts
            createMonthlyTrendsChart(data.monthly_trends);
            createQuarterlyTopicTrendsChart(data.quarterly_topic_trends);
            createTopicChart(data.topic_comparison);
            createTopicCTChart(data.topic_comparison);
            createCoordinatorChart(data.coordinator_comparison);
            createSiteChart(data.site_comparison);
            createECNsOver100DaysChart(data.ecns_over_100_days);
            createECNTypeStateChart(data.ecn_type_state_open, data.open_ecns_info);
            createStateChart(data.state_distribution);
            createRushChart(data.rush_comparison);
            createVoidTrendChart(data.advanced_kpis.void_analysis.monthly_trend);
            createPercentile90thChart(data.percentile_90th_analysis);

            // Generate executive summary
            generateExecutiveSummary(data);

            // Generate advanced KPIs
            generateAdvancedKPIs(data);

            // Create KPI charts
            createFTRTrendChart(data.advanced_kpis.ftr_analysis.monthly_trend);
            createHoldTopicChart(data.advanced_kpis.hold_analysis.by_topic);
            createCoordWorkloadChart(data.advanced_kpis.coordinator_workload.workload_distribution);
            createRushTrendChart(data.advanced_kpis.rush_analysis.monthly_trend);
            createMfgSiteChart(data.advanced_kpis.mfg_site_analysis);

            // Highlight SYSTEM USERID after a short delay to ensure DOM is ready
            setTimeout(highlightSystemUserid, 100);
        }}

        function highlightSystemUserid() {{
            // Find all elements and replace SYSTEM USERID text with styled version
            try {{
                const elements = document.querySelectorAll('td, th, p, span, li');

                elements.forEach(element => {{
                    // Skip if element contains a canvas or is a script
                    if (element.querySelector && element.querySelector('canvas')) {{
                        return;
                    }}

                    if (element.textContent &&
                        element.textContent.includes('SYSTEM USERID') &&
                        !element.innerHTML.includes('system-userid')) {{
                        element.innerHTML = element.innerHTML.replace(
                            /SYSTEM USERID/g,
                            '<span class="system-userid">SYSTEM USERID</span>'
                        );
                    }}
                }});
            }} catch(e) {{
                console.error('Error highlighting SYSTEM USERID:', e);
            }}
        }}

        function generateExecutiveSummary(data) {{
            const stats = data.overall_stats;
            const voidData = data.advanced_kpis.void_analysis;
            const holdData = data.advanced_kpis.hold_analysis;
            const ftrData = data.advanced_kpis.ftr_analysis;
            const rushData = data.advanced_kpis.rush_analysis;

            let html = `
                <h3 class="summary-subtitle">Period Overview</h3>
                <p class="summary-text">
                    This dashboard analyzes <span class="metric-highlight">${{stats.total_requests.toLocaleString()}}</span>
                    ECN requests spanning from <span class="metric-highlight">${{stats.date_range.start}}</span> to
                    <span class="metric-highlight">${{stats.date_range.end}}</span>.
                </p>

                <h3 class="summary-subtitle">Cycle Time Performance</h3>
                <div class="insight-box">
                    <p class="summary-text">
                        <strong>Average Processing Cycle Time:</strong> <span class="metric-highlight">${{stats.avg_proc_ct.toFixed(2)}} days</span><br>
                        <strong>Average Total Cycle Time:</strong> <span class="metric-highlight">${{stats.avg_total_ct.toFixed(2)}} days</span><br>
                        <strong>Median Total Cycle Time:</strong> <span class="metric-highlight">${{stats.median_total_ct.toFixed(2)}} days</span>
                    </p>
                </div>

                <h3 class="summary-subtitle">Processing Time Percentiles (Closed ECNs Only)</h3>
                <div class="insight-box success">
                    <p class="summary-text">
                        <strong>50th Percentile (Median):</strong> <span class="metric-highlight">${{stats.percentile_50_proc_ct.toFixed(2)}} days</span><br>
                        <strong>75th Percentile:</strong> <span class="metric-highlight">${{stats.percentile_75_proc_ct.toFixed(2)}} days</span><br>
                        <strong>90th Percentile:</strong> <span class="metric-highlight">${{stats.percentile_90_proc_ct.toFixed(2)}} days</span>
                    </p>
                    <p class="summary-text" style="margin-top: 10px; font-size: 0.95em; color: #666;">
                        <strong>Based on ${{stats.total_closed.toLocaleString()}} Closed ECNs</strong> (excluding ${{stats.total_void.toLocaleString()}} Void ECNs):<br>
                        50% of closed requests processed within ${{stats.percentile_50_proc_ct.toFixed(2)}} days,
                        75% within ${{stats.percentile_75_proc_ct.toFixed(2)}} days, and
                        90% within ${{stats.percentile_90_proc_ct.toFixed(2)}} days.
                    </p>
                </div>

                <h3 class="summary-subtitle">Top 3 ECN Categories by Percentile</h3>

                <div class="insight-box" style="background: #f0f7ff; margin-bottom: 15px;">
                    <p class="summary-text">
                        <strong style="color: #667eea;">50th Percentile (0 - ${{stats.percentile_50_proc_ct.toFixed(2)}} days) - Fastest Processing:</strong>
                    </p>
                    <ul style="margin-left: 20px; margin-top: 10px;">
            `;

            stats.top_categories_50th.forEach((cat, idx) => {{
                html += `<li style="margin-bottom: 5px;"><strong>${{idx + 1}}. ${{cat['ECN Topic']}}</strong> - ${{cat.RequestNum.toLocaleString()}} requests (Avg: ${{cat['ProcCT(days)'].toFixed(2)}} days)</li>`;
            }});

            html += `
                    </ul>
                </div>

                <div class="insight-box" style="background: #fff8e1; margin-bottom: 15px;">
                    <p class="summary-text">
                        <strong style="color: #f59e0b;">75th Percentile (${{stats.percentile_50_proc_ct.toFixed(2)}} - ${{stats.percentile_75_proc_ct.toFixed(2)}} days) - Medium Processing:</strong>
                    </p>
                    <ul style="margin-left: 20px; margin-top: 10px;">
            `;

            stats.top_categories_75th.forEach((cat, idx) => {{
                html += `<li style="margin-bottom: 5px;"><strong>${{idx + 1}}. ${{cat['ECN Topic']}}</strong> - ${{cat.RequestNum.toLocaleString()}} requests (Avg: ${{cat['ProcCT(days)'].toFixed(2)}} days)</li>`;
            }});

            html += `
                    </ul>
                </div>

                <div class="insight-box" style="background: #fef2f2; margin-bottom: 15px;">
                    <p class="summary-text">
                        <strong style="color: #ef4444;">90th Percentile (${{stats.percentile_75_proc_ct.toFixed(2)}} - ${{stats.percentile_90_proc_ct.toFixed(2)}} days) - Slower Processing:</strong>
                    </p>
                    <ul style="margin-left: 20px; margin-top: 10px;">
            `;

            stats.top_categories_90th.forEach((cat, idx) => {{
                html += `<li style="margin-bottom: 5px;"><strong>${{idx + 1}}. ${{cat['ECN Topic']}}</strong> - ${{cat.RequestNum.toLocaleString()}} requests (Avg: ${{cat['ProcCT(days)'].toFixed(2)}} days)</li>`;
            }});

            html += `
                    </ul>
                </div>

                <h3 class="summary-subtitle">Quality Metrics</h3>
                <div class="insight-box ${{voidData.void_rate > 10 ? 'warning' : 'success'}}">
                    <p class="summary-text">
                        <strong>Void Rate:</strong> <span class="metric-highlight">${{voidData.void_rate.toFixed(2)}}%</span>
                        (${{voidData.void_count.toLocaleString()}} requests)<br>
                        ${{voidData.void_rate > 10 ? '⚠️ Above recommended threshold of 10%' : '[OK] Within acceptable range'}}
                    </p>
                </div>

                <div class="insight-box ${{holdData.hold_rate > 15 ? 'warning' : 'success'}}">
                    <p class="summary-text">
                        <strong>Hold Rate:</strong> <span class="metric-highlight">${{holdData.hold_rate.toFixed(2)}}%</span>
                        (${{holdData.hold_count.toLocaleString()}} requests)<br>
                        ${{holdData.hold_rate > 15 ? '⚠️ Consider reviewing hold reasons' : '[OK] Acceptable hold rate'}}
                    </p>
                </div>

                <div class="insight-box ${{ftrData.ftr_rate < 70 ? 'warning' : 'success'}}">
                    <p class="summary-text">
                        <strong>First-Time-Right Rate:</strong> <span class="metric-highlight">${{ftrData.ftr_rate.toFixed(2)}}%</span><br>
                        ${{ftrData.ftr_rate < 70 ? '⚠️ Below target of 70% - improvement opportunity' : '[OK] Meeting FTR target'}}
                    </p>
                </div>

                <h3 class="summary-subtitle">Rush Requests</h3>
                <div class="insight-box">
                    <p class="summary-text">
                        <strong>Rush Request Rate:</strong> <span class="metric-highlight">${{rushData.rush_rate.toFixed(2)}}%</span>
                        (${{rushData.rush_count.toLocaleString()}} requests)
                    </p>
                </div>

                <h3 class="summary-subtitle">Key Recommendations</h3>
                <ul class="recommendations">
            `;

            // Add recommendations based on metrics
            if (voidData.void_rate > 10) {{
                html += `<li>Investigate top void reasons to reduce waste and improve efficiency</li>`;
            }}
            if (holdData.hold_rate > 15) {{
                html += `<li>Analyze hold reasons and implement preventive measures to reduce delays</li>`;
            }}
            if (ftrData.ftr_rate < 70) {{
                html += `<li>Focus on improving first-time-right rate through training and process improvements</li>`;
            }}
            if (rushData.rush_rate > 20) {{
                html += `<li>High rush rate indicates potential planning issues - review request workflows</li>`;
            }}

            html += `
                    <li>Monitor cycle time trends and identify outliers for process optimization</li>
                    <li>Balance coordinator workload to maintain consistent processing times</li>
                    <li>Continue tracking KPIs monthly to identify improvement opportunities</li>
                </ul>
            `;

            document.getElementById('execSummaryContent').innerHTML = html;
        }}

        function generateAdvancedKPIs(data) {{
            const voidData = data.advanced_kpis.void_analysis;
            const holdData = data.advanced_kpis.hold_analysis;
            const ftrData = data.advanced_kpis.ftr_analysis;
            const rushData = data.advanced_kpis.rush_analysis;
            const coordData = data.advanced_kpis.coordinator_workload;
            const mfgData = data.advanced_kpis.mfg_site_analysis;
            const requestorData = data.advanced_kpis.requestor_analysis;

            let html = `
                <h3 class="summary-subtitle">Performance Metrics Summary</h3>
                <table class="kpi-table">
                    <thead>
                        <tr>
                            <th>Metric</th>
                            <th>Value</th>
                            <th>Status</th>
                            <th>Details</th>
                        </tr>
                    </thead>
                    <tbody>
                        <tr>
                            <td><strong>Void Rate</strong></td>
                            <td>${{voidData.void_rate.toFixed(2)}}%</td>
                            <td><span class="badge ${{voidData.void_rate > 10 ? 'warning' : 'good'}}">
                                ${{voidData.void_rate > 10 ? 'Needs Attention' : 'Good'}}
                            </span></td>
                            <td>${{voidData.void_count.toLocaleString()}} requests voided</td>
                        </tr>
                        <tr>
                            <td><strong>Hold Rate</strong></td>
                            <td>${{holdData.hold_rate.toFixed(2)}}%</td>
                            <td><span class="badge ${{holdData.hold_rate > 15 ? 'warning' : 'good'}}">
                                ${{holdData.hold_rate > 15 ? 'Needs Attention' : 'Good'}}
                            </span></td>
                            <td>${{holdData.hold_count.toLocaleString()}} requests on hold</td>
                        </tr>
                        <tr>
                            <td><strong>First-Time-Right Rate</strong></td>
                            <td>${{ftrData.ftr_rate.toFixed(2)}}%</td>
                            <td><span class="badge ${{ftrData.ftr_rate < 70 ? 'warning' : 'good'}}">
                                ${{ftrData.ftr_rate < 70 ? 'Below Target' : 'On Target'}}
                            </span></td>
                            <td>${{ftrData.ftr_count.toLocaleString()}} requests completed first-time-right</td>
                        </tr>
                        <tr>
                            <td><strong>Rush Request Rate</strong></td>
                            <td>${{rushData.rush_rate.toFixed(2)}}%</td>
                            <td><span class="badge ${{rushData.rush_rate > 20 ? 'warning' : 'good'}}">
                                ${{rushData.rush_rate > 20 ? 'High' : 'Normal'}}
                            </span></td>
                            <td>${{rushData.rush_count.toLocaleString()}} rush requests</td>
                        </tr>
                        <tr>
                            <td><strong>Active Coordinators</strong></td>
                            <td>${{coordData.workload_distribution.length}}</td>
                            <td><span class="badge good">Active</span></td>
                            <td>Processing requests across sites</td>
                        </tr>
                        <tr>
                            <td><strong>Manufacturing Sites</strong></td>
                            <td>${{mfgData.length}}</td>
                            <td><span class="badge good">Active</span></td>
                            <td>Top sites by volume tracked</td>
                        </tr>
                        <tr>
                            <td><strong>Active Requestors</strong></td>
                            <td>${{requestorData.length}}</td>
                            <td><span class="badge good">Active</span></td>
                            <td>Requestors with 10+ requests</td>
                        </tr>
                    </tbody>
                </table>

                <h3 class="summary-subtitle">
                    Top Void Reasons
                    <br>
                    <a href="{sharepoint_base_url}/ECN_Void_by_Reason.xlsx" download="ECN_Void_by_Reason.xlsx"
                       class="download-btn"
                       style="display: inline-block; margin-top: 8px; padding: 6px 12px; background: #10b981; color: white; border-radius: 4px; text-decoration: none; font-size: 12px;">
                        📥 Download Void ECNs Excel Report
                    </a>
                </h3>
                <table class="kpi-table">
                    <thead>
                        <tr>
                            <th>#</th>
                            <th>Void Reason</th>
                            <th>Count</th>
                            <th>% of Total Voids</th>
                        </tr>
                    </thead>
                    <tbody>
            `;

            const topVoidReasons = voidData.by_reason.slice(0, 5);
            const totalVoids = voidData.void_count;

            topVoidReasons.forEach((reason, index) => {{
                const pct = ((reason.RequestNum / totalVoids) * 100).toFixed(1);
                html += `
                    <tr>
                        <td>${{index + 1}}</td>
                        <td>${{reason.VoidReason || 'Not Specified'}}</td>
                        <td>${{reason.RequestNum.toLocaleString()}}</td>
                        <td>${{pct}}%</td>
                    </tr>
                `;
            }});

            html += `
                    </tbody>
                </table>

                <h3 class="summary-subtitle">Top Hold Reasons</h3>
                <table class="kpi-table">
                    <thead>
                        <tr>
                            <th>#</th>
                            <th>Hold Reason</th>
                            <th>Count</th>
                            <th>Avg Total CT (days)</th>
                        </tr>
                    </thead>
                    <tbody>
            `;

            const topHoldReasons = holdData.by_reason.slice(0, 5);

            topHoldReasons.forEach((reason, index) => {{
                html += `
                    <tr>
                        <td>${{index + 1}}</td>
                        <td>${{reason.HoldReason || 'Not Specified'}}</td>
                        <td>${{reason.RequestNum.toLocaleString()}}</td>
                        <td>${{reason['TotalCT(Days)'].toFixed(2)}}</td>
                    </tr>
                `;
            }});

            html += `
                    </tbody>
                </table>

                <h3 class="summary-subtitle">Top Coordinators by Performance</h3>
                <table class="kpi-table">
                    <thead>
                        <tr>
                            <th>#</th>
                            <th>Coordinator</th>
                            <th>Volume</th>
                            <th>% of Total</th>
                            <th>Avg Processing CT (days)</th>
                        </tr>
                    </thead>
                    <tbody>
            `;

            const topCoords = coordData.workload_distribution.slice(0, 10);

            topCoords.forEach((coord, index) => {{
                html += `
                    <tr>
                        <td>${{index + 1}}</td>
                        <td>${{coord.ECNCoordinator || 'Unassigned'}}</td>
                        <td>${{coord.RequestNum.toLocaleString()}}</td>
                        <td>${{coord.Percentage.toFixed(1)}}%</td>
                        <td>${{coord['ProcCT(days)'].toFixed(2)}}</td>
                    </tr>
                `;
            }});

            html += `
                    </tbody>
                </table>
            `;

            document.getElementById('kpisContent').innerHTML = html;
        }}

        // Chart creation functions
        function createMonthlyTrendsChart(data) {{
            const ctx = document.getElementById('monthlyTrendsChart').getContext('2d');

            // Show last 12 months of data
            const filteredData = data.slice(-12);

            new Chart(ctx, {{
                type: 'line',
                data: {{
                    labels: filteredData.map(d => d.YearMonth),
                    datasets: [
                        {{
                            label: 'Processing CT (days)',
                            data: filteredData.map(d => d['ProcCT(days)']),
                            borderColor: '#667eea',
                            backgroundColor: 'rgba(102, 126, 234, 0.1)',
                            tension: 0.4,
                            fill: true
                        }},
                        {{
                            label: 'Total CT (days)',
                            data: filteredData.map(d => d['TotalCT(Days)']),
                            borderColor: '#48bb78',
                            backgroundColor: 'rgba(72, 187, 120, 0.1)',
                            tension: 0.4,
                            fill: true
                        }}
                    ]
                }},
                options: {{
                    responsive: true,
                    maintainAspectRatio: true,
                    plugins: {{
                        legend: {{
                            display: true,
                            position: 'top'
                        }}
                    }},
                    scales: {{
                        y: {{
                            beginAtZero: true,
                            title: {{
                                display: true,
                                text: 'Days'
                            }}
                        }}
                    }}
                }}
            }});
        }}

        function createQuarterlyTopicTrendsChart(data) {{
            const ctx = document.getElementById('quarterlyTopicTrendsChart');
            if (!ctx) return;

            // Get unique quarters and topics
            const quarters = [...new Set(data.map(d => d.YearQuarter))].sort();
            const topics = [...new Set(data.map(d => d['ECN Topic']))];

            // Color palette for different ECN types
            const colors = [
                {{ border: '#667eea', bg: 'rgba(102, 126, 234, 0.1)' }},
                {{ border: '#48bb78', bg: 'rgba(72, 187, 120, 0.1)' }},
                {{ border: '#ed8936', bg: 'rgba(237, 137, 54, 0.1)' }},
                {{ border: '#9f7aea', bg: 'rgba(159, 122, 234, 0.1)' }},
                {{ border: '#38b2ac', bg: 'rgba(56, 178, 172, 0.1)' }}
            ];

            // Create dataset for each topic - showing Processing Cycle Time
            const datasets = topics.map((topic, idx) => {{
                const topicData = quarters.map(quarter => {{
                    const record = data.find(d => d.YearQuarter === quarter && d['ECN Topic'] === topic);
                    return record ? record['ProcCT(days)'] : null;
                }});

                const color = colors[idx % colors.length];

                return {{
                    label: topic,
                    data: topicData,
                    borderColor: color.border,
                    backgroundColor: color.bg,
                    tension: 0.4,
                    fill: true,
                    borderWidth: 2,
                    spanGaps: false
                }};
            }});

            new Chart(ctx, {{
                type: 'line',
                data: {{
                    labels: quarters,
                    datasets: datasets
                }},
                options: {{
                    responsive: true,
                    maintainAspectRatio: true,
                    interaction: {{
                        mode: 'index',
                        intersect: false
                    }},
                    plugins: {{
                        legend: {{
                            display: true,
                            position: 'bottom',
                            labels: {{
                                boxWidth: 12,
                                padding: 10,
                                font: {{
                                    size: 11
                                }}
                            }}
                        }},
                        title: {{
                            display: true,
                            text: 'Processing Cycle Time by Quarter',
                            font: {{
                                size: 14
                            }}
                        }},
                        tooltip: {{
                            callbacks: {{
                                label: function(context) {{
                                    let label = context.dataset.label || '';
                                    if (label) {{
                                        label += ': ';
                                    }}
                                    if (context.parsed.y !== null) {{
                                        label += context.parsed.y.toFixed(2) + ' days';
                                    }}
                                    return label;
                                }}
                            }}
                        }}
                    }},
                    scales: {{
                        y: {{
                            beginAtZero: true,
                            title: {{
                                display: true,
                                text: 'Processing Cycle Time (days)'
                            }},
                            ticks: {{
                                callback: function(value) {{
                                    return value.toFixed(1);
                                }}
                            }}
                        }},
                        x: {{
                            title: {{
                                display: true,
                                text: 'Quarter'
                            }}
                        }}
                    }}
                }}
            }});
        }}

        function createTopicChart(data) {{
            const ctx = document.getElementById('topicChart').getContext('2d');
            const sortedData = [...data].sort((a, b) => b.RequestNum - a.RequestNum).slice(0, 10);

            new Chart(ctx, {{
                type: 'bar',
                data: {{
                    labels: sortedData.map(d => truncateLabel(d['ECN Topic'], 30)),
                    datasets: [{{
                        label: 'Request Count',
                        data: sortedData.map(d => d.RequestNum),
                        backgroundColor: '#667eea'
                    }}]
                }},
                options: {{
                    responsive: true,
                    maintainAspectRatio: true,
                    indexAxis: 'y',
                    plugins: {{
                        legend: {{
                            display: false
                        }}
                    }},
                    scales: {{
                        x: {{
                            beginAtZero: true
                        }}
                    }}
                }}
            }});
        }}

        function createTopicCTChart(data) {{
            const ctx = document.getElementById('topicCTChart').getContext('2d');
            const sortedData = [...data].sort((a, b) => b.RequestNum - a.RequestNum).slice(0, 10);

            new Chart(ctx, {{
                type: 'bar',
                data: {{
                    labels: sortedData.map(d => truncateLabel(d['ECN Topic'], 30)),
                    datasets: [{{
                        label: 'Avg Total CT (days)',
                        data: sortedData.map(d => d['TotalCT(Days)']),
                        backgroundColor: '#48bb78'
                    }}]
                }},
                options: {{
                    responsive: true,
                    maintainAspectRatio: true,
                    indexAxis: 'y',
                    plugins: {{
                        legend: {{
                            display: false
                        }}
                    }},
                    scales: {{
                        x: {{
                            beginAtZero: true
                        }}
                    }}
                }}
            }});
        }}

        function createCoordinatorChart(data) {{
            const ctx = document.getElementById('coordinatorChart').getContext('2d');

            // Merge SYSTEM USERID entries
            const mergedData = {{}};
            data.forEach(item => {{
                const coordinator = item.ECNCoordinator || 'Unassigned';
                if (mergedData[coordinator]) {{
                    mergedData[coordinator] += item.RequestNum;
                }} else {{
                    mergedData[coordinator] = item.RequestNum;
                }}
            }});

            // Convert back to array and sort
            const sortedData = Object.entries(mergedData)
                .map(([name, count]) => ({{ ECNCoordinator: name, RequestNum: count }}))
                .sort((a, b) => b.RequestNum - a.RequestNum)
                .slice(0, 10);

            // Create colors array - red for SYSTEM USERID, orange for others
            const backgroundColors = sortedData.map(d =>
                d.ECNCoordinator === 'SYSTEM USERID' ? '#dc2626' : '#f59e0b'
            );

            new Chart(ctx, {{
                type: 'bar',
                data: {{
                    labels: sortedData.map(d => d.ECNCoordinator),
                    datasets: [{{
                        label: 'Request Count',
                        data: sortedData.map(d => d.RequestNum),
                        backgroundColor: backgroundColors
                    }}]
                }},
                options: {{
                    responsive: true,
                    maintainAspectRatio: true,
                    indexAxis: 'y',
                    plugins: {{
                        legend: {{
                            display: false
                        }}
                    }},
                    scales: {{
                        x: {{
                            beginAtZero: true
                        }}
                    }}
                }}
            }});
        }}

        function createSiteChart(data) {{
            const ctx = document.getElementById('siteChart');
            if (!ctx) return;

            // Use separated data if available, otherwise fall back to regular data
            const separatedData = (globalData && globalData.site_comparison_separated) ? globalData.site_comparison_separated : null;

            // If we don't have separated data, use the old pie chart
            if (!separatedData) {{
                new Chart(ctx, {{
                    type: 'pie',
                    data: {{
                        labels: data.map(d => d.ECNCoordinatorSite),
                        datasets: [{{
                            data: data.map(d => d.RequestNum),
                            backgroundColor: ['#667eea', '#48bb78', '#f59e0b', '#ef4444', '#8b5cf6']
                        }}]
                    }},
                    options: {{
                        responsive: true,
                        maintainAspectRatio: true,
                        plugins: {{ legend: {{ position: 'right' }} }}
                    }}
                }});
                return;
            }}

            // Get unique sites
            const sites = [...new Set(separatedData.map(d => d.ECNCoordinatorSite))];

            // Separate SYSTEM USERID and Others
            const systemData = sites.map(site => {{
                const record = separatedData.find(d => d.ECNCoordinatorSite === site && d.Coordinator === 'SYSTEM USERID');
                return record ? record.RequestNum : 0;
            }});

            const otherData = sites.map(site => {{
                const record = separatedData.find(d => d.ECNCoordinatorSite === site && d.Coordinator === 'Other Coordinators');
                return record ? record.RequestNum : 0;
            }});

            new Chart(ctx, {{
                type: 'bar',
                data: {{
                    labels: sites,
                    datasets: [
                        {{
                            label: 'SYSTEM USERID',
                            data: systemData,
                            backgroundColor: '#dc2626',
                            borderColor: '#991b1b',
                            borderWidth: 1
                        }},
                        {{
                            label: 'Other Coordinators',
                            data: otherData,
                            backgroundColor: '#667eea',
                            borderColor: '#4c51bf',
                            borderWidth: 1
                        }}
                    ]
                }},
                options: {{
                    responsive: true,
                    maintainAspectRatio: true,
                    plugins: {{
                        legend: {{
                            position: 'top'
                        }},
                        title: {{
                            display: true,
                            text: 'Site Distribution: SYSTEM USERID vs Other Coordinators'
                        }}
                    }},
                    scales: {{
                        y: {{
                            beginAtZero: true,
                            title: {{
                                display: true,
                                text: 'Request Count'
                            }}
                        }},
                        x: {{
                            title: {{
                                display: true,
                                text: 'Site'
                            }}
                        }}
                    }}
                }}
            }});
        }}

        function createECNsOver100DaysChart(data) {{
            const ctx = document.getElementById('ecnsOver100DaysChart').getContext('2d');

            if (!data || !data.by_type_and_state || data.by_type_and_state.length === 0) {{
                ctx.font = '16px Arial';
                ctx.fillStyle = '#10b981';
                ctx.textAlign = 'center';
                ctx.fillText('✓ No ECNs open for more than 100 days!', ctx.canvas.width / 2, ctx.canvas.height / 2);
                return;
            }}

            // Prepare data - group by ECN Type and State
            const ecnTypes = [...new Set(data.by_type_and_state.map(d => d['ECN Topic']))];
            const states = [...new Set(data.by_type_and_state.map(d => d.State))];

            // Get top 10 ECN types by count
            const ecnTypeTotals = {{}};
            ecnTypes.forEach(type => {{
                ecnTypeTotals[type] = data.by_type_and_state.filter(d => d['ECN Topic'] === type)
                    .reduce((sum, d) => sum + d.Count, 0);
            }});
            const topEcnTypes = Object.entries(ecnTypeTotals)
                .sort((a, b) => b[1] - a[1])
                .slice(0, 10)
                .map(e => e[0]);

            // Create datasets for each state with unique colors
            const stateColors = {{
                'Hold': '#ec4899',                                   // Pink
                'Hold_Response': '#f97316',                          // Orange
                'PENDING_SAP_CREATION': '#a3e635',                   // Light Lime
                'Pending ADPH FG Input': '#14b8a6',                  // Teal
                'Pending Subcon Input': '#06b6d4',                   // Cyan
                'Pending Subcon Input (Close ECN)': '#0891b2',       // Dark Cyan
                'Submitted': '#3b82f6',                              // Blue
                'Processing': '#8b5cf6',                             // Purple
                'Pending Approval': '#eab308',                       // Yellow
                'ABS Status Pending': '#84cc16',                     // Lime
                'Error': '#ef4444',                                  // Red
                'Pending ADML FG Input': '#10b981',                  // Green
                'Processing (Auto:ABS)': '#7c3aed',                  // Dark Purple
                'Processing (AutoClose:3AA)': '#6d28d9',             // Darker Purple
                'Issue_Review_BOM': '#f59e0b',                       // Amber
                'Issue_Review_Material': '#fb923c'                   // Light Orange
            }};

            // Generate unique colors for any states not in the predefined list
            const colorPalette = [
                '#6366f1', '#f472b6', '#38bdf8', '#fbbf24', '#34d399',
                '#a78bfa', '#fb7185', '#22d3ee', '#fcd34d', '#6ee7b7',
                '#c084fc', '#fda4af', '#7dd3fc', '#fde68a', '#86efac'
            ];

            const datasets = states.map((state, index) => {{
                const stateData = topEcnTypes.map(type => {{
                    const item = data.by_type_and_state.find(d => d['ECN Topic'] === type && d.State === state);
                    return item ? item.Count : 0;
                }});

                // Get color from predefined map, or use palette color if not found
                let color = stateColors[state];
                if (!color) {{
                    color = colorPalette[index % colorPalette.length];
                }}

                return {{
                    label: state,
                    data: stateData,
                    backgroundColor: color
                }};
            }});

            new Chart(ctx, {{
                type: 'bar',
                data: {{
                    labels: topEcnTypes,
                    datasets: datasets
                }},
                options: {{
                    responsive: true,
                    maintainAspectRatio: true,
                    plugins: {{
                        legend: {{
                            display: true,
                            position: 'top'
                        }},
                        title: {{
                            display: true,
                            text: `Total: ${{data.total_count}} ECNs | Avg Days Open: ${{data.avg_days_open.toFixed(0)}} | Oldest: ${{data.max_days_open}} days`,
                            font: {{
                                size: 14,
                                weight: 'bold'
                            }},
                            color: '#dc2626'
                        }}
                    }},
                    scales: {{
                        x: {{
                            stacked: true,
                            title: {{
                                display: true,
                                text: 'ECN Type'
                            }}
                        }},
                        y: {{
                            stacked: true,
                            title: {{
                                display: true,
                                text: 'Count of ECNs Open > 100 Days'
                            }}
                        }}
                    }}
                }}
            }});
        }}

        function createECNTypeStateChart(data, openEcnsInfo) {{
            const ctx = document.getElementById('ecnTypeStateChart').getContext('2d');

            // Prepare data - group by ECN Type and State
            const ecnTypes = [...new Set(data.map(d => d['ECN Topic']))];
            const states = [...new Set(data.map(d => d.State))];

            // Get top 10 ECN types by total count
            const ecnTypeTotals = {{}};
            ecnTypes.forEach(type => {{
                ecnTypeTotals[type] = data.filter(d => d['ECN Topic'] === type)
                    .reduce((sum, d) => sum + d.Count, 0);
            }});
            const topEcnTypes = Object.entries(ecnTypeTotals)
                .sort((a, b) => b[1] - a[1])
                .slice(0, 10)
                .map(e => e[0]);

            // Create datasets for each state with unique colors
            const stateColors = {{
                'Submitted': '#3b82f6',                              // Blue
                'Processing': '#8b5cf6',                             // Purple
                'Processing (Auto:ABS)': '#7c3aed',                  // Dark Purple
                'Processing (AutoClose:3AA)': '#6d28d9',             // Darker Purple
                'Hold': '#ec4899',                                   // Pink
                'Hold_Response': '#f97316',                          // Orange
                'Pending Approval': '#eab308',                       // Yellow
                'Pending Subcon Input': '#06b6d4',                   // Cyan
                'Pending Subcon Input (Close ECN)': '#0891b2',       // Dark Cyan
                'Pending ADPH FG Input': '#14b8a6',                  // Teal
                'Pending ADML FG Input': '#10b981',                  // Green
                'ABS Status Pending': '#84cc16',                     // Lime
                'PENDING_SAP_CREATION': '#a3e635',                   // Light Lime
                'Error': '#ef4444',                                  // Red
                'Issue_Review_BOM': '#f59e0b',                       // Amber
                'Issue_Review_Material': '#fb923c'                   // Light Orange
            }};

            // Generate unique colors for any states not in the predefined list
            const colorPalette = [
                '#6366f1', '#f472b6', '#38bdf8', '#fbbf24', '#34d399',
                '#a78bfa', '#fb7185', '#22d3ee', '#fcd34d', '#6ee7b7',
                '#c084fc', '#fda4af', '#7dd3fc', '#fde68a', '#86efac'
            ];

            const datasets = states.map((state, index) => {{
                const stateData = topEcnTypes.map(type => {{
                    const item = data.find(d => d['ECN Topic'] === type && d.State === state);
                    return item ? item.Count : 0;
                }});

                // Get color from predefined map, or use palette color if not found
                let color = stateColors[state];
                if (!color) {{
                    color = colorPalette[index % colorPalette.length];
                }}

                return {{
                    label: state,
                    data: stateData,
                    backgroundColor: color
                }};
            }});

            new Chart(ctx, {{
                type: 'bar',
                data: {{
                    labels: topEcnTypes,
                    datasets: datasets
                }},
                options: {{
                    responsive: true,
                    maintainAspectRatio: true,
                    plugins: {{
                        legend: {{
                            display: true,
                            position: 'top',
                            labels: {{
                                boxWidth: 12,
                                font: {{
                                    size: 10
                                }}
                            }}
                        }},
                        title: {{
                            display: true,
                            text: `Top 10 ECN Types - Open ECNs as of ${{openEcnsInfo.as_of_date}} - Total: ${{openEcnsInfo.total_ecns}} ECNs`,
                            font: {{
                                size: 14,
                                weight: 'bold'
                            }}
                        }}
                    }},
                    scales: {{
                        x: {{
                            stacked: true,
                            title: {{
                                display: true,
                                text: 'ECN Type'
                            }}
                        }},
                        y: {{
                            stacked: true,
                            title: {{
                                display: true,
                                text: 'Count'
                            }}
                        }}
                    }}
                }}
            }});
        }}

        function createStateChart(data) {{
            const ctx = document.getElementById('stateChart').getContext('2d');

            new Chart(ctx, {{
                type: 'bar',
                data: {{
                    labels: data.map(d => d.State),
                    datasets: [
                        {{
                            label: 'Request Count',
                            data: data.map(d => d.RequestNum),
                            backgroundColor: '#667eea',
                            yAxisID: 'y'
                        }},
                        {{
                            label: 'Avg Total CT (days)',
                            data: data.map(d => d['TotalCT(Days)']),
                            backgroundColor: '#48bb78',
                            yAxisID: 'y1'
                        }}
                    ]
                }},
                options: {{
                    responsive: true,
                    maintainAspectRatio: true,
                    plugins: {{
                        legend: {{
                            display: true
                        }}
                    }},
                    scales: {{
                        y: {{
                            type: 'linear',
                            position: 'left',
                            title: {{
                                display: true,
                                text: 'Count'
                            }}
                        }},
                        y1: {{
                            type: 'linear',
                            position: 'right',
                            title: {{
                                display: true,
                                text: 'Days'
                            }},
                            grid: {{
                                drawOnChartArea: false
                            }}
                        }}
                    }}
                }}
            }});
        }}

        function createRushChart(data) {{
            const ctx = document.getElementById('rushChart').getContext('2d');

            new Chart(ctx, {{
                type: 'bar',
                data: {{
                    labels: data.map(d => d.RushRequest || 'Regular'),
                    datasets: [
                        {{
                            label: 'Avg Processing CT',
                            data: data.map(d => d['ProcCT(days)']),
                            backgroundColor: '#667eea'
                        }},
                        {{
                            label: 'Avg Total CT',
                            data: data.map(d => d['TotalCT(Days)']),
                            backgroundColor: '#48bb78'
                        }}
                    ]
                }},
                options: {{
                    responsive: true,
                    maintainAspectRatio: true,
                    plugins: {{
                        legend: {{
                            display: true
                        }}
                    }},
                    scales: {{
                        y: {{
                            beginAtZero: true,
                            title: {{
                                display: true,
                                text: 'Days'
                            }}
                        }}
                    }}
                }}
            }});
        }}

        function createVoidTrendChart(data) {{
            const ctx = document.getElementById('voidTrendChart').getContext('2d');

            new Chart(ctx, {{
                type: 'line',
                data: {{
                    labels: data.map(d => d.YearMonth),
                    datasets: [{{
                        label: 'Void Rate (%)',
                        data: data.map(d => d.VoidRate),
                        borderColor: '#ef4444',
                        backgroundColor: 'rgba(239, 68, 68, 0.1)',
                        tension: 0.4,
                        fill: true
                    }}]
                }},
                options: {{
                    responsive: true,
                    maintainAspectRatio: true,
                    plugins: {{
                        legend: {{
                            display: true
                        }}
                    }},
                    scales: {{
                        y: {{
                            beginAtZero: true,
                            title: {{
                                display: true,
                                text: 'Void Rate (%)'
                            }}
                        }}
                    }}
                }}
            }});
        }}

        function createPercentile90thChart(data) {{
            const ctx = document.getElementById('percentile90thChart');
            if (!ctx) return;

            // Sort by count descending
            const sortedData = [...data].sort((a, b) => b.Count - a.Count);

            // Create color arrays - highlight 3Z entries in purple/violet
            const countColors = sortedData.map(d =>
                d.Is3Z ? 'rgba(139, 92, 246, 0.8)' : 'rgba(239, 68, 68, 0.7)'
            );
            const countBorderColors = sortedData.map(d =>
                d.Is3Z ? '#8b5cf6' : '#dc2626'
            );
            const ctColors = sortedData.map(d =>
                d.Is3Z ? 'rgba(168, 85, 247, 0.8)' : 'rgba(251, 191, 36, 0.7)'
            );
            const ctBorderColors = sortedData.map(d =>
                d.Is3Z ? '#a855f7' : '#f59e0b'
            );

            new Chart(ctx, {{
                type: 'bar',
                data: {{
                    labels: sortedData.map(d => d['ECN Topic']),
                    datasets: [
                        {{
                            label: 'Number of ECNs in 90th Percentile',
                            data: sortedData.map(d => d.Count),
                            backgroundColor: countColors,
                            borderColor: countBorderColors,
                            borderWidth: 2,
                            yAxisID: 'y'
                        }},
                        {{
                            label: 'Average Processing CT (days)',
                            data: sortedData.map(d => d.AvgProcCT),
                            backgroundColor: ctColors,
                            borderColor: ctBorderColors,
                            borderWidth: 2,
                            yAxisID: 'y1'
                        }}
                    ]
                }},
                options: {{
                    responsive: true,
                    maintainAspectRatio: true,
                    interaction: {{
                        mode: 'index',
                        intersect: false
                    }},
                    plugins: {{
                        legend: {{
                            display: true,
                            position: 'top'
                        }},
                        title: {{
                            display: true,
                            text: 'Slowest 10% of ECNs - Count and Cycle Time by Type',
                            font: {{
                                size: 14
                            }}
                        }},
                        tooltip: {{
                            callbacks: {{
                                afterLabel: function(context) {{
                                    const dataIndex = context.dataIndex;
                                    const topic = sortedData[dataIndex];
                                    if (context.datasetIndex === 0) {{
                                        return `Avg CT: ${{topic.AvgProcCT.toFixed(2)}} days`;
                                    }} else {{
                                        return `Count: ${{topic.Count}} ECNs`;
                                    }}
                                }}
                            }}
                        }}
                    }},
                    scales: {{
                        y: {{
                            type: 'linear',
                            position: 'left',
                            beginAtZero: true,
                            title: {{
                                display: true,
                                text: 'Number of ECNs',
                                color: '#dc2626'
                            }},
                            ticks: {{
                                color: '#dc2626',
                                callback: function(value) {{
                                    return value.toLocaleString();
                                }}
                            }}
                        }},
                        y1: {{
                            type: 'linear',
                            position: 'right',
                            beginAtZero: true,
                            title: {{
                                display: true,
                                text: 'Processing Cycle Time (days)',
                                color: '#f59e0b'
                            }},
                            ticks: {{
                                color: '#f59e0b',
                                callback: function(value) {{
                                    return value.toFixed(1);
                                }}
                            }},
                            grid: {{
                                drawOnChartArea: false
                            }}
                        }},
                        x: {{
                            ticks: {{
                                maxRotation: 45,
                                minRotation: 45,
                                font: {{
                                    size: 10
                                }}
                            }}
                        }}
                    }}
                }}
            }});
        }}

        // Advanced KPI Charts
        function createFTRTrendChart(data) {{
            const ctx = document.getElementById('ftrTrendChart').getContext('2d');

            new Chart(ctx, {{
                type: 'line',
                data: {{
                    labels: data.map(d => d.YearMonth),
                    datasets: [{{
                        label: 'FTR Rate (%)',
                        data: data.map(d => d.FTR_Rate),
                        borderColor: '#48bb78',
                        backgroundColor: 'rgba(72, 187, 120, 0.1)',
                        tension: 0.4,
                        fill: true
                    }}]
                }},
                options: {{
                    responsive: true,
                    maintainAspectRatio: true,
                    plugins: {{
                        legend: {{
                            display: true
                        }}
                    }},
                    scales: {{
                        y: {{
                            beginAtZero: true,
                            max: 100,
                            title: {{
                                display: true,
                                text: 'FTR Rate (%)'
                            }}
                        }}
                    }}
                }}
            }});
        }}

        function createHoldTopicChart(data) {{
            const ctx = document.getElementById('holdTopicChart').getContext('2d');
            const sortedData = [...data].sort((a, b) => b.HoldRate - a.HoldRate).slice(0, 10);

            new Chart(ctx, {{
                type: 'bar',
                data: {{
                    labels: sortedData.map(d => truncateLabel(d['ECN Topic'], 25)),
                    datasets: [{{
                        label: 'Hold Rate (%)',
                        data: sortedData.map(d => d.HoldRate),
                        backgroundColor: '#f59e0b'
                    }}]
                }},
                options: {{
                    responsive: true,
                    maintainAspectRatio: true,
                    indexAxis: 'y',
                    plugins: {{
                        legend: {{
                            display: false
                        }}
                    }},
                    scales: {{
                        x: {{
                            beginAtZero: true,
                            title: {{
                                display: true,
                                text: 'Hold Rate (%)'
                            }}
                        }}
                    }}
                }}
            }});
        }}

        function createCoordWorkloadChart(data) {{
            const ctx = document.getElementById('coordWorkloadChart').getContext('2d');
            const sortedData = [...data].sort((a, b) => b.RequestNum - a.RequestNum).slice(0, 10);

            new Chart(ctx, {{
                type: 'doughnut',
                data: {{
                    labels: sortedData.map(d => d.ECNCoordinator || 'Unassigned'),
                    datasets: [{{
                        data: sortedData.map(d => d.RequestNum),
                        backgroundColor: [
                            '#667eea', '#48bb78', '#f59e0b', '#ef4444', '#8b5cf6',
                            '#ec4899', '#06b6d4', '#84cc16', '#f97316', '#14b8a6'
                        ]
                    }}]
                }},
                options: {{
                    responsive: true,
                    maintainAspectRatio: true,
                    plugins: {{
                        legend: {{
                            position: 'right'
                        }}
                    }}
                }}
            }});
        }}

        function createRushTrendChart(data) {{
            const ctx = document.getElementById('rushTrendChart').getContext('2d');

            new Chart(ctx, {{
                type: 'line',
                data: {{
                    labels: data.map(d => d.YearMonth),
                    datasets: [{{
                        label: 'Rush Rate (%)',
                        data: data.map(d => d.RushRate),
                        borderColor: '#8b5cf6',
                        backgroundColor: 'rgba(139, 92, 246, 0.1)',
                        tension: 0.4,
                        fill: true
                    }}]
                }},
                options: {{
                    responsive: true,
                    maintainAspectRatio: true,
                    plugins: {{
                        legend: {{
                            display: true
                        }}
                    }},
                    scales: {{
                        y: {{
                            beginAtZero: true,
                            title: {{
                                display: true,
                                text: 'Rush Rate (%)'
                            }}
                        }}
                    }}
                }}
            }});
        }}

        function createMfgSiteChart(data) {{
            const ctx = document.getElementById('mfgSiteChart').getContext('2d');
            const sortedData = [...data].sort((a, b) => b.RequestNum - a.RequestNum).slice(0, 10);

            new Chart(ctx, {{
                type: 'bar',
                data: {{
                    labels: sortedData.map(d => d.MFGSiteCode),
                    datasets: [{{
                        label: 'Request Count',
                        data: sortedData.map(d => d.RequestNum),
                        backgroundColor: '#06b6d4'
                    }}]
                }},
                options: {{
                    responsive: true,
                    maintainAspectRatio: true,
                    indexAxis: 'y',
                    plugins: {{
                        legend: {{
                            display: false
                        }}
                    }},
                    scales: {{
                        x: {{
                            beginAtZero: true
                        }}
                    }}
                }}
            }});
        }}

        function truncateLabel(label, maxLength) {{
            if (!label) return '';
            return label.length > maxLength ? label.substring(0, maxLength) + '...' : label;
        }}
    </script>
</body>
</html>
"""

# Save enhanced dashboard
with open('ECN_Metrics_Dashboard.html', 'w', encoding='utf-8') as f:
    f.write(html_content)

print("Enhanced dashboard created successfully!")
print("File saved as: ECN_Metrics_Dashboard.html")
print("\nFeatures:")
print("  [OK] Tab navigation (Overview, Advanced KPIs, Executive Summary)")
print("  [OK] Executive summary with insights and recommendations")
print("  [OK] Advanced KPI tables and charts")
print("  [OK] All data embedded - works offline")
