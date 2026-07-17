# ECN Metrics Dashboard Generator - Setup Guide

Professional dashboard for analyzing ECN (Engineering Change Notice) cycle time metrics with interactive visualizations and PDF reporting.

## Features

- **Interactive Dashboard** with real-time data visualization
- **Cycle Time Analysis** across sites, coordinators, and topics
- **SYSTEM USERID Highlighting** in red for easy identification
- **Site Distribution** separated by SYSTEM USERID vs Other Coordinators
- **PDF Executive Summary** with color-coded performance metrics
- **Multiple Export Options**: Excel, JSON, and PDF downloads
- **Responsive Charts** using Chart.js
- **Advanced KPIs** including percentile analysis and quality metrics

## Prerequisites

- **Python 3.7+** installed on your system
- **Internet connection** (for CDN-hosted libraries in dashboard)
- **Modern web browser** (Chrome, Edge, or Firefox recommended)

## Installation

### Step 1: Install Python Dependencies

```powershell
pip install pandas openpyxl
```

Or use the provided requirements file:

```powershell
pip install -r requirements.txt
```

### Step 2: Verify Installation

```powershell
python --version
# Should show Python 3.7 or higher
```

## Required Files

Place these files in the same directory:

1. **ECN Coordinators.xlsx** (provided)
   - Maps user IDs to coordinator full names
   - Format: Username | Coordinator columns

2. **Your ECN Data File** (Excel format)
   - Must contain these columns:
     - `RequestNum` - ECN request number
     - `ECNCoordinator` - Coordinator user ID
     - `ECNCoordinatorSite` - Site location
     - `ECN Topic` - Category/topic
     - `State` - ECN state
     - `ProcCT(days)` - Processing cycle time
     - `TotalCT(Days)` - Total cycle time
     - `RushRequest` - Rush flag (optional)
     - Date columns for time-based analysis

## Usage

### Quick Start (3 Steps)

1. **Process Your Data**
   ```powershell
   python process_data.py
   ```
   - Reads your ECN Excel file
   - Maps coordinator names
   - Generates `data.json`
   - Creates `source_data.xlsx` for downloads

2. **Generate Dashboard**
   ```powershell
   python create_enhanced_dashboard.py
   ```
   - Creates interactive `dashboard.html`
   - Embeds all data for standalone use
   - Ready to open in browser

3. **Open Dashboard**
   - Double-click `dashboard.html`
   - OR for best experience, use local server:
     ```powershell
     python -m http.server 8000
     ```
     Then open: http://localhost:8000/dashboard.html

### Customization

#### Modify Data File Name

Edit `process_data.py` line 11:
```python
df = pd.read_excel('YOUR_FILE_NAME.xlsx')
```

#### Update Coordinator Mapping

Edit `ECN Coordinators.xlsx`:
- Column A: Username (user ID)
- Column B: Coordinator (full name)

#### Filter Data by Date Range

Edit `process_data.py` to add date filtering:
```python
# After reading the Excel file
df = df[df['YourDateColumn'] >= '2025-01-01']
```

## Dashboard Features

### Overview Tab
- **Key Metrics Cards**: Total requests, average CT, void rate, FTR rate
- **Monthly Trends Chart**: Cycle time trends over time
- **Site Distribution**: Separated SYSTEM USERID vs Other Coordinators
- **Coordinator Performance**: Top performers ranked
- **ECN Topic Analysis**: Most common categories
- **State Distribution**: ECN status breakdown

### Executive Summary Tab
- High-level KPIs and trends
- Percentile analysis (50th, 75th, 90th)
- Top categories by performance tier
- Key insights and recommendations

### Advanced KPIs Tab
- Detailed coordinator comparison
- Rush request analysis
- Cycle time distribution
- Quality metrics breakdown

### Download Options

1. **Download Full ECN Data (Excel)** - Green button
   - Complete source data in Excel format
   - File: `source_data.xlsx`

2. **Download Executive Summary (PDF)** - Purple button
   - Professional PDF report
   - Color-coded performance metrics
   - Ready for stakeholder presentations

3. **Download Metrics (JSON)** - Blue button
   - Raw metrics data
   - For further analysis or integration

## Troubleshooting

### Dashboard Shows "Loading Dashboard..."

**Solution 1**: Hard refresh browser
```
Press Ctrl + Shift + R
```

**Solution 2**: Check browser console (F12)
- Look for red error messages
- Verify data.json exists

**Solution 3**: Use local server
```powershell
python -m http.server 8000
# Open http://localhost:8000/dashboard.html
```

### PDF Download Not Working

**Check**: Internet connection (jsPDF loads from CDN)

**Solution**: Use local web server
```powershell
python -m http.server 8000
```

**Test**: In browser console (F12), type:
```javascript
window.jspdf
```
Should show object, not "undefined"

### Excel Download Button Does Nothing

**Cause**: `source_data.xlsx` file missing

**Solution**: Re-run data processing
```powershell
python process_data.py
```

### Coordinator Names Not Showing

**Cause**: `ECN Coordinators.xlsx` file missing or format wrong

**Solution**:
1. Verify file exists in same folder
2. Check Excel format:
   - Column A: Username
   - Column B: Coordinator
   - No extra headers

### SYSTEM USERID Not Highlighted in Red

**Solution**: Refresh browser with hard reload
```
Press Ctrl + Shift + R
```

## File Structure

```
project-folder/
├── process_data.py              # Step 1: Data processing script
├── create_enhanced_dashboard.py # Step 2: Dashboard generator
├── ECN Coordinators.xlsx        # Coordinator name mapping
├── YOUR_ECN_DATA.xlsx          # Your input data (not included)
├── requirements.txt             # Python dependencies
├── SETUP_GUIDE.md              # This file
│
├── Generated Files (created by scripts):
├── data.json                    # Processed metrics
├── source_data.xlsx            # Data for Excel download
└── dashboard.html              # Interactive dashboard
```

## Best Practices

### For Best Performance

1. **Use Local Server** instead of opening file:// directly
   ```powershell
   python -m http.server 8000
   ```

2. **Chrome Browser** recommended for best PDF generation

3. **Wait for Full Load** before clicking download buttons (2-3 seconds)

### For Sharing Dashboard

**Option 1**: Share standalone HTML
- Send `dashboard.html` file only
- Recipient opens in browser
- All data embedded, no other files needed

**Option 2**: Share complete package
- Include `data.json` and `source_data.xlsx`
- Enables all download features
- Better for team collaboration

**Option 3**: Host on web server
- Upload files to internal web server
- Share URL instead of file
- Best for large teams

### For Regular Updates

1. Save your ECN data file with consistent name
2. Run processing script regularly:
   ```powershell
   python process_data.py && python create_enhanced_dashboard.py
   ```
3. Dashboard automatically updates with latest data

## Technical Details

### Technologies Used
- **Python**: pandas, openpyxl
- **JavaScript**: Chart.js v4, jsPDF v2.5.1
- **CSS**: Custom responsive design
- **HTML5**: Standalone dashboard

### Browser Compatibility
- Chrome 90+
- Edge 90+
- Firefox 88+
- Safari (PDF download may have issues)
- IE11 (not supported)

### Data Processing
- Automatic coordinator name mapping
- Percentile calculations (50th, 75th, 90th)
- Void rate and FTR (First Time Right) calculations
- Time-based trend analysis
- Site and topic aggregations

### Security Notes
- Dashboard runs entirely client-side
- No data sent to external servers
- CDN libraries (Chart.js, jsPDF) loaded from trusted sources
- Safe to use with confidential data

## Common Questions

**Q: Can I use this with different data formats?**
A: Yes, modify `process_data.py` to match your column names

**Q: How do I add more coordinators?**
A: Edit `ECN Coordinators.xlsx` and add new rows

**Q: Can I customize the dashboard colors?**
A: Yes, edit `create_enhanced_dashboard.py` CSS section

**Q: Dashboard is slow with large datasets?**
A: Filter data by date range in `process_data.py`

**Q: Can I automate this with a scheduled task?**
A: Yes, use Windows Task Scheduler to run scripts daily/weekly

---

## Quick Reference Card

```
GENERATE DASHBOARD:
1. python process_data.py
2. python create_enhanced_dashboard.py
3. Open dashboard.html

BEST VIEWING:
python -m http.server 8000
http://localhost:8000/dashboard.html

DOWNLOADS:
Green = Excel | Purple = PDF | Blue = JSON

FEATURES:
✓ Red SYSTEM USERID highlighting
✓ Separated site metrics
✓ Interactive charts
✓ PDF export
✓ Advanced KPIs
```

---

Created for ECN cycle time analysis and reporting.
Built with Python, Chart.js, and jsPDF.
