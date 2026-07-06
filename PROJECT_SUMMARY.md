# SEA Insurance MIS Dashboard — Complete Project Summary

**Document Version**: Final (Handoff to Claude Code)  
**Status**: Production-ready (7 phases complete, 1 pending)  
**Last Updated**: June 2026

---

## 1. PROJECT OVERVIEW

### What We're Building
A live, hostable HTML dashboard for SEA Insurance MIS reporting covering Philippines (PH) and Indonesia (INDO) markets. The dashboard tracks insurance channel performance metrics across monthly and QTD (Quarter-to-Date) views against AOP (Annual Operating Plan) targets.

### Key Features
- **Dual Country Support**: Separate tabs for PH (8 channels) and INDO (6 channels)
- **Custom Date Range Picker**: Start/end date selection (Apr-26 to Mar-27 FY27)
- **Dynamic QTD Assembly**: Automatically combines closed/reported months + live current month
- **Previous Quarter Data**: Pulls from Insurance Actuals file (handles all FY quarters)
- **Freeze/Unfreeze**: Locks data snapshots to Excel, no auto-advance
- **Next Period Helper**: Opt-in button to suggest day after last frozen end date (stored in localStorage)

### Metrics Tracked Per Channel
- **NOPs**: Number of Policies
- **NP**: Net Premium (USD Thousands)
- **Rev**: Revenue (USD Thousands)
- **GM**: Gross Margin (USD Thousands)
- **Rev%**: Revenue % of Premium (computed)
- **GM%**: Gross Margin % of Premium (computed)

### Dashboard Columns (PH: 7, INDO: 5)
| Column | PH | INDO | Source |
|--------|:--:|:----:|--------|
| Current Month Actuals | ✓ | ✓ | NAV Summary (live) |
| w/ Passthrough | ✓ | — | NAV Summary (live) |
| AOP | ✓ | ✓ | AOP File |
| QTD Actuals | ✓ | ✓ | NAV + Insurance Actuals |
| QTD w/ Passthrough | ✓ | — | NAV + Insurance Actuals |
| QTD AOP | ✓ | ✓ | AOP File |
| Prev Quarter Actuals | ✓ | ✓ | Insurance Actuals |

---

## 2. ARCHITECTURE

### Technology Stack
- **Frontend**: HTML5 + Vanilla JavaScript (no frameworks)
- **Backend**: Google Apps Script (standalone web app, single project)
- **Data Sources**: 
  - **NAV Summary Sheets** (live, updated daily): `_PH_Summary`, `_INDO_Summary`
  - **Insurance Actuals File** (reported/frozen historical): `PH Actuals FY27/FY26`, `Indo Actuals FY27/FY26`
  - **AOP File** (targets): `PH - AOP27`, `INDO - AOP27`
- **Storage**: 
  - **Server**: CacheService (10 min TTL, automatic cache busting)
  - **Client**: localStorage (freeze state, next-period hint)
- **Export**: SheetJS XLSX + CSV fallback

### Request Flow Diagram
```
User selects date range (e.g., May-26 to May-26)
         ↓
buildQTDDescriptor() analyzes end date
  - Determines which months are closed (before end date in same Q)
  - Determines live month (the end date itself)
  - Looks up previous quarter definition
         ↓
loadCountry() executes sequential API calls:
  1. getActualsFromFile(prev quarter, useQtrCol=true) → Insurance Actuals
  2. For each closed month: getActualsFromFile(month, useQtrCol=false) → Insurance Actuals
  3. getActuals(start, end, forceFresh=false) → NAV Summary (user's range)
  4. getAOP(lastMonth) → AOP File
         ↓
sumChannels() combines reported + live into QTD
         ↓
renderPH() or renderINDO() displays dashboard with:
  - Data rows with color coding
  - QTD header showing composition
  - Frozen overlay (if frozen)
  - Control bar (Freeze/Unfreeze/Encode/Next Period)
```

### State Machine (Per Country, Per Date Range)
```
[LIVE] ← default, fetch fresh data every time
  ↓ on Freeze
[FROZEN] ← snapshot saved to localStorage, export Excel
  ↓ on Next Period / Unfreeze
[LIVE] ← clear snapshot, fetch fresh
```

### localStorage Keys
| Key | Value | Set By | Used For |
|-----|-------|--------|----------|
| `mis_state_{country}_{start}_{end}` | `LIVE` or `FROZEN` | setState() | Track freeze state |
| `mis_snap_{country}_{start}_{end}` | JSON string (channel data) | saveSnap() | Store frozen snapshot |
| `mis_next_start_{country}` | ISO date string | doFreeze() | Remember next period hint |

---

## 3. FILE REFERENCES

### Google Sheets Files (Source of Truth)

| File | ID | Workbook/Tabs | Purpose | Updated |
|---|---|---|---|---|
| **PH NAV Workbook** | `1_r-TfN7c4yrXcANEZP9eVKYf7oOJhLU2Z7wRBw7bPb4` | `_PH_Summary` | Live policy data summary (8 channels × SUMIFS) | Daily |
| **INDO NAV Workbook** | `1925KcTeBptHNkDvRLTkKiL5MsikwpMrwK8GQLLPHRDA` | `_INDO_Summary` | Live policy data summary (6 channels × SUMIFS) | Daily |
| **AOP File** | `10KsfE-shzOlFG_u1WLAPCgcYZx6FeJNnoiWFQ4YfoTw` | `PH - AOP27`, `INDO - AOP27` | Monthly & quarterly targets | Q-end |
| **Insurance Actuals** | `1saDApEjljXF_1fYui_UXnL2LYYaQ2Mz6YmAdtoW8UsA` | `PH Actuals FY27`, `PH Actuals FY26`, `Indo Actuals FY27`, `Indo Actuals FY26` | Reported/closed historical | Monthly |

### Apps Script Files

- **Code.gs** (530 lines)
  - Handles all backend logic
  - `doGet()` router
  - Data fetching + caching
  - Health check

- **index.html** (991 lines)
  - Complete dashboard UI + JavaScript
  - `buildQTDDescriptor()` logic
  - `loadCountry()` orchestration
  - Freeze/Unfreeze/Next Period state machine
  - `renderPH()` and `renderINDO()` table builders

---

## 4. PROGRESS & PHASES (COMPLETED)

### Phase 1 ✅ — Pre-aggregation Summary Sheets
**Status**: Complete. Summary sheets built with SUMIFS formulas in existing NAV workbooks.

**Details**:
- `_PH_Summary`: 8 channels (B2C, Renewal, Branch, UC Attached, B2B & Affiliates, Agency, B2B2C Subtotal, Philippines Insurance)
- `_INDO_Summary`: 6 channels (Tele-Sales, Agency, Dealer, B2B/SAU, B2B2C Subtotal, INDO Insurance Total)
- Data rows: PH (6–13), INDO (7–12)
- Columns: A=Channel, B=NOPs, C=NP, D=Rev, E=GM
- Filter rule: **Include all statuses except "Cancelled" and "Not Issued"** (important: "Pending Policy" is included)
- Currency conversion: PHP÷58÷1000, IDR÷16200÷1000 = USD Thousands

### Phase 2 ✅ — AOP Mapping & Debugging
**Status**: Complete. Correct row maps verified through multiple test iterations.

**PH AOP Row Map** (0-indexed, from `PH - AOP27` tab):
```
B2C (Tele-Sales in file):           NOPs=1,  NP=2,  Rev=3,  GM=5
Renewal:                            NOPs=7,  NP=8,  Rev=9,  GM=11
Branch:                             NOPs=24, NP=25, Rev=26, GM=28
B2B & Affiliates:                   NOPs=47, NP=48, Rev=49, GM=51
Agency:                             NOPs=53, NP=54, Rev=55, GM=57
UC Attached:                        null (always dash)
```

**INDO AOP Row Map** (0-indexed, from `INDO - AOP27` tab):
```
Tele-Sales:                         NOPs=31, NP=32, Rev=33, GM=35
Agency:                             NOPs=14, NP=15, Rev=16, GM=18
B2B/SAU:                            NOPs=8,  NP=9,  Rev=10, GM=12
Dealer:                             null (always dash)
```

**Monthly Columns** (0-indexed):
- Apr-26=2, May-26=3, Jun-26=4, Jul-26=5, Aug-26=6, Sep-26=7, Oct-26=8, Nov-26=9, Dec-26=10, Jan-27=11, Feb-27=12, Mar-27=13

**QTD Columns** (0-indexed):
- PH: Q1=16, Q2=17, Q3=18, Q4=19
- INDO: Q1=2, Q2=3, Q3=4, Q4=5

### Phase 3 ✅ — Apps Script API with Cache Layer
**Status**: Complete. Single standalone project with doGet() router and three main endpoints.

**Core Functions**:
- `getActuals(country, startDate, endDate, forceFresh)` — reads NAV summary
- `getAOP(country, month)` — reads AOP targets
- `ping()` — health check (verifies all 4 file accesses)

**Cache**: CacheService with 10 min TTL. Keys include country + date range.
**Warm Trigger**: `warmCache()` function (for time-based trigger, every 5 min during business hours).

### Phase 4 ✅ — HtmlService + Date Range Support
**Status**: Complete. Switched from fetch() to google.script.run (CORS resolved).

**Implementation**:
- Dashboard served via HtmlService from same Apps Script project
- No external API calls; all data routed through Apps Script
- Summary sheets accept ISO date ranges in B1/B2 (auto-populated by getActuals)

### Phase 5 ✅ — QTD Assembly Logic
**Status**: Complete. `buildQTDDescriptor()` determines closed vs live months.

**Logic**:
- End date → current quarter
- Months before end date in same Q → "closed" (reported)
- End date itself → "live" (real-time NAV)
- QTD = sum(closed reported) + (live from NAV)

**Example** (May-26 selected in Q1 Apr-Jun):
- Apr-26 (closed) → from Insurance Actuals
- May-26 (live) → from NAV Summary
- QTD = Apr + May

### Phase 6 ✅ — Insurance Actuals Integration
**Status**: Complete. New `getActualsFromFile()` handles reported data + prev quarters.

**Function Signature**:
```javascript
getActualsFromFile(country, month, fy, useQtrCol, qtr)
```

**Two Modes**:
1. `useQtrCol=true`: Reads quarterly aggregate (Q1/Q2/Q3/Q4 columns)
2. `useQtrCol=false`: Reads individual month column (Apr-26, May-26, etc.)

**Used For**:
- Previous quarter (always useQtrCol=true)
- Closed months within current quarter (useQtrCol=false)

**Previous Quarter Map** (by current quarter):
```
If Q1 (Apr-Jun):  prev = FY26 Q4 (Jan-Mar 26)  → read col Q4 from Insurance Actuals FY26
If Q2 (Jul-Sep):  prev = FY27 Q1 (Apr-Jun 26)  → read col Q1 from Insurance Actuals FY27
If Q3 (Oct-Dec):  prev = FY27 Q2 (Jul-Sep 26)  → read col Q2 from Insurance Actuals FY27
If Q4 (Jan-Mar):  prev = FY27 Q3 (Oct-Dec 26)  → read col Q3 from Insurance Actuals FY27
```

### Phase 7 ✅ — Freeze/Unfreeze/Next Period State Machine
**Status**: Complete. Freeze keeps date range fixed; Next Period is opt-in button.

**Key Change from Previous**:
- **Old behavior**: Freeze auto-advanced dates (confusing workflow)
- **New behavior**: Freeze keeps dates fixed, stores next-start hint, user clicks button to advance

**doFreeze()**:
1. Save state = FROZEN
2. Save data snapshot to localStorage
3. Export Excel
4. Store next-start hint (day after end date) in localStorage
5. Re-render same range in frozen state
6. **Does NOT change date inputs**

**applyNextPeriod()**:
1. Retrieve stored next-start from localStorage
2. Set start = that date, end = end of that month
3. Load fresh live data for new range
4. User can edit dates afterward if needed

**doUnfreeze()**:
1. Clear frozen state
2. Clear localStorage snapshot
3. Force fresh fetch from server (`forceFresh=true`)
4. Re-render live data

---

## 5. INSURANCE ACTUALS FILE STRUCTURE

### Layout (All 4 tabs: PH FY27/FY26, INDO FY27/FY26)

```
Col A              Col B               Cols C-N            Col O-P          Cols Q-T
Channel            Metric              Individual months   FY total + gap   Quarterly
(merged by block)  label               Apr-26...Mar-27    (skipped)        Q1 Q2 Q3 Q4
                                      (cols 2-13)        (cols 14-15)     (cols 16-19)

Row 1:  Header
Row 2+: Data (channel blocks)
```

### Metric Offsets Per Channel Block
```
+0  NOPs (number of policies)
+1  Premium (net premium)
+2  Revenue
+3  Rev% (percentage — skip, computed)
+4  Gross Margin
+5  GM% (percentage — skip, computed)
+6-9: P&O, Marketing, CM, CM% (skip, not in dashboard)
```

### PH Actuals Row Maps (0-indexed start row of channel)
```
B2C (labeled Tele-Sales in file):   1
Renewal:                            7
Branch (labeled B2C in file):       13
B2B & Affiliates:                   46
Agency:                             52
UC Attached:                        noData (null)
```

### INDO Actuals Row Maps
```
Tele-Sales:   1
Agency:       36
B2B/SAU:      30
Dealer:       noData (null)
```

### Month Columns (Insurance Actuals, 0-indexed)
```
Apr-26: 2,  May-26: 3,  Jun-26: 4,  Jul-26: 5,   Aug-26: 6,  Sep-26: 7,
Oct-26: 8,  Nov-26: 9,  Dec-26: 10, Jan-27: 11,  Feb-27: 12, Mar-27: 13
```

### Quarter Columns
```
Q1: 16, Q2: 17, Q3: 18, Q4: 19
```

---

## 6. CODE COMPONENTS IN DETAIL

### Apps Script: Core Functions

#### `doGet(e)` — Router
- Checks `e.parameter.action` (ping, getActuals, getActualsFromFile, getAOP)
- Returns HtmlService for index (no action param)
- Routes to appropriate function

#### `getActuals(country, startDate, endDate, forceFresh)`
**Reads**: Live data from NAV summary sheets

**Process**:
1. Check cache (skip if forceFresh=true)
2. Open NAV workbook (PH or INDO)
3. Set B1 = startDate, B2 = endDate (triggers SUMIFS recalc)
4. Flush
5. Read 8 rows (PH) or 6 rows (INDO) from columns A-E
6. Parse channels into {nops, np, rev, gm}
7. Cache result (TTL 600s)
8. Return

**Return Format**:
```javascript
{
  success: true,
  country: "PH",
  startDate: "2026-05-01",
  endDate: "2026-05-31",
  fromCache: false,
  fetchedAt: "2026-06-13T10:30:00.000Z",
  channels: {
    "B2C": {nops: 100, np: 50.5, rev: 75.2, gm: 25.1},
    ...
  }
}
```

#### `getActualsFromFile(country, month, fy, useQtrCol, qtr)`
**Reads**: Reported historical data from Insurance Actuals file

**Parameters**:
- `country`: 'PH' | 'INDO'
- `month`: 'Apr-26', 'May-26', etc. (ignored if useQtrCol=true)
- `fy`: 'FY27' | 'FY26'
- `useQtrCol`: true = read quarterly column | false = read monthly
- `qtr`: 'Q1' | 'Q2' | 'Q3' | 'Q4' (required if useQtrCol=true)

**Process**:
1. Build cache key from country + fy + (qtr or month)
2. Check cache
3. Resolve tab name (PH Actuals FY27, etc.)
4. Open Insurance Actuals file, read full sheet
5. Determine column (ACTUALS_MONTH_COL or ACTUALS_QTR_COL)
6. For each channel in row map:
   - Extract nops, np, rev, gm from correct row + metric offsets
   - Handle noData channels (null)
7. Add subtotals (same logic as AOP)
8. Cache result
9. Return

#### `getAOP(country, month)`
**Reads**: AOP targets from AOP file (monthly + quarterly)

**Returns**:
```javascript
{
  success: true,
  country: "PH",
  month: "May-26",
  fromCache: false,
  fetchedAt: "...",
  channels: { /* same structure as actuals */ },
  qtd: { /* quarterly data */ }
}
```

#### Cache Functions
- `cacheGet(key)`: Retrieve from CacheService, parse JSON
- `cacheSet(key, data)`: Serialize + store with TTL=600s
- `cacheBust(key)`: Remove single key
- `cacheBustAll(country, start, end)`: Remove all keys for a range (monthly + quarterly combos)

### HTML: Core Functions

#### `buildQTDDescriptor(endDate)`
**Input**: ISO end date (e.g., '2026-05-31')

**Output**:
```javascript
{
  qLabel: "Q1",              // current quarter
  qShort: "AMJ",             // month initials
  closedMonths: ["Apr-26"],  // months before end date
  liveMonth: "May-26",       // end date itself
  monthIdx: 1,               // position in quarter (0=all, 1=2nd month, etc.)
  prevDef: {
    fy: "FY26",
    qtrShort: "Q4",
    months: ["Jan-26", "Feb-26", "Mar-26"],
    isoStart: "2026-01-01",
    isoEnd: "2026-03-31"
  }
}
```

**Logic**: Analyzes which quarter the end date falls in, determines closed vs live months within that quarter, and looks up previous quarter definition.

#### `loadCountry(country, start, end, forceFresh)`
**Orchestrates** all data fetches in sequence:

1. Call `buildQTDDescriptor(end)` → get closedMonths, liveMonth, prevDef
2. Call `getActualsFromFile(country, '', prevDef.fy, true, prevDef.qtrShort)` → prev quarter
3. For each month in closedMonths: `getActualsFromFile(country, month, 'FY27', false, '')` → each closed month
4. Call `getActuals(country, start, end, forceFresh)` → range actuals (NAV, LAST so B1/B2 stay on user dates)
5. Call `getAOP(country, liveMonth)` → targets
6. Combine via `sumChannels()` into QTD
7. Call `renderTable()` with assembled data
8. Call `updateControlBar()`

**Key**: Sequential order ensures B1/B2 writes don't race. Range actuals is always last so user's dates remain in the cells.

#### `doFreeze()`
1. Verify data loaded (`dataCache[key]` exists)
2. Call `setState(activeTab, s, e, 'FROZEN')`
3. Call `saveSnap(activeTab, s, e, dataCache[key])`
4. Call `exportToExcel(activeTab, s, e, dataCache[key])`
5. Store next-start hint: `localStorage.setItem('mis_next_start_' + activeTab, addDays(e, 1))`
6. Re-render table in frozen state: `renderTable(activeTab, dataCache[key], s, e)`
7. Call `updateControlBar()` (swaps button states)
8. Show toast: "frozen & exported ✓"

**Does NOT**: Change date inputs, load new data, or auto-advance

#### `applyNextPeriod()`
1. Retrieve hint: `localStorage.getItem('mis_next_start_' + activeTab)`
2. If no hint: show toast "No frozen period yet"
3. Set ns = hint, ne = endOfMonth(ns)
4. Guard: ns > '2027-03-31' → ns = '2027-03-31'
5. Set date inputs: `document.getElementById('startDate').value = ns`
6. Set active vars: `activeStart = ns; activeEnd = ne;`
7. Clear cache for new range
8. Call `loadAll()` (loads data for new range)

#### `doUnfreeze()`
1. Call `setState(activeTab, s, e, 'LIVE')`
2. Clear localStorage snapshot: `localStorage.removeItem(snapKey(...))`
3. Clear in-memory cache
4. Call `updateControlBar()`
5. Call `loadCountry(activeTab, s, e, true)` with forceFresh=true

#### `updateControlBar()`
Sets button disabled states based on freeze state + data availability:
- Freeze: enabled if state=LIVE AND hasData
- Unfreeze: enabled if state=FROZEN
- Encode: enabled if state=FROZEN
- Next: enabled if next-start hint exists

#### `renderPH()` / `renderINDO()`
Generates HTML table with:
- Channel rows (8 for PH, 6 for INDO)
- 6 metric rows per channel (NOPs, NP, Rev, Rev%, GM, GM%)
- 7 columns for PH (Current, Passthrough, AOP, QTD Actuals, QTD PT, QTD AOP, Prev Qtr)
- 5 columns for INDO (Current, AOP, QTD Actuals, QTD AOP, Prev Qtr)
- Color coding per column (e.g., #E2EFDA for Actuals, #F2F2F2 for AOP)
- FROZEN overlay if frozen
- QTD header with composition ("Apr reported + May live")
- Footnote with QTD breakdown + prev quarter source

---

## 7. KNOWN ISSUES & FIXES

### Issue 1: Cache Race Condition on B1/B2 Writes
**Problem**: Concurrent getActuals calls write to shared B1/B2 simultaneously, causing SUMIFS to use wrong date.

**Fix Applied**: Load sequence is now strictly sequential. Range actuals is always LAST, so user's dates rest in the cells at end.

### Issue 2: Q1 Previous Quarter (Jan-Mar 26) Not in Live NAV
**Problem**: FY26 Q4 doesn't have live source in current NAV workbooks.

**Fix Applied**: Always read prev quarter from Insurance Actuals file (handles all FY quarters).

### Issue 3: Frozen Range Auto-Advanced (Old Behavior)
**Problem**: Freeze auto-advanced dates to next month, confusing workflow. User had to reset picker.

**Fix Applied**: Freeze keeps dates fixed. Stores next-start hint separately. User clicks "Next Period" button when ready.

### Issue 4: QTD Composition Unclear
**Problem**: User didn't know which months contributed to QTD.

**Fix Applied**: QTD header shows composition (e.g., "Apr reported + May live"). Footnote lists each month's source (Insurance Actuals or NAV).

---

## 8. DEPLOYMENT INSTRUCTIONS

### Step 1: Create/Update Apps Script Project
1. Go to [script.google.com](https://script.google.com)
2. Create new project (or open existing)
3. Replace all `Code.gs` content with provided `SEA_MIS_API.gs`
4. Create new file (File → New → Html file)
5. Name it `index`
6. Paste entire `index.html` content into it
7. Save all

### Step 2: Deploy as Web App
1. Click "Deploy" (top right)
2. Select "New deployment"
3. Type: "Web app"
4. Execute as: "Me" (your Google account)
5. Who has access: "Anyone"
6. Click Deploy
7. Copy the deployment URL (looks like `https://script.google.com/macros/d/{deploymentId}/userweb`)

### Step 3: Verify Health Check
Call: `https://your-deployment-url?action=ping`

Expected response:
```json
{
  "ok": true,
  "ts": "2026-06-13T10:30:00.000Z",
  "checks": {
    "PH": "ok",
    "INDO": "ok",
    "AOP": "ok",
    "ACTUALS": "ok"
  }
}
```

### Step 4: Open Dashboard
Navigate to: `https://your-deployment-url` (no parameters)

### Step 5 (Optional): Set Up Cache Warm Trigger
1. Apps Script → Triggers (left sidebar)
2. Click "Create trigger"
3. Function: `warmCache`
4. Type: Time-driven
5. Frequency: Every 5 minutes
6. Save

This keeps cache warm during business hours so first user of a period gets instant response.

---

## 9. API CALLING CONVENTIONS

### From HTML to Apps Script
All calls via `runScript(funcName, ...args)`:

```javascript
// Get live data
await runScript('getActuals', 'PH', '2026-05-01', '2026-05-31', false);

// Get from Insurance Actuals (mode 1: quarterly)
await runScript('getActualsFromFile', 'PH', '', 'FY26', true, 'Q4');

// Get from Insurance Actuals (mode 2: individual month)
await runScript('getActualsFromFile', 'PH', 'Apr-26', 'FY27', false, '');

// Get AOP targets
await runScript('getAOP', 'PH', 'May-26');

// Health check
await runScript('ping');
```

### Return Format (All Endpoints)
```javascript
{
  success: true,
  country: "PH",
  channels: {
    "B2C": {nops: 100, np: 50.5, rev: 75.2, gm: 25.1},
    "Renewal": { ... },
    ...
  },
  // Plus endpoint-specific fields (month, fy, col, qtd, etc.)
  fromCache: false,
  fetchedAt: "2026-06-13T10:30:00.000Z"
}
```

---

## 10. RECENT CHANGES (Latest Session)

### Feature: Freeze No Longer Auto-Advances Dates
- **Reason**: Previous behavior (freeze → auto-advance → load next month) was confusing. Users had to fight the auto-advance to stay frozen longer or go back.
- **New Behavior**: 
  - Freeze locks the data snapshot + exports Excel
  - Date range stays exactly where user left it
  - Table shows FROZEN overlay
  - On unfreeze, dates stay same, data re-fetches fresh
  - New opt-in "Next Period" button applies stored next-start when user clicks

### Code Changes:
1. **doFreeze()**: Removed lines that set date inputs. Added localStorage store for next-start hint.
2. **applyNextPeriod()** (new function): Retrieves hint, applies to inputs, loads new range
3. **updateControlBar()**: Added next-button enable logic (enabled when hint exists)
4. **CSS**: Added `.btn-next` style (purple #5e35b1)
5. **HTML**: Added Next Period button after Encode button
6. **doUnfreeze()**: Changed to use `loadCountry(activeTab, s, e, true)` (forceFresh=true)

---

## 11. QUICK REFERENCE

### File IDs (Copy-Paste Ready)
```
PH NAV:         1_r-TfN7c4yrXcANEZP9eVKYf7oOJhLU2Z7wRBw7bPb4
INDO NAV:       1925KcTeBptHNkDvRLTkKiL5MsikwpMrwK8GQLLPHRDA
AOP:            10KsfE-shzOlFG_u1WLAPCgcYZx6FeJNnoiWFQ4YfoTw
Actuals:        1saDApEjljXF_1fYui_UXnL2LYYaQ2Mz6YmAdtoW8UsA
```

### Sheet Names
```
PH NAV:         _PH_Summary
INDO NAV:       _INDO_Summary
PH AOP:         PH - AOP27
INDO AOP:       INDO - AOP27
PH Actuals:     PH Actuals FY27, PH Actuals FY26
INDO Actuals:   Indo Actuals FY27, Indo Actuals FY26
```

### Quarter Calendar (FY27)
```
Q1 (Apr-Jun 26):  Prev = FY26 Q4 (Jan-Mar 26)
Q2 (Jul-Sep 26):  Prev = FY27 Q1 (Apr-Jun 26)
Q3 (Oct-Dec 26):  Prev = FY27 Q2 (Jul-Sep 26)
Q4 (Jan-Mar 27):  Prev = FY27 Q3 (Oct-Dec 26)
```

### Test Scenarios
- [ ] Load May-26 (Q1, live month) → should show Apr reported + May live in QTD
- [ ] Load Jun-26 (Q1, live month) → should show Apr + May reported + Jun live in QTD
- [ ] Load Jul-26 (Q2, first month) → should show Q1 prev quarter total
- [ ] Freeze May-26, click Next Period → dates should jump to Jun-01 to Jun-30
- [ ] Freeze May-26 in PH, switch to INDO, freeze Jun-26 in INDO, switch back to PH, click Next Period → should respect PH's last frozen date
- [ ] Export Excel → should have sheet with date range in title
- [ ] Switch country tabs → should preserve frozen state per tab

---

## 12. GLOSSARY

| Term | Definition |
|---|---|
| **NAV** | Net Asset Value workbooks; contain raw policy data summarized via SUMIFS in `_Summary` sheets |
| **AOP** | Annual Operating Plan; target numbers for each channel/month/quarter |
| **QTD** | Quarter-to-Date; rolling aggregate of all months from start of quarter through current month |
| **Closed Month** | Month that has ended and is before the current/live month; uses reported data from Insurance Actuals |
| **Live Month** | Current reporting month; uses real-time data from NAV summary |
| **Prev Quarter** | Previous quarter's complete data (all 3 months); sourced from Insurance Actuals file |
| **Freeze** | Lock current date range's data snapshot + export to Excel; disables date picker, enables Unfreeze |
| **Unfreeze** | Return frozen range to LIVE mode; clears snapshot, fetches fresh data |
| **Encode** | (Phase 8) Write frozen snapshot back to Insurance Actuals file |
| **FX** | Foreign exchange; PHP÷58÷1000 = USD K, IDR÷16200÷1000 = USD K |
| **USD K** | US Dollars in thousands (1.5 K = $1,500) |
| **Passthrough** | (PH only) Revenue column accounting for commission pass-through to partner |
| **B2B2C** | Business-to-Business-to-Consumer; aggregate of UC Attached + B2B + Agency channels |

---

## 13. NEXT STEPS FOR CLAUDE CODE

### Priority Fixes
1. **QTD Edge Case**: Verify single-month selection (monthIdx=0) correctly sums all to QTD
2. **Insurance Actuals Access**: Test reading FY26 Q4 (previous quarter for Q1)
3. **Excel Export**: Verify metadata (date range, source notes) appears on export
4. **Cache Invalidation**: Stress test with rapid date changes; verify cache busts properly

### Pending Features (Phase 8)
1. **Encode Button**: Implement `doEncode()` to write frozen snapshot to Insurance Actuals file
2. **Time Trigger**: Confirm `warmCache()` trigger setup works (test by clearing cache, verifying auto-warm)
3. **Mobile UX**: Optional responsive adjustments for narrower screens

### Testing Checklist
- [ ] Load PH Q1 (Apr-Jun), verify prev quarter from FY26 Q4 Insurance Actuals
- [ ] Load INDO Q2 (Jul-Sep), verify prev quarter from FY27 Q1 Insurance Actuals
- [ ] Freeze May-26, check localStorage keys set correctly
- [ ] Unfreeze, verify forceFresh fetch bypasses cache
- [ ] Switch PH ↔ INDO, verify freeze state independent per tab
- [ ] Click Next Period from frozen state, verify dates advance correctly
- [ ] Export Excel from both countries
- [ ] Verify health check passes all 4 file accesses
- [ ] Load 2 consecutive ranges back-to-back, check cache reuse (should be instant)

---

**End of Summary. Ready for Claude Code!**
