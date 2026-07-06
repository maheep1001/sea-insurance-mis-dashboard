// ============================================================
//  SEA MIS DASHBOARD — Apps Script API
//  Single project handling both PH and INDO
//  Deploy as: Web App → Execute as Me → Anyone can access
// ============================================================
//
//  ENDPOINTS (routed by doGet ?action=… , or called via google.script.run):
//    ping                — health check: can we open all 4 files?
//    getActuals          — LIVE channel actuals from the NAV summary sheets,
//                          for any [startDate,endDate] (drives the summary's
//                          B1/B2 → SUMIFS). Returns { channels:{...} }.
//    getActualsFromFile  — REPORTED actuals from the Insurance Actuals file,
//                          by month column OR quarter column. Used for the
//                          Prev-Qtr column (quarter mode). Returns { channels }.
//    getAOP              — AOP targets from the AOP file. Returns BOTH a
//                          monthly bucket (.channels) AND a quarter bucket
//                          (.qtd). NOTE: .qtd is a FLAT channel map (no
//                          `channels` wrapper) — the dashboard wraps it.
//    createFlash         — builds a fresh monthly Flash workbook (see bottom).
//
//  RETURN-SHAPE CONVENTION: data endpoints return { channels: { '<Channel>':
//  {nops,np,rev,gm}, ... } }. getAOP additionally returns .qtd as a FLAT map.
//  The dashboard's getC() reads `.channels[key]`, so anything handed to it
//  must be wrapped accordingly (this is a common "whole column is blank" bug).
//
//  CURRENCY: NAV/AOP/Actuals are pre-converted to USD thousands
//  (PHP ÷ 58 ÷ 1000, IDR ÷ 16200 ÷ 1000) at source.
// ============================================================


// ── FILE IDs ─────────────────────────────────────────────────
//  The four Google files the dashboard reads. (Flash generation uses its own
//  source IDs in FLASH_CONFIG near the bottom.)
const FILE_IDS = {
  PH_NAV:   '1_r-TfN7c4yrXcANEZP9eVKYf7oOJhLU2Z7wRBw7bPb4', // live PH policy summary (_PH_Summary)
  INDO_NAV: '1925KcTeBptHNkDvRLTkKiL5MsikwpMrwK8GQLLPHRDA', // live INDO policy summary (_INDO_Summary)
  AOP:      '10KsfE-shzOlFG_u1WLAPCgcYZx6FeJNnoiWFQ4YfoTw', // AOP targets (monthly + quarterly)
  ACTUALS:  '1saDApEjljXF_1fYui_UXnL2LYYaQ2Mz6YmAdtoW8UsA'  // reported/closed history (Prev-Qtr source)
};


// ── SHEET NAMES ───────────────────────────────────────────────
//  Tab names inside the files above. *** If AOP / Prev-Qtr numbers don't show
//  up, first confirm these tab names match the workbook EXACTLY *** — a renamed
//  tab (e.g. a dedicated "… for Claude" tab) means the read silently misses.
//  This is the ONE place to repoint the AOP / Actuals source tabs.
const SHEET_NAMES = {
  PH_SUMMARY:    '_PH_Summary',
  INDO_SUMMARY:  '_INDO_Summary',
  PH_AOP:        'PH - AOP27',       // ← PH AOP source tab (monthly + QTD columns)
  INDO_AOP:      'INDO - AOP27',     // ← INDO AOP source tab
  PH_ACT_FY27:   'PH Actuals FY27',  // ← PH reported actuals, current FY  (Prev-Qtr Q2–Q4)
  PH_ACT_FY26:   'PH Actuals FY26 (recast)',  // ← PH reported actuals, prior FY    (Prev-Qtr for Q1)
  INDO_ACT_FY27: 'Indo Actuals FY27',
  INDO_ACT_FY26: 'Indo Actuals FY26'
};


// ── CACHE TTL (seconds) ───────────────────────────────────────
const CACHE_TTL = 600; // 10 minutes


// ── INSURANCE ACTUALS FILE: LAYOUT CONSTANTS ─────────────────
//  The Insurance Actuals tabs are laid out as: each channel occupies a BLOCK of
//  rows starting at `startRow` (0-based into the full sheet values array), and
//  within that block the metrics sit at fixed offsets (METRIC_OFFSETS). Each
//  month has its own column (ACTUALS_MONTH_COL) and each quarter an aggregate
//  column (ACTUALS_QTR_COL). getActualsFromFile() reads value[startRow+offset][col].

// Month → 0-based column index in the Insurance Actuals tabs.
const ACTUALS_MONTH_COL = {
  'Apr-26': 2,  'May-26': 3,  'Jun-26': 4,
  'Jul-26': 5,  'Aug-26': 6,  'Sep-26': 7,
  'Oct-26': 8,  'Nov-26': 9,  'Dec-26': 10,
  'Jan-27': 11, 'Feb-27': 12, 'Mar-27': 13
};

// Quarter → 0-based column index of the quarterly aggregate (used for Prev-Qtr).
const ACTUALS_QTR_COL = {
  'Q1': 16, 'Q2': 17, 'Q3': 18, 'Q4': 19
};

// PH channel → 0-based START ROW of its block in the Insurance Actuals tab.
// (noData:true → channel has no reported row, always returns nulls.)
// VERIFIED against live PH Actuals FY27 tab row structure:
//   Row 0:  Header
//   Row 1:  Tele-Sales (= B2C dashboard) NOPs, +1 NP, +2 Rev, +4 GM
//   Row 7:  Renewal
//   Row 13: B2C subtotal (Tele-Sales + Renewal — skip, we read components)
//   Row 24: Branch
//   Row 35: Dealer (no data)
//   Row 42: B2B & Affiliates   ← was incorrectly 46 (was reading GM row as NOPs)
//   Row 49: Agency              ← was incorrectly 52 (same off-by-block error)
const PH_ACT_ROW_MAP = {
  'B2C':              { startRow: 1  },
  'Renewal':          { startRow: 7  },
  'Branch':           { startRow: 24 },
  'B2B & Affiliates': { startRow: 41 }, // FIXED: was 46 (was off by one channel block)                                   
  'Agency':           { startRow: 47 }, // FIXED: was 52 (same root cause)                                                
  'UC Attached':      { startRow: 35 }
};

// INDO channel → 0-based START ROW of its block in the Insurance Actuals tab.
// VERIFIED against live Indo Actuals FY27 tab (readSpreadsheet, UNFORMATTED_VALUE):
//   Row 1:  Tele-Sales
//   Row 7:  Renewal (no data, all blank — not mapped, not a dashboard channel)
//   Row 13: B2C (= Tele-Sales, duplicate block — not mapped)
//   Row 24: Dealer   ← FIXED: was incorrectly flagged noData:true; block exists
//                        and is readable (values are legitimately 0 for months
//                        with no Dealer activity, not structurally missing).
//   Row 30: B2B (dashboard key 'B2B/SAU')
//   Row 36: Agency
const INDO_ACT_ROW_MAP = {
  'Tele-Sales': { startRow: 1  },
  'Agency':     { startRow: 36 },
  'B2B/SAU':    { startRow: 30 },
  'Dealer':     { startRow: 24 }
};

// Row offset of each metric WITHIN a channel block (note: Rev% sits at +3, GM%
// at +5, so gm is +4 — those % rows are skipped/recomputed downstream).
const METRIC_OFFSETS = { nops: 0, np: 1, rev: 2, gm: 4 };


// ── FRONT DOOR ────────────────────────────────────────────────
function doGet(e) {
  try {
    const action  = (e.parameter.action  || '').trim();
    const country = (e.parameter.country || '').trim().toUpperCase();
    const month   = (e.parameter.month   || '').trim();

    if (!action) {
      return HtmlService.createHtmlOutputFromFile('index')
        .setTitle('SEA Insurance MIS Dashboard')
        .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
    }

    switch (action) {

      case 'ping':
        return respond({ ok: true, ts: new Date().toISOString() });

      case 'getActuals': {
        const startDate  = (e.parameter.startDate  || '').trim();
        const endDate    = (e.parameter.endDate    || '').trim();
        const forceFresh = e.parameter.forceFresh === 'true';
        if (!country || !startDate || !endDate) return respondError('country, startDate and endDate are required');
        if (!['PH','INDO'].includes(country))   return respondError('country must be PH or INDO');
        return respond(getActuals(country, startDate, endDate, forceFresh));
      }

      case 'getActualsFromFile': {
        const fileMon    = (e.parameter.month      || '').trim();
        const fy         = (e.parameter.fy         || 'FY27').trim();
        const useQtrCol  = e.parameter.useQtrCol   === 'true';
        const qtr        = (e.parameter.qtr        || '').trim();
        if (!country || (!fileMon && !useQtrCol)) return respondError('country and month (or useQtrCol+qtr) required');
        if (!['PH','INDO'].includes(country))     return respondError('country must be PH or INDO');
        return respond(getActualsFromFile(country, fileMon, fy, useQtrCol, qtr));
      }

      case 'getAOP': {
        if (!country || !month) return respondError('country and month are required');
        if (!['PH','INDO'].includes(country)) return respondError('country must be PH or INDO');
        return respond(getAOP(country, month));
      }

      case 'createFlash': {
        const startDate = (e.parameter.startDate || '').trim();
        const endDate   = (e.parameter.endDate   || '').trim();
        if (!country || !startDate || !endDate) return respondError('country, startDate and endDate are required');
        if (!['PH','INDO'].includes(country))   return respondError('country must be PH or INDO');
        return respond(createFlash(country, startDate, endDate));
      }

      default:
        return respondError('Unknown action: ' + action);
    }

  } catch (err) {
    return respondError(err.message);
  }
}


// ── RESPONSE HELPERS ─────────────────────────────────────────
function respond(data) {
  return ContentService
    .createTextOutput(JSON.stringify(data))
    .setMimeType(ContentService.MimeType.JSON);
}
// Wrap an error message in the standard JSON envelope { success:false, error }.
function respondError(message) {
  return respond({ success: false, error: message });
}


// ============================================================
//  getActuals(country, startDate, endDate, forceFresh)
// ============================================================

function getActuals(country, startDate, endDate, forceFresh) {
  const cacheKey = 'actuals_' + country + '_' + startDate + '_' + endDate;

  if (!forceFresh) {
    const cached = cacheGet(cacheKey);
    if (cached) return Object.assign(cached, { fromCache: true });
  }

  const fileId    = country === 'PH' ? FILE_IDS.PH_NAV : FILE_IDS.INDO_NAV;
  const sheetName = country === 'PH' ? SHEET_NAMES.PH_SUMMARY : SHEET_NAMES.INDO_SUMMARY;

  const ss    = SpreadsheetApp.openById(fileId);
  const sheet = ss.getSheetByName(sheetName);
  if (!sheet) throw new Error(sheetName + ' not found in ' + country + ' workbook');

  // The summary sheet computes its channel rows with SUMIFS that reference the
  // date range in cells B1 (start) and B2 (end). So to get actuals for a range
  // we WRITE the dates into B1/B2, flush() to force recalculation, then read.
  //  ⚠ Because B1/B2 are shared, concurrent getActuals() calls would clobber
  //    each other — callers must run these SEQUENTIALLY (see loadCountry).
  sheet.getRange('B1').setValue(new Date(startDate + 'T00:00:00'));
  sheet.getRange('B2').setValue(new Date(endDate   + 'T00:00:00'));
  SpreadsheetApp.flush();

  // Channel rows live at a fixed block: PH = rows 6–13 (8 channels),
  // INDO = rows 7–12 (6 channels); columns A–E = [Channel, NOPs, NP, Rev, GM].
  const dataRange = country === 'PH'
    ? sheet.getRange(6, 1, 8, 5)   // (startRow=6, startCol=1, 8 rows, 5 cols)
    : sheet.getRange(7, 1, 6, 5);

  const rows = dataRange.getValues();
  const channels = {};
  rows.forEach(function(row) {
    const label = String(row[0]).trim();   // col A = channel name (key)
    if (!label) return;                     // skip blank separator rows  
    channels[label] = {
      nops: parseNum(row[1]),  // col B
      np:   parseNum(row[2]),  // col C
      rev:  parseNum(row[3]),  // col D
      gm:   parseNum(row[4])   // col E
    };
  });

  // B2C Subtotal (B2C + Renewal) has no row of its own in the live
  // _PH_Summary sheet's fixed 8-row block, so it can't come from the
  // SUMIFS read above — derive it here in code instead. B2B2C Subtotal
  // and Philippines Insurance already exist as real rows in that block
  // and are left as-is (untouched) to avoid diverging from the sheet's
  // own formulas.
  if (country === 'PH' && channels['B2C'] && channels['Renewal']) {
    channels['B2C Subtotal'] = {
      nops: (channels['B2C'].nops || 0) + (channels['Renewal'].nops || 0),
      np:   (channels['B2C'].np   || 0) + (channels['Renewal'].np   || 0),
      rev:  (channels['B2C'].rev  || 0) + (channels['Renewal'].rev  || 0),
      gm:   (channels['B2C'].gm   || 0) + (channels['Renewal'].gm   || 0)
    };
  }

  // INDO's B2B2C Subtotal row DOES exist in the live _INDO_Summary sheet's
  // fixed 6-row block, but that row's own formula only sums Agency+Dealer
  // (B2B/SAU excluded) — the same definition addSubtotals_INDO() used to
  // have before it was corrected to Agency+Dealer+B2B/SAU. Since the sheet
  // formula can't be edited from here, override the sheet-derived value
  // with the correct 3-channel sum so the "Actuals" column matches every
  // other column (AOP/QTD/prior months/Prev Qtr, which all go through
  // addSubtotals_INDO()).
  if (country === 'INDO' && channels['Agency'] && channels['Dealer'] && channels['B2B/SAU']) {
    const M = ['nops','np','rev','gm'];
    const b2b2cChannels = ['Agency','Dealer','B2B/SAU'];
    channels['B2B2C Subtotal'] = {};
    M.forEach(function(m) {
      channels['B2B2C Subtotal'][m] = b2b2cChannels.reduce(function(s,c){
        return s + (channels[c] && channels[c][m] ? channels[c][m] : 0);
      }, 0);
    });
  }

  const result = {
    success:   true,
    country:   country,
    startDate: startDate,
    endDate:   endDate,
    fetchedAt: new Date().toISOString(),
    fromCache: false,
    channels:  channels        // ← getC-ready { channels: {...} }
  };

  cacheSet(cacheKey, result);
  return result;
}


// ============================================================
//  getActualsFromFile(country, month, fy, useQtrCol, qtr)
//  Reads REPORTED actuals from the Insurance Actuals file.
//    • useQtrCol=true  → read the quarter aggregate column (qtr) → Prev-Qtr col
//    • useQtrCol=false → read a single month column (month)      → closed months
//  Returns { channels: { '<Channel>': {nops,np,rev,gm} } }.
// ============================================================

function getActualsFromFile(country, month, fy, useQtrCol, qtr) {

  // Cache key distinguishes month-mode vs quarter-mode and the fiscal year tab.
  const colTag   = useQtrCol ? qtr : month;
  const cacheKey = 'actfile_' + country + '_' + fy + '_' + colTag;

  const cached = cacheGet(cacheKey);
  if (cached) return Object.assign(cached, { fromCache: true });

  // Pick the right FY tab. Prev-Qtr for Q1 needs the PRIOR fiscal year (FY26).
  let sheetName;
  if (country === 'PH') {
    sheetName = fy === 'FY26' ? SHEET_NAMES.PH_ACT_FY26 : SHEET_NAMES.PH_ACT_FY27;
  } else {
    sheetName = fy === 'FY26' ? SHEET_NAMES.INDO_ACT_FY26 : SHEET_NAMES.INDO_ACT_FY27;
  }

  //Opening the right sheet tab 
  const ss    = SpreadsheetApp.openById(FILE_IDS.ACTUALS);
  const sheet = ss.getSheetByName(sheetName);
  if (!sheet) throw new Error(sheetName + ' not found in Insurance Actuals file');

  //Parsing values from the right sheet tab 
  const data = sheet.getDataRange().getValues();   // full grid (0-based [row][col])

  let col;
  if (useQtrCol) {
    col = ACTUALS_QTR_COL[qtr];
    if (col === undefined) throw new Error('Unknown quarter: ' + qtr);
  } else {
    col = ACTUALS_MONTH_COL[month];
    if (col === undefined) throw new Error('Unknown month: ' + month);
  }
 //mapping the channel rows
  const rowMap   = country === 'PH' ? PH_ACT_ROW_MAP : INDO_ACT_ROW_MAP;
  const channels = {};

  Object.keys(rowMap).forEach(function(label) {
    const entry = rowMap[label];
    if (entry.noData) {
      channels[label] = { nops: null, np: null, rev: null, gm: null };
      return;
    }

  // mapping the channel metrics in a matrix  
    const r = entry.startRow;
    channels[label] = {
      nops: parseNum(data[r + METRIC_OFFSETS.nops][col]),
      np:   parseNum(data[r + METRIC_OFFSETS.np  ][col]),
      rev:  parseNum(data[r + METRIC_OFFSETS.rev ][col]),
      gm:   parseNum(data[r + METRIC_OFFSETS.gm  ][col])
    };
  });

  if (country === 'PH') {
    addSubtotals_PH(channels);
  } else {
    addSubtotals_INDO(channels);
  }

  const result = {
    success:   true,
    country:   country,
    fy:        fy,
    col:       useQtrCol ? qtr : month,
    fetchedAt: new Date().toISOString(),
    fromCache: false,
    channels:  channels
  };

  cacheSet(cacheKey, result);
  return result;
}


// ── HELPER: safe number parse ────────────────────────────────
function parseNum(val) {
  if (val === null || val === undefined || val === '' || val === '-' || val === '  -   ') return 0;
  if (typeof val === 'number') return Math.round(val * 1000) / 1000;
  const cleaned = String(val).replace(/,/g, '').trim();
  const n = parseFloat(cleaned);
  return isNaN(n) ? 0 : Math.round(n * 1000) / 1000;
}


// ============================================================
//  getAOP(country, month)
//  Reads AOP targets for ONE month from the AOP file and returns BOTH:
//    • channels = that month's AOP            (per-month column)
//    • qtd      = the whole quarter's AOP     (per-quarter column)
//
//  ⚠ SHAPE ASYMMETRY (caused a real bug): `channels` is a flat channel map and
//    so is `qtd` (qtd additionally carries a `quarter` field). NEITHER is
//    wrapped in { channels: ... }. The dashboard wraps BOTH before handing them
//    to getC(): aopRange = { channels: aopData.channels }, and
//    aopQTD = { channels: aopData.qtd }. Forgetting to wrap qtd made the whole
//    QTD-AOP column render as dashes.
// ============================================================

function getAOP(country, month) {
  const cacheKey = 'aop_' + country + '_' + month;
  const cached   = cacheGet(cacheKey);
  if (cached) return Object.assign(cached, { fromCache: true });

  const ss        = SpreadsheetApp.openById(FILE_IDS.AOP);
  const sheetName = country === 'PH' ? SHEET_NAMES.PH_AOP : SHEET_NAMES.INDO_AOP;
  const sheet     = ss.getSheetByName(sheetName);
  if (!sheet) throw new Error(sheetName + ' not found in AOP file');

  const data = sheet.getDataRange().getValues();

  // Which COLUMN to read: one for the month, one for the quarter aggregate.
  const quarter  = monthToQuarter(month);
  const monthCol = country === 'PH' ? PH_MONTHLY_COL(month)   : INDO_MONTHLY_COL(month);
  const qtdCol   = country === 'PH' ? PH_QTD_COL(quarter)     : INDO_QTD_COL(quarter);

  const channels = {};                  // → monthly AOP, flat channel map
  const qtd      = { quarter: quarter };// → quarter AOP, flat channel map (+ quarter tag)

  // extract* fills BOTH `channels` (from monthCol) and `qtd` (from qtdCol),
  // reading each channel's NOPs/NP/Rev/GM from its fixed AOP rows.
  if (country === 'PH') {
    extractPH(data, monthCol, qtdCol, channels, qtd);
  } else {
    extractINDO(data, monthCol, qtdCol, channels, qtd);
  }

  const result = {
    success:   true,
    country:   country,
    month:     month,
    fetchedAt: new Date().toISOString(),
    fromCache: false,
    channels:  channels,   // monthly  (dashboard wraps as aopRange)
    qtd:       qtd          // quarter  (dashboard MUST wrap as { channels: qtd })
  };

  cacheSet(cacheKey, result);
  return result;
}


// ── PH AOP EXTRACTION ────────────────────────────────────────
// AOP sheet values are stored in USD THOUSANDS (e.g. a cell showing 500
// means $500,000) — same convention as the Insurance Actuals file. np/rev/gm
// are scaled ×1000 to raw USD to match the live NAV actuals unit; nops is a
// count and is never scaled.
function extractPH(data, mCol, qCol, channels, qtd) {
  const map = [
    { label: 'B2C',              rows: [1,  2,  3,  5]  },
    { label: 'Renewal',          rows: [7,  8,  9,  11] },
    { label: 'Branch',           rows: [24, 25, 26, 28] },
    { label: 'B2B & Affiliates', rows: [47, 48, 49, 51] },
    { label: 'Agency',           rows: [53, 54, 55, 57] }
  ];
  //rows maps the channel metrics columns maps the month
  map.forEach(function(ch) {
    channels[ch.label] = {
      nops: parseNum(data[ch.rows[0]][mCol]),
      np:   parseNum(data[ch.rows[1]][mCol]) * 1000,
      rev:  parseNum(data[ch.rows[2]][mCol]) * 1000,
      gm:   parseNum(data[ch.rows[3]][mCol]) * 1000
    };
    qtd[ch.label] = {
      nops: parseNum(data[ch.rows[0]][qCol]),
      np:   parseNum(data[ch.rows[1]][qCol]) * 1000,
      rev:  parseNum(data[ch.rows[2]][qCol]) * 1000,
      gm:   parseNum(data[ch.rows[3]][qCol]) * 1000
    };
  });
  channels['UC Attached'] = { nops: null, np: null, rev: null, gm: null };
  qtd['UC Attached']      = { nops: null, np: null, rev: null, gm: null };
  addSubtotals_PH(channels);
  addSubtotals_PH(qtd);
}


// ── INDO AOP EXTRACTION ──────────────────────────────────────
// Row maps VERIFIED against live INDO - AOP27 tab (0-based, including header row 0):
//   Row 0:  Header
//   Row 1:  Dealer  NOPs   → Dealer has no AOP target (always dash), kept as noData
//   Row 2:  Dealer  Premium
//   ...
//   Row 7:  B2B     NOPs   ← B2B/SAU = rows [7,8,9,11]
//   Row 8:  B2B     Premium
//   Row 9:  B2B     Revenue
//   Row 10: B2B     Rev%   (skip)
//   Row 11: B2B     GM
//   Row 12: B2B     GM%    (skip)
//   Row 13: blank
//   Row 14: Agency  NOPs   ← rows [14,15,16,18]
//   Row 15: Agency  Premium
//   Row 16: Agency  Revenue
//   Row 17: Agency  Rev%   (skip)
//   Row 18: Agency  GM
//   Tele-Sales does NOT appear in this tab's AOP structure — it's mapped from B2C.
//   The INDO dashboard's "Tele-Sales" AOP comes from the B2C block (rows [1,2,3,5] = Dealer).
//   NOTE: From live data, the correct Tele-Sales AOP is in the Overall/B2C section.
//         For now mapped to the first channel block (rows 1-6 = Dealer in the file,
//         but Dealer has no AOP so Tele-Sales AOP is null too per current spec).
// AOP sheet values are stored in USD THOUSANDS (e.g. a cell showing 500
// means $500,000) — same convention as the Insurance Actuals file. np/rev/gm
// are scaled ×1000 to raw USD to match the live NAV actuals unit; nops is a
// count and is never scaled.
function extractINDO(data, mCol, qCol, channels, qtd) {
  const map = [
    { label: 'Tele-Sales', rows: [2,3,4,6]  },
    // B2B/SAU: rows [7,8,9,11] verified from live data (NOPs=80 Q1, matching AOP file)
    { label: 'B2B/SAU',   rows: [19,20,21,23] },
    // Agency: rows [14,15,16,18] verified (NOPs=2677 Q1, matching live AOP)
    { label: 'Agency',    rows: [25,26,27,29] }
  ];
  map.forEach(function(ch) {
    channels[ch.label] = {
      nops: parseNum(data[ch.rows[0]][mCol]),
      np:   parseNum(data[ch.rows[1]][mCol]) * 1000,
      rev:  parseNum(data[ch.rows[2]][mCol]) * 1000,
      gm:   parseNum(data[ch.rows[3]][mCol]) * 1000
    };
    qtd[ch.label] = {
      nops: parseNum(data[ch.rows[0]][qCol]),
      np:   parseNum(data[ch.rows[1]][qCol]) * 1000,
      rev:  parseNum(data[ch.rows[2]][qCol]) * 1000,
      gm:   parseNum(data[ch.rows[3]][qCol]) * 1000
    };
  });
  // Tele-Sales and Dealer have no AOP in this file
  // channels['Tele-Sales'] = { nops: null, np: null, rev: null, gm: null };
  channels['Dealer']     = { nops: null, np: null, rev: null, gm: null };
  // qtd['Tele-Sales']      = { nops: null, np: null, rev: null, gm: null };
  qtd['Dealer']          = { nops: null, np: null, rev: null, gm: null };
  addSubtotals_INDO(channels);
  addSubtotals_INDO(qtd);
}


// ── SUBTOTAL HELPERS ─────────────────────────────────────────
function addSubtotals_PH(obj) {
  const METRICS        = ['nops','np','rev','gm'];
  const b2cChannels    = ['B2C','Renewal'];
  const b2b2cChannels  = ['Branch','UC Attached','B2B & Affiliates','Agency'];
  const totalChannels  = ['B2C','Renewal','Branch','UC Attached','B2B & Affiliates','Agency'];
  obj['B2C Subtotal']          = {};
  obj['B2B2C Subtotal']        = {};
  obj['Philippines Insurance'] = {};
  METRICS.forEach(function(m) {
    obj['B2C Subtotal'][m]          = b2cChannels.reduce(function(s,c){ return s + (obj[c] && obj[c][m] ? obj[c][m] : 0); }, 0);
    obj['B2B2C Subtotal'][m]        = b2b2cChannels.reduce(function(s,c){ return s + (obj[c] && obj[c][m] ? obj[c][m] : 0); }, 0);
    obj['Philippines Insurance'][m] = totalChannels.reduce(function(s,c){ return s + (obj[c] && obj[c][m] ? obj[c][m] : 0); }, 0);
  });
}

// Add INDO roll-ups to an obj: B2B2C Subtotal = Agency+Dealer+B2B/SAU; INDO Total = all 4 channels.
function addSubtotals_INDO(obj) {
  const METRICS       = ['nops','np','rev','gm'];
  const b2b2cChannels = ['Agency','Dealer','B2B/SAU'];
  const totalChannels = ['Tele-Sales','Agency','Dealer','B2B/SAU'];
  obj['B2B2C Subtotal']       = {};
  obj['INDO Insurance Total'] = {};
  METRICS.forEach(function(m) {
    obj['B2B2C Subtotal'][m]       = b2b2cChannels.reduce(function(s,c){ return s + (obj[c] && obj[c][m] ? obj[c][m] : 0); }, 0);
    obj['INDO Insurance Total'][m] = totalChannels.reduce(function(s,c){ return s + (obj[c] && obj[c][m] ? obj[c][m] : 0); }, 0);
  });
}


// ── MONTH → QUARTER ──────────────────────────────────────────
function monthToQuarter(month) {
  const Q1 = ['Apr-26','May-26','Jun-26'];
  const Q2 = ['Jul-26','Aug-26','Sep-26'];
  const Q3 = ['Oct-26','Nov-26','Dec-26'];
  const Q4 = ['Jan-27','Feb-27','Mar-27'];
  if (Q1.includes(month)) return 'Q1';
  if (Q2.includes(month)) return 'Q2';
  if (Q3.includes(month)) return 'Q3';
  if (Q4.includes(month)) return 'Q4';
  throw new Error('Cannot determine quarter for month: ' + month);
}


// ── AOP COLUMN RESOLVERS ─────────────────────────────────────
function PH_MONTHLY_COL(month) {
  const map = {
    'Apr-26':2,'May-26':3,'Jun-26':4,'Jul-26':5,
    'Aug-26':6,'Sep-26':7,'Oct-26':8,'Nov-26':9,
    'Dec-26':10,'Jan-27':11,'Feb-27':12,'Mar-27':13
  };
  if (map[month] === undefined) throw new Error('PH: unknown month ' + month);
  return map[month];
}
// PH AOP sheet column index (0-based) for a quarter's QTD bucket.
function PH_QTD_COL(quarter) {
  const map = { 'Q1':16, 'Q2':17, 'Q3':18, 'Q4':19 };
  if (map[quarter] === undefined) throw new Error('PH: unknown quarter ' + quarter);
  return map[quarter];
}
// INDO AOP sheet column index (0-based) for a given month.
// VERIFIED against live INDO - AOP27 tab: header row has Channel(0), metric(1),
// then quarterly buckets AMJ/JAS/OND/JFM at cols 2-5, FY'27 at 6, blank,
// then monthly Apr-26 at col 7 through Mar-27 at col 18.
function INDO_MONTHLY_COL(month) {
  const map = {
    'Apr-26':8, 'May-26':9,  'Jun-26':10,  'Jul-26':11,
    'Aug-26':12,'Sep-26':13, 'Oct-26':14, 'Nov-26':15,
    'Dec-26':16,'Jan-27':17, 'Feb-27':18, 'Mar-27':19
  };
  if (map[month] === undefined) throw new Error('INDO: unknown month ' + month);
  return map[month];
}
// INDO AOP sheet column index (0-based) for a quarter's QTD bucket.
// VERIFIED: AMJ=2, JAS=3, OND=4, JFM=5 in the INDO - AOP27 tab.
function INDO_QTD_COL(quarter) {
  const map = { 'Q1':2, 'Q2':3, 'Q3':4, 'Q4':5 };
  if (map[quarter] === undefined) throw new Error('INDO: unknown quarter ' + quarter);
  return map[quarter];
}


// ============================================================
//  CACHE LAYER
// ============================================================

function cacheGet(key) {
  try {
    const hit = CacheService.getScriptCache().get(key);
    if (!hit) return null;
    return JSON.parse(hit);
  } catch(e) { return null; }
}

// Store a value in CacheService as JSON with the standard TTL.
function cacheSet(key, data) {
  try {
    CacheService.getScriptCache().put(key, JSON.stringify(data), CACHE_TTL);
  } catch(e) {}
}

// Remove a single cache key.
function cacheBust(key) {
  try { CacheService.getScriptCache().remove(key); } catch(e) {}
}

// Remove every cache key tied to a (country, range): actuals + AOP + all monthly/quarterly actuals-file keys.
function cacheBustAll(country, start, end) {
  cacheBust('actuals_' + country + '_' + start + '_' + end);
  cacheBust('aop_'     + country + '_' + isoToMonthStr_gs(end));
  const MONTHS = ['Apr-26','May-26','Jun-26','Jul-26','Aug-26','Sep-26',
                  'Oct-26','Nov-26','Dec-26','Jan-27','Feb-27','Mar-27'];
  const QTRS   = ['Q1','Q2','Q3','Q4'];
  MONTHS.forEach(function(m) {
    cacheBust('actfile_' + country + '_FY27_' + m);
    cacheBust('actfile_' + country + '_FY26_' + m);
  });
  QTRS.forEach(function(q) {
    cacheBust('actfile_' + country + '_FY27_' + q);
    cacheBust('actfile_' + country + '_FY26_' + q);
  });
}

// Server-side ISO date -> 'Mon-YY' (e.g. '2026-05-01' -> 'May-26').
function isoToMonthStr_gs(iso) {
  const NAMES = ['Jan','Feb','Mar','Apr','May','Jun','Jul','Aug','Sep','Oct','Nov','Dec'];
  const d = new Date(iso + 'T00:00:00');
  return NAMES[d.getMonth()] + '-' + String(d.getFullYear()).slice(-2);
}


// ============================================================
//  PING — health check
// ============================================================

function ping() {
  const checks = {};
  try {
    const s = SpreadsheetApp.openById(FILE_IDS.PH_NAV).getSheetByName(SHEET_NAMES.PH_SUMMARY);
    checks.PH = s ? 'ok' : 'sheet not found';
  } catch(e) { checks.PH = 'error: ' + e.message; }

  try {
    const s = SpreadsheetApp.openById(FILE_IDS.INDO_NAV).getSheetByName(SHEET_NAMES.INDO_SUMMARY);
    checks.INDO = s ? 'ok' : 'sheet not found';
  } catch(e) { checks.INDO = 'error: ' + e.message; }

  try {
    const f = SpreadsheetApp.openById(FILE_IDS.AOP);
    checks.AOP = f ? 'ok' : 'file not found';
  } catch(e) { checks.AOP = 'error: ' + e.message; }

  try {
    const f = SpreadsheetApp.openById(FILE_IDS.ACTUALS);
    checks.ACTUALS = f ? 'ok' : 'file not found';
  } catch(e) { checks.ACTUALS = 'error: ' + e.message; }

  return {
    ok:     Object.values(checks).every(function(v){ return v === 'ok'; }),
    ts:     new Date().toISOString(),
    checks: checks
  };
}


// Time-trigger target: pre-warm the cache for the current quarter's months so the first user load is instant.
function warmCache() {
  const CURRENT_MONTHS = ['Apr-26','May-26','Jun-26'];
  const COUNTRIES      = ['PH','INDO'];
  COUNTRIES.forEach(function(country) {
    CURRENT_MONTHS.forEach(function(month) {
      try {
        const key = 'actuals_' + country + '_' + month;
        if (!cacheGet(key)) getActuals(country, month);
      } catch(e) {}
    });
  });
}
function clearAllCacheNow() {
  cacheBustAll('PH',   '2026-04-01', '2026-05-31');
  cacheBustAll('INDO', '2026-04-01', '2026-05-31');
  Logger.log('✓ All cache cleared. Refresh the dashboard now.');
}

// ============================================================
// ============================================================
//  FLASH FILE GENERATION  (from scratch — no template)
//
//  The "Create Flash" button (Frozen state) calls createFlash()
//  with the frozen window. It CREATES A BRAND-NEW workbook named
//  INS <Market>_<Mon>_<YY> FY_<FY>, writes the detail tabs from the
//  live source sheets for the window, then builds the Summary tabs
//  (Summary local + Summary USD) from scratch, and returns the URL.
//
//  The two consolidators below are ported verbatim from the existing
//  container-bound flash scripts. The ONLY changes: they write into
//  the passed target workbook (not getActiveSpreadsheet), the date
//  filter is parameterised by the window, and menus/alerts are gone.
//
//  WHAT IS / ISN'T GENERATED:
//   • Detail tabs (INS_* for PH; Telesales/Agency/Dealer + SAU for INDO) — yes.
//   • Summary (local) + Summary (USD): channel->product rows, subtotals,
//     grand total, COGS = Rev - GM, Units, W/ and W/O cancellations — yes,
//     computed in code (robust to the grouped/repeated-header detail tabs).
//   • Management table AOP / AOP x80% / Recast / hardcoded history, and the
//     manual "Sharing Deal Data" passthrough rows — NOT generated; these are
//     manual planning inputs with no derivable source.
//
//  DEPLOYMENT: the web app's executing identity needs read access to every
//  source workbook below. Drive scope is required (DriveApp) for create/move.
// ============================================================
// ============================================================

const FLASH_CONFIG = {
  PH: {
    market:       'PH',
    fx:           60,
    b2b2cNavId:   '1NJB5s8mmjZuk_QiGjIJV-WbpBndAHBu9Ml0T7w3u74Y', // B2B2C NAV (B2B/Agency/UC/Ridesecure&Coverex)
    b2cChannelId: '1YC7z2msZSbzo2T4lIGJrDodB7--QgaigZL2VpjuWljQ'  // B2C Channel (B2C/Renewal)
  },
  INDO: {
    market:    'Indo',
    fx:        16200,
    indoNavId: '1itzybrZYAASOJf70_-ji6T1GhEZJX2TF_Pzh6TiI98c',    // INDO NAV (Indo_Raw)
    sauId:     '12WA8iG6NIoKkB5_65PKVmSyxAdBQ11DMu1oX_4pjoEU'     // B2B SAU (Sheet1)
  },
  // Optional Drive folder for generated flash files. '' = My Drive root.
  outputFolderId: ''
};

// Policy status values that mark a non-live INDO policy (lower-cased compare).
// Rows matching ANY of these are excluded from the W/O cancellations bucket.
const FLASH_INDO_EXCLUDE_STATUSES = { 'cancelled': true, 'rejected': true };


// ── Window → reporting month / quarter / FY (Apr–Mar fiscal year) ──
// PH anchors on the START month (the reporting month); extended end days
// that spill into the next month are backdated by the encoding team, so the
// reporting month is always the start month. INDO uses the full date range.
function flashWindow_(startIso, endIso) {
  const MON = ['Jan','Feb','Mar','Apr','May','Jun','Jul','Aug','Sep','Oct','Nov','Dec'];
  const s = new Date(startIso + 'T00:00:00');
  const e = new Date(endIso   + 'T00:00:00');
  const year   = s.getFullYear();
  const month0 = s.getMonth();
  const fyNum  = (month0 >= 3) ? year + 1 : year;   // Apr..Dec → next-year FY
  const Q = ['Q4 JFM','Q4 JFM','Q4 JFM','Q1 AMJ','Q1 AMJ','Q1 AMJ',
             'Q2 JAS','Q2 JAS','Q2 JAS','Q3 OND','Q3 OND','Q3 OND'][month0];
  return {
    startDate: s, endDate: e,
    year: year, month0: month0,
    monName: MON[month0], yy: String(year).slice(-2),
    fy: 'FY' + String(fyNum).slice(-2), fyNum: fyNum, quarter: Q
  };
}


// ── ORCHESTRATOR ─────────────────────────────────────────────
function createFlash(country, startIso, endIso) {
  try {
    const cfg = FLASH_CONFIG[country];
    if (!cfg) return { success:false, error:'No flash config for ' + country };

    const w    = flashWindow_(startIso, endIso);
    const name = 'INS ' + cfg.market + '_' + w.monName + '_' + w.yy + ' FY_' + String(w.fyNum).slice(-2);

    // 1. Create a BRAND-NEW workbook (no template)
    const targetSS = SpreadsheetApp.create(name);
    if (FLASH_CONFIG.outputFolderId) {
      try { DriveApp.getFileById(targetSS.getId()).moveTo(DriveApp.getFolderById(FLASH_CONFIG.outputFolderId)); } catch(e) {}
    }

    // 2. Write the detail tabs from the live sources for the window
    let counts, summary;
    if (country === 'PH') {
      counts  = flashConsolidatePH_(targetSS, cfg, w.startDate, w.endDate);
      summary = flashBuildSummaryPH_(targetSS, cfg, w);
    } else {
      const indo = flashConsolidateINDO_(targetSS, cfg, w.startDate, w.endDate);
      const sau  = flashConsolidateSAU_(targetSS, cfg, w.startDate, w.endDate);
      counts = Object.assign({}, indo, sau);
      summary = flashBuildSummaryINDO_(targetSS, cfg, w);
    }

    // 3. Drop the default empty "Sheet1" that create() adds
    const def = targetSS.getSheetByName('Sheet1');
    if (def && targetSS.getSheets().length > 1) targetSS.deleteSheet(def);

    SpreadsheetApp.flush();

    return {
      success: true,
      country: country,
      name:    name,
      url:     targetSS.getUrl(),
      fileId:  targetSS.getId(),
      window:  { start: startIso, end: endIso, month: w.monName + '-' + w.yy, quarter: w.quarter, fy: w.fy },
      counts:  counts,
      summary: summary,
      ts:      new Date().toISOString()
    };
  } catch (err) {
    return { success:false, error: err.message };
  }
}


// ============================================================
//  PH CONSOLIDATOR  (ported from importAllWorkbooks)
// ============================================================

const FLASH_PH_TARGETS = {
  b2c:    'INS_B2C',
  b2b:    'INS_B2B',
  agency: 'INS_Agency',
  uc:     'INS_UC',
  renew:  'INS_Renewals',
  xdeal:  'INS_XDeal_Branch'
};

const FLASH_PH_REQUIRED_COLUMNS = [
  'Channel','INS POLICY CODE','Product','Case Received Date','Policy Number',
  "Assured's Name",'Provider','Lead Source','Policy Status','Status Date/ Cancellation Date',
  'Vehicle Type','Car Brand','Car Year','Car Model','Plate Number','Effectivity Start Date',
  'FMV/Coverage','APA','BI/PD','Issue Rate','FMV Premium','BI','PD','AUTO PA','RSCC',
  'Net_Premium','DST','VAT','Other fee','FST','LGT','Road Side Assistance','Gross Premium',
  'Discount/CWT','Subagent Incentive (Payout)',"Dealer's Mark Up",'Premium Agreed w/ Client',
  'Commission/Net Rate',"OTO's Commission (Vat Inclusive)","OTO's Commission (Vat Exclusive)",
  'Commision Share % (Agent/Agency)','Commission Share','Withholding Tax (10%)',
  'MIS_Revenue','Gross_Margin',
  'Payment Status'
  // INS_* output tab column map (A=Workbook, B=Sheet, data from C):
  //   C=Channel  E=Product  G=Policy Number  K=Policy Status
  //   AB=Net_Premium  AT=MIS_Revenue  AU=Gross_Margin  AV=Payment Status
];

// Build the six PH detail tabs (INS_*) in targetSS for the reporting month, by reading + date-filtering
// the policy-level source workbooks and grouping each tab by Channel.
function flashConsolidatePH_(targetSS, cfg, startDate, endDate) {
  const WORKBOOKS = [
    { id: cfg.b2b2cNavId, sheets: ['B2B'],                  channel: 'b2b'    },
    { id: cfg.b2b2cNavId, sheets: ['Agency'],               channel: 'agency' },
    { id: cfg.b2b2cNavId, sheets: ['UC'],                   channel: 'uc'     },
    { id: cfg.b2b2cNavId, sheets: ['Ridesecure & Coverex'], channel: 'agency' }
  ];
  const NEW_WORKBOOK = { id: cfg.b2cChannelId, sheets: ['B2C','Renewal'], channel: 'b2c' };

  const sheetB2C    = flashGetOrCreateSheet_(targetSS, FLASH_PH_TARGETS.b2c);
  const sheetB2B    = flashGetOrCreateSheet_(targetSS, FLASH_PH_TARGETS.b2b);
  const sheetAgency = flashGetOrCreateSheet_(targetSS, FLASH_PH_TARGETS.agency);
  const sheetUC     = flashGetOrCreateSheet_(targetSS, FLASH_PH_TARGETS.uc);
  const sheetRenew  = flashGetOrCreateSheet_(targetSS, FLASH_PH_TARGETS.renew);
  const sheetXDeal  = flashGetOrCreateSheet_(targetSS, FLASH_PH_TARGETS.xdeal);

  sheetB2C.clear(); sheetB2B.clear(); sheetAgency.clear();
  sheetUC.clear();  sheetRenew.clear(); sheetXDeal.clear();

  const allRowsB2C = [], allRowsB2B = [], allRowsAgency = [],
        allRowsUC  = [], allRowsRenew = [], allRowsXDeal = [];

  WORKBOOKS.forEach(function(book) {
    let wb;
    try { wb = SpreadsheetApp.openById(book.id); } catch(e) { return; }
    const wbName = wb.getName();
    book.sheets.forEach(function(sheetName) {
      const rows = flashProcessSheetPH_(wb, wbName, sheetName, startDate, endDate);
      if      (book.channel === 'b2b')    rows.forEach(function(r){ allRowsB2B.push(r); });
      else if (book.channel === 'agency') rows.forEach(function(r){ allRowsAgency.push(r); });
      else if (book.channel === 'uc')     rows.forEach(function(r){ allRowsUC.push(r); });
      else                                rows.forEach(function(r){ allRowsB2C.push(r); });
    });
  });

  let newWb;
  try { newWb = SpreadsheetApp.openById(NEW_WORKBOOK.id); } catch(e) {}
  if (newWb) {
    const wbName  = newWb.getName();
    const chanIdx = FLASH_PH_REQUIRED_COLUMNS.indexOf('Channel') + 2; // +2 for Workbook + Sheet prefix
    NEW_WORKBOOK.sheets.forEach(function(sheetName) {
      const rows = flashProcessSheetPH_(newWb, wbName, sheetName, startDate, endDate);
      rows.forEach(function(r) {
        const ch         = (r[chanIdx] || '').toString().trim().toLowerCase();
        const sheetLower = sheetName.trim().toLowerCase();
        if (sheetLower.indexOf('renewal') === 0 || ch.indexOf('renewal') === 0) {
          allRowsRenew.push(r);
        } else if (ch.indexOf('xdeal') === 0 || ch.indexOf('branch') === 0) {
          allRowsXDeal.push(r);
        } else {
          allRowsB2C.push(r);
        }
      });
    });
  }

  flashWriteGroupedBy_(sheetB2C,    allRowsB2C);
  flashWriteGroupedBy_(sheetB2B,    allRowsB2B);
  flashWriteGroupedBy_(sheetAgency, allRowsAgency);
  flashWriteGroupedBy_(sheetUC,     allRowsUC);
  flashWriteGroupedBy_(sheetRenew,  allRowsRenew);
  flashWriteGroupedBy_(sheetXDeal,  allRowsXDeal);

  return {
    INS_B2C:          allRowsB2C.length,
    INS_B2B:          allRowsB2B.length,
    INS_Agency:       allRowsAgency.length,
    INS_UC:           allRowsUC.length,
    INS_Renewals:     allRowsRenew.length,
    INS_XDeal_Branch: allRowsXDeal.length
  };
}

// Read one source sheet; keep rows whose Case Received Date falls within
// [startDate, endDate] INCLUSIVE (full range comparison, not a calendar-month
// match) — so extended ranges spanning into the next month (e.g. May 1–Jun 3)
// correctly include every row in between, not just rows in the single
// calendar month of `startDate`. endDate is treated as inclusive through
// end-of-day (23:59:59.999) so a Case Received Date stamped exactly on the
// end date, at any time of day, is still included.
// Maps to the REQUIRED_COLUMNS order; prefixes [Workbook, Sheet]. Returns kept rows.
function flashProcessSheetPH_(wb, wbName, sheetName, startDate, endDate) {
  const sh = wb.getSheetByName(sheetName);
  if (!sh) return [];
  const lastRow = sh.getLastRow();
  if (lastRow < 2) return [];

  const headers = sh.getRange(1, 1, 1, sh.getLastColumn()).getValues()[0];
  const data    = sh.getRange(2, 1, lastRow - 1, sh.getLastColumn()).getValues();
  if (!data.length) return [];

  const normalize = function(s){ return s.toString().trim().toLowerCase(); };
  const headerMap = {};
  headers.forEach(function(h, i){ headerMap[normalize(h)] = i; });

  const colIndexes = FLASH_PH_REQUIRED_COLUMNS.map(function(col) {
    const idx = headerMap[normalize(col)];
    return idx !== undefined ? idx : null;
  });
  const caseIdx = FLASH_PH_REQUIRED_COLUMNS.indexOf('Case Received Date');
  const outputRows = [];

  // endDate is passed as midnight of the end-date's calendar day; extend it
  // to 23:59:59.999 so a Case Received Date stamped anywhere on that day is
  // included, not just rows exactly at midnight.
  const endDateEOD = new Date(endDate.getTime());
  endDateEOD.setHours(23, 59, 59, 999);

  data.forEach(function(row) {
    let mappedRow = colIndexes.map(function(idx){ return idx !== null ? row[idx] : ''; });

    const caseRaw = mappedRow[caseIdx];
    if (!caseRaw) return;
    let caseDate = (caseRaw instanceof Date) ? caseRaw : new Date(caseRaw);
    if (isNaN(caseDate.getTime())) return;

    // Full date-range comparison (not a calendar-month match) — so an
    // extended range like May 1–Jun 3 correctly includes every row in
    // between, not just rows falling in the single calendar month of start.
    if (caseDate < startDate || caseDate > endDateEOD) return;

    const tz = Session.getScriptTimeZone();
    mappedRow = mappedRow.map(function(cell) {
      if (cell instanceof Date) return Utilities.formatDate(cell, tz, 'MM/dd/yyyy');
      return cell;
    });
    outputRows.push([wbName, sheetName].concat(mappedRow));
  });

  return outputRows;
}

// Write rows into a detail tab grouped by Channel: a header row per channel block, its rows
// below, a blank line between blocks (batched writes).
function flashWriteGroupedBy_(targetSheet, allRows) {
  if (!allRows.length) return;
  const groupIdx  = FLASH_PH_REQUIRED_COLUMNS.indexOf('Channel') + 2;
  const headerRow = ['Workbook','Sheet'].concat(FLASH_PH_REQUIRED_COLUMNS);
  const totalCols = headerRow.length;
  const blankRow  = [Array(totalCols).fill('')];

  const groupOrder  = [];
  const groupBucket = {};
  allRows.forEach(function(row) {
    const key = (row[groupIdx] || '').toString().trim() || '(Unknown)';
    if (!groupBucket[key]) { groupBucket[key] = []; groupOrder.push(key); }
    groupBucket[key].push(row);
  });

  let currentRow = 1;
  const BATCH = 500;
  groupOrder.forEach(function(key, i) {
    if (i > 0) {
      targetSheet.getRange(currentRow, 1, 1, totalCols).setValues(blankRow);
      currentRow++;
    }
    targetSheet.getRange(currentRow, 1, 1, totalCols).setValues([headerRow]);
    currentRow++;
    const groupRows = groupBucket[key];
    for (let j = 0; j < groupRows.length; j += BATCH) {
      const batch = groupRows.slice(j, j + BATCH);
      targetSheet.getRange(currentRow, 1, batch.length, totalCols).setValues(batch);
      currentRow += batch.length;
    }
  });
}

// Return a sheet by name in ss, creating it if missing.
function flashGetOrCreateSheet_(ss, name) {
  let sh = ss.getSheetByName(name);
  if (!sh) sh = ss.insertSheet(name);
  return sh;
}


// ============================================================
//  INDO + SAU CONSOLIDATOR  (ported from consolidateIndoFlash / consolidateSAU)
// ============================================================

const FLASH_INDO_RAW_SHEET  = 'Indo_Raw';
const FLASH_COL_ORDER_DATE  = 1;   // col B
const FLASH_COL_BIZ_CHANNEL = 6;   // col G
const FLASH_COL_LAST_INDO   = 54;  // read A..(COL+1)
const FLASH_CHANNEL_TABS    = { 'Telesales':'Telesales', 'Agency':'Agency', 'Dealer':'Dealer' };

const FLASH_SAU_SHEET        = 'Sheet1';
const FLASH_SAU_USDIDR       = 16200;
const FLASH_SAU_COL_PROD_MTH = 1;
const FLASH_SAU_COL_POLICY   = 2;
const FLASH_SAU_COL_RNB      = 12;
const FLASH_SAU_COL_AGENT    = 10;
const FLASH_SAU_COL_GWP      = 28;
const FLASH_SAU_COL_REV      = 29;
const FLASH_SAU_COL_GP       = 30;
const FLASH_SAU_COL_PREM_IDR = 15;
const FLASH_SAU_COL_FX       = 27;
const FLASH_SAU_NUM_COLS     = 31;
const FLASH_SAU_OUTPUT_COLS  = [0,1,2,3,4,5,6,11,12,13,15,16,17,18,19,20,21,23,24,25,27,28,29,30];
const FLASH_SAU_TAB_NB       = 'B2B(SAU)-New Business';
const FLASH_SAU_TAB_REN      = 'B2B(SAU)-Renewal';

// Pick only the wanted column indices out of a full row (missing -> '').
function flashExtractCols_(row, colIndices) {
  const out = [];
  for (let i = 0; i < colIndices.length; i++) {
    out.push(row[colIndices[i]] !== undefined ? row[colIndices[i]] : '');
  }
  return out;
}

// Parse a SAU "Production Month" string to the 1st of that month (strips a leading "YYYY, "
// prefix); null if unparseable.
function flashParseProdMonth_(val) {
  if (!val) return null;
  let s = val.toString().trim();
  s = s.replace(/^\d{4},\s*/, '').trim();
  const MONTHS = { 'jan':0,'feb':1,'mar':2,'apr':3,'may':4,'jun':5,
                   'jul':6,'aug':7,'sep':8,'oct':9,'nov':10,'dec':11 };
  const parts = s.split(/\s+/);
  if (parts.length < 2) return null;
  const mon = parts[0].toLowerCase().substring(0, 3);
  const yr  = parseInt(parts[1], 10);
  if (MONTHS[mon] === undefined || isNaN(yr)) return null;
  return new Date(yr, MONTHS[mon], 1);
}

// Parse a messy numeric string (strips IDR/spaces/dots/commas; ()->negative); NaN if not numeric.
function flashParseNumeric_(val) {
  if (val === null || val === undefined || val === '' || val === '-') return NaN;
  const s = val.toString()
    .replace(/IDR/gi, '').replace(/\s/g, '')
    .replace(/\./g, '').replace(/,/g, '')
    .replace(/\(/g, '-').replace(/\)/g, '').trim();
  return parseFloat(s);
}

// Replace a tab's contents with [headerRow, ...dataRows] (creates the tab if missing).
function flashWriteToTab_(targetSS, tabName, headerRow, dataRows) {
  let sheet = targetSS.getSheetByName(tabName);
  if (!sheet) sheet = targetSS.insertSheet(tabName);
  else        sheet.clearContents();
  const output = [headerRow].concat(dataRows);
  if (output.length > 0) sheet.getRange(1, 1, output.length, headerRow.length).setValues(output);
}

// Build the INDO Telesales/Agency/Dealer detail tabs in targetSS by filtering Indo_Raw on Order
// Date within [startDate, endDate] and grouping by business channel.
function flashConsolidateINDO_(targetSS, cfg, startDate, endDate) {
  const indoSS   = SpreadsheetApp.openById(cfg.indoNavId);
  const rawSheet = indoSS.getSheetByName(FLASH_INDO_RAW_SHEET);
  if (!rawSheet) throw new Error("Sheet '" + FLASH_INDO_RAW_SHEET + "' not found.");

  const lastRow = rawSheet.getLastRow();
  if (lastRow < 2) return { Telesales:0, Agency:0, Dealer:0 };

  const headerRow = rawSheet.getRange(1, 1, 1, FLASH_COL_LAST_INDO + 1).getDisplayValues()[0];
  const rawData   = rawSheet.getRange(2, 1, lastRow - 1, FLASH_COL_LAST_INDO + 1).getDisplayValues();

  // endDate is passed as midnight of the end-date's calendar day; extend it
  // to 23:59:59.999 so an Order Date stamped anywhere on that day is
  // included, not just rows exactly at midnight.
  const endDateEOD = new Date(endDate.getTime());
  endDateEOD.setHours(23, 59, 59, 999);

  const grouped = { 'Telesales': [], 'Agency': [], 'Dealer': [] };
  for (let i = 0; i < rawData.length; i++) {
    const row = rawData[i];
    const dateStr = row[FLASH_COL_ORDER_DATE];
    if (!dateStr) continue;
    const d = new Date(dateStr);
    if (isNaN(d) || d < startDate || d > endDateEOD) continue;
    const ch = row[FLASH_COL_BIZ_CHANNEL];
    if (!ch || !grouped.hasOwnProperty(ch)) continue;
    grouped[ch].push(row);
  }

  for (const ch in FLASH_CHANNEL_TABS) {
    flashWriteToTab_(targetSS, FLASH_CHANNEL_TABS[ch], headerRow, grouped[ch]);
  }
  return { Telesales: grouped['Telesales'].length, Agency: grouped['Agency'].length, Dealer: grouped['Dealer'].length };
}

// Build the two B2B(SAU) detail tabs (New Business / Renewal) in targetSS, applying the SAU rules
// (CSI -> HJK policies only; renewal negative-filter + Gross Profit override) for the window month.
function flashConsolidateSAU_(targetSS, cfg, startDate, endDate) {
  const sauSS    = SpreadsheetApp.openById(cfg.sauId);
  const sauSheet = sauSS.getSheetByName(FLASH_SAU_SHEET);
  if (!sauSheet) throw new Error("Sheet '" + FLASH_SAU_SHEET + "' not found.");

  const lastRow = sauSheet.getLastRow();
  if (lastRow < 2) return { 'B2B(SAU)-New Business':0, 'B2B(SAU)-Renewal':0 };

  const headerRow = sauSheet.getRange(1, 1, 1, FLASH_SAU_NUM_COLS).getDisplayValues()[0];
  const rawData   = sauSheet.getRange(2, 1, lastRow - 1, FLASH_SAU_NUM_COLS).getDisplayValues();

  const newBizRows = [], renewalRows = [];
  const sY = startDate.getFullYear(), sM = startDate.getMonth();
  const eY = endDate.getFullYear(),   eM = endDate.getMonth();

  for (let i = 0; i < rawData.length; i++) {
    const row = rawData[i];

    let allBlank = true;
    for (let c = 0; c < row.length; c++) {
      if (row[c] !== '' && row[c] !== null && row[c] !== undefined) { allBlank = false; break; }
    }
    if (allBlank) continue;

    const prodDate = flashParseProdMonth_(row[FLASH_SAU_COL_PROD_MTH]);
    if (!prodDate) continue;
    const pY = prodDate.getFullYear(), pM = prodDate.getMonth();
    if (pY < sY || (pY === sY && pM < sM) || pY > eY || (pY === eY && pM > eM)) continue;

    const agent   = (row[FLASH_SAU_COL_AGENT]  || '').toString().trim();
    const rnb     = (row[FLASH_SAU_COL_RNB]    || '').toString().trim();
    const policy  = (row[FLASH_SAU_COL_POLICY] || '').toString().trim();
    const isCSI   = agent.toUpperCase() === 'CSI';
    const isRenew = rnb.toLowerCase().indexOf('renewal') !== -1;

    if (isCSI) {
      if (policy.toUpperCase().indexOf('HJK') !== 0) continue;
      (isRenew ? renewalRows : newBizRows).push(flashExtractCols_(row, FLASH_SAU_OUTPUT_COLS));
      continue;
    }

    const gwp = flashParseNumeric_(row[FLASH_SAU_COL_GWP]);
    const rev = flashParseNumeric_(row[FLASH_SAU_COL_REV]);
    const gp  = flashParseNumeric_(row[FLASH_SAU_COL_GP]);

    if (!isRenew) {
      newBizRows.push(flashExtractCols_(row, FLASH_SAU_OUTPUT_COLS));
    } else {
      if ((!isNaN(gwp) && gwp < 0) || (!isNaN(rev) && rev < 0) || (!isNaN(gp) && gp < 0)) continue;
      const prem = flashParseNumeric_(row[FLASH_SAU_COL_PREM_IDR]);
      let fx     = flashParseNumeric_(row[FLASH_SAU_COL_FX]);
      if (isNaN(fx) || fx <= 0) fx = FLASH_SAU_USDIDR;
      const overRow = row.slice();
      if (!isNaN(prem) && fx > 0) {
        overRow[FLASH_SAU_COL_GP] = String(Math.round((prem / fx) * 0.01 * 100) / 100);
      }
      renewalRows.push(flashExtractCols_(overRow, FLASH_SAU_OUTPUT_COLS));
    }
  }

  const sauOutHeader = flashExtractCols_(headerRow, FLASH_SAU_OUTPUT_COLS);
  flashWriteToTab_(targetSS, FLASH_SAU_TAB_NB,  sauOutHeader, newBizRows);
  flashWriteToTab_(targetSS, FLASH_SAU_TAB_REN, sauOutHeader, renewalRows);
  return { 'B2B(SAU)-New Business': newBizRows.length, 'B2B(SAU)-Renewal': renewalRows.length };
}


// ============================================================
//  SUMMARY BUILDERS
//
//  PH  → flashBuildSummaryPH_()
//          Writes live SUMIFS/COUNTIFS formulas into Summary (PHP).
//          Summary USD = formula mirror of PHP tab, every monetary cell ÷ FX.
//
//  INDO → flashBuildSummaryINDO_()
//          Aggregates in code (header-based column lookup, same as PH).
//          Uses flashBlocksFromAgg_() + flashWriteSummaryTab_() to render
//          Summary (IDR) and Summary USD.
//
//  COGS is always Revenue − Gross Margin (never a raw source column).
// ============================================================

// Coerce a cell to a number: '' / null → 0; strip thousands separators and
// treat (parentheses) as negative (accounting style).
function flashCellNum_(v) {
  if (v === '' || v === null || v === undefined) return 0;
  if (typeof v === 'number') return v;
  const n = parseFloat(String(v).replace(/[, ]/g, '').replace(/\(/g, '-').replace(/\)/g, ''));
  return isNaN(n) ? 0 : n;
}
// A fresh accumulator (one per W and per W/O bucket).
function flashZero_() { return { nops:0, np:0, rev:0, gm:0 }; }
// Add one policy's metrics into an accumulator (n = unit count, usually 1).
function flashAcc_(o, np, rev, gm, n) { o.np += np; o.rev += rev; o.gm += gm; o.nops += n; }

// Generic writer: blocks = [{ name, rows:[{label,w,wo}], subW, subWO }], plus a grand total.
// Writes one tab; monetary values ÷ divisor; COGS = Rev − GM. W/O cancellations
// ONLY — the W/ cancellations block has been removed per spec (each row's `w`
// accumulator is still passed in by the caller but is no longer written out).
// Layout per row: label | Net Premium | Revenue | Units | COGS | Gross Margin.
function flashWriteSummaryTab_(ss, tabName, blocks, grand, divisor, ccyLabel, w, fxText) {
  let sh = ss.getSheetByName(tabName);
  if (!sh) sh = ss.insertSheet(tabName);
  else     sh.clearContents();

  const m = function(x){ return Math.round((x / divisor) * 100) / 100; };
  const line = function(label, a) {
    const cogsWO = a.wo.rev - a.wo.gm;
    return [ label, m(a.wo.np), m(a.wo.rev), Math.round(a.wo.nops), m(cogsWO), m(a.wo.gm) ];
  };

  const out = [];
  out.push([ 'INS ' + (ccyLabel === 'USD' ? 'Summary (USD)' : 'Summary (Local ' + ccyLabel + ')') +
             ' — ' + w.monName + '-' + w.yy + ' ' + w.fy ]);
  out.push([ 'Window: ' + w.quarter, '', 'FX: ' + fxText ]);
  out.push([]);
  out.push([ 'Channel / Product','Net Premium','Revenue','Units','COGS','Gross Margin' ]);

  blocks.forEach(function(b) {
    out.push([ b.name ]);
    b.rows.forEach(function(r) { out.push(line('  ' + r.label, r)); });
    out.push(line('  Total ' + b.name, { w:b.subW, wo:b.subWO }));
    out.push([]);
  });
  out.push(line('GRAND TOTAL', { w:grand.w, wo:grand.wo }));

  const width = 6;
  const padded = out.map(function(r){ while (r.length < width) r.push(''); return r; });
  sh.getRange(1, 1, padded.length, width).setValues(padded);
  return out.length;
}


// ── PH Summary ───────────────────────────────────────────────
//  Builds Summary (PHP) with live SUMIFS/COUNTIFS formulas pointing at the
//  INS_* detail tabs, matching the reference Summary (PHP) tab layout.
//
//  INS_* column schema (A=Workbook, B=Sheet, data from C):
//    C=Channel   E=Product   G=Policy Number (unit count)
//    K=Policy Status
//    AB=Net_Premium   AT=MIS_Revenue   AU=Gross_Margin   AV=Payment Status
//
//  W/ cancellations  = all statuses (no status filter).
//  W/O cancellations = Payment Status ≠ "Cancelled" (applied to all tabs).
//  Summary USD       = formula mirror of Summary (PHP), every monetary cell ÷ FX.
//
//  PH_LAYOUT defines the locked row order matching the reference sheet.
//  Each entry is one of:
//    { tab, ch, pr, label }          — data row: SUMIFS/COUNTIFS into that tab
//    { isSubtotal, label, members }  — SUM of named member rows
//    { isPassthrough, label }        — blank cells (manual entry)
//    { isGrandTotal, label }         — SUM of all subtotal rows
//
//  ch = Channel criterion (null = omit, match all rows in tab)
//  pr = Product criterion (null = omit, match all products)
function flashBuildSummaryPH_(ss, cfg, w) {

  const PH_LAYOUT = [
    // ── INS_B2C ────────────────────────────────────────────────────────────────
    // Channel values in source: 'b2c-MOT', 'b2c-Others', 'b2c-CTPL'
    // W/O filter: Payment Status <> Cancelled (AV col)
    { tab:'INS_B2C',          ch:'b2c-MOT',         pr:null, gmPr:null,            label:'b2c-MOT'         },
    { tab:'INS_B2C',          ch:'b2c-Others',       pr:null, gmPr:null,            label:'b2c-Others'       },
    { tab:'INS_B2C',          ch:'b2c-CTPL',         pr:null, gmPr:null,            label:'b2c-CTPL'         },
    { isSubtotal:true, label:'Total Ins_Open_B2C',
      members:['b2c-MOT','b2c-Others','b2c-CTPL'] },

    // ── xDeal (INS_XDeal_Branch) ───────────────────────────────────────────────
    // Channel values in source: 'xDeal-MOT', 'xDeal-CTPL', 'xDeal-Others'
    // W/O filter: Payment Status <> Cancelled (AV col)
    { tab:'INS_XDeal_Branch', ch:'xDeal-MOT',        pr:null, gmPr:null,            label:'xDeal-MOT'       },
    { tab:'INS_XDeal_Branch', ch:'xDeal-CTPL',       pr:null, gmPr:null,            label:'xDeal-CTPL'      },
    { tab:'INS_XDeal_Branch', ch:'xDeal-Others',     pr:null, gmPr:null,            label:'xDeal-Others'    },

    // ── Branch (INS_XDeal_Branch) ───────────────────────────────────────────────
    // Channel values in source: 'Branch-LIPA-MOT', 'Branch-MM-MOT', 'Branch-NAGA-MOT', 'Branch-Others'
    // W/O filter: Payment Status <> Cancelled (AV col)
    { tab:'INS_XDeal_Branch', ch:'Branch-LIPA-MOT',  pr:null, gmPr:null,            label:'Branch-LIPA-MOT' },
    { tab:'INS_XDeal_Branch', ch:'Branch-MM-MOT',    pr:null, gmPr:null,            label:'Branch-MM-MOT'   },
    { tab:'INS_XDeal_Branch', ch:'Branch-NAGA-MOT',  pr:null, gmPr:null,            label:'Branch-NAGA-MOT' },
    { tab:'INS_XDeal_Branch', ch:'Branch-Others',    pr:null, gmPr:null,            label:'Branch-Others'   },
    { isSubtotal:true, label:'Total Ins_Open_Branch',
      members:['xDeal-MOT','xDeal-CTPL','xDeal-Others',
               'Branch-LIPA-MOT','Branch-MM-MOT','Branch-NAGA-MOT','Branch-Others'] },

    // ── Affiliates (INS_B2B) ─────────────────────────────────────────────────────
    // W/O filter: Payment Status <> Cancelled (AV col); gmPr kept for label context only
    { tab:'INS_B2B',          ch:'B2B-MOT',          pr:null, gmPr:'Comprehensive', label:'Comprehensive'   },
    { tab:'INS_B2B',          ch:'B2B-Others',       pr:null, gmPr:null,            label:'B2B-Others'      },
    { tab:'INS_B2B',          ch:'B2B-CTPL',         pr:null, gmPr:'CTPL',          label:'CTPL'            },
    { isPassthrough:true, label:'Sharing Deal Data (MOTOR)'                                                  },
    { isPassthrough:true, label:'Sharing Deal Data (NON MOTOR)'                                              },
    { isSubtotal:true, label:'Total Ins_Open_Affiliates',
      members:['Comprehensive','B2B-Others','CTPL',
               'Sharing Deal Data (MOTOR)','Sharing Deal Data (NON MOTOR)'] },

    // ── Renewals (INS_Renewals) ──────────────────────────────────────────────────
    // W/O filter: Payment Status <> Cancelled (AV col); gmPr kept for label context only
    { tab:'INS_Renewals',     ch:'Renewal-MOT',      pr:null, gmPr:'Comprehensive', label:'Comprehensive'   },
    { tab:'INS_Renewals',     ch:'Renewal-CTPL',     pr:null, gmPr:'CTPL',          label:'CTPL'            },
    { isSubtotal:true, label:'Total Ins_Renewals',
      members:['Comprehensive','CTPL'] },

    // ── Agency (INS_Agency) ────────────────────────────────────────────────────────
    // W/O filter: Payment Status <> Cancelled (AV col); gmPr kept for label context only
    // (gmPr field retained in layout for reference but no longer drives a Policy Status filter)
    { tab:'INS_Agency',       ch:'Agency-MOT',        pr:null, gmPr:'Comprehensive', label:'Comprehensive'   },
    { tab:'INS_Agency',       ch:'Agency-Others',     pr:null, gmPr:null,            label:'Agency-Others'   },
    { tab:'INS_Agency',       ch:'Agency-CTPL',       pr:null, gmPr:'CTPL',          label:'CTPL'            },
    { tab:'INS_Agency',       ch:'Agency-Coverex',    pr:null, gmPr:'Coverex',       label:'Coverex NAV'     },
    { isPassthrough:true, label:'Sharing Deal Data (MOTOR)'                                                  },
    { isPassthrough:true, label:'Sharing Deal Data (NON MOTOR)'                                              },
    { isSubtotal:true, label:'Total Ins_Agency',
      members:['Comprehensive','Agency-Others','CTPL','Coverex NAV',
               'Sharing Deal Data (MOTOR)','Sharing Deal Data (NON MOTOR)'] },

    // ── UC Attached (INS_UC) ─────────────────────────────────────────────────────
    // W/O filter: Payment Status <> Cancelled (AV col); gmPr kept for label context only
    { tab:'INS_UC',           ch:'UC-MOT',            pr:null, gmPr:'Comprehensive', label:'Comprehensive'   },
    { isSubtotal:true, label:'Total Ins_UC Attached',
      members:['Comprehensive'] },

    // ── Grand Total ────────────────────────────────────────────────────────────────
    { isGrandTotal:true, label:'GRAND TOTAL' }
  ];

  flashWriteSummaryPHFormulas_(ss, 'Summary (PHP)', PH_LAYOUT, w, 'PHP', 'local currency');
  flashWriteSummaryUSDMirror_(ss, 'Summary USD', 'Summary (PHP)', w, cfg.fx);
  return { note: 'PH Summary (PHP) built with SUMIFS formulas; Summary USD = PHP ÷ ' + cfg.fx + '.' };
}


// ── flashWriteSummaryPHFormulas_ ──────────────────────────────
//  Writes Summary (PHP) with live SUMIFS/COUNTIFS formulas. W/O
//  cancellations ONLY — the W/ cancellations block has been removed per
//  spec. 6 columns (A–F):
//
//    A=channel label (ch value when set, else rowDef.label)
//    B=Net Premium   C=Revenue   D=Units   E=COGS   F=Gross Margin
//
//  COGS (E) = Revenue − Gross Margin, always a formula.
//
//  Row label (col A): the channel string (ch) is used as the label so the
//  SUMIFS can reference $A{row} dynamically. Editing the label cell updates
//  all formulas in that row automatically.
//
//  Filter: Payment Status (col AV) ≠ "Cancelled" AND Policy Status (col K)
//  ≠ "Not Issued" — applied uniformly to ALL INS_* tabs. "Policy Issued"
//  as a required value on Policy Status is NOT used (only the "Not Issued"
//  exclusion applies there).
function flashWriteSummaryPHFormulas_(ss, tabName, layout, w, ccyLabel, fxText) {
  let sh = ss.getSheetByName(tabName);
  if (!sh) sh = ss.insertSheet(tabName);
  else     sh.clearContents();

  // ── INS_* detail-tab column letters ─────────────────────────────────────
  const CC  = 'C';   // Channel
  const CP  = 'E';   // Product
  const CU  = 'G';   // Policy Number  (non-blank = 1 unit in COUNTIFS)
  const CST = 'K';   // Policy Status  — filter: <>Not Issued
  const CPS = 'AV';  // Payment Status — filter: <>Cancelled
  const CNP = 'AB';  // Net_Premium
  const CRV = 'AT';  // MIS_Revenue
  const CGM = 'AU';  // Gross_Margin

  // ── Filter rule (applied uniformly to ALL tabs) ───────────────────────
  // Payment Status (col AV) <> "Cancelled".
  // Policy Status (col K) <> "Not Issued".
  // These are two separate columns/conditions, both applied as exclusions.
  // "Policy Issued" as a required value is NOT used.
  // This applies to every INS_* tab — B2C, XDeal_Branch, B2B, Renewals, Agency, UC.

  // ── Formula builders ─────────────────────────────────────────────────────

  // SUMIFS filtered by Payment Status <> Cancelled AND Policy Status <> Not Issued.
  // ch = channel criterion string or null; pr = product criterion string or null.
  // When ch is a cell reference like '$A5' it is passed WITHOUT surrounding quotes.
  function fSumifs(tab, col, ch, pr, chIsRef) {
    const sum  = "'" + tab + "'!" + col + ':' + col;
    const crit = [];
    if (ch) {
      const chVal = chIsRef ? ch : '"' + ch + '"';
      crit.push("'" + tab + "'!" + CC + ':' + CC + ',' + chVal);
    }
    if (pr) crit.push("'" + tab + "'!" + CP + ':' + CP + ',"' + pr + '"');
    crit.push("'" + tab + "'!" + CPS + ':' + CPS + ',"<>Cancelled"');
    crit.push("'" + tab + "'!" + CST + ':' + CST + ',"<>Not Issued"');
    return 'SUMIFS(' + sum + ',' + crit.join(',') + ')';
  }

  // COUNTIFS filtered by Payment Status <> Cancelled AND Policy Status <> Not Issued.
  function fCountifs(tab, ch, pr, chIsRef) {
    const unitR = "'" + tab + "'!" + CU + ':' + CU;
    const crit  = [unitR + ',"<>"'];
    if (ch) {
      const chVal = chIsRef ? ch : '"' + ch + '"';
      crit.push("'" + tab + "'!" + CC + ':' + CC + ',' + chVal);
    }
    if (pr) crit.push("'" + tab + "'!" + CP + ':' + CP + ',"' + pr + '"');
    crit.push("'" + tab + "'!" + CPS + ':' + CPS + ',"<>Cancelled"');
    crit.push("'" + tab + "'!" + CST + ':' + CST + ',"<>Not Issued"');
    return 'COUNTIFS(' + crit.join(',') + ')';
  }

  const labelRow = {};
  const out      = [];
  const WIDTH    = 6; // A–F

  // ── Header rows ──────────────────────────────────────────────────────────
  out.push(['INS Summary (' + ccyLabel + ') \u2014 ' + w.monName + '-' + w.yy + ' ' + w.fy,
            '','','','','']);
  out.push(['Window: ' + w.quarter,'','FX: ' + fxText,'','','']);
  out.push(['','','','','','']);
  out.push(['Channel / Product','Net Premium','Revenue','Units','COGS','Gross Margin']);

  // ── Data rows ─────────────────────────────────────────────────────────────
  layout.forEach(function(rowDef) {
    const sr = out.length + 1;

    // Passthrough: label only, data left blank for manual entry
    if (rowDef.isPassthrough) {
      labelRow[rowDef.label] = sr;
      out.push([rowDef.label,'','','','','']);
      return;
    }

    // Subtotal: SUM over named member rows
    if (rowDef.isSubtotal) {
      labelRow[rowDef.label] = sr;
      const refs = rowDef.members.map(function(m){ return labelRow[m]; }).filter(Boolean);
      function sRef(c) {
        return refs.length ? 'SUM('+refs.map(function(r){return c+r;}).join(',')+')' : '0';
      }
      out.push([rowDef.label,
        '='+sRef('B'), '='+sRef('C'), '='+sRef('D'), '=C'+sr+'-F'+sr, '='+sRef('F')]);
      return;
    }

    // Grand total: SUM over all subtotal rows
    if (rowDef.isGrandTotal) {
      const stRefs = layout
        .filter(function(r){ return r.isSubtotal; })
        .map(function(r){ return labelRow[r.label]; })
        .filter(Boolean);
      function gtRef(c) {
        return stRefs.length ? 'SUM('+stRefs.map(function(r){return c+r;}).join(',')+')' : '0';
      }
      out.push([rowDef.label,
        '='+gtRef('B'), '='+gtRef('C'), '='+gtRef('D'), '=C'+sr+'-F'+sr, '='+gtRef('F')]);
      return;
    }

    // Normal data row
    // Col A label: use the channel string (ch) when it exists, otherwise rowDef.label.
    // This means B2B-MOT, b2c-MOT, Renewal-MOT etc. appear as row headers.
    // The SUMIFS then reference $A{row} (the col-A cell) instead of a hardcoded
    // string, so editing the label cell automatically updates all formulas.
    const ch   = rowDef.ch  || null;
    const pr   = rowDef.pr  || null;
    const tab  = rowDef.tab;
    const rowLabel = ch || rowDef.label;       // channel string preferred as label
    labelRow[rowDef.label] = sr;               // still track by logical label for subtotals

    // chRef: reference to this row's col-A cell, used as the channel criterion.
    // chIsRef=true tells fSumifs/fCountifs NOT to wrap it in extra quotes.
    const chRef  = '$A' + sr;
    const chIsRef = !!ch;   // only use cell-ref criterion when a channel is defined

    out.push([rowLabel,
      '=' + fSumifs(tab, CNP, chIsRef ? chRef : null, pr, chIsRef),   // B: Net Premium
      '=' + fSumifs(tab, CRV, chIsRef ? chRef : null, pr, chIsRef),   // C: Revenue
      '=' + fCountifs(tab, chIsRef ? chRef : null, pr, chIsRef),      // D: Units
      '=C' + sr + '-F' + sr,                                          // E: COGS = Rev - GM
      '=' + fSumifs(tab, CGM, chIsRef ? chRef : null, pr, chIsRef)    // F: Gross Margin
    ]);
  });

  const padded = out.map(function(r){ while (r.length < WIDTH) r.push(''); return r; });
  sh.getRange(1, 1, padded.length, WIDTH).setValues(padded);
}


// ── flashWriteSummaryUSDMirror_ ───────────────────────────────
//  Builds Summary USD as a formula mirror of Summary (PHP).
//  Monetary cells = PHP cell ÷ FX; units and labels referenced directly.
//  6-col layout (A-F, 0-based): B=1 (Net Premium), C=2 (Revenue),
//  D=3 (Units, NOT monetary), E=4 (COGS), F=5 (Gross Margin).
function flashWriteSummaryUSDMirror_(ss, tabName, phpTabName, w, fx) {
  const phpSh = ss.getSheetByName(phpTabName);
  if (!phpSh) return;
  let sh = ss.getSheetByName(tabName);
  if (!sh) sh = ss.insertSheet(tabName);
  else     sh.clearContents();

  const numRows = phpSh.getLastRow();
  const WIDTH   = 6;
  // 6-col layout (A-F, 0-based): B=1 (Net Premium), C=2 (Revenue), E=4 (COGS),
  // F=5 (Gross Margin) are monetary; D=3 (Units) is a count, not monetary.
  const MONEY   = [1,2,4,5];
  const out = [];
  for (let r = 0; r < numRows; r++) {
    const phpRow = r + 1;
    const outRow = [];
    for (let c = 0; c < WIDTH; c++) {
      const ref = "'" + phpTabName + "'!" + colLetter_(c+1) + phpRow;
      if      (r===0 && c===0) outRow.push('INS Summary (USD) — '+w.monName+'-'+w.yy+' '+w.fy);
      else if (r===1 && c===2) outRow.push('FX: PHP ÷ '+fx);
      else if (MONEY.indexOf(c)!==-1) outRow.push('='+ref+'/'+fx);
      else outRow.push('='+ref);
    }
    out.push(outRow);
  }
  sh.getRange(1, 1, out.length, WIDTH).setValues(out);
}


// ── colLetter_ ───────────────────────────────────────────────
//  Convert a 1-based column number to a spreadsheet letter (A…Z, AA, AB…).
function colLetter_(n) {
  let s = '';
  while (n>0) { const r=(n-1)%26; s=String.fromCharCode(65+r)+s; n=Math.floor((n-1)/26); }
  return s;
}

// Converts the { order:[blockName...], blocks:{ [blockName]: {order:[prodLabel...],
// prod:{ [prodLabel]: {label,w,wo} } } } } aggregation structure (built by
// flashBuildSummaryINDO_) into the { blocks, grand } shape flashWriteSummaryTab_
// expects: blocks = [{ name, rows:[{label,w,wo}], subW, subWO }], grand = {w,wo}
// summed across every row in every block. (Previously called but never
// defined — this was causing INDO flash generation to throw
// "flashBlocksFromAgg_ is not defined".)
function flashBlocksFromAgg_(merged) {
  const blocks = [];
  const grand  = { w: flashZero_(), wo: flashZero_() };

  merged.order.forEach(function(blockName) {
    const agg = merged.blocks[blockName];
    if (!agg || !agg.order.length) return;

    const rows  = [];
    const subW  = flashZero_();
    const subWO = flashZero_();

    agg.order.forEach(function(prodLabel) {
      const p = agg.prod[prodLabel];
      rows.push(p);
      flashAcc_(subW,  p.w.np,  p.w.rev,  p.w.gm,  p.w.nops);
      flashAcc_(subWO, p.wo.np, p.wo.rev, p.wo.gm, p.wo.nops);
    });

    flashAcc_(grand.w,  subW.np,  subW.rev,  subW.gm,  subW.nops);
    flashAcc_(grand.wo, subWO.np, subWO.rev, subWO.gm, subWO.nops);

    blocks.push({ name: blockName, rows: rows, subW: subW, subWO: subWO });
  });

  return { blocks: blocks, grand: grand };
}

// ── INDO Summary ─────────────────────────────────────────────
//  Builds Summary (IDR) + Summary USD by resolving metric columns from the
//  header row of each detail tab — exactly the same approach as PH.
//  No manual column index configuration needed.
//
//  Telesales / Agency / Dealer tabs (Indo_Raw columns written by flashConsolidateINDO_):
//    Channel      : tab name is used directly (each tab = one channel)
//    Product      : 'Product Name'
//    Net Premium  : 'Gross Written Premium'  (IDR)
//    Revenue      : 'Net Revenue'
//    Gross Profit : 'Gross Profit'
//    Policy Status: 'Policy Status'  (W/O filter: excludes 'cancelled' and 'rejected')
//
//  B2B(SAU) tabs (24-col output written by flashConsolidateSAU_):
//    Product      : 'Business Class'
//    Net Premium  : 'GWP (USD)'              (already in USD — no FX divide needed)
//    Revenue      : 'Revenue (USD)'
//    Gross Profit : 'Gross Profit (USD)'
//    No status col in SAU output → all SAU rows always count as W/O (non-cancelled)
//
//  W/O cancellations = Policy Status (lowercased) not in FLASH_INDO_EXCLUDE_STATUSES
function flashBuildSummaryINDO_(ss, cfg, w) {

  // Helper: resolve a named column to its 0-based index in a header array.
  // Returns -1 if not found (mirrors String.indexOf behaviour).
  function colIdx(headers, name) {
    const n = name.trim().toLowerCase();
    for (var i = 0; i < headers.length; i++) {
      if (String(headers[i]).trim().toLowerCase() === n) return i;
    }
    return -1;
  }

  const merged = { order:[], blocks:{} };

  // ── Telesales / Agency / Dealer ──────────────────────────────────────────
  // Each tab is a single flat table (one header row, data rows below).
  // Channel = the tab name itself.
  ['Telesales','Agency','Dealer'].forEach(function(tab) {
    const sh = ss.getSheetByName(tab);
    if (!sh) return;
    const values = sh.getDataRange().getValues();
    if (values.length < 2) return;

    const hdr = values[0];
    const iPr  = colIdx(hdr, 'Product Name');
    const iNP  = colIdx(hdr, 'Gross Written Premium');
    const iRev = colIdx(hdr, 'Net Revenue');
    const iGM  = colIdx(hdr, 'Gross Profit');
    const iSt  = colIdx(hdr, 'Policy Status');

    // Warn in log if key columns not found — won't throw, just produces zeros
    if (iNP < 0 || iRev < 0 || iGM < 0) {
      Logger.log('INDO Summary: missing metric column(s) in tab "' + tab + '" — check headers');
    }

    const blk = { order:[], prod:{} };
    values.slice(1).forEach(function(row) {
      const pr   = (iPr  >= 0 ? String(row[iPr]  || '').trim() : '') || 'Comprehensive';
      const np   = iNP  >= 0 ? flashCellNum_(row[iNP])  : 0;
      const rev  = iRev >= 0 ? flashCellNum_(row[iRev]) : 0;
      const gm   = iGM  >= 0 ? flashCellNum_(row[iGM])  : 0;
      const st   = iSt  >= 0 ? String(row[iSt] || '').trim().toLowerCase() : '';
      const inWO = (iSt >= 0) ? !FLASH_INDO_EXCLUDE_STATUSES[st] : true;

      if (!blk.prod[pr]) { blk.prod[pr] = { label:pr, w:flashZero_(), wo:flashZero_() }; blk.order.push(pr); }
      flashAcc_(blk.prod[pr].w, np, rev, gm, 1);
      if (inWO) flashAcc_(blk.prod[pr].wo, np, rev, gm, 1);
    });

    if (blk.order.length) {
      merged.blocks[tab] = blk;
      merged.order.push(tab);
    }
  });

  // ── B2B (SAU) — New Business + Renewal ───────────────────────────────────
  // SAU tabs use a 24-col output schema written by flashConsolidateSAU_.
  // Metric values are already in USD (GWP/Revenue/Gross Profit columns).
  // For the IDR Summary we multiply back by FX; the USD Summary divides by FX
  // again, which nets to 1 — effectively keeping USD values in the IDR tab too.
  // This matches how SAU has always been reported: in USD, not IDR.
  const sauBlk = { order:[], prod:{} };
  [['B2B(SAU)-New Business','New Business'], ['B2B(SAU)-Renewal','Renewal']].forEach(function(pair) {
    const sh = ss.getSheetByName(pair[0]);
    if (!sh) return;
    const values = sh.getDataRange().getValues();
    if (values.length < 2) return;

    const hdr  = values[0];
    const iPr  = colIdx(hdr, 'Business Class');
    const iNP  = colIdx(hdr, 'GWP (USD)');
    const iRev = colIdx(hdr, 'Revenue (USD)');
    const iGM  = colIdx(hdr, 'Gross Profit (USD)');

    if (iNP < 0 || iRev < 0 || iGM < 0) {
      Logger.log('INDO Summary: missing metric column(s) in SAU tab "' + pair[0] + '" — check headers');
    }

    const pr = pair[1]; // 'New Business' or 'Renewal'
    if (!sauBlk.prod[pr]) { sauBlk.prod[pr] = { label:pr, w:flashZero_(), wo:flashZero_() }; sauBlk.order.push(pr); }

    values.slice(1).forEach(function(row) {
      // SAU values are in USD — multiply by FX to get IDR-equivalent for the IDR Summary.
      // The USD Summary will then divide by FX, returning the original USD value.
      const np  = (iNP  >= 0 ? flashCellNum_(row[iNP])  : 0) * cfg.fx;
      const rev = (iRev >= 0 ? flashCellNum_(row[iRev]) : 0) * cfg.fx;
      const gm  = (iGM  >= 0 ? flashCellNum_(row[iGM])  : 0) * cfg.fx;
      flashAcc_(sauBlk.prod[pr].w,  np, rev, gm, 1);
      flashAcc_(sauBlk.prod[pr].wo, np, rev, gm, 1); // SAU has no cancellation status
    });
  });
  if (sauBlk.order.length) {
    merged.blocks['B2B (SAU)'] = sauBlk;
    merged.order.push('B2B (SAU)');
  }

  const bg = flashBlocksFromAgg_(merged);
  flashWriteSummaryTab_(ss, 'Summary (IDR)', bg.blocks, bg.grand, 1,      'IDR', w, 'local currency');
  flashWriteSummaryTab_(ss, 'Summary USD',   bg.blocks, bg.grand, cfg.fx, 'USD', w, 'IDR ÷ ' + cfg.fx);
  return { blocks: bg.blocks.length, note: 'INDO Summary (IDR) + Summary USD built from headers.' };
}