/**
 * ============================================================
 * ANGELVEREIN WASSERQUALITÄTS-TRACKER
 * Fishing Club Water Quality Tracker
 * ============================================================
 *
 * SETUP INSTRUCTIONS:
 *   1. Open a new Google Sheet
 *   2. Extensions → Apps Script
 *   3. Delete all existing code, paste this file's contents
 *   4. Create a new HTML file named "DatePicker" (+ button → HTML),
 *      paste the contents of DatePicker.html into it
 *   5. Save (Ctrl+S), then run setupSpreadsheet() once
 *   6. Authorize the script when prompted
 *
 * After setup, use the "🎣 Wasserdaten" menu in your spreadsheet.
 * ============================================================
 */

// ============================================================
// CONFIGURATION — edit pond names / limits here if needed
// ============================================================

const PONDS = [
  'Dungeteich',
  'Kl. Lesumteich',
  'Gr. Lesumteich',
  'Tietjenteich',
  'Deichkämpe',
  'Ihlebecken',
  'Schlossteich'
];

const WEATHER_OPTIONS = ['Sonnig', 'Bewölkt', 'Regnerisch', 'Windig', 'Neblig'];

/**
 * Parameter definitions.
 * dataCol: 0-based column index in the Messdaten sheet
 *   0=Date 1=Pond 2=Time 3=WaterTemp 4=AirTemp 5=Weather 6=Depth
 *   7=O2  8=pH  9=NH4  10=Carb  11=SVB  12=NO3  13=NO2  14=PO4  15=GH
 *
 * limitMin / limitMax: null means no limit in that direction.
 * SVB is auto-calculated from Carbonathärte ÷ 2.8 (no manual limits).
 */
const PARAMS = [
  { key: 'o2',   label: 'Sauerstoff O²',  unit: 'mg/l',       limitMin: 5.0,  limitMax: null, dataCol: 7  },
  { key: 'ph',   label: 'pH-Wert',         unit: '',           limitMin: 6.5,  limitMax: 8.0,  dataCol: 8  },
  { key: 'nh4',  label: 'Ammonium NH4+',   unit: 'mg/l',       limitMin: null, limitMax: 2.0,  dataCol: 9  },
  { key: 'carb', label: 'Carbonathärte',   unit: 'D°',         limitMin: 1.0,  limitMax: null, dataCol: 10 },
  { key: 'svb',  label: 'SVB',             unit: 'mmol',       limitMin: null, limitMax: null, dataCol: 11 },
  { key: 'no3',  label: 'Nitrat NO3⁻',     unit: 'mg/l',       limitMin: null, limitMax: 50.0, dataCol: 12 },
  { key: 'no2',  label: 'Nitrit NO2⁻',     unit: 'mg/l',       limitMin: null, limitMax: 0.01, dataCol: 13 },
  { key: 'po4',  label: 'Phosphat PO4²⁻',  unit: 'mg/l',       limitMin: null, limitMax: 0.3,  dataCol: 14 },
  { key: 'gh',   label: 'Gesamthärte',     unit: 'DH° mmol/l', limitMin: null, limitMax: null, dataCol: 15 }
];

// Colour per pond (used in charts + dashboard row tinting)
const POND_COLORS = [
  '#4285F4', // Blue        – Dungeteich
  '#34A853', // Green       – Kl. Lesumteich
  '#F9A825', // Amber       – Gr. Lesumteich
  '#EA4335', // Red         – Tietjenteich
  '#9C27B0', // Purple      – Deichkämpe
  '#FF6D00', // Orange      – Ihlebecken
  '#00BCD4'  // Cyan        – Schlossteich
];

// Sheet names
const SHEET = {
  data:      'Messdaten',
  chartData: 'Diagramm-Daten',
  charts:    'Diagramme',
  dashboard: 'Übersicht'
};

// Each parameter block in the ChartData sheet is this many columns wide:
// [Date] + [7 ponds] + [LimitMin] + [LimitMax] = 10
const BLOCK_COLS    = 10;
const BLOCK_SPACER  = 1;   // 1 blank column between blocks
const MAX_DATA_ROWS = 300;  // max unique visit dates (~6 years of weekly visits)

// ============================================================
// MENU
// ============================================================

function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu('🎣 Wasserdaten')
    .addItem('📅 Neuen Eintrag hinzufügen', 'showAddEntryDialog')
    .addSeparator()
    .addItem('⬆️  Älteste zuerst (aufsteigend)',  'sortAscending')
    .addItem('⬇️  Neueste zuerst (absteigend)',   'sortDescending')
    .addSeparator()
    .addItem('🔄 Diagramme aktualisieren',  'refreshChartData')
    .addItem('📊 Übersicht aktualisieren',  'refreshDashboard')
    .addSeparator()
    .addItem('⚙️  Ersteinrichtung (einmalig)', 'setupSpreadsheet')
    .addToUi();
}

// ============================================================
// ONE-TIME SETUP
// ============================================================

function setupSpreadsheet() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const ui = SpreadsheetApp.getUi();

  // Get or create sheets
  const dataSheet      = _getOrCreate(ss, SHEET.data);
  const chartDataSheet = _getOrCreate(ss, SHEET.chartData);
  const chartsSheet    = _getOrCreate(ss, SHEET.charts);
  const dashboardSheet = _getOrCreate(ss, SHEET.dashboard);

  // Remove default "Sheet1" / "Tabelle1" if it still exists
  ['Sheet1', 'Tabelle1'].forEach(name => {
    const s = ss.getSheetByName(name);
    if (s) try { ss.deleteSheet(s); } catch(e) {}
  });

  // Order sheets: Dashboard | Messdaten | Diagramme | (ChartData hidden)
  ss.setActiveSheet(dashboardSheet);
  ss.moveActiveSheet(1);
  ss.setActiveSheet(dataSheet);
  ss.moveActiveSheet(2);
  ss.setActiveSheet(chartsSheet);
  ss.moveActiveSheet(3);
  ss.setActiveSheet(chartDataSheet);
  ss.moveActiveSheet(4);

  // Build each sheet
  _setupDataSheet(dataSheet);
  _setupChartDataSheet(chartDataSheet);
  _setupDashboardSheet(dashboardSheet);

  // Populate ChartData + charts
  refreshChartData();
  _setupCharts(chartsSheet);
  refreshDashboard();

  // Install the onEdit trigger (needed for auto-date & SVB calc)
  _setupTriggers();

  // Activate the data entry sheet
  ss.setActiveSheet(dataSheet);

  ui.alert(
    '✅ Einrichtung abgeschlossen!',
    'Der Tracker ist bereit.\n\n' +
    '• Verwende "🎣 Wasserdaten → Neuen Eintrag hinzufügen" um Messungen zu erfassen.\n' +
    '• Die Diagramme befinden sich im Tab "Diagramme".\n' +
    '• Die Übersicht zeigt immer die aktuellsten Werte.',
    ui.ButtonSet.OK
  );
}

function _getOrCreate(ss, name) {
  return ss.getSheetByName(name) || ss.insertSheet(name);
}

// ============================================================
// DATA SHEET
// ============================================================

function _setupDataSheet(sheet) {
  sheet.clear();
  sheet.setTabColor('#1E3A5F');
  sheet.setFrozenRows(1);

  // ── Headers ──────────────────────────────────────────────
  const headers = [
    'Datum', 'Teich', 'Uhrzeit',
    'Wassertemp.\n(°C)', 'Lufttemp.\n(°C)', 'Wetter', 'Messtiefe\n(m)',
    'Sauerstoff O²\n(mg/l)', 'pH-Wert',
    'Ammonium\nNH4+ (mg/l)', 'Carbonathärte\n(D°)', 'SVB\n(mmol) auto',
    'Nitrat NO3⁻\n(mg/l)', 'Nitrit NO2⁻\n(mg/l)',
    'Phosphat PO4²⁻\n(mg/l)', 'Gesamthärte\n(DH° mmol/l)'
  ];

  const headerRange = sheet.getRange(1, 1, 1, headers.length);
  headerRange.setValues([headers])
    .setBackground('#1E3A5F')
    .setFontColor('#FFFFFF')
    .setFontWeight('bold')
    .setWrap(true)
    .setVerticalAlignment('middle')
    .setHorizontalAlignment('center');
  sheet.setRowHeight(1, 52);

  // ── Column widths ─────────────────────────────────────────
  const colWidths = [105, 135, 70, 80, 80, 90, 80, 95, 75, 95, 105, 100, 95, 95, 110, 130];
  colWidths.forEach((w, i) => sheet.setColumnWidth(i + 1, w));

  // ── Data validation ───────────────────────────────────────
  // Pond dropdown
  sheet.getRange(2, 2, 999, 1)
    .setDataValidation(
      SpreadsheetApp.newDataValidation()
        .requireValueInList(PONDS, true)
        .setAllowInvalid(false)
        .build()
    );

  // Weather dropdown
  sheet.getRange(2, 6, 999, 1)
    .setDataValidation(
      SpreadsheetApp.newDataValidation()
        .requireValueInList(WEATHER_OPTIONS, true)
        .setAllowInvalid(true)
        .build()
    );

  // Date type (gives native calendar picker on mobile too)
  sheet.getRange(2, 1, 999, 1)
    .setDataValidation(
      SpreadsheetApp.newDataValidation()
        .requireDate()
        .setAllowInvalid(true)
        .build()
    );

  // Time dropdown (Uhrzeit) — every 30 min from 06:00 to 22:30
  const times = [];
  for (let h = 6; h <= 22; h++) {
    times.push(String(h).padStart(2, '0') + ':00');
    times.push(String(h).padStart(2, '0') + ':30');
  }
  sheet.getRange(2, 3, 999, 1)
    .setDataValidation(
      SpreadsheetApp.newDataValidation()
        .requireValueInList(times, true)
        .setAllowInvalid(true)   // still allows manual typing
        .build()
    );

  // ── Number formats ────────────────────────────────────────
  sheet.getRange(2, 1, 999, 1).setNumberFormat('DD.MM.YYYY');
  sheet.getRange(2, 3, 999, 1).setNumberFormat('@');        // plain text e.g. "14:30"
  sheet.getRange(2, 4, 999, 2).setNumberFormat('0.0');
  sheet.getRange(2, 7, 999, 1).setNumberFormat('0.0');
  sheet.getRange(2, 8, 999, 1).setNumberFormat('0.00');
  sheet.getRange(2, 9, 999, 1).setNumberFormat('0.00');
  sheet.getRange(2, 10, 999, 1).setNumberFormat('0.000');
  sheet.getRange(2, 11, 999, 1).setNumberFormat('0.0');
  sheet.getRange(2, 12, 999, 1).setNumberFormat('0.000');
  sheet.getRange(2, 13, 999, 1).setNumberFormat('0.0');
  sheet.getRange(2, 14, 999, 1).setNumberFormat('0.0000');
  sheet.getRange(2, 15, 999, 1).setNumberFormat('0.000');
  sheet.getRange(2, 16, 999, 1).setNumberFormat('0.00');

  // ── Alternating row banding ───────────────────────────────
  if (sheet.getBandings().length === 0) {
    sheet.getRange(1, 1, 1000, 16)
      .applyRowBanding(SpreadsheetApp.BandingTheme.LIGHT_GREY)
      .setHeaderRowColor('#1E3A5F')
      .setFirstRowColor('#F8F9FA')
      .setSecondRowColor('#FFFFFF');
  }

  // ── Native filter row (click ▼ on any header to sort/filter) ──
  if (!sheet.getFilter()) {
    sheet.getRange(1, 1, 1, 16).createFilter();
  }

  // ── Conditional formatting (violation = red cell) ─────────
  _setupConditionalFormatting(sheet);
}

function _setupConditionalFormatting(sheet) {
  const rules = [];

  // Helper: add a rule for violations
  const addRule = (col, condition) => {
    rules.push(
      condition
        .setBackground('#FFCCCC')
        .setFontColor('#B71C1C')
        .setRanges([sheet.getRange(2, col, 999, 1)])
        .build()
    );
  };

  // Sauerstoff O² col 8: < 5.0 bad
  addRule(8,  SpreadsheetApp.newConditionalFormatRule().whenNumberLessThan(5.0));

  // pH col 9: < 6.5 or > 8.0 bad
  addRule(9,  SpreadsheetApp.newConditionalFormatRule().whenNumberLessThan(6.5));
  addRule(9,  SpreadsheetApp.newConditionalFormatRule().whenNumberGreaterThan(8.0));

  // Ammonium col 10: > 2.0 bad
  addRule(10, SpreadsheetApp.newConditionalFormatRule().whenNumberGreaterThan(2.0));

  // Carbonathärte col 11: < 1.0 bad
  addRule(11, SpreadsheetApp.newConditionalFormatRule().whenNumberLessThan(1.0));

  // Nitrat col 13: > 50.0 bad
  addRule(13, SpreadsheetApp.newConditionalFormatRule().whenNumberGreaterThan(50.0));

  // Nitrit col 14: > 0.01 bad
  addRule(14, SpreadsheetApp.newConditionalFormatRule().whenNumberGreaterThan(0.01));

  // Phosphat col 15: > 0.3 bad
  addRule(15, SpreadsheetApp.newConditionalFormatRule().whenNumberGreaterThan(0.3));

  sheet.setConditionalFormatRules(rules);
}

// ============================================================
// CHART DATA SHEET  (hidden helper sheet)
// ============================================================

function _setupChartDataSheet(sheet) {
  sheet.clear();
  sheet.setTabColor('#CCCCCC');
  sheet.hideSheet();
  // Content is written by refreshChartData()
}

/**
 * Returns the 1-based start column for a given parameter block.
 * Layout: each block is BLOCK_COLS wide + BLOCK_SPACER gap.
 */
function _blockStartCol(paramIndex) {
  return paramIndex * (BLOCK_COLS + BLOCK_SPACER) + 1;
}

/**
 * Reads Messdaten, builds one pivot table per parameter:
 *   [Date | Pond1..7 | LimitMin | LimitMax]
 * and writes each block to the ChartData sheet.
 */
function refreshChartData() {
  const ss             = SpreadsheetApp.getActiveSpreadsheet();
  const dataSheet      = ss.getSheetByName(SHEET.data);
  const chartDataSheet = ss.getSheetByName(SHEET.chartData);
  if (!dataSheet || !chartDataSheet) return;

  const allValues = dataSheet.getDataRange().getValues();
  if (allValues.length <= 1) return; // header only

  // Filter out empty rows, sort by date ascending
  const rows = allValues.slice(1)
    .filter(r => r[0] !== '' && r[1] !== '')
    .sort((a, b) => _toDate(a[0]) - _toDate(b[0]));

  if (rows.length === 0) return;

  // Unique dates as "DD.MM.YYYY" strings (preserving sort order)
  const datesSeen = new Set();
  const uniqueDates = [];
  rows.forEach(r => {
    const ds = _fmtDate(_toDate(r[0]));
    if (!datesSeen.has(ds)) { datesSeen.add(ds); uniqueDates.push(ds); }
  });

  // Build a lookup: { "DD.MM.YYYY|PondName" → row }
  const lookup = {};
  rows.forEach(r => {
    const key = `${_fmtDate(_toDate(r[0]))}|${r[1]}`;
    lookup[key] = r;
  });

  // Clear entire chart data sheet first for a clean slate
  chartDataSheet.clearContents();

  PARAMS.forEach((param, pIdx) => {
    const startCol = _blockStartCol(pIdx);

    // Header: Datum | Pond1..7 | Min-Grenzwert | Max-Grenzwert
    const blockHeader = ['Datum', ...PONDS, 'Min-Grenzwert', 'Max-Grenzwert'];

    // Data rows
    const dataRows = uniqueDates.map(dateStr => {
      const row = [dateStr];
      PONDS.forEach(pond => {
        const r = lookup[`${dateStr}|${pond}`];
        if (r && r[param.dataCol] !== '') {
          const v = parseFloat(r[param.dataCol]);
          row.push(isNaN(v) ? '' : v);
        } else {
          row.push('');
        }
      });
      // Constant limit lines (same value repeated for every date row)
      row.push(param.limitMin !== null ? param.limitMin : '');
      row.push(param.limitMax !== null ? param.limitMax : '');
      return row;
    });

    const writeData = [blockHeader, ...dataRows];
    chartDataSheet.getRange(1, startCol, writeData.length, BLOCK_COLS)
      .setValues(writeData);

    // Style header row
    chartDataSheet.getRange(1, startCol, 1, BLOCK_COLS)
      .setBackground('#2C5F8A')
      .setFontColor('#FFFFFF')
      .setFontWeight('bold');
  });

  SpreadsheetApp.flush();
}

// ============================================================
// CHARTS
// ============================================================

// Run these three in order (one after another) to create all 9 charts:
function setupCharts_Part1() { _setupChartsRange(0, 3, true);  } // O2, pH, NH4
function setupCharts_Part2() { _setupChartsRange(3, 6, false); } // Carb, SVB, NO3
function setupCharts_Part3() { _setupChartsRange(6, 9, false); } // NO2, PO4, GH

/** Called once during full setup — only creates first 3 charts to avoid timeout */
function _setupCharts(chartsSheet) {
  chartsSheet.setTabColor('#34A853');
  _setupChartsRange(0, 3, true);
}

function _setupChartsRange(fromIdx, toIdx, clearFirst) {
  const ss             = SpreadsheetApp.getActiveSpreadsheet();
  const chartsSheet    = ss.getSheetByName(SHEET.charts);
  const chartDataSheet = ss.getSheetByName(SHEET.chartData);

  if (clearFirst) {
    chartsSheet.getCharts().forEach(c => chartsSheet.removeChart(c));
  }

  const CHART_W      = 620;
  const CHART_H      = 400;
  const PER_ROW      = 2;
  const ROW_HEIGHT   = 24;
  const COL_WIDTH    = 10;
  for (let pIdx = fromIdx; pIdx < toIdx; pIdx++) {
    const param      = PARAMS[pIdx];
    const startCol   = _blockStartCol(pIdx);
    const dataRange  = chartDataSheet.getRange(1, startCol, MAX_DATA_ROWS + 1, BLOCK_COLS);

    // ── Series configuration ─────────────────────────────────
    // Columns: [Date] [Pond0..6] [LimitMin] [LimitMax]
    // Series indices (0-based, Date col is excluded by chart):
    //   0-6  → pond lines   (coloured)
    //   7    → Min limit    (orange dashed)
    //   8    → Max limit    (red dashed)
    const series = {};

    PONDS.forEach((_, i) => {
      series[i] = {
        color:           POND_COLORS[i],
        lineWidth:       2,
        pointSize:       5,
        visibleInLegend: true
      };
    });

    // LimitMin series (index 7)
    series[7] = {
      color:           '#FF8C00',
      lineWidth:       2,
      pointSize:       0,
      lineDashStyle:   [6, 3],
      visibleInLegend: param.limitMin !== null,
      labelInLegend:   param.limitMin !== null ? `Min: ${param.limitMin}` : ''
    };

    // LimitMax series (index 8)
    series[8] = {
      color:           '#CC0000',
      lineWidth:       2,
      pointSize:       0,
      lineDashStyle:   [6, 3],
      visibleInLegend: param.limitMax !== null,
      labelInLegend:   param.limitMax !== null ? `Max: ${param.limitMax}` : ''
    };

    // ── Position in sheet ─────────────────────────────────────
    const chartRow    = Math.floor(pIdx / PER_ROW);
    const chartColPos = pIdx % PER_ROW;
    const anchorRow   = chartRow * (ROW_HEIGHT + 2) + 1;
    const anchorCol   = chartColPos * (COL_WIDTH + 1) + 1;

    const title     = param.unit ? `${param.label} (${param.unit})` : param.label;
    let   limitNote = '';
    if (param.limitMin !== null && param.limitMax !== null)
      limitNote = `Sollwert: ${param.limitMin} – ${param.limitMax}`;
    else if (param.limitMin !== null)
      limitNote = `Mindestens: ${param.limitMin}`;
    else if (param.limitMax !== null)
      limitNote = `Höchstens: ${param.limitMax}`;

    const chart = chartsSheet.newChart()
      .setChartType(Charts.ChartType.LINE)
      .addRange(dataRange)
      .setNumHeaders(1)
      .setTransposeRowsAndColumns(false)
      .setPosition(anchorRow, anchorCol, 5, 0)
      .setOption('title',    title)
      .setOption('subtitle', limitNote)
      .setOption('width',    CHART_W)
      .setOption('height',   CHART_H)
      .setOption('hAxis', {
        title:            'Datum',
        slantedText:      true,
        slantedTextAngle: 45,
        textStyle:        { fontSize: 10 }
      })
      .setOption('vAxis', {
        title:     param.unit || param.label,
        textStyle: { fontSize: 10 }
      })
      .setOption('series',           series)
      .setOption('interpolateNulls', false)
      .setOption('legend',           { position: 'bottom', maxLines: 3 })
      .setOption('backgroundColor',  '#FFFFFF')
      .setOption('chartArea',        { left: 65, top: 55, width: '72%', height: '58%' })
      .setOption('lineWidth',        2)
      .build();

    chartsSheet.insertChart(chart);
  }

  SpreadsheetApp.flush();
}

// ============================================================
// DASHBOARD SHEET
// ============================================================

function _setupDashboardSheet(sheet) {
  sheet.clear();
  sheet.setTabColor('#34A853');
  sheet.setFrozenRows(4);

  // Title
  sheet.getRange('A1').setValue('🎣 Angelverein – Wasserqualitäts-Übersicht');
  sheet.getRange('A1')
    .setFontSize(16)
    .setFontWeight('bold')
    .setFontColor('#1E3A5F');

  sheet.getRange('A2').setValue('Letzte gemessene Werte je Teich')
    .setFontColor('#666666').setFontSize(10);

  sheet.getRange('A3').setValue('') // timestamp filled by refreshDashboard

  // Legend note
  sheet.getRange('A4').setValue(
    '🟢 = im Sollbereich   🔴 = außerhalb Grenzwert   ⬜ = kein Grenzwert definiert'
  ).setFontColor('#555555').setFontSize(9).setFontStyle('italic');
  sheet.getRange(4, 1, 1, 12).merge();

  // Column headers (row 5)
  const colHeaders = ['Teich', 'Datum'];
  PARAMS.forEach(p => colHeaders.push(p.unit ? `${p.label}\n(${p.unit})` : p.label));

  sheet.getRange(5, 1, 1, colHeaders.length)
    .setValues([colHeaders])
    .setBackground('#1E3A5F')
    .setFontColor('#FFFFFF')
    .setFontWeight('bold')
    .setWrap(true)
    .setVerticalAlignment('middle')
    .setHorizontalAlignment('center');
  sheet.setRowHeight(5, 52);

  // Pond rows (6–12)
  PONDS.forEach((pond, i) => {
    const row = sheet.getRange(6 + i, 1);
    row.setValue(pond).setFontWeight('bold');
    // Tint row with pond colour at low opacity
    sheet.getRange(6 + i, 1, 1, colHeaders.length)
      .setBackground(POND_COLORS[i] + '22');
  });

  // Column widths
  sheet.setColumnWidth(1, 140);
  sheet.setColumnWidth(2, 105);
  PARAMS.forEach((_, i) => sheet.setColumnWidth(3 + i, 95));

  sheet.setFrozenRows(5);
}

function refreshDashboard() {
  const ss        = SpreadsheetApp.getActiveSpreadsheet();
  const dataSheet = ss.getSheetByName(SHEET.data);
  const dashSheet = ss.getSheetByName(SHEET.dashboard);
  if (!dataSheet || !dashSheet) return;

  const allValues = dataSheet.getDataRange().getValues();
  if (allValues.length <= 1) return;

  const rows = allValues.slice(1).filter(r => r[0] !== '' && r[1] !== '');

  PONDS.forEach((pond, pIdx) => {
    const pondRows = rows.filter(r => r[1] === pond);
    if (pondRows.length === 0) return;

    // Sort descending → first element is latest
    pondRows.sort((a, b) => _toDate(b[0]) - _toDate(a[0]));
    const latest = pondRows[0];

    const sheetRow = 6 + pIdx;
    dashSheet.getRange(sheetRow, 2).setValue(_fmtDate(_toDate(latest[0])));

    PARAMS.forEach((param, paramIdx) => {
      const cell = dashSheet.getRange(sheetRow, 3 + paramIdx);
      const raw  = latest[param.dataCol];

      if (raw !== '' && raw !== null && raw !== undefined) {
        const v = parseFloat(raw);
        cell.setValue(v);
        if (!isNaN(v)) {
          if (_isViolation(v, param)) {
            cell.setBackground('#FFCCCC').setFontColor('#B71C1C').setFontWeight('bold');
          } else if (param.limitMin !== null || param.limitMax !== null) {
            cell.setBackground('#C8E6C9').setFontColor('#1B5E20').setFontWeight('normal');
          } else {
            cell.setBackground(POND_COLORS[pIdx] + '18')
                .setFontColor('#333333').setFontWeight('normal');
          }
        }
      } else {
        cell.setValue('–').setFontColor('#AAAAAA').setBackground(null);
      }
    });
  });

  dashSheet.getRange('A3')
    .setValue('Aktualisiert: ' + new Date().toLocaleString('de-DE', {
      day: '2-digit', month: '2-digit', year: 'numeric',
      hour: '2-digit', minute: '2-digit'
    }))
    .setFontColor('#999999').setFontSize(9);

  SpreadsheetApp.flush();
}

// ============================================================
// DATA ENTRY — date picker dialog
// ============================================================

function showAddEntryDialog() {
  const html = HtmlService.createHtmlOutputFromFile('DatePicker')
    .setWidth(400)
    .setHeight(220);
  SpreadsheetApp.getUi().showModalDialog(html, '📅 Datum für neue Messung');
}

/**
 * Called by the HTML dialog after the user picks a date.
 * Creates 7 new rows (one per pond) and navigates to them.
 */
function startNewEntry(dateStr) {
  // dateStr = "YYYY-MM-DD"
  const ss        = SpreadsheetApp.getActiveSpreadsheet();
  const dataSheet = ss.getSheetByName(SHEET.data);

  const parts = dateStr.split('-').map(Number);
  const date  = new Date(parts[0], parts[1] - 1, parts[2]);

  const lastRow    = dataSheet.getLastRow();
  const insertRow  = lastRow + 1;
  const numCols    = 16;

  const newRows = PONDS.map(pond => {
    const row    = new Array(numCols).fill('');
    row[0] = date;
    row[1] = pond;
    return row;
  });

  dataSheet.getRange(insertRow, 1, newRows.length, numCols).setValues(newRows);

  // Sort everything by date
  sortDataByDate(true);

  // Navigate to data sheet (charts update is done by auto-trigger)
  ss.setActiveSheet(dataSheet);
}

// ============================================================
// SORTING
// ============================================================

function sortAscending()  { sortDataByDate(true);  }
function sortDescending() { sortDataByDate(false); }

function sortDataByDate(ascending) {
  const ss        = SpreadsheetApp.getActiveSpreadsheet();
  const sheet     = ss.getSheetByName(SHEET.data);
  const lastRow   = sheet.getLastRow();
  if (lastRow <= 1) return;

  sheet.getRange(2, 1, lastRow - 1, 16)
    .sort([
      { column: 1, ascending: ascending }, // primary:   Date
      { column: 2, ascending: true }       // secondary: Pond name
    ]);

  refreshChartData();
  refreshDashboard();
}

// ============================================================
// INSTALLABLE TRIGGER (onEdit equivalent with full permissions)
// ============================================================

function _setupTriggers() {
  // Remove any existing trigger with the same handler to avoid duplicates
  ScriptApp.getProjectTriggers().forEach(t => {
    if (t.getHandlerFunction() === 'onEditInstallable') {
      ScriptApp.deleteTrigger(t);
    }
  });

  ScriptApp.newTrigger('onEditInstallable')
    .forSpreadsheet(SpreadsheetApp.getActive())
    .onEdit()
    .create();
}

/**
 * Runs on every edit.
 * 1. Auto-fills today's date when a user starts typing in a new row.
 * 2. Auto-calculates SVB when Carbonathärte (col 11) is entered.
 */
function onEditInstallable(e) {
  if (!e || !e.range) return;

  const sheet = e.range.getSheet();
  if (sheet.getName() !== SHEET.data) return;

  const row = e.range.getRow();
  const col = e.range.getColumn();
  if (row <= 1) return; // header row

  // Auto-fill date if the Date cell (col 1) is empty and user edits another col
  const dateCell = sheet.getRange(row, 1);
  if (col !== 1 && dateCell.getValue() === '') {
    dateCell.setValue(new Date());
  }

  // Auto-calculate SVB = Carbonathärte / 2.8  (col 11 → col 12)
  if (col === 11) {
    const carbVal = parseFloat(e.range.getValue());
    if (!isNaN(carbVal)) {
      sheet.getRange(row, 12).setValue(Math.round((carbVal / 2.8) * 10000) / 10000);
    }
  }
}

// ============================================================
// UTILITY FUNCTIONS
// ============================================================

function _toDate(v) {
  return v instanceof Date ? v : new Date(v);
}

function _fmtDate(d) {
  const dd   = String(d.getDate()).padStart(2, '0');
  const mm   = String(d.getMonth() + 1).padStart(2, '0');
  const yyyy = d.getFullYear();
  return `${dd}.${mm}.${yyyy}`;
}

function _isViolation(val, param) {
  if (isNaN(val)) return false;
  if (param.limitMin !== null && val < param.limitMin) return true;
  if (param.limitMax !== null && val > param.limitMax) return true;
  return false;
}
