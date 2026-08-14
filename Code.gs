// ╔══════════════════════════════════════════════════════════════════╗
//   IBI STAFF ATTENDANCE — Google Apps Script Backend
//   Updated: v2.0 (Jul 2026)  ·  Month-Wise Edition
//
//   NEW IN v2.0
//   • SERVER-SIDE GEOFENCE: punches are now refused server-side unless the
//     phone is within WH_LAT/WH_LON ± MAX_DIST_M with GPS accuracy ≤ MAX_ACC_M.
//     The old build recorded any coordinates (or none) — home-punching is closed.
//   • STAFF ALLOW-LIST: only names in STAFF_NAMES may punch or be read.
//   • ERP SALARY FEED: doGet?action=getStaffSalaries returns per-staff/per-month
//     salary totals; the IBI ERP System pulls it as de-duplicated "Salary" expenses.
//     >>> After pasting this file, Deploy → Manage deployments → Edit → New version. <<<
//
//   WHAT'S NEW IN THIS VERSION
//   • Each staff member gets a NEW TAB EVERY MONTH, auto-created
//     on the first punch of that month.
//     Naming pattern:  "S. Aswin - May 2026", "M. Ajay - May 2026"
//   • Headers, colors (#00c5ff on black), Roboto font, column widths,
//     number formats — all applied automatically.
//   • Salary columns auto-recalculate after every punch.
//   • Cumulative salary resets to 0 at the start of each month.
//
//   FIRST-TIME SETUP (one click)
//   1. Reload the spreadsheet — an "IBI Tools" menu appears.
//   2. Click  IBI Tools → Migrate Legacy Tabs
//        (renames "S. Aswin" → "S. Aswin - April 2026", etc.)
//   3. Click  IBI Tools → Setup Current Month
//        (creates this month's tabs in advance)
//   That's it — system runs itself from now on.
// ╚══════════════════════════════════════════════════════════════════╝

// ── Configuration ────────────────────────────────────────────────
var AUTH_TOKEN  = 'IBI_ATTEND_2026';                  // Must match CFG.TOKEN in index.html
var TZ          = 'Asia/Kolkata';
var TIME_FORMAT = 'hh:mm:ss a';                       // "08:51:00 AM"
var DATE_FORMAT = 'dd-MM-yyyy';                       // "01-05-2026"
var HEADER_ROWS = 1;
var HOURLY_RATE = 100;                                // Rs per hour
var STAFF_NAMES = ['S. Aswin', 'M. Ajay'];            // ← add new staff names here (also the server-side allow-list)
var APP_VERSION = 'v2.3';                             // shown in About; keep in sync with index.html badge

// ── Anti-repeat-punch guard (v2.3) ───────────────────────────────
//   The v2.0 firewall only enforced IN/OUT *alternation*. It had no notion of
//   TIME, so a staff member who tapped OUT and then re-tapped a few seconds
//   later (thinking the first tap "didn't work") produced a perfectly legal
//   alternating punch that opened a brand-new shift column — e.g. 11-08-2026
//   OUT-2 7:35:02 PM then IN-3 7:35:43 PM, 41 seconds apart. Those are the
//   repeated entries seen in the sheet; they are mis-taps, not staff fraud.
//   A punch landing within MIN_PUNCH_GAP_SEC of the previous one is now
//   refused, which is what every real attendance system does.
var MIN_PUNCH_GAP_SEC = 120;   // seconds: reject a punch this soon after the previous one
var STALE_PUNCH_SEC   = 900;   // seconds: refuse an offline-queued punch older than this (stops
                               //          a punch queued in the morning replaying in the evening)
var REPLAY_CACHE_SEC  = 1800;  // seconds to remember a punchId for idempotent retries
var LOCK_WAIT_MS      = 25000; // how long a concurrent punch waits for the script lock

// ── Server-side geofence (tamper-proofing, HARD REJECT) ──────────
//   The browser geofence in index.html is only for UX. THESE checks are
//   the authoritative gate — a punch is refused unless the phone proves it
//   is physically at the warehouse with a good-enough GPS fix. This blocks
//   punching from home / via a hand-crafted request, which the old build
//   allowed (the server used to record any coordinates, or none at all).
var WH_LAT     = 8.202612;    // warehouse latitude  (must match index.html CFG.WAREHOUSE_LAT)
var WH_LON     = 77.240076;   // warehouse longitude (must match index.html CFG.WAREHOUSE_LON)
var MAX_DIST_M = 75;          // metres: a punch beyond this from the warehouse is REJECTED (client locks at 50; server adds boundary tolerance)
var MAX_ACC_M  = 150;         // metres: reject only clearly-garbage GPS fixes. Kept lenient so the network-location
                              //         fallback (for cheap phones that can't get precise GPS) still works on-site;
                              //         the DISTANCE check above is the real anti-home-punch guard.

// Haversine distance in metres (server-side geofence math)
function havM(la1, lo1, la2, lo2) {
  var R  = 6371000;
  var f1 = la1 * Math.PI / 180, f2 = la2 * Math.PI / 180;
  var df = (la2 - la1) * Math.PI / 180, dl = (lo2 - lo1) * Math.PI / 180;
  var a  = Math.sin(df / 2) * Math.sin(df / 2) +
           Math.cos(f1) * Math.cos(f2) * Math.sin(dl / 2) * Math.sin(dl / 2);
  return R * 2 * Math.atan2(Math.sqrt(a), Math.sqrt(1 - a));
}

var COLUMN_HEADERS = [
  'DATE','IN-1','OUT-1','IN-2','OUT-2','IN-3','OUT-3',
  'IN-4','OUT-4','IN-5','OUT-5',
  'Total Hours','Daily Salary (Rs)','Cumulative Salary (Rs)','GPS'
];
// Column indices (1-based):
//  A=1  DATE
//  B=2  IN-1     C=3  OUT-1
//  D=4  IN-2     E=5  OUT-2
//  F=6  IN-3     G=7  OUT-3
//  H=8  IN-4     I=9  OUT-4
//  J=10 IN-5     K=11 OUT-5
//  L=12 Total Hours
//  M=13 Daily Salary
//  N=14 Cumulative Salary
//  O=15 GPS audit log

// ╔══════════════════════════════════════════════════════════════════╗
//   MENU — appears at the top of the spreadsheet on every open
// ╚══════════════════════════════════════════════════════════════════╝
function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu('IBI Tools')
    .addItem('🆕  Setup Current Month',      'menuSetupCurrentMonth')
    .addItem('🔄  Recalculate Current Month','menuRecalcCurrentMonth')
    .addItem('🔁  Recalculate All Months',   'menuRecalcAllMonths')
    .addSeparator()
    .addItem('🧹  Fix Duplicate Date Rows',  'menuFixDuplicateRows')
    .addItem('🔍  Scan for Repeated Punches','menuScanRepeatedPunches')
    .addItem('🧽  Remove Repeated Punches',  'menuRemoveRepeatedPunches')
    .addItem('📦  Migrate Legacy Tabs',      'menuMigrateLegacy')
    .addItem('ℹ️  About',                    'menuAbout')
    .addToUi();
}

// ╔══════════════════════════════════════════════════════════════════╗
//   onEdit — fires automatically on EVERY cell edit
//   If the user manually edits a date or time in a staff monthly tab,
//   the salary columns (L, M, N) auto-recalculate instantly.
// ╚══════════════════════════════════════════════════════════════════╝
function onEdit(e) {
  try {
    if (!e || !e.range) return;

    var sheet     = e.range.getSheet();
    var sheetName = sheet.getName();

    // Only act on staff monthly tabs (e.g. "S. Aswin - May 2026")
    var isStaffTab = STAFF_NAMES.some(function(s) { return sheetName.indexOf(s) === 0; });
    if (!isStaffTab) return;

    // Ignore edits to header row
    if (e.range.getRow() <= HEADER_ROWS) return;

    // Only react to edits in columns A through K (date + 5 IN/OUT pairs)
    // Columns: A=1 (DATE), B=2 (IN-1) ... K=11 (OUT-5)
    var col = e.range.getColumn();
    if (col < 1 || col > 11) return;

    // Don't loop on our own writes (recalcSalary writes to columns L, M, N
    // which are filtered out above, so this is safe)
    recalcSalary(sheet);

  } catch (err) {
    // Silently swallow — we don't want to break the user's edit experience
    // with popup errors. Errors are logged instead.
    console.error('onEdit error: ' + err.toString());
  }
}

// Cleans up sheets that have multiple rows for the same date
// (caused by the locale bug). Merges punch times into a single row,
// keeps the earliest date row, deletes the extras.
function menuFixDuplicateRows() {
  var ss     = SpreadsheetApp.getActiveSpreadsheet();
  var sheets = ss.getSheets();
  var report = [];

  sheets.forEach(function(sh) {
    var name = sh.getName();
    var isStaffTab = STAFF_NAMES.some(function(s) { return name.indexOf(s) === 0; });
    if (!isStaffTab) return;

    var merged = mergeDuplicateRowsInSheet(sh);
    if (merged > 0) {
      ensureHeaders(sh);
      recalcSalary(sh);
      report.push(name + ': merged ' + merged + ' duplicate row(s)');
    }
  });

  var msg = report.length
    ? '✓ Cleanup done:\n  • ' + report.join('\n  • ')
    : 'No duplicate rows found.';
  SpreadsheetApp.getUi().alert('Fix Duplicate Date Rows', msg, SpreadsheetApp.getUi().ButtonSet.OK);
}

// Returns count of duplicate rows merged.
// Algorithm: group rows by Y/M/D, keep the first row, merge all
// non-empty IN/OUT cells from later rows into it (in original order),
// delete the now-empty later rows from bottom up.
function mergeDuplicateRowsInSheet(sheet) {
  var lastRow = sheet.getLastRow();
  if (lastRow < 3) return 0;

  var data    = sheet.getRange(2, 1, lastRow - 1, 11).getValues();   // A..K
  var groups  = {};   // dateKey → array of row indices (0-based within data)
  var order   = [];   // dateKey insertion order

  for (var i = 0; i < data.length; i++) {
    var cell = data[i][0];
    if (cell === '' || cell === null) continue;

    var key = null;
    if (cell instanceof Date) {
      key = Utilities.formatDate(cell, TZ, 'yyyy-MM-dd');
    } else if (typeof cell === 'string') {
      var s  = cell.trim();
      var m1 = s.match(/^(\d{1,2})[-\/](\d{1,2})[-\/](\d{4})$/);
      if (m1) {
        key = m1[3] + '-' + ('0'+m1[2]).slice(-2) + '-' + ('0'+m1[1]).slice(-2);
      }
    }
    if (!key) continue;

    if (!groups[key]) { groups[key] = []; order.push(key); }
    groups[key].push(i);
  }

  // Find duplicates and merge
  var rowsToDelete = [];
  var mergedCount  = 0;

  order.forEach(function(key) {
    var rows = groups[key];
    if (rows.length < 2) return;

    // Collect all non-empty IN/OUT times across all duplicate rows, in order
    var times = [];
    rows.forEach(function(rIdx) {
      for (var c = 1; c <= 10; c++) {
        var v = data[rIdx][c];
        if (v !== '' && v !== null) times.push(v);
      }
    });

    // Cap at 10 (max 5 IN/OUT pairs)
    times = times.slice(0, 10);

    // Write merged times to first row
    var firstSheetRow = rows[0] + 2;   // +2 because data starts at sheet row 2
    var rowOut = times.slice();
    while (rowOut.length < 10) rowOut.push('');
    sheet.getRange(firstSheetRow, 2, 1, 10).setValues([rowOut]);

    // Schedule the rest for deletion
    for (var k = 1; k < rows.length; k++) {
      rowsToDelete.push(rows[k] + 2);   // sheet row number
    }
    mergedCount += (rows.length - 1);
  });

  // Delete duplicate rows from BOTTOM up so indices don't shift
  rowsToDelete.sort(function(a, b) { return b - a; });
  rowsToDelete.forEach(function(sr) { sheet.deleteRow(sr); });

  return mergedCount;
}

// ── MENU HANDLERS ────────────────────────────────────────────────

function menuSetupCurrentMonth() {
  var now    = new Date();
  var made   = [];
  var exists = [];
  STAFF_NAMES.forEach(function(name) {
    var sheetName = monthlyTabName(name, now);
    var ss        = SpreadsheetApp.getActiveSpreadsheet();
    var s         = ss.getSheetByName(sheetName);
    if (s) { exists.push(sheetName); }
    else   { getOrCreateMonthlySheet(name, now); made.push(sheetName); }
  });
  var msg = '';
  if (made.length)   msg += '✓ Created:\n  • ' + made.join('\n  • ') + '\n\n';
  if (exists.length) msg += 'Already existed:\n  • ' + exists.join('\n  • ');
  SpreadsheetApp.getUi().alert('Setup Current Month', msg, SpreadsheetApp.getUi().ButtonSet.OK);
}

function menuRecalcCurrentMonth() {
  var now  = new Date();
  var done = [];
  var miss = [];
  STAFF_NAMES.forEach(function(name) {
    var sheetName = monthlyTabName(name, now);
    var s = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(sheetName);
    if (s) { recalcSalary(s);  done.push(sheetName); }
    else   { miss.push(sheetName); }
  });
  var msg = '';
  if (done.length) msg += '✓ Recalculated:\n  • ' + done.join('\n  • ') + '\n\n';
  if (miss.length) msg += 'Not found:\n  • ' + miss.join('\n  • ');
  SpreadsheetApp.getUi().alert('Recalculate Current Month', msg, SpreadsheetApp.getUi().ButtonSet.OK);
}

function menuRecalcAllMonths() {
  var ss      = SpreadsheetApp.getActiveSpreadsheet();
  var sheets  = ss.getSheets();
  var done    = [];
  sheets.forEach(function(sh) {
    var n = sh.getName();
    // Match any tab beginning with one of our staff names
    var isStaffTab = STAFF_NAMES.some(function(name) { return n.indexOf(name) === 0; });
    if (isStaffTab) {
      // Make sure headers exist before recalculating
      ensureHeaders(sh);
      recalcSalary(sh);
      done.push(n);
    }
  });
  var msg = done.length
    ? '✓ Recalculated ' + done.length + ' tabs:\n  • ' + done.join('\n  • ')
    : 'No staff tabs found.';
  SpreadsheetApp.getUi().alert('Recalculate All Months', msg, SpreadsheetApp.getUi().ButtonSet.OK);
}

function menuMigrateLegacy() {
  var ss      = SpreadsheetApp.getActiveSpreadsheet();
  var renamed = [];
  var skipped = [];

  STAFF_NAMES.forEach(function(name) {
    var legacy = ss.getSheetByName(name);
    if (!legacy) return;

    // Look at column A row 2 to find the dominant month/year of this tab's data
    var d = detectDominantMonth(legacy);
    if (!d) { skipped.push(name + ' (no data found)'); return; }

    var newName = name + ' - ' + monthName(d.getMonth()) + ' ' + d.getFullYear();
    if (ss.getSheetByName(newName)) {
      skipped.push(name + ' (target "' + newName + '" already exists)');
      return;
    }
    legacy.setName(newName);
    ensureHeaders(legacy);
    recalcSalary(legacy);
    renamed.push(name + '  →  ' + newName);
  });

  var msg = '';
  if (renamed.length) msg += '✓ Renamed:\n  • ' + renamed.join('\n  • ') + '\n\n';
  if (skipped.length) msg += 'Skipped:\n  • ' + skipped.join('\n  • ');
  if (!renamed.length && !skipped.length) msg = 'No legacy tabs found — nothing to migrate.';
  SpreadsheetApp.getUi().alert('Migrate Legacy Tabs', msg, SpreadsheetApp.getUi().ButtonSet.OK);
}

// ╔══════════════════════════════════════════════════════════════════╗
//   REPEATED-PUNCH REPAIR (v2.3)
//
//   Cleans up the mis-taps the old build let through: a staff member taps
//   OUT, thinks nothing happened, and taps IN again seconds later. That
//   alternates correctly, so v2.0 accepted it and split one real shift into
//   two — the repeated entries visible in the sheet.
//
//   A "cluster" is a run of consecutive punches each landing within
//   MIN_PUNCH_GAP_SEC of the one before it. Punches must stay strictly
//   IN,OUT,IN,OUT..., so an even number must always be dropped:
//     • even cluster {OUT,IN}        → drop all  (the shift never really broke)
//     • odd  cluster {IN,OUT,IN}     → keep the first, drop the rest
//   The GPS audit log in column O is deliberately left untouched, so the
//   original taps remain on record even after the times are collapsed.
// ╚══════════════════════════════════════════════════════════════════╝

function menuScanRepeatedPunches()   { repeatedPunchTool(false); }

function menuRemoveRepeatedPunches() {
  var ui = SpreadsheetApp.getUi();
  var ok = ui.alert(
    'Remove Repeated Punches',
    'This rewrites the punch times on affected rows and recalculates salary.\n\n' +
    'Run "Scan for Repeated Punches" first and check the list.\n' +
    'The GPS audit log (column O) is kept intact.\n\nProceed?',
    ui.ButtonSet.YES_NO
  );
  if (ok !== ui.Button.YES) return;
  repeatedPunchTool(true);
}

function repeatedPunchTool(apply) {
  var ss     = SpreadsheetApp.getActiveSpreadsheet();
  var report = [];
  var fixed  = 0;
  var touchedSheets = [];

  ss.getSheets().forEach(function(sh) {
    var name = sh.getName();
    if (!STAFF_NAMES.some(function(s) { return name.indexOf(s) === 0; })) return;

    var last = sh.getLastRow();
    if (last <= HEADER_ROWS) return;

    var values      = sh.getRange(HEADER_ROWS + 1, 1, last - HEADER_ROWS, 11).getValues();
    var sheetTouched = false;

    for (var r = 0; r < values.length; r++) {
      var rowVals = values[r];
      if (rowVals[0] === '' || rowVals[0] == null) continue;

      // Collect this day's punches in order
      var punches = [];
      for (var c = 1; c <= 10; c++) {
        if (rowVals[c] === '' || rowVals[c] == null) break;
        punches.push({ text: punchCellToText(rowVals[c]), sec: punchCellToSecOfDay(rowVals[c]) });
      }
      if (punches.length < 2) continue;

      // Mark punches to drop, cluster by cluster
      var drop = [];
      for (var i = 0; i < punches.length; i++) drop.push(false);

      var i2 = 0;
      while (i2 < punches.length) {
        var end = i2;   // last index still inside this cluster
        while (end + 1 < punches.length &&
               punches[end].sec !== null && punches[end + 1].sec !== null &&
               (punches[end + 1].sec - punches[end].sec) < MIN_PUNCH_GAP_SEC) {
          end++;
        }
        var size = end - i2 + 1;
        if (size > 1) {
          // even cluster → drop all; odd cluster → keep the first
          var from = (size % 2 === 0) ? i2 : i2 + 1;
          for (var k = from; k <= end; k++) drop[k] = true;
        }
        i2 = end + 1;
      }

      var kept    = [];
      var removed = [];
      for (var j = 0; j < punches.length; j++) {
        if (drop[j]) removed.push(punches[j].text);
        else         kept.push(punches[j].text);
      }
      if (!removed.length) continue;

      var dateTxt = punchDateText(rowVals[0]);
      report.push('• ' + name + ' · ' + dateTxt +
                  '\n    was:     ' + punches.map(function(p) { return p.text; }).join('  ') +
                  '\n    becomes: ' + (kept.length ? kept.join('  ') : '(no punches left)') +
                  '\n    removed: ' + removed.join('  '));
      fixed++;

      if (apply) {
        var out = [];
        for (var w = 0; w < 10; w++) out.push(w < kept.length ? kept[w] : '');
        sh.getRange(HEADER_ROWS + 1 + r, 2, 1, 10).setValues([out]);
        sheetTouched = true;
      }
    }

    if (apply && sheetTouched) { recalcSalary(sh); touchedSheets.push(name); }
  });

  var ui    = SpreadsheetApp.getUi();
  var title = apply ? 'Repeated Punches Removed' : 'Scan for Repeated Punches';

  if (!fixed) {
    ui.alert(title, 'No repeated punches found.\n\nAny two punches closer together than ' +
                    MIN_PUNCH_GAP_SEC + ' seconds would be flagged here.', ui.ButtonSet.OK);
    return;
  }

  var head = apply
    ? '✓ Cleaned ' + fixed + ' day(s).\nSalary recalculated on: ' + touchedSheets.join(', ') + '\n\n'
    : 'Found ' + fixed + ' day(s) with repeated punches.\nNothing has been changed — run ' +
      '"Remove Repeated Punches" to apply.\n\n';

  // Alert bodies get truncated, so keep the dialog short and log the rest
  var body = report.join('\n\n');
  Logger.log(head + body);
  ui.alert(title, head + (body.length > 1200 ? body.slice(0, 1200) + '\n\n…full list in View → Logs.' : body), ui.ButtonSet.OK);
}

// Date cell -> "dd-MM-yyyy" text for the repair report
function punchDateText(cell) {
  if (cell instanceof Date) return Utilities.formatDate(cell, TZ, DATE_FORMAT);
  return String(cell == null ? '' : cell).trim();
}

function menuAbout() {
  SpreadsheetApp.getUi().alert(
    'IBI Staff Attendance · Month-Wise Edition\n' +
    '──────────────────────────────────────\n' +
    'Version: ' + APP_VERSION + '\n' +
    'Hourly rate: Rs ' + HOURLY_RATE + '\n' +
    'Staff: ' + STAFF_NAMES.join(', ') + '\n' +
    'Time zone: ' + TZ + '\n' +
    'Geofence: ' + MAX_DIST_M + 'm · max GPS error ' + MAX_ACC_M + 'm (server-enforced)\n' +
    'Min gap between punches: ' + MIN_PUNCH_GAP_SEC + 's (server-enforced)\n\n' +
    'How it works:\n' +
    '  • Every staff member gets a fresh tab each month\n' +
    '  • Tabs are auto-created on the first punch of the month\n' +
    '  • Punches are refused server-side unless the phone is inside\n' +
    '    the warehouse geofence with a good GPS fix (anti-tamper)\n' +
    '  • Headers, colors (#00c5ff on black), Roboto font, salary\n' +
    '    formulas — all applied automatically\n' +
    '  • Cumulative salary resets to 0 at the start of each month\n' +
    '  • Salary feed (getStaffSalaries) is pulled by the IBI ERP System'
  );
}

// ╔══════════════════════════════════════════════════════════════════╗
//   MONTHLY SHEET LIFECYCLE
// ╚══════════════════════════════════════════════════════════════════╝

function monthName(monthIndex) {
  return ['January','February','March','April','May','June',
          'July','August','September','October','November','December'][monthIndex];
}

function monthlyTabName(staffName, date) {
  return staffName + ' - ' + monthName(date.getMonth()) + ' ' + date.getFullYear();
}

// Returns the sheet for this staff/date, creating it (with headers + formatting) if needed
function getOrCreateMonthlySheet(staffName, date) {
  var ss        = SpreadsheetApp.getActiveSpreadsheet();
  var sheetName = monthlyTabName(staffName, date);
  var sheet     = ss.getSheetByName(sheetName);
  if (sheet) {
    ensureHeaders(sheet);   // safety: re-apply headers if they were deleted
    return sheet;
  }
  // Create brand new sheet
  sheet = ss.insertSheet(sheetName);
  setupSheetFormatting(sheet);
  return sheet;
}

// Apply full formatting to a freshly created sheet
function setupSheetFormatting(sheet) {
  // ── Headers ───────────────────────────────────────
  sheet.getRange(1, 1, 1, COLUMN_HEADERS.length).setValues([COLUMN_HEADERS]);

  var hdr = sheet.getRange(1, 1, 1, COLUMN_HEADERS.length);
  hdr.setFontFamily('Roboto');
  hdr.setFontWeight('bold');
  hdr.setFontSize(11);
  hdr.setBackground('#000000');
  hdr.setFontColor('#00c5ff');
  hdr.setHorizontalAlignment('center');
  hdr.setVerticalAlignment('middle');
  sheet.setRowHeight(1, 34);
  sheet.setFrozenRows(1);

  // ── Column widths ─────────────────────────────────
  sheet.setColumnWidth(1,  110);   // DATE
  for (var c = 2; c <= 11; c++) sheet.setColumnWidth(c, 95);   // IN/OUT × 5
  sheet.setColumnWidth(12, 100);   // Total Hours
  sheet.setColumnWidth(13, 130);   // Daily Salary
  sheet.setColumnWidth(14, 170);   // Cumulative Salary
  sheet.setColumnWidth(15, 420);   // GPS

  // ── Body styling (apply to ALL rows for future writes) ──
  var maxRow = sheet.getMaxRows();
  if (maxRow > 1) {
    var body = sheet.getRange(2, 1, maxRow - 1, COLUMN_HEADERS.length);
    body.setFontFamily('Roboto');
    body.setFontSize(10);
    body.setHorizontalAlignment('center');
    body.setVerticalAlignment('middle');

    // Column A — date format dd-MM-yyyy (locale-immune display)
    sheet.getRange(2, 1, maxRow - 1, 1).setNumberFormat('dd-MM-yyyy');

    // Salary numeric format — 2 decimals
    sheet.getRange(2, 12, maxRow - 1, 3).setNumberFormat('0.00');

    // GPS column left-aligned (it's long text)
    sheet.getRange(2, 15, maxRow - 1, 1).setHorizontalAlignment('left');
  }

  // ── Light alternating row colors for readability ──
  // (skip — keep clean white background)

  // ── Hide unused columns past the GPS column ───────
  if (sheet.getMaxColumns() > 15) {
    sheet.hideColumns(16, sheet.getMaxColumns() - 15);
  }
}

// Re-apply only the headers (safe to call repeatedly)
function ensureHeaders(sheet) {
  var existing = sheet.getRange(1, 1, 1, COLUMN_HEADERS.length).getValues()[0];
  var needsFix = false;
  for (var i = 0; i < COLUMN_HEADERS.length; i++) {
    if (existing[i] !== COLUMN_HEADERS[i]) { needsFix = true; break; }
  }
  if (needsFix) {
    sheet.getRange(1, 1, 1, COLUMN_HEADERS.length).setValues([COLUMN_HEADERS]);
    var hdr = sheet.getRange(1, 1, 1, COLUMN_HEADERS.length);
    hdr.setFontFamily('Roboto');
    hdr.setFontWeight('bold');
    hdr.setFontSize(11);
    hdr.setBackground('#000000');
    hdr.setFontColor('#00c5ff');
    hdr.setHorizontalAlignment('center');
    hdr.setVerticalAlignment('middle');
    sheet.setRowHeight(1, 34);
    sheet.setFrozenRows(1);
  }
}

// Find the most common month/year among existing dates (for migration)
function detectDominantMonth(sheet) {
  var last = sheet.getLastRow();
  if (last < 2) return null;
  var values = sheet.getRange(2, 1, last - 1, 1).getValues();
  var counts = {};
  for (var i = 0; i < values.length; i++) {
    var d = parseSheetDate(values[i][0]);
    if (!d) continue;
    var key = d.getFullYear() + '-' + d.getMonth();
    counts[key] = (counts[key] || 0) + 1;
  }
  var bestKey = null, bestCount = 0;
  for (var k in counts) {
    if (counts[k] > bestCount) { bestCount = counts[k]; bestKey = k; }
  }
  if (!bestKey) return null;
  var parts = bestKey.split('-');
  return new Date(parseInt(parts[0]), parseInt(parts[1]), 1);
}

// ╔══════════════════════════════════════════════════════════════════╗
//   SALARY RECALCULATION
//   Reads all data rows, computes Total Hours / Daily / Cumulative,
//   writes everything in ONE batch call.
// ╚══════════════════════════════════════════════════════════════════╝
function recalcSalary(sheet) {
  var lastRow = sheet.getLastRow();
  if (lastRow < 2) return;

  var data    = sheet.getRange(2, 1, lastRow - 1, 11).getValues();   // A..K
  var output  = [];
  var cumSal  = 0;

  for (var i = 0; i < data.length; i++) {
    var row  = data[i];
    var date = row[0];

    if (date === '' || date === null) {
      output.push(['', '', '']);
      continue;
    }

    // Compute total hours by summing the 5 IN/OUT pairs
    var pairs      = [[1,2], [3,4], [5,6], [7,8], [9,10]];
    var totalHours = 0;
    for (var p = 0; p < pairs.length; p++) {
      var inV  = toDecimalDay(row[pairs[p][0]]);
      var outV = toDecimalDay(row[pairs[p][1]]);
      if (inV !== null && outV !== null && outV > inV) {
        totalHours += (outV - inV) * 24;
      }
    }
    totalHours = Math.round(totalHours * 100) / 100;

    var daily = Math.round(totalHours * HOURLY_RATE * 100) / 100;
    cumSal    = Math.round((cumSal + daily) * 100) / 100;

    output.push([totalHours, daily, cumSal]);
  }

  // Single batch write
  sheet.getRange(2, 12, output.length, 3).setValues(output);
  sheet.getRange(2, 12, output.length, 3).setNumberFormat('0.00');
}

// ╔══════════════════════════════════════════════════════════════════╗
//   TIME PARSER
//   Accepts Date objects, decimal-day numbers, or text like
//   "08:51:00 AM" / "8:51 AM" / "13:45"
// ╚══════════════════════════════════════════════════════════════════╝
function toDecimalDay(val) {
  if (val === '' || val === null || val === undefined) return null;
  if (val instanceof Date) {
    return (val.getHours() * 3600 + val.getMinutes() * 60 + val.getSeconds()) / 86400;
  }
  if (typeof val === 'number') {
    return (val >= 0 && val < 2) ? val : (val - Math.floor(val));
  }
  if (typeof val === 'string') {
    var s  = val.trim().toUpperCase();
    var m1 = s.match(/^(\d{1,2}):(\d{2})(?::(\d{2}))?\s*(AM|PM)$/);
    if (m1) {
      var h = parseInt(m1[1]), mi = parseInt(m1[2]), se = m1[3] ? parseInt(m1[3]) : 0;
      if (m1[4] === 'PM' && h !== 12) h += 12;
      if (m1[4] === 'AM' && h === 12) h = 0;
      return (h * 3600 + mi * 60 + se) / 86400;
    }
    var m2 = s.match(/^(\d{1,2}):(\d{2})(?::(\d{2}))?$/);
    if (m2) {
      var h = parseInt(m2[1]), mi = parseInt(m2[2]), se = m2[3] ? parseInt(m2[3]) : 0;
      return (h * 3600 + mi * 60 + se) / 86400;
    }
  }
  return null;
}

// Parses a date cell value into a JS Date (for reading column A)
function parseSheetDate(val) {
  if (val instanceof Date) return val;
  if (typeof val === 'string' && val.trim() !== '') {
    var m = val.trim().match(/^(\d{1,2})-(\d{1,2})-(\d{4})$/);
    if (m) return new Date(parseInt(m[3]), parseInt(m[2]) - 1, parseInt(m[1]));
  }
  return null;
}

// ╔══════════════════════════════════════════════════════════════════╗
//   ERP INTEGRATION — Staff salary feed (read-only)
//   Returns ONE record per staff member per month, summed straight from
//   the monthly tabs. The IBI ERP System pulls this server-to-server and
//   books it as a de-duplicated "Salary" expense (keyed by staff+month),
//   so re-syncs never double-count.
//   Response: { status:'success', ok:true, records:[...], total:N }
// ╚══════════════════════════════════════════════════════════════════╝
function getStaffSalariesFeed() {
  var ss      = SpreadsheetApp.getActiveSpreadsheet();
  var sheets  = ss.getSheets();
  var records = [];

  sheets.forEach(function(sh) {
    var name  = sh.getName();
    // Match tabs named "<Staff Name> - <Month> <Year>"  e.g. "S. Aswin - May 2026"
    var staff = null;
    for (var s = 0; s < STAFF_NAMES.length; s++) {
      if (name.indexOf(STAFF_NAMES[s] + ' - ') === 0) { staff = STAFF_NAMES[s]; break; }
    }
    if (!staff) return;

    var monthLabel = name.substring(staff.length + 3);   // text after " - "  ->  "May 2026"

    ensureHeaders(sh);
    recalcSalary(sh);                                    // make sure L/M/N are current

    var lastRow = sh.getLastRow();
    if (lastRow < 2) return;

    // Columns A..N  (DATE=1 … Total Hours=12[L] · Daily=13[M] · Cumulative=14[N])
    var data        = sh.getRange(2, 1, lastRow - 1, 14).getValues();
    var days        = 0;
    var totalHours  = 0;
    var totalSalary = 0;
    var firstDate   = null, lastDate = null;

    for (var i = 0; i < data.length; i++) {
      var d = data[i][0];
      if (d === '' || d === null) continue;

      var dateStr = null;
      if (d instanceof Date) {
        dateStr = Utilities.formatDate(d, TZ, 'yyyy-MM-dd');
      } else if (typeof d === 'string') {
        var m = d.trim().match(/^(\d{1,2})[-\/](\d{1,2})[-\/](\d{4})$/);
        if (m) dateStr = m[3] + '-' + ('0' + m[2]).slice(-2) + '-' + ('0' + m[1]).slice(-2);
      }

      var th = parseFloat(data[i][11]);   // Total Hours  (col L)
      var dl = parseFloat(data[i][12]);   // Daily Salary (col M)
      if (!isNaN(th) && th > 0) days++;
      if (!isNaN(th)) totalHours  += th;
      if (!isNaN(dl)) totalSalary += dl;

      if (dateStr) { if (!firstDate) firstDate = dateStr; lastDate = dateStr; }
    }

    var monthKey = lastDate ? lastDate.slice(0, 7) : monthLabelToKey(monthLabel);

    records.push({
      staffName:   staff,
      monthLabel:  monthLabel,
      monthKey:    monthKey,                              // 'YYYY-MM' (local, never toISOString)
      daysWorked:  days,
      totalHours:  Math.round(totalHours * 100) / 100,
      ratePerHour: HOURLY_RATE,
      totalSalary: Math.round(totalSalary * 100) / 100,
      firstDate:   firstDate,
      lastDate:    lastDate
    });
  });

  return jsonOut({ status: 'success', ok: true, records: records, total: records.length });
}

// "May 2026" -> "2026-05"  (fallback month key when a tab has no date rows yet)
function monthLabelToKey(label) {
  var parts = String(label).trim().split(/\s+/);
  if (parts.length < 2) return '';
  var months = ['january','february','march','april','may','june',
                'july','august','september','october','november','december'];
  var mi = months.indexOf(parts[0].toLowerCase());
  if (mi === -1) return '';
  return parts[1] + '-' + ('0' + (mi + 1)).slice(-2);
}

// ╔══════════════════════════════════════════════════════════════════╗
//   doGet — returns today's state and punch list
// ╚══════════════════════════════════════════════════════════════════╝
function doGet(e) {
  try {
    if (!checkToken(e.parameter.token)) {
      return jsonOut({ status: 'error', message: 'Unauthorized' });
    }
    var action = (e.parameter.action || '');

    // ── ERP salary feed: every staff member, every month (read-only) ──
    if (action === 'getStaffSalaries') {
      return getStaffSalariesFeed();
    }

    if (action !== 'getState') {
      return jsonOut({ status: 'error', message: 'Unknown action' });
    }
    var staffName = (e.parameter.staffName || '').trim();
    if (!staffName) return jsonOut({ status: 'error', message: 'staffName is required' });
    if (STAFF_NAMES.indexOf(staffName) === -1) {
      return jsonOut({ status: 'error', message: 'Unknown staff member.' });
    }

    var now   = new Date();
    var sheet = getOrCreateMonthlySheet(staffName, now);

    // ── Locale-immune date components (in IST) ──────
    var todayY = parseInt(Utilities.formatDate(now, TZ, 'yyyy'));
    var todayM = parseInt(Utilities.formatDate(now, TZ, 'MM'));
    var todayD = parseInt(Utilities.formatDate(now, TZ, 'dd'));

    var data = sheet.getDataRange().getValues();
    var ri   = findTodayRow(data, todayY, todayM, todayD);

    if (ri === -1) {
      return jsonOut({ status: 'success', currentState: 'OUT', punches: [], punchCount: 0 });
    }

    var row     = data[ri];
    var labels  = ['IN','OUT','IN','OUT','IN','OUT','IN','OUT','IN','OUT'];
    var punches = [];
    for (var c = 1; c <= 10; c++) {
      var cell = row[c];
      if (cell === '' || cell == null) break;
      var t = (cell instanceof Date)
        ? Utilities.formatDate(cell, TZ, TIME_FORMAT)
        : cell.toString().trim();
      punches.push({ type: labels[c - 1], time: t });
    }

    var currentState = 'OUT';
    if (punches.length > 0) {
      currentState = punches[punches.length - 1].type === 'IN' ? 'IN' : 'OUT';
    }
    return jsonOut({
      status: 'success', currentState: currentState,
      punches: punches, punchCount: punches.length,
    });
  } catch (err) {
    return jsonOut({ status: 'error', message: 'doGet error: ' + err.toString() });
  }
}

// ╔══════════════════════════════════════════════════════════════════╗
//   doPost — records a punch into the correct monthly sheet
// ╚══════════════════════════════════════════════════════════════════╝
function doPost(e) {
  // ── Serialise every write (v2.3) ──────────────────────────────
  //   Without a lock, two punches arriving together both read the sheet
  //   BEFORE either writes, so both compute the same "next empty" target —
  //   or, if the second read lands just after the first write, it appends a
  //   SECOND row for the same date. The lock makes the read-decide-write
  //   sequence atomic, which is the only way the checks below can hold.
  var lock = LockService.getScriptLock();
  try {
    lock.waitLock(LOCK_WAIT_MS);
  } catch (lockErr) {
    return jsonOut({ status: 'error', message: 'Server busy — please tap again in a moment.' });
  }
  try {
    return doPostLocked(e);
  } finally {
    try { lock.releaseLock(); } catch (ignore) {}
  }
}

function doPostLocked(e) {
  try {
    if (!checkToken(e.parameter.token)) {
      return jsonOut({ status: 'error', message: 'Unauthorized' });
    }

    var staffName = (e.parameter.staffName || '').trim();
    var action    = (e.parameter.action    || '').trim().toUpperCase();
    var lat       = e.parameter.lat      || '';
    var lon       = e.parameter.lon      || '';
    var accuracy  = e.parameter.accuracy || '';

    if (!staffName || (action !== 'IN' && action !== 'OUT')) {
      return jsonOut({ status: 'error', message: 'Invalid parameters.' });
    }

    // ── Allow-list: only known staff may punch (blocks junk tabs) ──
    if (STAFF_NAMES.indexOf(staffName) === -1) {
      return jsonOut({ status: 'error', message: 'Unknown staff member.' });
    }

    // ── Idempotent retry guard (v2.3) ─────────────────────────────
    //   The phone queues a punch whenever the fetch rejects — but a fetch also
    //   rejects when the request DID reach us and only the reply was lost on a
    //   flaky mobile connection. Replaying that punch used to write a second
    //   entry. The client now stamps each punch with a punchId, so a replay of
    //   one already recorded is answered "success" without touching the sheet.
    var punchId = (e.parameter.punchId || '').trim();
    var cache   = CacheService.getScriptCache();
    if (punchId) {
      if (cache.get('punch_' + punchId)) {
        return jsonOut({ status: 'success', duplicate: true, message: 'Punch already recorded.' });
      }
    }

    // ── Stale-queue guard (v2.3) ──────────────────────────────────
    //   A queued punch that never got through must not surface hours later and
    //   land in a completely different shift slot. Past this age it is dropped.
    var clientTs = parseInt(e.parameter.clientTs || '', 10);
    if (!isNaN(clientTs) && clientTs > 0) {
      var ageSec = (new Date().getTime() - clientTs) / 1000;
      if (ageSec > STALE_PUNCH_SEC) {
        return jsonOut({
          status: 'error', stale: true,
          message: 'This queued punch is ' + Math.round(ageSec / 60) + ' minutes old and was discarded. Please punch again.'
        });
      }
    }

    // ── AUTHORITATIVE geofence + accuracy gate (hard reject) ──
    //   Refuse the punch unless the phone is really at the warehouse with a
    //   good-enough GPS fix. The browser geofence in index.html is only UX;
    //   this server check is what actually blocks home / API-bypass punches.
    var latN = parseFloat(lat), lonN = parseFloat(lon), accN = parseFloat(accuracy);
    if (isNaN(latN) || isNaN(lonN)) {
      return jsonOut({ status: 'error', message: 'Location required to punch.' });
    }
    if (isNaN(accN) || accN > MAX_ACC_M) {
      return jsonOut({ status: 'error', message: 'GPS accuracy too low (±' + (isNaN(accN) ? '?' : Math.round(accN)) + 'm). Move to an open area and retry.' });
    }
    var distM = havM(latN, lonN, WH_LAT, WH_LON);
    if (distM > MAX_DIST_M) {
      return jsonOut({ status: 'error', message: 'Outside warehouse zone (' + Math.round(distM) + 'm away). Punch not recorded.' });
    }

    var now   = new Date();
    var sheet = getOrCreateMonthlySheet(staffName, now);

    // ── Locale-immune date components (in IST) ──────
    var todayY = parseInt(Utilities.formatDate(now, TZ, 'yyyy'));
    var todayM = parseInt(Utilities.formatDate(now, TZ, 'MM'));
    var todayD = parseInt(Utilities.formatDate(now, TZ, 'dd'));

    var timeNow = Utilities.formatDate(now, TZ, TIME_FORMAT);
    var data    = sheet.getDataRange().getValues();
    var ri      = findTodayRow(data, todayY, todayM, todayD);

    // ── New day in this monthly tab ─────────────────
    if (ri === -1) {
      if (action === 'OUT') {
        return jsonOut({ status: 'error', message: 'Cannot punch OUT — no punch IN recorded today.' });
      }
      var newRow = nextEmptyRow(sheet);

      // Write date as a proper Date object (locale-immune) with an
      // explicit dd-MM-yyyy display format on the cell
      var dateOnly = new Date(todayY, todayM - 1, todayD);  // local midnight
      sheet.getRange(newRow, 1)
           .setValue(dateOnly)
           .setNumberFormat('dd-MM-yyyy');

      sheet.getRange(newRow, 2).setValue(timeNow);
      logGPS(sheet, newRow, action, lat, lon, accuracy, now);
      recalcSalary(sheet);
      if (punchId) cache.put('punch_' + punchId, '1', REPLAY_CACHE_SEC);
      return jsonOut({ status: 'success' });
    }

    // ── Existing row: anti-double-punch firewall ────
    var row         = data[ri];
    var emptyColIdx = -1;
    for (var c = 1; c <= 10; c++) {
      if (row[c] === '' || row[c] == null) { emptyColIdx = c; break; }
    }
    if (emptyColIdx === -1) {
      return jsonOut({ status: 'error', message: 'Maximum shifts (5) reached for today.' });
    }

    var expectIN = (emptyColIdx % 2 === 1);
    if (expectIN && action === 'OUT') {
      return jsonOut({ status: 'error', message: 'You are currently clocked OUT. Cannot punch OUT again.' });
    }
    if (!expectIN && action === 'IN') {
      return jsonOut({ status: 'error', message: 'You are currently clocked IN. Cannot punch IN again.' });
    }

    // ── Minimum-gap guard (v2.3) ───────────────────────────────────
    //   Alternation alone is not enough: a re-tap seconds after a real punch
    //   alternates correctly and so used to be accepted, opening a spurious
    //   shift. Anything closer together than MIN_PUNCH_GAP_SEC is a mis-tap.
    //   Both sides are reduced to seconds-since-IST-midnight, so no Date is
    //   rebuilt from a formatted string and the comparison cannot drift.
    var prevSec = punchCellToSecOfDay(row[emptyColIdx - 1]);
    var nowSec  = hmsToSecOfDay(Utilities.formatDate(now, TZ, 'HH:mm:ss'));
    if (prevSec !== null && nowSec !== null) {
      var gapSec = nowSec - prevSec;
      if (gapSec >= 0 && gapSec < MIN_PUNCH_GAP_SEC) {
        var waitSec = Math.ceil(MIN_PUNCH_GAP_SEC - gapSec);
        return jsonOut({
          status: 'error', tooSoon: true, retryAfterSec: waitSec,
          message: 'Your ' + (expectIN ? 'OUT' : 'IN') + ' punch at ' +
                   punchCellToText(row[emptyColIdx - 1]) + ' is already recorded. ' +
                   'Please wait ' + waitSec + ' more second' + (waitSec === 1 ? '' : 's') + ' before punching again.'
        });
      }
    }

    var sheetRow = ri + 1;
    sheet.getRange(sheetRow, 1 + emptyColIdx).setValue(timeNow);
    logGPS(sheet, sheetRow, action, lat, lon, accuracy, now);
    recalcSalary(sheet);
    if (punchId) cache.put('punch_' + punchId, '1', REPLAY_CACHE_SEC);
    return jsonOut({ status: 'success' });

  } catch (err) {
    return jsonOut({ status: 'error', message: 'doPost error: ' + err.toString() });
  }
}

// ╔══════════════════════════════════════════════════════════════════╗
//   HELPERS
// ╚══════════════════════════════════════════════════════════════════╝

function checkToken(token) { return token === AUTH_TOKEN; }

// Returns the 0-based row index in `data` for today's row.
// Compares by Y/M/D numbers (immune to locale-driven format errors).
// Handles: Date objects (any locale interpretation) and legacy text dates.
function findTodayRow(data, todayY, todayM, todayD) {
  for (var i = HEADER_ROWS; i < data.length; i++) {
    var cell = data[i][0];
    if (cell === '' || cell === null) continue;

    if (cell instanceof Date) {
      // Use sheet timezone for the comparison
      var cY = parseInt(Utilities.formatDate(cell, TZ, 'yyyy'));
      var cM = parseInt(Utilities.formatDate(cell, TZ, 'MM'));
      var cD = parseInt(Utilities.formatDate(cell, TZ, 'dd'));
      if (cY === todayY && cM === todayM && cD === todayD) return i;

    } else if (typeof cell === 'string') {
      var s = cell.trim();
      // dd-MM-yyyy  or  dd/MM/yyyy
      var m1 = s.match(/^(\d{1,2})[-\/](\d{1,2})[-\/](\d{4})$/);
      if (m1) {
        var d = parseInt(m1[1]), mo = parseInt(m1[2]), y = parseInt(m1[3]);
        if (d === todayD && mo === todayM && y === todayY) return i;
      }
    }
  }
  return -1;
}

// ── Punch-time helpers (v2.3) ────────────────────────────────────
//   A punch cell holds either the text this script writes ("08:51:00 AM") or,
//   on hand-edited rows, a real Date. Both reduce to seconds-since-midnight.

// "HH:mm:ss" (24h) -> seconds since midnight
function hmsToSecOfDay(hms) {
  var m = String(hms).match(/^(\d{1,2}):(\d{2}):(\d{2})$/);
  if (!m) return null;
  return parseInt(m[1], 10) * 3600 + parseInt(m[2], 10) * 60 + parseInt(m[3], 10);
}

// Punch cell -> seconds since midnight (null if the cell is empty/unparseable)
function punchCellToSecOfDay(cell) {
  if (cell === '' || cell == null) return null;

  if (cell instanceof Date) {
    return hmsToSecOfDay(Utilities.formatDate(cell, TZ, 'HH:mm:ss'));
  }

  var s = String(cell).trim();
  // "8:51:00 AM" / "08:51:00 PM"  (12-hour, as written by this script)
  var m = s.match(/^(\d{1,2}):(\d{2}):(\d{2})\s*([AaPp])\.?[Mm]\.?$/);
  if (m) {
    var h = parseInt(m[1], 10) % 12;
    if (m[4].toLowerCase() === 'p') h += 12;
    return h * 3600 + parseInt(m[2], 10) * 60 + parseInt(m[3], 10);
  }
  // "17:48:05" / "17:48"  (24-hour fallback)
  m = s.match(/^(\d{1,2}):(\d{2})(?::(\d{2}))?$/);
  if (m) {
    return parseInt(m[1], 10) * 3600 + parseInt(m[2], 10) * 60 + parseInt(m[3] || '0', 10);
  }
  return null;
}

// Punch cell -> display text for messages
function punchCellToText(cell) {
  if (cell instanceof Date) return Utilities.formatDate(cell, TZ, TIME_FORMAT);
  return String(cell == null ? '' : cell).trim();
}

function nextEmptyRow(sheet) {
  var last = sheet.getLastRow();
  if (last <= HEADER_ROWS) return HEADER_ROWS + 1;
  var vals = sheet.getRange(HEADER_ROWS + 1, 1, last - HEADER_ROWS, 1).getValues();
  for (var i = 0; i < vals.length; i++) {
    if (vals[i][0] === '' || vals[i][0] == null) return HEADER_ROWS + 1 + i;
  }
  return last + 1;
}

// GPS audit log written to Column O (col 15) — matches the screenshot layout
function logGPS(sheet, sheetRow, action, lat, lon, accuracy, now) {
  if (!lat || !lon) return;
  var ts    = Utilities.formatDate(now, TZ, 'dd-MM-yyyy HH:mm:ss');
  var entry = action + ' | ' + lat + ', ' + lon + ' ±' + accuracy + 'm @ ' + ts;
  var prev  = sheet.getRange(sheetRow, 15).getValue() || '';
  sheet.getRange(sheetRow, 15).setValue(prev ? prev + ' || ' + entry : entry);
  // Re-apply formatting to this row in case it's a fresh row in a fresh column
  sheet.getRange(sheetRow, 15).setHorizontalAlignment('left').setFontFamily('Roboto').setFontSize(10);
}

function jsonOut(obj) {
  return ContentService
    .createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}
