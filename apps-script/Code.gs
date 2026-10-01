/**
 * WFM Transitions Dashboard — Live Sync from monday.com
 * ------------------------------------------------------
 * Pulls the "Transitions Summary" board (Active + On Hold projects) and
 * cross-references checklist % from the "Active Transitions" board,
 * then rewrites the Raw Data tab. The Dashboard tab's formulas read from
 * Raw Data, so once this runs, the whole dashboard updates itself.
 *
 * ONE-TIME SETUP (about 5 minutes):
 *
 * 1. Open your Google Sheet → Extensions → Apps Script.
 * 2. Delete any starter code in Code.gs, paste this whole file in, save.
 * 3. Get a monday.com API token:
 *      monday.com → Avatar (bottom-left) → Administration → Connections →
 *      API → "Generate" (or reuse an existing personal token).
 * 4. In the Apps Script editor: Project Settings (gear icon, left sidebar)
 *      → Script Properties → Add script property
 *      Name:  MONDAY_API_TOKEN
 *      Value: <paste your token>
 * 5. Back in the editor, select the function "refreshFromMonday" in the
 *      dropdown at the top, click Run once. The first run will ask you
 *      to authorize the script (it needs permission to call external
 *      APIs and edit this spreadsheet) — approve it.
 * 6. Set it to run automatically: left sidebar → Triggers (clock icon)
 *      → + Add Trigger → function: refreshFromMonday → Event source:
 *      Time-driven → Hour timer → Every hour (or whatever cadence you want)
 *      → Save.
 *
 * That's it — from then on, this Sheet refreshes itself on that schedule
 * without anyone needing to open Apps Script again. To change *what* it
 * pulls (columns, boards, mappings), edit the CONFIG block below.
 */

// ============================== CONFIG ==============================
var SUMMARY_BOARD_ID = '9668126921';   // "Transitions Summary" board
var CHECKLIST_BOARD_ID = '9655713663'; // "Active Transitions" checklist board
var SHEET_NAME = 'Raw Data';
var HEADER_ROW = 5;   // matches the masthead layout already in the sheet
var FIRST_DATA_ROW = HEADER_ROW + 1;

// Which groups on the Summary board count as "Active" vs "On Hold".
// (Closed Projects group is intentionally excluded from Raw Data.)
var GROUP_LABELS = {
  'topics': 'Active',
  'group_mkr08ks3': 'On Hold'
};

// WFM PM email -> display name, as used on the Summary board.
var PM_NAMES = {
  'gerson.jimenez@telusdigital.com': 'Gerson Jiménez',
  'elle.kumari@telusdigital.com': 'Elle Kumari',
  'michael.vergara@telusdigital.com': 'Michael Vergara'
};

var SUMMARY_COLUMN_IDS = [
  'color_mkt5f0m',      // Category (NCNB/ECNB/ECEB)
  'color_mkt6j79k',      // Stage
  'text_mkt6122g',       // WFM PM (email)
  'date_mkt6pcxx',       // Kickoff date
  'date_mkt6mqx1',       // Target Go-Live
  'status9',             // SOW Status
  'long_text_mm3bxnqh'   // Notes (Overall Status)
];

var CHECKLIST_STATUS_COLUMN_ID = 'status';

// ============================ ENTRY POINT ============================
function refreshFromMonday() {
  var token = getMondayToken_();

  var summaryItems = fetchAllItems_(SUMMARY_BOARD_ID, SUMMARY_COLUMN_IDS, token);
  var checklistItems = fetchAllItems_(CHECKLIST_BOARD_ID, [CHECKLIST_STATUS_COLUMN_ID], token);
  var checklistByGroup = buildChecklistIndex_(checklistItems);

  var rows = [];
  summaryItems.forEach(function (item) {
    var groupLabel = GROUP_LABELS[item.group.id];
    if (!groupLabel) return; // skip Closed Projects and anything else

    var col = colMap_(item.column_values);
    var pmEmail = col['text_mkt6122g'] || '';
    var pmName = PM_NAMES[pmEmail] || pmEmail;

    var checklistPct = lookupChecklistPct_(item.name, checklistByGroup);

    rows.push([
      item.name,
      pmName,
      col['color_mkt5f0m'] || '',
      col['color_mkt6j79k'] || '',
      groupLabel,
      col['date_mkt6pcxx'] || '',
      col['date_mkt6mqx1'] || '',
      col['status9'] || '',
      checklistPct,                 // number 0..1, or '' if no checklist tracked
      (col['long_text_mm3bxnqh'] || '').replace(/\n/g, ' ')
    ]);
  });

  writeRows_(rows);
  Logger.log('Refreshed Raw Data with ' + rows.length + ' projects at ' + new Date());
}

// ============================ HELPERS ============================

function getMondayToken_() {
  var token = PropertiesService.getScriptProperties().getProperty('MONDAY_API_TOKEN');
  if (!token) {
    throw new Error('Missing MONDAY_API_TOKEN. Set it under Project Settings > Script Properties.');
  }
  return token;
}

function mondayQuery_(query, token) {
  var resp = UrlFetchApp.fetch('https://api.monday.com/v2', {
    method: 'post',
    contentType: 'application/json',
    headers: { 'Authorization': token },
    payload: JSON.stringify({ query: query }),
    muteHttpExceptions: true
  });
  var body = JSON.parse(resp.getContentText());
  if (body.errors) {
    throw new Error('monday.com API error: ' + JSON.stringify(body.errors));
  }
  return body.data;
}

// Fetches every item on a board (paginated), requesting `text` for each
// column so we don't have to parse type-specific JSON per column type.
function fetchAllItems_(boardId, columnIds, token) {
  var colIdsGql = columnIds.map(function (c) { return '"' + c + '"'; }).join(',');
  var items = [];
  var cursor = null;
  var first = true;

  while (first || cursor) {
    first = false;
    var query = cursor
      ? 'query { next_items_page(limit: 500, cursor: "' + cursor + '") { cursor items { id name group { id title } column_values(ids: [' + colIdsGql + ']) { id text } } } }'
      : 'query { boards(ids: [' + boardId + ']) { items_page(limit: 500) { cursor items { id name group { id title } column_values(ids: [' + colIdsGql + ']) { id text } } } } }';

    var data = mondayQuery_(query, token);
    var page = cursor ? data.next_items_page : data.boards[0].items_page;
    items = items.concat(page.items);
    cursor = page.cursor || null;
  }
  return items;
}

function colMap_(columnValues) {
  var map = {};
  columnValues.forEach(function (cv) { map[cv.id] = cv.text; });
  return map;
}

function normalize_(name) {
  return (name || '').toLowerCase().replace(/[^a-z0-9]/g, '');
}

// Groups the checklist board's items by their group (one checklist group
// per project) and reduces each group to a "% Done" number.
function buildChecklistIndex_(checklistItems) {
  var byGroup = {}; // normalized group title -> { done, total }
  checklistItems.forEach(function (item) {
    var title = item.group.title;
    var key = normalize_(title);
    if (!byGroup[key]) byGroup[key] = { title: title, done: 0, total: 0 };
    byGroup[key].total += 1;
    var statusText = colMap_(item.column_values)[CHECKLIST_STATUS_COLUMN_ID];
    if (statusText === 'Done') byGroup[key].done += 1;
  });
  return byGroup;
}

// Matches a Summary-board project name to a checklist group, tolerating
// small naming differences (e.g. "Sparklight/Cable One - IN / SV" vs
// "Sparklight  - IN - SV") via exact-then-contains matching.
function lookupChecklistPct_(projectName, checklistByGroup) {
  var key = normalize_(projectName);
  if (checklistByGroup[key]) {
    var g = checklistByGroup[key];
    return g.total > 0 ? Math.round((g.done / g.total) * 100) / 100 : '';
  }
  for (var k in checklistByGroup) {
    if (k.indexOf(key) !== -1 || key.indexOf(k) !== -1) {
      var g2 = checklistByGroup[k];
      return g2.total > 0 ? Math.round((g2.done / g2.total) * 100) / 100 : '';
    }
  }
  return ''; // no matching checklist group -> "No checklist tracked"
}

function writeRows_(rows) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(SHEET_NAME);
  if (!sheet) throw new Error('Sheet "' + SHEET_NAME + '" not found.');

  var lastRow = sheet.getLastRow();
  if (lastRow >= FIRST_DATA_ROW) {
    sheet.getRange(FIRST_DATA_ROW, 1, lastRow - FIRST_DATA_ROW + 1, 10).clearContent();
  }
  if (rows.length > 0) {
    sheet.getRange(FIRST_DATA_ROW, 1, rows.length, 10).setValues(rows);
    // Re-apply percent format to the Checklist % column (col 9)
    sheet.getRange(FIRST_DATA_ROW, 9, rows.length, 1).setNumberFormat('0%');
  }
}
