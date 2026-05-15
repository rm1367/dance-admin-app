// ===================================================================
// AWA Admin Portal — Attendance Analytics
// File: Attendance.gs
//
// All constants (SPREADSHEET_ID, CHECKIN_START_ROW, DATE_TAB_PATTERN)
// are defined in Code.gs and available here because GAS compiles all
// .gs files in the project together. No changes to Code.gs needed.
// ===================================================================

// Rows whose column-A name matches this pattern are template placeholders
// (e.g. "SKIP THIS ROW (DO NOT ENTER TEXT)") — excluded everywhere.
const SKIP_NAME_PATTERN_ = /skip/i;

// Class text that identifies a Social Dance Only purchase.
// These attendees are counted in totals but shown in their own
// top-10 leaderboard and excluded from the main top-10.
const SOCIAL_ONLY_PATTERN_ = /social dance only/i;

// Parses a column-B class string and returns a flag for each category.
// A single entry can count for multiple categories — for example,
// "Drop-In: Two Classes (Social & Level 1)" increments both Social and Level 1.
function parseClassCounts_(classText) {
  const t = (classText || '').toLowerCase();
  return {
    level2: t.includes('level 2') ? 1 : 0,
    level1: t.includes('level 1') ? 1 : 0,
    level0: t.includes('level 0') ? 1 : 0,
    social: t.includes('social')  ? 1 : 0,
  };
}

// Returns true if this row should be skipped entirely (template placeholder rows).
function isSkipRow_(name) {
  return SKIP_NAME_PATTERN_.test(name);
}

// Returns true if the class option is Social Dance Only.
function isSocialOnly_(classText) {
  return SOCIAL_ONLY_PATTERN_.test(classText || '');
}

// Reads check-in rows for a single date tab and returns attendance counts.
function getAttendanceForDate_(ss, tabName) {
  const empty = { total: 0, level0: 0, level1: 0, level2: 0, social: 0 };
  if (!tabName) return empty;

  const sheet = ss.getSheetByName(tabName);
  if (!sheet)  return empty;

  const lastRow = sheet.getLastRow();
  if (lastRow < CHECKIN_START_ROW) return empty;

  const numRows = lastRow - CHECKIN_START_ROW + 1;
  // Columns A (name) and B (class) starting at check-in row
  const data   = sheet.getRange(CHECKIN_START_ROW, 1, numRows, 2).getValues();
  const counts = { total: 0, level0: 0, level1: 0, level2: 0, social: 0 };

  data.forEach(function(row) {
    const name = String(row[0] || '').trim();
    if (!name || isSkipRow_(name)) return; // skip blank and placeholder rows
    counts.total++;
    const c = parseClassCounts_(row[1]);
    counts.level0 += c.level0;
    counts.level1 += c.level1;
    counts.level2 += c.level2;
    counts.social += c.social;
  });

  return counts;
}

// Scans ALL date tabs and builds year-to-date totals + two top-10 leaderboards:
//   top10        — all attendees EXCEPT social-dance-only purchases
//   top10Social  — only social-dance-only attendees
// Uses DATE_TAB_PATTERN (defined in Code.gs) to identify date tabs.
function getYTDStats_(ss) {
  const dateTabs = ss.getSheets().filter(function(s) {
    return DATE_TAB_PATTERN.test(s.getName().trim());
  });

  const totals   = { total: 0, level0: 0, level1: 0, level2: 0, social: 0 };
  const byName   = {}; // regular attendees (non-social-only)
  const bySocial = {}; // social-dance-only attendees

  dateTabs.forEach(function(sheet) {
    const lastRow = sheet.getLastRow();
    if (lastRow < CHECKIN_START_ROW) return;

    const numRows = lastRow - CHECKIN_START_ROW + 1;
    const data    = sheet.getRange(CHECKIN_START_ROW, 1, numRows, 2).getValues();

    data.forEach(function(row) {
      const name      = String(row[0] || '').trim();
      const classText = String(row[1] || '').trim();

      if (!name || isSkipRow_(name)) return; // skip blanks and placeholders

      // Always count toward totals
      totals.total++;
      const c = parseClassCounts_(classText);
      totals.level0 += c.level0;
      totals.level1 += c.level1;
      totals.level2 += c.level2;
      totals.social += c.social;

      // Route to the correct leaderboard (case-insensitive key, preserve casing)
      const key = name.toLowerCase();
      if (isSocialOnly_(classText)) {
        if (!bySocial[key]) bySocial[key] = { name: name, count: 0 };
        bySocial[key].count++;
      } else {
        if (!byName[key]) byName[key] = { name: name, count: 0 };
        byName[key].count++;
      }
    });
  });

  const top10 = Object.values(byName)
    .sort(function(a, b) { return b.count - a.count; })
    .slice(0, 10);

  const top10Social = Object.values(bySocial)
    .sort(function(a, b) { return b.count - a.count; })
    .slice(0, 10);

  return { totals: totals, top10: top10, top10Social: top10Social };
}

// ── Public entry point ───────────────────────────────────────────────
// Called via google.script.run.getAttendanceData(tabName) from the app.
// Returns tonight's counts for the selected date tab, YTD totals, and
// both top-10 leaderboards — all in a single round trip.
function getAttendanceData(tabName) {
  const ss      = SpreadsheetApp.openById(SPREADSHEET_ID);
  const tonight = getAttendanceForDate_(ss, tabName);
  const ytd     = getYTDStats_(ss);

  Logger.log('getAttendanceData: tab=' + tabName +
             ' tonight=' + JSON.stringify(tonight) +
             ' ytdTotal=' + ytd.totals.total +
             ' top10Count=' + ytd.top10.length +
             ' top10SocialCount=' + ytd.top10Social.length);

  return {
    tonight:     tonight,
    ytd:         ytd.totals,
    top10:       ytd.top10,
    top10Social: ytd.top10Social
  };
}
