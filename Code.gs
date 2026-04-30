// ============================================================
//  Almentor Quality Dashboard — Code.gs
//  Google Apps Script Backend
// ============================================================

var SPREADSHEET_ID = "1qdE3LUOj8vrtPK1CL99Y7UPnvmFh8TcON8cYU4zXm-E";

// ── Entry point ──────────────────────────────────────────────
function doGet(e) {
  return HtmlService
    .createHtmlOutputFromFile("Index")
    .setTitle("Almentor Quality Dashboard")
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL)
    .addMetaTag("viewport", "width=device-width, initial-scale=1");
}

// ── Master data fetcher ──────────────────────────────────────
function getSheetData(sheetName) {
  try {
    var ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    var sheet = sheetName ? ss.getSheetByName(sheetName) : ss.getSheets()[0];

    if (!sheet) return { error: "Sheet not found: " + sheetName };

    var data   = sheet.getDataRange().getValues();
    if (data.length < 2) return { headers: [], rows: [], sheetName: sheet.getName() };

    var headers = data[0].map(String);
    var rows    = data.slice(1).map(function(row) {
      return row.map(function(cell) {
        // Convert Date objects to ISO strings
        if (cell instanceof Date) return Utilities.formatDate(cell, Session.getScriptTimeZone(), "yyyy-MM-dd");
        return cell;
      });
    });

    return {
      sheetName : sheet.getName(),
      headers   : headers,
      rows      : rows
    };
  } catch (err) {
    return { error: err.message };
  }
}

// ── All sheet names ──────────────────────────────────────────
function getSheetNames() {
  try {
    var ss     = SpreadsheetApp.openById(SPREADSHEET_ID);
    var sheets = ss.getSheets();
    return sheets.map(function(s) { return s.getName(); });
  } catch (err) {
    return { error: err.message };
  }
}

// ── KPI aggregation helper ───────────────────────────────────
function getKPIs(sheetName) {
  try {
    var result  = getSheetData(sheetName);
    if (result.error) return result;

    var headers = result.headers;
    var rows    = result.rows;

    // Detect team column
    var teamCol = _findCol(headers, ["Team", "team name", "department"]);

    // Detect score/quality column
    var scoreCol = _findCol(headers, ["Score", "Quality", "Rating", "Grade", "Mark", "Result", "Percentage"]);

    // Detect status column
    var statusCol = _findCol(headers, ["Status", "State", "Completion", "Progress"]);

    var teamScores = {};
    var statusCounts = {};
    var total = rows.length;

    rows.forEach(function(row) {
      // Teams
      if (teamCol > -1) {
        var team  = String(row[teamCol]).trim() || "Unknown";
        var score = scoreCol > -1 ? parseFloat(row[scoreCol]) : 0;
        if (!teamScores[team]) teamScores[team] = { sum: 0, count: 0 };
        teamScores[team].sum   += isNaN(score) ? 0 : score;
        teamScores[team].count += 1;
      }

      // Statuses
      if (statusCol > -1) {
        var st = String(row[statusCol]).trim() || "Unknown";
        statusCounts[st] = (statusCounts[st] || 0) + 1;
      }
    });

    // Best / worst team by average score
    var bestTeam = "N/A", worstTeam = "N/A";
    var bestAvg = -Infinity, worstAvg = Infinity;
    Object.keys(teamScores).forEach(function(t) {
      var avg = teamScores[t].sum / teamScores[t].count;
      if (avg > bestAvg)  { bestAvg  = avg; bestTeam  = t; }
      if (avg < worstAvg) { worstAvg = avg; worstTeam = t; }
    });

    return {
      totalRecords   : total,
      bestTeam       : bestTeam,
      worstTeam      : worstTeam,
      teamScores     : teamScores,
      statusCounts   : statusCounts
    };
  } catch (err) {
    return { error: err.message };
  }
}

// ── Internal column finder ───────────────────────────────────
function _findCol(headers, candidates) {
  var lower = headers.map(function(h) { return h.toLowerCase(); });
  for (var i = 0; i < candidates.length; i++) {
    var idx = lower.findIndex(function(h) {
      return h.indexOf(candidates[i].toLowerCase()) > -1;
    });
    if (idx > -1) return idx;
  }
  return -1;
}
