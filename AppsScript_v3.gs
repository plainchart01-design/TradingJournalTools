// ============================================================
// TRADING JOURNAL — GOOGLE APPS SCRIPT (LIVE FINAL v3)
// Matches tradejournal_v5.html + TradeJournalEA.mq5 v2.31+
//
// WHAT'S NEW IN v3:
// - logTrade() now matches trades by TICKET (new column Q) when
//   the EA sends one (v2.31+). Fixes two bugs from the old
//   pair+time_open+direction key:
//     1. Two trades opened on the same pair in the same minute
//        were wrongly bounced as duplicates.
//     2. A close report always matched its own open row's key and
//        was silently dropped — closes never reached the sheet.
//   A close report (data.time_close present) now UPDATES the
//   matching open row in place instead of being rejected.
// - Older EA builds without a ticket field still work — logTrade
//   falls back to the original coarse key as a safety net.
// - Requires TradeJournalEA.mq5 v2.31+ (sends "ticket" in JSON).
//   This is backend-only otherwise: no changes needed to your
//   dashboard or your existing sheet data.
//
// WHAT'S NEW IN v2:
// - doPost now handles "addTrade" and "updateTrade" (sent by the
//   "Log Trade Manually" / "Edit Trade" feature in the web app).
//   Previously these fell through to logTrade() with the wrong
//   shape of data and either silently failed or wrote blank rows.
// - New column P "LOCAL_ID" lets manually-added trades be found
//   again later for editing, even after their sheet row changes.
// - getTrades() now returns local_id so the web app can match
//   manual trades correctly after a refresh.
//
// HOW TO USE:
// 1. Open Google Sheets → Extensions → Apps Script
// 2. Select ALL existing code → Delete
// 3. Paste this entire file
// 4. Press Ctrl+S to save
// 5. Do NOT run setupSheet() on your existing sheet — it clears
//    the Journal tab's contents. Column Q (TICKET) doesn't need a
//    header to work; the code writes/reads it by column number.
//    (setupSheet is still safe to use for a brand-new sheet.)
// 6. Deploy → Manage deployments → Edit → New version → Deploy
//    (you do NOT need a new URL — the existing /exec URL keeps working)
// 7. Update the EA to v2.31+ (sends the "ticket" field) and
//    recompile (F7) — without it, this backend still runs fine but
//    falls back to the old coarse matching that has the two bugs.
// ============================================================

const SHEET_NAME = "Journal";

// ─────────────────────────────────────────────────────────────
// doGet — web app reads trades + health check
// ─────────────────────────────────────────────────────────────
function doGet(e) {
  const action   = (e && e.parameter && e.parameter.action)   || "";
  const callback = (e && e.parameter && e.parameter.callback) || "";

  let result;
  if (action === "getTrades") {
    result = getTrades();
  } else if (action === "getAnalysisHistory") {
    result = getAnalysisHistory();
  } else {
    result = { status: "alive", message: "Trading Journal webhook is running!" };
  }

  // JSONP support for cross-origin requests from web app
  if (callback) {
    return ContentService
      .createTextOutput(callback + "(" + JSON.stringify(result) + ")")
      .setMimeType(ContentService.MimeType.JAVASCRIPT);
  }
  return respond(result);
}

// ─────────────────────────────────────────────────────────────
// doPost — receives new trades, manual trades, edits, and updates
// ─────────────────────────────────────────────────────────────
function doPost(e) {
  try {
    const raw  = e.postData.contents;
    const data = JSON.parse(raw);

    if (data.action === "update_row") {
      return respond(updateRow(data));
    }

    if (data.action === "saveAnalysis") {
      return respond(saveAnalysis(data));
    }

    // Manual trade logged from the web app's "Log Trade Manually" form.
    // The trade fields are nested under data.trade.
    if (data.action === "addTrade") {
      return respond(addTrade(data.trade || data));
    }

    // Manual trade edited from the web app's "Edit Trade" form.
    if (data.action === "updateTrade") {
      return respond(updateTrade(data.trade || data));
    }

    // Default: new trade coming in from MT5 EA
    return respond(logTrade(data));

  } catch (err) {
    return respond({ status: "error", message: err.message });
  }
}

// ─────────────────────────────────────────────────────────────
// RESPOND — always returns JSON (CORS handled by Google)
// ─────────────────────────────────────────────────────────────
function respond(obj) {
  return ContentService
    .createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}

// ─────────────────────────────────────────────────────────────
// GET TRADES
// Returns all trade rows with EXACT key names the web app uses:
// pair, day, time_open, time_close, duration, session,
// direction, entry, stop_loss, outcome, rr, entry_type,
// emotion, source, row, local_id, ticket
// ─────────────────────────────────────────────────────────────
function getTrades() {
  try {
    const ss    = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = ss.getSheetByName(SHEET_NAME);
    if (!sheet) return { trades: [], error: "Journal sheet not found" };

    const data = sheet.getDataRange().getValues();
    if (data.length <= 1) return { trades: [] };

    const trades = [];

    for (let i = 1; i < data.length; i++) {
      const r = data[i];

      // Skip completely empty rows (no pair AND no direction)
      if (!r[0] && !r[6] && !r[7]) continue;

      // Format the open time cleanly as "YYYY-MM-DD HH:MM"
      const timeOpen  = formatDateCell(r[2]);
      const timeClose = formatDateCell(r[3]);

      trades.push({
        row:        i + 1,       // ← actual Google Sheet row (header=row1, data starts row2)
        pair:       String(r[0]  || ""),
        day:        String(r[1]  || ""),
        time_open:  timeOpen,
        time_close: timeClose,
        duration:   String(r[4]  || ""),
        session:    String(r[5]  || ""),
        direction:  String(r[6]  || ""),
        entry:      String(r[7]  || ""),
        stop_loss:  String(r[8]  || ""),
        outcome:    String(r[9]  || ""),
        rr:         String(r[10] || ""),
        entry_type: String(r[11] || ""),
        emotion:    String(r[12] || ""),
        source:     String(r[13] || ""),
        timestamp:  String(r[14] || ""),
        local_id:   String(r[15] || ""),  // used to find manual trades for editing
        ticket:     String(r[16] || "")   // NEW — EA position ticket, used to match open↔close
      });
    }

    return { trades: trades };

  } catch (err) {
    return { trades: [], error: err.message };
  }
}

// ─────────────────────────────────────────────────────────────
// FORMAT DATE CELL
// Converts Google Sheets date objects → "YYYY-MM-DD HH:MM"
// Handles: Date objects, strings, numbers, empty values
// ─────────────────────────────────────────────────────────────
function formatDateCell(val) {
  if (!val || val === "") return "";

  // Already a string — check if it looks like a date
  if (typeof val === "string") {
    // Already in our format
    if (/^\d{4}-\d{2}-\d{2}/.test(val)) return val.slice(0, 16);
    // Try parsing
    const d = new Date(val);
    if (!isNaN(d)) return Utilities.formatDate(d, Session.getScriptTimeZone(), "yyyy-MM-dd HH:mm");
    return val;
  }

  // Google Sheets Date object
  if (val instanceof Date) {
    return Utilities.formatDate(val, Session.getScriptTimeZone(), "yyyy-MM-dd HH:mm");
  }

  // Number (serial date)
  if (typeof val === "number") {
    const d = new Date(val);
    if (!isNaN(d)) return Utilities.formatDate(d, Session.getScriptTimeZone(), "yyyy-MM-dd HH:mm");
  }

  return String(val);
}

// ─────────────────────────────────────────────────────────────
// BUILD TRADE ROW
// Shared row-builder used by logTrade (EA), addTrade (manual),
// and updateTrade (manual edit). Returns a 16-element array
// matching columns A–P (P = LOCAL_ID).
// ─────────────────────────────────────────────────────────────
function buildTradeRow(data) {
  const now       = new Date();
  const openTime  = data.time_open  ? parseIncomingTime(data.time_open)  : now;
  const closeTime = data.time_close ? parseIncomingTime(data.time_close) : null;
  const day       = data.day || getDayOfWeek(openTime);
  const session   = data.session || detectSession(openTime);

  // Prefer a duration sent by the client; otherwise compute it
  const duration = (data.duration !== undefined && data.duration !== "")
    ? data.duration
    : ((openTime && closeTime) ? calcDuration(openTime, closeTime) : "");

  // Prefer an RR sent by the client (non-empty); otherwise compute it
  const rr = (data.rr !== undefined && data.rr !== "")
    ? data.rr
    : calcRR(data.entry, data.stop_loss, data.take_profit, data.direction);

  const outcome   = (data.outcome || "OPEN").toUpperCase();
  const direction = (data.direction || "").toUpperCase();

  return [
    data.pair       || "",                                              // A - PAIR
    day,                                                                 // B - DAY
    Utilities.formatDate(openTime, Session.getScriptTimeZone(), "yyyy-MM-dd HH:mm"),  // C - TIME OPEN
    closeTime ? Utilities.formatDate(closeTime, Session.getScriptTimeZone(), "yyyy-MM-dd HH:mm") : "", // D - TIME CLOSE
    duration,                                                            // E - DURATION
    session,                                                             // F - SESSION
    direction,                                                           // G - LONG/SHORT
    data.entry      || "",                                               // H - ENTRY
    data.stop_loss  || "",                                               // I - STOP LOSS
    outcome,                                                             // J - OUTCOME
    rr,                                                                  // K - RR
    data.entry_type || "",                                               // L - ENTRY TYPE
    data.emotion    || "",                                               // M - EMOTION
    data.source     || "MetaTrader5",                                    // N - SOURCE
    now.toISOString(),                                                   // O - TIMESTAMP
    data.local_id !== undefined && data.local_id !== null ? String(data.local_id) : "", // P - LOCAL_ID
    data.ticket   !== undefined && data.ticket   !== null ? String(data.ticket)   : ""  // Q - TICKET (EA position ticket, used to match open↔close)
  ];
}

// ─────────────────────────────────────────────────────────────
// LOG TRADE — called when MT5 EA sends a new trade (open or close)
// ─────────────────────────────────────────────────────────────
// v3: trades are matched by TICKET (column Q) when the EA provides
// one. This fixes two bugs that the old pair+time_open+direction
// key caused:
//   1. Two genuinely different trades opened on the same pair in
//      the same minute were wrongly flagged as duplicates.
//   2. A close report (same pair/time_open/direction as its open,
//      plus time_close) always matched its own open row and was
//      silently dropped — closes never reached the sheet.
// With a ticket: an OPEN report checks for an existing ticket
// (true duplicate → reject); a CLOSE report (data.time_close is
// set) looks up that ticket's row and UPDATES it in place.
// Without a ticket (older EA build, or test/manual data): falls
// back to the original coarse key as a safety net.
// ─────────────────────────────────────────────────────────────
function logTrade(data) {
  try {
    const ss    = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = ss.getSheetByName(SHEET_NAME);
    if (!sheet) return { status: "error", message: "Journal sheet not found" };

    const ticket = (data.ticket !== undefined && data.ticket !== null && data.ticket !== "")
      ? String(data.ticket) : "";
    const isClose = !!data.time_close;

    // ── TICKET-BASED MATCHING ──────────────────────────────────
    if (ticket) {
      const matchRow = findRowByTicket(sheet, ticket);

      if (isClose) {
        if (matchRow !== -1) {
          return updateRowOnClose(sheet, matchRow, data);
        }
        // No matching open row found — log the close anyway so the
        // data isn't lost (e.g. the open report failed earlier).
        const row = buildTradeRow(data);
        row[16] = ticket;
        sheet.appendRow(row);
        return { status: "success", message: "Trade logged (close, no matching open found)" };
      }

      // Open report.
      if (matchRow !== -1) {
        Logger.log("Duplicate open skipped: ticket " + ticket);
        return { status: "duplicate", message: "Trade already exists" };
      }
      const row = buildTradeRow(data);
      row[16] = ticket;
      sheet.appendRow(row);
      return { status: "success", message: "Trade logged" };
    }

    // ── FALLBACK: no ticket sent ────────────────────────────────
    if (data.pair && data.time_open && data.direction) {
      const key = (data.pair + "|" + String(data.time_open).slice(0,16) + "|" + data.direction).toUpperCase();
      const lastRow = sheet.getLastRow();
      if (lastRow >= 2) {
        const existing = sheet.getRange(2, 1, lastRow - 1, 7).getValues();
        for (let i = 0; i < existing.length; i++) {
          const r = existing[i];
          const existingKey = (String(r[0]) + "|" + String(formatDateCell(r[2])).slice(0,16) + "|" + String(r[6])).toUpperCase();
          if (existingKey === key) {
            Logger.log("Duplicate trade skipped: " + key);
            return { status: "duplicate", message: "Trade already exists" };
          }
        }
      }
    }

    sheet.appendRow(buildTradeRow(data));
    return { status: "success", message: "Trade logged" };

  } catch (err) {
    return { status: "error", message: err.message };
  }
}

// Find the sheet row whose TICKET column (Q = column 17) matches.
// Returns -1 if not found.
function findRowByTicket(sheet, ticket) {
  const lastRow = sheet.getLastRow();
  if (lastRow < 2) return -1;
  const tickets = sheet.getRange(2, 17, lastRow - 1, 1).getValues();
  for (let i = 0; i < tickets.length; i++) {
    if (String(tickets[i][0]) === ticket) return i + 2;
  }
  return -1;
}

// Update an existing OPEN row in place with close-time data.
function updateRowOnClose(sheet, row, data) {
  const closeTime = data.time_close ? parseIncomingTime(data.time_close) : null;
  sheet.getRange(row, 4).setValue(
    closeTime ? Utilities.formatDate(closeTime, Session.getScriptTimeZone(), "yyyy-MM-dd HH:mm") : ""
  ); // D - TIME CLOSED
  if (data.duration  !== undefined && data.duration  !== "") sheet.getRange(row, 5).setValue(data.duration);   // E - DURATION
  if (data.stop_loss !== undefined && data.stop_loss !== "") sheet.getRange(row, 9).setValue(data.stop_loss);  // I - STOP LOSS (in case it moved before close)
  if (data.outcome) sheet.getRange(row, 10).setValue(String(data.outcome).toUpperCase());                     // J - OUTCOME
  if (data.rr !== undefined && data.rr !== "") sheet.getRange(row, 11).setValue(data.rr);                      // K - RR
  return { status: "updated", row: row, message: "Trade closed" };
}

// ─────────────────────────────────────────────────────────────
// ADD TRADE — called when the web app logs a trade manually.
// Same as logTrade but always tags the row as "Manual" (unless
// the client explicitly says otherwise) and stores LOCAL_ID so
// it can be found again later if the user edits it.
// ─────────────────────────────────────────────────────────────
function addTrade(trade) {
  try {
    const ss  = SpreadsheetApp.getActiveSpreadsheet();
    let sheet = ss.getSheetByName(SHEET_NAME);
    if (!sheet) { setupSheet(); sheet = ss.getSheetByName(SHEET_NAME); }
    if (!sheet) return { status: "error", message: "Could not create Journal sheet" };

    trade.source = trade.source || "Manual";
    sheet.appendRow(buildTradeRow(trade));
    return { status: "success", message: "Trade added", local_id: trade.local_id || "" };

  } catch (err) {
    return { status: "error", message: err.message };
  }
}

// ─────────────────────────────────────────────────────────────
// UPDATE TRADE — called when the user edits a manually-logged
// trade in the web app. Finds the existing row by LOCAL_ID
// (column P). If LOCAL_ID isn't found (e.g. the trade was never
// successfully synced), falls back to the sheet row number sent
// by the client, and if that's not usable either, appends a new
// row so the edit is never lost.
// ─────────────────────────────────────────────────────────────
function updateTrade(trade) {
  try {
    const ss    = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = ss.getSheetByName(SHEET_NAME);
    if (!sheet) return { status: "error", message: "Journal sheet not found" };

    const localId = (trade.local_id !== undefined && trade.local_id !== null) ? String(trade.local_id) : "";
    let targetRow = -1;

    // 1) Try to find the row by LOCAL_ID (column P = column 16)
    if (localId) {
      const lastRow = sheet.getLastRow();
      if (lastRow >= 2) {
        const ids = sheet.getRange(2, 16, lastRow - 1, 1).getValues();
        for (let i = 0; i < ids.length; i++) {
          if (String(ids[i][0]) === localId) { targetRow = i + 2; break; }
        }
      }
    }

    // 2) Fall back to a real sheet row number if the client sent one
    //    (anything that looks like a Date.now() ID — 13+ digits — is
    //    NOT a real row number and is ignored here)
    if (targetRow === -1 && trade.row) {
      const rowNum = Number(trade.row);
      if (!isNaN(rowNum) && rowNum >= 2 && rowNum < 100000) {
        targetRow = rowNum;
      }
    }

    trade.source = trade.source || "Manual";
    const rowData = buildTradeRow(trade);
    if (localId) rowData[15] = localId;

    // 3) Not found anywhere — append as a new row instead of losing the edit
    if (targetRow === -1) {
      sheet.appendRow(rowData);
      return { status: "added", local_id: localId };
    }

    sheet.getRange(targetRow, 1, 1, rowData.length).setValues([rowData]);
    return { status: "updated", row: targetRow };

  } catch (err) {
    return { status: "error", message: err.message };
  }
}

// ─────────────────────────────────────────────────────────────
// UPDATE ROW — called when user saves emotion/session/type
// from the web app emotion modal
// ─────────────────────────────────────────────────────────────
function updateRow(data) {
  try {
    const ss    = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = ss.getSheetByName(SHEET_NAME);
    const row   = parseInt(data.row);

    if (isNaN(row) || row < 2) {
      return { status: "error", message: "Invalid row number: " + data.row };
    }

    // Only update fields that were actually provided
    if (data.emotion    !== undefined && data.emotion    !== "") sheet.getRange(row, 13).setValue(data.emotion);
    if (data.session    !== undefined && data.session    !== "") sheet.getRange(row, 6).setValue(data.session);
    if (data.entry_type !== undefined && data.entry_type !== "") sheet.getRange(row, 12).setValue(data.entry_type);

    return { status: "updated", row: row };

  } catch (err) {
    return { status: "error", message: err.message };
  }
}

// ─────────────────────────────────────────────────────────────
// TEST TRADE — select this in dropdown and click ▶ Play
// Adds a fake trade so you can test everything works
// ─────────────────────────────────────────────────────────────
function testTrade() {
  const testData = {
    pair:        "XAUUSD",
    direction:   "LONG",
    entry:       "2345.50",
    stop_loss:   "2330.00",
    take_profit: "2376.00",
    outcome:     "WIN",
    time_open:   "2026-04-25 09:30",
    time_close:  "2026-04-25 11:45",
    profit:      "150.00",
    source:      "MetaTrader5"
  };

  const result = logTrade(testData);
  Logger.log("testTrade result: " + JSON.stringify(result));

  if (result.status === "success") {
    Logger.log("✅ Test trade added! Check your Journal sheet — a new XAUUSD row should be there.");
    Logger.log("✅ Now go to your web app and click Refresh to see it appear.");
  } else {
    Logger.log("❌ Error: " + result.message);
  }
}

// ─────────────────────────────────────────────────────────────
// SETUP SHEET — run ONCE to create headers and stats
// Select setupSheet in dropdown → click ▶ Play
// (Not required on your existing sheet — only needed for a
// brand-new Journal sheet, or if you want to reset it.)
// ─────────────────────────────────────────────────────────────
function setupSheet() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();

  // If Journal sheet already exists, clear its contents but keep it (preserves row history)
  // If it doesn't exist, create it fresh
  let sheet = ss.getSheetByName(SHEET_NAME);
  if (sheet) {
    // Remove any existing banding before clearing
    const bandings = sheet.getBandings();
    bandings.forEach(b => b.remove());
    sheet.clearContents();
    sheet.clearFormats();
  } else {
    sheet = ss.insertSheet(SHEET_NAME);
  }

  // Set headers (row 1)
  const headers = [
    "PAIR",           // A
    "DAY",            // B
    "TIME (OPEN)",    // C
    "TIME (CLOSED)",  // D
    "DURATION",       // E
    "SESSION",        // F
    "LONG/SHORT",     // G
    "ENTRY",          // H
    "STOP LOSS",      // I
    "OUTCOME",        // J
    "RR - RISK REWARD", // K
    "ENTRY TYPE",     // L
    "EMOTION",        // M
    "SOURCE",         // N
    "TIMESTAMP",      // O
    "LOCAL_ID",       // P — persistent ID for editing manual trades
    "TICKET"          // Q — NEW: EA position ticket, matches open↔close reports
  ];

  sheet.getRange(1, 1, 1, headers.length).setValues([headers]);

  // Style header row
  sheet.getRange(1, 1, 1, headers.length)
    .setBackground("#0c1428")
    .setFontColor("#00f5d4")
    .setFontWeight("bold")
    .setHorizontalAlignment("center")
    .setFontSize(10);

  sheet.setFrozenRows(1);

  // Column widths
  const widths = [100,90,140,140,90,100,90,100,100,90,130,110,200,120,160,140,100];
  widths.forEach((w, i) => sheet.setColumnWidth(i+1, w));

  // Alternating row colors for readability
  sheet.getRange(2, 1, 100, headers.length).applyRowBanding(SpreadsheetApp.BandingTheme.LIGHT_GREY);

  // Create Stats sheet
  let stats = ss.getSheetByName("Stats");
  if (!stats) stats = ss.insertSheet("Stats");
  stats.clearContents();

  const statsRows = [
    ["TRADING STATISTICS", ""],
    ["", ""],
    ["WINS",         `=COUNTIF(Journal!J:J,"WIN")`],
    ["LOSSES",       `=COUNTIF(Journal!J:J,"LOSS")`],
    ["TOTAL TRADES", `=C3+C4`],
    ["WIN %",        `=IF(C5=0,"",TEXT(C3/C5,"0.00%"))`],
    ["", ""],
    ["ENTRY TYPE BREAKDOWN", ""],
    ["FLIP",         `=COUNTIF(Journal!L:L,"FLIP")`],
    ["EXTREME",      `=COUNTIF(Journal!L:L,"EXTREME")`],
    ["BREAKOUT",     `=COUNTIF(Journal!L:L,"BREAKOUT")`],
    ["PULLBACK",     `=COUNTIF(Journal!L:L,"PULLBACK")`],
    ["REVERSAL",     `=COUNTIF(Journal!L:L,"REVERSAL")`],
    ["OTHER",        `=COUNTIF(Journal!L:L,"OTHER")`],
  ];

  stats.getRange(1, 2, statsRows.length, 2).setValues(statsRows);
  stats.getRange("B1").setFontWeight("bold").setFontSize(13).setFontColor("#0c1428");
  stats.getRange("B8").setFontWeight("bold").setFontSize(11);
  stats.getRange("B3:B14").setFontWeight("bold");
  stats.setColumnWidth(2, 180);
  stats.setColumnWidth(3, 120);

  // Create Analysis sheet
  let analysis = ss.getSheetByName("Analysis");
  if (!analysis) {
    analysis = ss.insertSheet("Analysis");
    const aHeaders = ["DATE", "TIME", "TRADES ANALYSED", "WIN RATE", "SUMMARY", "FULL CONVERSATION"];
    analysis.getRange(1, 1, 1, aHeaders.length).setValues([aHeaders]);
    analysis.getRange(1, 1, 1, aHeaders.length)
      .setBackground("#0c1428")
      .setFontColor("#a855f7")
      .setFontWeight("bold")
      .setHorizontalAlignment("center")
      .setFontSize(10);
    analysis.setFrozenRows(1);
    analysis.setColumnWidth(1, 100);
    analysis.setColumnWidth(2, 80);
    analysis.setColumnWidth(3, 130);
    analysis.setColumnWidth(4, 100);
    analysis.setColumnWidth(5, 300);
    analysis.setColumnWidth(6, 500);
  }

  Logger.log("✅ Journal setup complete! Your sheet is ready.");
  Logger.log("✅ Next: click Deploy → New deployment → Web app → Anyone → Deploy → Copy URL.");
}

// ─────────────────────────────────────────────────────────────
// SAVE ANALYSIS — saves a completed AI analysis to the sheet
// Called from web app after each analysis session
// ─────────────────────────────────────────────────────────────
function saveAnalysis(data) {
  try {
    const ss    = SpreadsheetApp.getActiveSpreadsheet();
    let sheet   = ss.getSheetByName("Analysis");

    // Auto-create the Analysis sheet if it doesn't exist yet
    if (!sheet) {
      sheet = ss.insertSheet("Analysis");
      const headers = ["DATE", "TIME", "TRADES ANALYSED", "WIN RATE", "SUMMARY", "FULL CONVERSATION"];
      sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
      sheet.getRange(1, 1, 1, headers.length)
        .setBackground("#0c1428")
        .setFontColor("#a855f7")
        .setFontWeight("bold")
        .setHorizontalAlignment("center")
        .setFontSize(10);
      sheet.setFrozenRows(1);
      sheet.setColumnWidth(1, 100);
      sheet.setColumnWidth(2, 80);
      sheet.setColumnWidth(3, 130);
      sheet.setColumnWidth(4, 100);
      sheet.setColumnWidth(5, 300);
      sheet.setColumnWidth(6, 500);
    }

    const now     = new Date();
    const dateStr = Utilities.formatDate(now, Session.getScriptTimeZone(), "yyyy-MM-dd");
    const timeStr = Utilities.formatDate(now, Session.getScriptTimeZone(), "HH:mm");

    const row = [
      dateStr,
      timeStr,
      data.tradeCount  || "",
      data.winRate     || "",
      data.summary     || "",       // short 1-2 line summary of key findings
      data.fullChat    || ""        // full JSON of the conversation messages
    ];

    sheet.appendRow(row);

    // Wrap text in the summary and full chat columns
    const lastRow = sheet.getLastRow();
    sheet.getRange(lastRow, 5).setWrap(true);
    sheet.getRange(lastRow, 6).setWrap(false); // keep full chat compact

    return { status: "saved", date: dateStr, time: timeStr };

  } catch (err) {
    return { status: "error", message: err.message };
  }
}

// ─────────────────────────────────────────────────────────────
// GET ANALYSIS HISTORY — returns past analyses for context
// Web app sends these to AI so it can reference past sessions
// ─────────────────────────────────────────────────────────────
function getAnalysisHistory() {
  try {
    const ss    = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = ss.getSheetByName("Analysis");
    if (!sheet) return { history: [] };

    const data = sheet.getDataRange().getValues();
    if (data.length <= 1) return { history: [] };

    // Return last 3 analyses (most recent first) for AI context
    const rows = data.slice(1).reverse().slice(0, 3);
    const history = rows.map(r => ({
      date:         String(r[0] || ""),
      time:         String(r[1] || ""),
      tradeCount:   String(r[2] || ""),
      winRate:      String(r[3] || ""),
      summary:      String(r[4] || "")
    }));

    return { history };

  } catch (err) {
    return { history: [], error: err.message };
  }
}

// ─────────────────────────────────────────────────────────────
// HELPERS
// ─────────────────────────────────────────────────────────────

function getDayOfWeek(date) {
  const days = ["Sunday","Monday","Tuesday","Wednesday","Thursday","Friday","Saturday"];
  return days[(date || new Date()).getDay()];
}

// Detects trading session from UTC hour
function detectSession(dateObj) {
  try {
    const hour = dateObj.getUTCHours();
    if (hour >= 22 || hour < 8)  return "Asia";
    if (hour >= 8  && hour < 12) return "London";
    if (hour >= 12 && hour < 17) return "New York";
    if (hour >= 17 && hour < 20) return "London Close";
    return "After Hours";
  } catch { return ""; }
}

// Parse incoming time string from MT5 EA → Date object
function parseIncomingTime(str) {
  if (!str) return new Date();
  // "2026.04.25 09:30" (MT5 format)
  if (/^\d{4}\.\d{2}\.\d{2}/.test(str)) {
    str = str.replace(/\./g, '-');
  }
  const d = new Date(str);
  return isNaN(d) ? new Date() : d;
}

// Calculate human-readable duration
function calcDuration(open, close) {
  try {
    const diff = close - open;
    if (diff < 0) return "";
    const h = Math.floor(diff / 3600000);
    const m = Math.floor((diff % 3600000) / 60000);
    return `${String(h).padStart(2,"0")}h ${String(m).padStart(2,"0")}m`;
  } catch { return ""; }
}

// Calculate Risk:Reward ratio
function calcRR(entry, sl, tp, direction) {
  try {
    const e = parseFloat(entry);
    const s = parseFloat(sl);
    const t = parseFloat(tp);
    if (!e || !s || !t || isNaN(e) || isNaN(s) || isNaN(t)) return "";
    const risk   = Math.abs(e - s);
    const reward = Math.abs(t - e);
    if (risk === 0) return "";
    return (reward / risk).toFixed(2);
  } catch { return ""; }
}
