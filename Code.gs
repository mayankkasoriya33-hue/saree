// ─────────────────────────────────────────
// SETUP — Ek baar manually run karo
// Apps Script Editor → dropdown mein "setupProperties" select → Run
// ─────────────────────────────────────────
function setupProperties() {
  PropertiesService.getScriptProperties().setProperty(
    'GEMINI_API_KEY',
    'APNI_GEMINI_API_KEY_YAHAN_DAALO'
  );
  Logger.log('✅ API key save ho gayi Script Properties mein!');
}

function authorizeMe() {
  UrlFetchApp.fetch("https://www.google.com");
  SpreadsheetApp.openById("19F_9d62ONF3RSwoI0qM_2H0fbun5IbGrLaj4QYBN-Fo");
  Logger.log("Authorization successful!");
}

// ─────────────────────────────────────────
// doGet — serves index.html as a template
// createTemplateFromFile processes <?!= ?> tags
// ─────────────────────────────────────────
function doGet() {
  return HtmlService.createTemplateFromFile('index')
    .evaluate()
    .setTitle("CRM KPI Dashboard")
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

// ─────────────────────────────────────────
// include() helper
// Usage inside index.html:
//   <?!= include('IndividualKPI') ?>
//   <?!= include('TeamKPI') ?>
//   <?!= include('Enrollment') ?>
//   <?!= include('OnFieldKPI') ?>
//   <?!= include('DeptList') ?>
// ─────────────────────────────────────────
function include(filename) {
  return HtmlService.createHtmlOutputFromFile(filename).getContent();
}

// ─────────────────────────────────────────
// PAGE 1 — Digi-Sales Individual Daily KPI
// ─────────────────────────────────────────
function getData(startDate, endDate) {
  var ss    = SpreadsheetApp.openById("19F_9d62ONF3RSwoI0qM_2H0fbun5IbGrLaj4QYBN-Fo");
  var sheet = ss.getSheetByName("Digi- Sales Induvisual Daily KPI Report");

  if (!sheet) return { error: "Sheet not found", teams: [], grandTotal: getEmptyTotal() };

  var lastRow = sheet.getLastRow();
  if (lastRow <= 1) return { teams: [], grandTotal: getEmptyTotal() };

  var data = sheet.getRange(1, 1, lastRow, 10).getValues();
  data.shift();

  var start = new Date(startDate); start.setHours(0, 0, 0, 0);
  var end   = new Date(endDate);   end.setHours(23, 59, 59, 999);

  var teams = {}, grandTotal = getEmptyTotal();

  data.forEach(function(row) {
    if (!row[0] && !row[1]) return;
    if (String(row[1]).indexOf('{') !== -1 || String(row[1]).indexOf(';') !== -1) return;

    var rowDate = new Date(row[8]);
    if (isNaN(rowDate.getTime())) return;
    if (rowDate < start || rowDate > end) return;

    var sr         = row[0] || "";
    var name       = String(row[1] || "").trim();
    var deposit    = safeNum(row[2]);
    var oldApp     = safeNum(row[3]);
    var newApp     = safeNum(row[4]);
    var withMOU    = safeNum(row[5]);
    var withoutMOU = safeNum(row[6]);
    var status     = String(row[7] || "Office").trim();
    var team       = String(row[9] || "").trim();
    var dateStr    = Utilities.formatDate(rowDate, Session.getScriptTimeZone(), "yyyy-MM-dd");

    if (!team || !name) return;

    if (!teams[team]) teams[team] = { teamName: team, members: {}, totals: getEmptyTotal() };

    var memberKey = sr + "_" + name.toLowerCase();
    if (!teams[team].members[memberKey]) {
      teams[team].members[memberKey] = {
        sr: sr, name: name,
        deposit: 0, oldApp: 0, newApp: 0, withMOU: 0, withoutMOU: 0,
        status: status, date: dateStr
      };
    }

    teams[team].members[memberKey].deposit     += deposit;
    teams[team].members[memberKey].oldApp      += oldApp;
    teams[team].members[memberKey].newApp      += newApp;
    teams[team].members[memberKey].withMOU     += withMOU;
    teams[team].members[memberKey].withoutMOU  += withoutMOU;
    if (status && status !== "Office") teams[team].members[memberKey].status = status;
  });

  var teamOrder   = ["Julie","Jinisha","Palak","Davish","Smriti"];
  var sortedTeams = Object.values(teams).sort(function(a, b) {
    var ai = teamOrder.findIndex(function(t){ return a.teamName.toLowerCase().indexOf(t.toLowerCase()) !== -1; });
    var bi = teamOrder.findIndex(function(t){ return b.teamName.toLowerCase().indexOf(t.toLowerCase()) !== -1; });
    if (ai === -1) ai = 999;
    if (bi === -1) bi = 999;
    if (ai !== bi) return ai - bi;
    return a.teamName.localeCompare(b.teamName);
  });

  grandTotal = getEmptyTotal();
  sortedTeams.forEach(function(team) {
    team.members = Object.values(team.members);
    team.members.sort(function(a, b){ return Number(a.sr) - Number(b.sr); });
    team.totals  = getEmptyTotal();
    team.members.forEach(function(m) {
      team.totals.deposit    += m.deposit;
      team.totals.oldApp     += m.oldApp;
      team.totals.newApp     += m.newApp;
      team.totals.withMOU    += m.withMOU;
      team.totals.withoutMOU += m.withoutMOU;
    });
    grandTotal.deposit    += team.totals.deposit;
    grandTotal.oldApp     += team.totals.oldApp;
    grandTotal.newApp     += team.totals.newApp;
    grandTotal.withMOU    += team.totals.withMOU;
    grandTotal.withoutMOU += team.totals.withoutMOU;
  });

  return { teams: sortedTeams, grandTotal: grandTotal };
}

// ─────────────────────────────────────────
// PAGE 2 — Digi-Sales Team Daily KPI Report
// ─────────────────────────────────────────
function getTeamData() {
  var ss    = SpreadsheetApp.openById("19F_9d62ONF3RSwoI0qM_2H0fbun5IbGrLaj4QYBN-Fo");
  var sheet = ss.getSheetByName("Digi- Sales Team Daily KPI Report");

  if (!sheet) return { error: "Sheet not found", rows: [], total: getEmptyTotal() };

  var lastRow = sheet.getLastRow();
  if (lastRow < 4) return { rows: [], total: getEmptyTotal() };

  var dateCell  = sheet.getRange(1, 1).getValue();
  var dataRange = sheet.getRange(4, 1, lastRow - 3, 6).getValues();

  var rows = [], total = getEmptyTotal();

  dataRange.forEach(function(row) {
    var teamName = String(row[0] || "").trim();
    if (!teamName || teamName.toLowerCase() === "total") return;

    var deposit    = safeNum(row[1]);
    var oldApp     = safeNum(row[2]);
    var newApp     = safeNum(row[3]);
    var withMOU    = safeNum(row[4]);
    var withoutMOU = safeNum(row[5]);

    rows.push({ teamName: teamName, deposit: deposit, oldApp: oldApp, newApp: newApp, withMOU: withMOU, withoutMOU: withoutMOU });

    total.deposit    += deposit;
    total.oldApp     += oldApp;
    total.newApp     += newApp;
    total.withMOU    += withMOU;
    total.withoutMOU += withoutMOU;
  });

  return { date: String(dateCell), rows: rows, total: total };
}

// ─────────────────────────────────────────
// PAGE 3 — Enrollment Sheet (no date = all)
// ─────────────────────────────────────────
function getEnrollmentData(startDate, endDate) {
  var ss    = SpreadsheetApp.openById("19F_9d62ONF3RSwoI0qM_2H0fbun5IbGrLaj4QYBN-Fo");
  var sheet = ss.getSheetByName("Enrollment sheet");

  if (!sheet) return { error: "Sheet not found", headers: [], rows: [], total: { active: 0, refund: 0, defer: 0, total: 0 } };

  var lastRow = sheet.getLastRow();
  var lastCol = sheet.getLastColumn();
  if (lastRow <= 1) return { headers: [], rows: [], total: { active: 0, refund: 0, defer: 0, total: 0 } };

  var useFilter = (startDate && endDate && String(startDate).trim() !== "" && String(endDate).trim() !== "");
  var start, end;
  if (useFilter) {
    start = new Date(startDate); start.setHours(0, 0, 0, 0);
    end   = new Date(endDate);   end.setHours(23, 59, 59, 999);
  }

  var headerRow = sheet.getRange(1, 1, 1, lastCol).getValues()[0];
  var headers   = headerRow.map(function(h){ return String(h || "").trim(); });

  function colIdx(keywords) {
    for (var i = 0; i < headers.length; i++) {
      var h = headers[i].toLowerCase();
      for (var k = 0; k < keywords.length; k++) {
        if (h.indexOf(keywords[k]) !== -1) return i;
      }
    }
    return -1;
  }

  var payDateCol  = colIdx(["payment date","pay date","paydate"]);
  var statusCol   = colIdx(["student status","status"]);
  var rmCol       = colIdx(["b2b rm","rm"]);
  var teamCol     = colIdx(["team leader","team"]);
  var onboardCol  = colIdx(["onboarded by","onboardedby","onboarded"]);
  var managedCol  = colIdx(["managed by","managedby","managed"]);

  if (payDateCol === -1) payDateCol = 3;
  if (statusCol  === -1) statusCol  = 1;
  if (rmCol      === -1) rmCol      = 6;
  if (teamCol    === -1) teamCol    = 7;

  var data   = sheet.getRange(2, 1, lastRow - 1, lastCol).getValues();
  var rows   = [];
  var active = 0, refund = 0, defer = 0;

  data.forEach(function(row) {
    var hasData = false;
    for (var ci = 0; ci < row.length; ci++) { if (row[ci] !== "") { hasData = true; break; } }
    if (!hasData) return;

    if (useFilter) {
      var payDate = new Date(row[payDateCol]);
      if (isNaN(payDate.getTime())) return;
      if (payDate < start || payDate > end) return;
    }

    var status = String(row[statusCol] || "").trim();
    var sl     = status.toLowerCase();
    if (sl.indexOf("active")  !== -1) active++;
    else if (sl.indexOf("refund") !== -1) refund++;
    else if (sl.indexOf("defer")  !== -1) defer++;

    var cells = row.map(function(c) {
      if (c instanceof Date) {
        return isNaN(c.getTime()) ? "" : Utilities.formatDate(c, Session.getScriptTimeZone(), "dd-MMM-yyyy");
      }
      return String(c || "").trim();
    });

    rows.push({
      cells:       cells,
      status:      status,
      rm:          String(row[rmCol]         || "").trim(),
      team:        String(row[teamCol]       || "").trim(),
      onboardedBy: onboardCol !== -1 ? String(row[onboardCol] || "").trim() : "",
      managedBy:   managedCol !== -1 ? String(row[managedCol] || "").trim() : ""
    });
  });

  return {
    headers: headers,
    rows:    rows,
    total:   { active: active, refund: refund, defer: defer, total: rows.length }
  };
}

// ─────────────────────────────────────────
// PAGE 4 — On-Field Team Daily KPI Report
// ─────────────────────────────────────────
function getOnFieldData(startDate, endDate) {
  var ss    = SpreadsheetApp.openById("19F_9d62ONF3RSwoI0qM_2H0fbun5IbGrLaj4QYBN-Fo");
  var sheet = ss.getSheetByName("On- Field Team Daily KPI Report");

  if (!sheet) return { error: "Sheet not found", teams: [], grandTotal: getEmptyTotal() };

  var lastRow = sheet.getLastRow();
  if (lastRow <= 1) return { teams: [], grandTotal: getEmptyTotal() };

  var start = new Date(startDate); start.setHours(0, 0, 0, 0);
  var end   = new Date(endDate);   end.setHours(23, 59, 59, 999);

  var data = sheet.getRange(1, 1, lastRow, 10).getValues();
  data.shift();

  var teams = {}, grandTotal = getEmptyTotal();

  data.forEach(function(row) {
    if (!row[0] && !row[1]) return;
    if (String(row[1]).indexOf('{') !== -1 || String(row[1]).indexOf(';') !== -1) return;

    var rowDate = new Date(row[8]);
    if (isNaN(rowDate.getTime())) return;
    if (rowDate < start || rowDate > end) return;

    var sr         = row[0] || "";
    var name       = String(row[1] || "").trim();
    var deposit    = safeNum(row[2]);
    var oldApp     = safeNum(row[3]);
    var newApp     = safeNum(row[4]);
    var withMOU    = safeNum(row[5]);
    var withoutMOU = safeNum(row[6]);
    var status     = String(row[7] || "Office").trim();
    var team       = String(row[9] || "").trim();
    var dateStr    = Utilities.formatDate(rowDate, Session.getScriptTimeZone(), "yyyy-MM-dd");

    if (!team || !name) return;

    if (!teams[team]) teams[team] = { teamName: team, members: {}, totals: getEmptyTotal() };

    var key = sr + "_" + name.toLowerCase();
    if (!teams[team].members[key]) {
      teams[team].members[key] = { sr: sr, name: name, deposit: 0, oldApp: 0, newApp: 0, withMOU: 0, withoutMOU: 0, status: status, date: dateStr };
    }

    teams[team].members[key].deposit     += deposit;
    teams[team].members[key].oldApp      += oldApp;
    teams[team].members[key].newApp      += newApp;
    teams[team].members[key].withMOU     += withMOU;
    teams[team].members[key].withoutMOU  += withoutMOU;
    if (status && status !== "Office") teams[team].members[key].status = status;
  });

  grandTotal = getEmptyTotal();
  var sortedTeams = Object.values(teams).sort(function(a, b){ return a.teamName.localeCompare(b.teamName); });

  sortedTeams.forEach(function(team) {
    team.members = Object.values(team.members);
    team.members.sort(function(a, b){ return Number(a.sr) - Number(b.sr); });
    team.totals  = getEmptyTotal();
    team.members.forEach(function(m) {
      team.totals.deposit    += m.deposit;
      team.totals.oldApp     += m.oldApp;
      team.totals.newApp     += m.newApp;
      team.totals.withMOU    += m.withMOU;
      team.totals.withoutMOU += m.withoutMOU;
    });
    grandTotal.deposit    += team.totals.deposit;
    grandTotal.oldApp     += team.totals.oldApp;
    grandTotal.newApp     += team.totals.newApp;
    grandTotal.withMOU    += team.totals.withMOU;
    grandTotal.withoutMOU += team.totals.withoutMOU;
  });

  return { teams: sortedTeams, grandTotal: grandTotal };
}

// ─────────────────────────────────────────
// PAGE 5 — Team List / Department Directory
// ─────────────────────────────────────────
function getTeamList() {
  var ss        = SpreadsheetApp.openById("19F_9d62ONF3RSwoI0qM_2H0fbun5IbGrLaj4QYBN-Fo");
  var allSheets = ss.getSheets();
  var sheet     = ss.getSheetByName("Team List");

  if (!sheet) {
    for (var i = 0; i < allSheets.length; i++) {
      if (allSheets[i].getName().toLowerCase().indexOf("team list") !== -1) { sheet = allSheets[i]; break; }
    }
  }
  if (!sheet) {
    for (var j = 0; j < allSheets.length; j++) {
      if (allSheets[j].getSheetId() === 747273934) { sheet = allSheets[j]; break; }
    }
  }
  if (!sheet) {
    var names = allSheets.map(function(s){ return s.getName(); }).join(", ");
    return { error: "Sheet not found. Available: " + names, digiSales: [], smriti: [], onField: [] };
  }

  var lastRow = sheet.getLastRow();
  if (lastRow <= 1) return { digiSales: [], smriti: [], onField: [] };

  var data      = sheet.getRange(2, 1, lastRow - 1, 9).getValues();
  var digiSales = [], smriti = [], onField = [];

  data.forEach(function(row) {
    if (!row[1]) return;
    var member = {
      sr: row[0] || "", name: String(row[1] || "").trim(),
      crmId: String(row[2] || "").trim(), designation: String(row[3] || "").trim(),
      team: String(row[4] || "").trim(), department: String(row[5] || "").trim(),
      email: String(row[6] || "").trim(), contact: String(row[7] || "").trim(),
      status: String(row[8] || "").trim()
    };

    var dept = member.department.toLowerCase();
    var team = member.team.toLowerCase();

    if (team.indexOf("smriti") !== -1) smriti.push(member);
    else if (dept.indexOf("on") !== -1 && dept.indexOf("field") !== -1) onField.push(member);
    else digiSales.push(member);
  });

  return { digiSales: digiSales, smriti: smriti, onField: onField };
}

// ─────────────────────────────────────────
// AI ANALYSIS — Gemini API
// ─────────────────────────────────────────
function getAIAnalysis(context, history) {
  var API_KEY = PropertiesService.getScriptProperties().getProperty('GEMINI_API_KEY');

  if (!API_KEY || API_KEY === 'APNI_GEMINI_API_KEY_YAHAN_DAALO') {
    return "⚠️ API Key setup nahi hui. Apps Script Editor mein 'setupProperties' function run karo.";
  }

  var contents = [];
  contents.push({
    role: "user",
    parts: [{ text: "Tu ek CRM KPI Dashboard ka AI assistant hai. Sirf Hindi/Hinglish mein jawab de — friendly aur professional tone mein. Yeh current data hai:\n\n" + context + "\n\nIs data ke basis pe sawaalon ke jawab de." }]
  });
  contents.push({
    role: "model",
    parts: [{ text: "Namaste! Main CRM KPI data samajh gaya hoon. Koi bhi analysis ya sawaal poochho! 😊" }]
  });

  if (history && history.length > 0) {
    history.forEach(function(msg) {
      if (msg.role === "user" || msg.role === "assistant") {
        contents.push({
          role: msg.role === "assistant" ? "model" : "user",
          parts: [{ text: msg.content }]
        });
      }
    });
  }

  var url     = "https://generativelanguage.googleapis.com/v1beta/models/gemini-1.5-flash:generateContent?key=" + API_KEY;
  var payload = {
    contents: contents,
    generationConfig: { temperature: 0.7, maxOutputTokens: 800, topP: 0.9 },
    safetySettings: [
      { category: "HARM_CATEGORY_HARASSMENT",        threshold: "BLOCK_NONE" },
      { category: "HARM_CATEGORY_HATE_SPEECH",       threshold: "BLOCK_NONE" },
      { category: "HARM_CATEGORY_SEXUALLY_EXPLICIT", threshold: "BLOCK_NONE" },
      { category: "HARM_CATEGORY_DANGEROUS_CONTENT", threshold: "BLOCK_NONE" }
    ]
  };

  try {
    var response = UrlFetchApp.fetch(url, {
      method: "post", contentType: "application/json",
      payload: JSON.stringify(payload), muteHttpExceptions: true
    });

    var code   = response.getResponseCode();
    var result = JSON.parse(response.getContentText());

    if (code !== 200) {
      return "❌ API Error (" + code + "): " + (result.error ? result.error.message : "Unknown error");
    }

    if (result.candidates && result.candidates[0] && result.candidates[0].content) {
      return result.candidates[0].content.parts[0].text;
    }

    return "⚠️ Response parse nahi hua. Please dobara try karo.";
  } catch (e) {
    return "❌ Network error: " + e.message;
  }
}

// ─────────────────────────────────────────
// On-Field RM Count (Pivot Modal — Page 3)
// ─────────────────────────────────────────
function getOnFieldRMDepositCount(startDate, endDate) {
  var ss        = SpreadsheetApp.openById("19F_9d62ONF3RSwoI0qM_2H0fbun5IbGrLaj4QYBN-Fo");
  var teamSheet = ss.getSheetByName("Team List");
  if (!teamSheet) return { error: "Team List sheet not found", rows: [] };

  var tlLastRow  = teamSheet.getLastRow();
  var tlData     = teamSheet.getRange(2, 1, tlLastRow - 1, 9).getValues();
  var rmList     = [];
  var onFieldSet = {};

  tlData.forEach(function(row) {
    var name  = String(row[1] || "").trim();
    var dept  = String(row[5] || "").trim().toLowerCase();
    var team  = String(row[4] || "").trim();
    if (!name) return;
    if (dept.indexOf("on") === -1 && dept.indexOf("field") === -1) return;
    rmList.push({ name: name, team: team, desig: String(row[3] || "").trim() });
    onFieldSet[name.toLowerCase()] = name;
  });

  var enSheet = ss.getSheetByName("Enrollment sheet");
  if (!enSheet) return { error: "Enrollment sheet not found", rows: [] };

  var enLastRow = enSheet.getLastRow();
  var enLastCol = enSheet.getLastColumn();
  if (enLastRow <= 1) return { rows: rmList.map(function(r){ return { name: r.name, team: r.team, count: 0 }; }) };

  var headers = enSheet.getRange(1, 1, 1, enLastCol).getValues()[0];

  var onboardedByCol = -1, managedByCol = -1, payDateCol = -1;
  headers.forEach(function(h, i) {
    var hl = String(h).toLowerCase().trim();
    if (hl.indexOf("onboarded by") !== -1 || hl === "onboarded") onboardedByCol = i;
    if (hl.indexOf("managed by")   !== -1 || hl === "managed")   managedByCol   = i;
    if (hl.indexOf("payment date") !== -1 || hl.indexOf("pay date") !== -1) payDateCol = i;
  });
  if (payDateCol === -1) payDateCol = 3;

  var useFilter = (startDate && endDate && String(startDate).trim() !== "" && String(endDate).trim() !== "");
  var start, end;
  if (useFilter) {
    start = new Date(startDate); start.setHours(0,  0,  0,   0);
    end   = new Date(endDate);   end.setHours(23, 59, 59, 999);
  }

  var enData   = enSheet.getRange(2, 1, enLastRow - 1, enLastCol).getValues();
  var countMap = {};

  enData.forEach(function(row) {
    if (useFilter) {
      var payDate = new Date(row[payDateCol]);
      if (isNaN(payDate.getTime())) return;
      if (payDate < start || payDate > end) return;
    }

    var ob    = onboardedByCol !== -1 ? String(row[onboardedByCol] || "").trim() : "";
    var mb    = managedByCol   !== -1 ? String(row[managedByCol]   || "").trim() : "";
    var obKey = ob.toLowerCase();
    var mbKey = mb.toLowerCase();

    if (ob && onFieldSet[obKey]) {
      countMap[obKey] = (countMap[obKey] || 0) + 1;
    }
    if (mb && onFieldSet[mbKey] && mbKey !== obKey) {
      countMap[mbKey] = (countMap[mbKey] || 0) + 1;
    }
  });

  var result = rmList.map(function(r) {
    return { name: r.name, team: r.team, desig: r.desig, count: countMap[r.name.toLowerCase()] || 0 };
  });

  result.sort(function(a, b) {
    if (a.count === 0 && b.count === 0) return a.name.localeCompare(b.name);
    if (a.count === 0) return 1;
    if (b.count === 0) return -1;
    return b.count - a.count;
  });

  return { rows: result };
}

// ─────────────────────────────────────────
// Helpers
// ─────────────────────────────────────────
function safeNum(val) {
  var n = Number(val);
  return isNaN(n) ? 0 : n;
}

function getEmptyTotal() {
  return { deposit: 0, oldApp: 0, newApp: 0, withMOU: 0, withoutMOU: 0 };
}

function listSheets() {
  var ss = SpreadsheetApp.openById("19F_9d62ONF3RSwoI0qM_2H0fbun5IbGrLaj4QYBN-Fo");
  Logger.log(ss.getSheets().map(function(s){ return s.getName(); }).join("\n"));
}
