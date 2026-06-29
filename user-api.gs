/**
 * 사용자 관리 API (Google Sheets 기반)
 *
 * 설정 방법:
 * 1. Google Sheets 새로 생성 → 시트 이름을 "users"로 변경
 * 2. 1행에 헤더 입력: email | name | picture | role | status | created
 * 3. 2행에 관리자 시드: tabin8437@gmail.com | HY J | (빈칸) | admin | approved | 2026-06-29
 * 4. 이 스크립트를 확장프로그램 > Apps Script에 붙여넣기
 * 5. 배포 > 새 배포 > 웹 앱 > 액세스: 모든 사용자 > 배포
 * 6. 배포 URL을 index.html의 USER_API_URL에 붙여넣기
 */

var SHEET_NAME = "user";

function getSheet() {
  return SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEET_NAME);
}

function getAllUsers() {
  var sheet = getSheet();
  var data = sheet.getDataRange().getValues();
  if (data.length <= 1) return [];
  var headers = data[0];
  var users = [];
  for (var i = 1; i < data.length; i++) {
    var row = data[i];
    if (!row[0]) continue;
    var user = {};
    for (var j = 0; j < headers.length; j++) {
      user[headers[j]] = row[j] || "";
    }
    users.push(user);
  }
  return users;
}

function findUserRow(email) {
  var sheet = getSheet();
  var data = sheet.getDataRange().getValues();
  for (var i = 1; i < data.length; i++) {
    if (data[i][0] && data[i][0].toString().toLowerCase() === email.toLowerCase()) {
      return i + 1;
    }
  }
  return -1;
}

function decB64(s){try{return Utilities.newBlob(Utilities.base64Decode(s)).getDataAsString('UTF-8');}catch(e){return s;}}

function doGet(e) {
  var p = (e && e.parameter) ? e.parameter : {};
  var action = (p.action || "list");
  var callback = p.callback || "";
  if(p.name) p.name = decB64(p.name);
  var result = {};

  if (action === "list") {
    result = { ok: true, users: getAllUsers() };

  } else if (action === "check") {
    var email = (p.email || "").toLowerCase();
    var users = getAllUsers();
    var found = null;
    for (var i = 0; i < users.length; i++) {
      if (users[i].email.toLowerCase() === email) { found = users[i]; break; }
    }
    result = { ok: true, user: found };

  } else if (action === "register") {
    var email = (p.email || "").toLowerCase();
    var row = findUserRow(email);
    if (row > 0) {
      result = { ok: false, msg: "already_exists" };
    } else {
      var sheet = getSheet();
      sheet.appendRow([email, p.name || email, p.picture || "", "viewer", "pending", new Date().toISOString()]);
      result = { ok: true, status: "pending" };
    }

  } else if (action === "update") {
    var email = (p.email || "").toLowerCase();
    var row = findUserRow(email);
    if (row < 0) {
      result = { ok: false, msg: "not_found" };
    } else {
      var sheet = getSheet();
      if (p.name) sheet.getRange(row, 2).setValue(p.name);
      if (p.picture) sheet.getRange(row, 3).setValue(p.picture);
      if (p.role) sheet.getRange(row, 4).setValue(p.role);
      if (p.status) sheet.getRange(row, 5).setValue(p.status);
      result = { ok: true };
    }

  } else if (action === "add") {
    var email = (p.email || "").toLowerCase();
    var row = findUserRow(email);
    if (row > 0) {
      result = { ok: false, msg: "already_exists" };
    } else {
      var sheet = getSheet();
      sheet.appendRow([email, p.name || email, "", p.role || "viewer", "approved", new Date().toISOString()]);
      result = { ok: true };
    }

  } else if (action === "log") {
    var email = (p.email || "").toLowerCase();
    var name = p.name || email;
    var logSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("logs");
    if (logSheet) {
      logSheet.appendRow([new Date().toISOString(), email, name, p.device || "", p.page || "login"]);
    }
    result = { ok: true };

  } else if (action === "logs") {
    var logSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("logs");
    var logs = [];
    if (logSheet && logSheet.getLastRow() > 0) {
      var data = logSheet.getDataRange().getValues();
      var start = Math.max(0, data.length - 100);
      for (var i = data.length - 1; i >= start; i--) {
        logs.push({time: data[i][0], email: data[i][1], name: data[i][2], device: data[i][3], page: data[i][4]});
      }
    }
    result = { ok: true, logs: logs };

  } else if (action === "delete") {
    var email = (p.email || "").toLowerCase();
    var row = findUserRow(email);
    if (row < 0) {
      result = { ok: false, msg: "not_found" };
    } else {
      var sheet = getSheet();
      sheet.deleteRow(row);
      result = { ok: true };
    }
  }

  var json = JSON.stringify(result);
  if (callback) {
    return ContentService.createTextOutput(callback + '(' + json + ')')
      .setMimeType(ContentService.MimeType.JAVASCRIPT);
  }
  return ContentService.createTextOutput(json)
    .setMimeType(ContentService.MimeType.JSON);
}
