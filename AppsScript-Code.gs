/**
 * My Schedule — Google Apps Script Backend
 * -----------------------------------------
 * This script turns your Google Sheet into a simple login + data-sync backend.
 *
 * SETUP INSTRUCTIONS:
 * 1. Open your Google Sheet:
 *    https://docs.google.com/spreadsheets/d/1nycQ2Caj3o2jD-Q1bpB7_We4ZaYh8zF6n99wjo1dCJE/edit
 * 2. Go to: Extensions > Apps Script
 * 3. Delete any starter code in Code.gs, then paste ALL of this file's contents in.
 * 4. Click "Deploy" (top right) > "New deployment"
 * 5. Click the gear icon next to "Select type" > choose "Web app"
 * 6. Set:
 *      - Description: "My Schedule API" (or anything)
 *      - Execute as: "Me"
 *      - Who has access: "Anyone"
 * 7. Click "Deploy". Google will ask you to authorize — approve it (you may
 *    see an "unverified app" warning since it's your own script; click
 *    "Advanced" > "Go to My Schedule API (unsafe)" > Allow).
 * 8. Copy the "Web app URL" it gives you (looks like:
 *    https://script.google.com/macros/s/AKfycb.../exec)
 * 9. Paste that URL into the HTML file, replacing the placeholder:
 *      const API_URL = "PASTE_YOUR_WEB_APP_URL_HERE";
 *
 * NOTE: Every time you edit this script, you must create a NEW deployment
 * (or "Manage deployments" > edit > new version) for changes to take effect.
 */

const SHEET_NAME = "Users";

function doGet(e) {
  return handleRequest(e.parameter);
}

function doPost(e) {
  var params = {};
  try {
    if (e.postData && e.postData.contents) {
      params = JSON.parse(e.postData.contents);
    }
  } catch (err) {
    // fall back to query params if body isn't JSON
    params = e.parameter;
  }
  return handleRequest(params);
}

function getSheet_() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(SHEET_NAME);
  if (!sheet) {
    sheet = ss.insertSheet(SHEET_NAME);
    sheet.appendRow(["username", "passwordHash", "data", "updatedAt"]);
  }
  return sheet;
}

function hashPassword_(pw) {
  var raw = Utilities.computeDigest(Utilities.DigestAlgorithm.SHA_256, pw, Utilities.Charset.UTF_8);
  return raw.map(function (b) {
    var v = (b < 0 ? b + 256 : b).toString(16);
    return v.length === 1 ? "0" + v : v;
  }).join("");
}

function findUserRow_(sheet, username) {
  var values = sheet.getDataRange().getValues();
  for (var i = 1; i < values.length; i++) {
    if (values[i][0] === username) return i + 1; // 1-indexed row number
  }
  return -1;
}

function jsonOut_(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}

function defaultAppData_(username) {
  return {
    profile: {
      name: username,
      nickname: username.charAt(0).toUpperCase(),
      theme: "light",
      anim: true,
      avatarType: "initial",
      avatarValue: ""
    },
    classes: [],
    activities: []
  };
}

function handleRequest(p) {
  var lock = LockService.getScriptLock();
  lock.waitLock(15000);
  try {
    var action = p.action;
    var sheet = getSheet_();

    if (action === "register") {
      var username = (p.username || "").trim();
      var password = p.password || "";
      if (!username || !password) {
        return jsonOut_({ success: false, message: "กรุณากรอกข้อมูลให้ครบ" });
      }
      if (username.length < 3) {
        return jsonOut_({ success: false, message: "ชื่อผู้ใช้ต้องมีอย่างน้อย 3 ตัวอักษร" });
      }
      if (findUserRow_(sheet, username) !== -1) {
        return jsonOut_({ success: false, message: "มีชื่อผู้ใช้นี้อยู่แล้ว กรุณาเลือกชื่ออื่น" });
      }
      var data = defaultAppData_(username);
      sheet.appendRow([username, hashPassword_(password), JSON.stringify(data), new Date().toISOString()]);
      return jsonOut_({ success: true, data: data });
    }

    if (action === "login") {
      var username = (p.username || "").trim();
      var password = p.password || "";
      var row = findUserRow_(sheet, username);
      if (row === -1) {
        return jsonOut_({ success: false, message: "ไม่พบชื่อผู้ใช้นี้ในระบบ" });
      }
      var storedHash = sheet.getRange(row, 2).getValue();
      if (storedHash !== hashPassword_(password)) {
        return jsonOut_({ success: false, message: "รหัสผ่านไม่ถูกต้อง" });
      }
      var raw = sheet.getRange(row, 3).getValue();
      var data;
      try { data = JSON.parse(raw); } catch (e2) { data = defaultAppData_(username); }
      return jsonOut_({ success: true, data: data });
    }

    if (action === "save") {
      var username = (p.username || "").trim();
      var password = p.password || "";
      var row = findUserRow_(sheet, username);
      if (row === -1) {
        return jsonOut_({ success: false, message: "ไม่พบผู้ใช้งานนี้" });
      }
      var storedHash = sheet.getRange(row, 2).getValue();
      if (storedHash !== hashPassword_(password)) {
        return jsonOut_({ success: false, message: "ยืนยันตัวตนไม่สำเร็จ" });
      }
      sheet.getRange(row, 3).setValue(p.data);
      sheet.getRange(row, 4).setValue(new Date().toISOString());
      return jsonOut_({ success: true });
    }

    return jsonOut_({ success: false, message: "ไม่รู้จักคำสั่งนี้ (unknown action)" });
  } finally {
    lock.releaseLock();
  }
}
