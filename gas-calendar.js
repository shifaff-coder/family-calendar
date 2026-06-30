const SHEET_NAME = "予定";
const IMG_SHEET_NAME = "画像";
const LOG_SHEET_NAME = "アクセスログ";

// ── シート取得・作成 ──
function getSheet(name) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName(name);
  if (!sheet) {
    sheet = ss.insertSheet(name);
    if (name === IMG_SHEET_NAME) {
      sheet.getRange(1,1,1,4).setValues([["id","yearMonth","filename","url"]]);
    }
    if (name === LOG_SHEET_NAME) {
      sheet.getRange(1,1,1,6).setValues([["日時","デバイスID","ブラウザ","画面サイズ","言語","初回"]]);
    }
  }
  return sheet;
}

// ── GETリクエスト ──
function doGet(e) {
  try {
    const action = e.parameter.action || "getEvents";
    if (action === "getEvents")          return getEvents();
    if (action === "getImages")          return getImages(e.parameter.yearMonth);
    if (action === "syncGoogleCalendar") return syncGoogleCalendar(e.parameter.yearMonth);
    return jsonResponse({ status: "error", message: "unknown action" });
  } catch(err) {
    return jsonResponse({ status: "error", message: err.message });
  }
}

// ── POSTリクエスト ──
function doPost(e) {
  try {
    const body = JSON.parse(e.postData.contents);
    const action = body.action;
    if (action === "saveEvents")  return saveEvents(body.events);
    if (action === "saveImage")   return saveImage(body);
    if (action === "deleteImage") return deleteImage(body.id);
    if (action === "accessLog")   return saveAccessLog(body);
    return jsonResponse({ status: "error", message: "unknown action" });
  } catch(err) {
    return jsonResponse({ status: "error", message: err.toString() });
  }
}

// ── 予定取得 ──
function getEvents() {
  const sheet = getSheet(SHEET_NAME);
  const data = sheet.getDataRange().getValues();
  if (data.length < 2) return jsonResponse({ status: "ok", events: [] });
  const headers = data[0];
  const rows = data.slice(1).filter(r => r[0] !== "");
  const events = rows.map(row => {
    const obj = {};
    headers.forEach((h, i) => obj[h] = row[i]);
    // シートのヘッダーは小文字 "memberids"。配列化して memberIds に詰め直す。
    try { obj.memberIds = JSON.parse(obj.memberids); } catch { obj.memberIds = []; }
    if (!Array.isArray(obj.memberIds)) obj.memberIds = [];
    return obj;
  });
  return jsonResponse({ status: "ok", events });
}

// ── 予定保存 ──
function saveEvents(events) {
  const sheet = getSheet(SHEET_NAME);

  // クライアントから送られたgcal_イベントをマップ化（メンバー設定などの変更を保持）
  const clientGcalMap = {};
  (events || []).filter(ev => String(ev.id).startsWith("gcal_")).forEach(ev => {
    clientGcalMap[String(ev.id)] = ev;
  });

  // シートのgcal_行を取得し、クライアントの変更があれば上書き
  const data = sheet.getDataRange().getValues();
  const gcalRows = data.slice(1)
    .filter(r => String(r[0]).startsWith("gcal_"))
    .map(r => {
      const id = String(r[0]);
      if (clientGcalMap[id]) {
        const ev = clientGcalMap[id];
        let memberIds = ev.memberIds || [];
        if (!Array.isArray(memberIds)) {
          try { memberIds = JSON.parse(memberIds); } catch { memberIds = []; }
        }
        return [id, r[1], r[2], JSON.stringify(memberIds), ev.categoryEmoji || r[4], ev.note || r[5]];
      }
      return r;
    });

  // 全データをクリア
  const lastRow = sheet.getLastRow();
  if (lastRow > 1) sheet.getRange(2, 1, lastRow - 1, 6).clearContent();

  // 手動入力の予定を書き込み
  const manualRows = (events || [])
    .filter(ev => !String(ev.id).startsWith("gcal_"))
    .map(ev => {
      let date = String(ev.date || "");
      if (date.includes("T")) {
        const d = new Date(date);
        const jst = new Date(d.getTime() + 9 * 60 * 60 * 1000);
        date = jst.toISOString().substring(0, 10);
      } else {
        date = date.substring(0, 10);
      }
      // memberIds を配列として受け取り、JSON文字列で保存
      let memberIds = ev.memberIds;
      if (!Array.isArray(memberIds)) {
        try { memberIds = JSON.parse(memberIds); } catch { memberIds = []; }
      }
      // ID は必ず文字列で書き込む（数値化による破損防止）
      return [String(ev.id), date, ev.title, JSON.stringify(memberIds), ev.categoryEmoji, ev.note || ""];
    });

  const allRows = [...manualRows, ...gcalRows];
  if (allRows.length > 0) {
    sheet.getRange(2, 1, allRows.length, 6).setValues(allRows);
    sheet.getRange(2, 1, allRows.length, 1).setNumberFormat("@"); // ID列
    sheet.getRange(2, 2, allRows.length, 1).setNumberFormat("@"); // 日付列
  }

  return jsonResponse({ status: "ok" });
}

// ── Googleカレンダー同期（追加・削除対応）──
function syncGoogleCalendar(yearMonth) {
  try {
    const [year, month] = yearMonth.split("-").map(Number);
    const startDate = new Date(year, month - 1, 1);
    const endDate = new Date(year, month, 1);

    const cal = CalendarApp.getDefaultCalendar();
    const gEvents = cal.getEvents(startDate, endDate);

    const sheet = getSheet(SHEET_NAME);
    const data = sheet.getDataRange().getValues();

    // 今月のgcal_行をすべて削除（ID形式の新旧混在を防ぐため全削除→再追加方式）
    for (let i = data.length - 1; i >= 1; i--) {
      const id = String(data[i][0]);
      if (!id.startsWith("gcal_")) continue;

      const rawDate = data[i][1];
      let dateStr = rawDate instanceof Date
        ? Utilities.formatDate(rawDate, "Asia/Tokyo", "yyyy-MM-dd")
        : String(rawDate).substring(0, 10);

      if (dateStr.startsWith(yearMonth)) {
        sheet.deleteRow(i + 1);
      }
    }

    // 全件再追加
    const existingGcalIds = new Set(); // 全削除後なので常に空

    const newRows = [];
    gEvents.forEach(ev => {
      const baseId = "gcal_" + ev.getId().replace(/[^a-zA-Z0-9]/g, "").substring(0, 28);
      const title = ev.getTitle() || "（無題）";

      const start = ev.getStartTime();
      const end = ev.getEndTime();

      const isAllDay = ev.isAllDayEvent();
      const endAdj = isAllDay ? new Date(end.getTime() - 1) : end;

      const startStr = Utilities.formatDate(start, "Asia/Tokyo", "yyyy-MM-dd");
      const endStr   = Utilities.formatDate(endAdj, "Asia/Tokyo", "yyyy-MM-dd");

      const cur = new Date(start);
      cur.setHours(0, 0, 0, 0);
      let dayIndex = 0;
      while (true) {
        const curStr = Utilities.formatDate(cur, "Asia/Tokyo", "yyyy-MM-dd");
        if (curStr > endStr) break;

        if (curStr.startsWith(yearMonth)) {
          const gcalId = baseId + "_d" + dayIndex;
          if (!existingGcalIds.has(gcalId)) {
            const dayLabel = startStr === endStr ? "" : "（" + (dayIndex + 1) + "日目）";
            newRows.push([gcalId, curStr, title + dayLabel, "[]", "📆", "Googleカレンダーより"]);
          }
        }

        cur.setDate(cur.getDate() + 1);
        dayIndex++;
        if (dayIndex > 31) break;
      }
    });

    if (newRows.length > 0) {
      const lastRow = sheet.getLastRow();
      sheet.getRange(lastRow + 1, 1, newRows.length, 6).setValues(newRows);
      sheet.getRange(lastRow + 1, 1, newRows.length, 1).setNumberFormat("@"); // ID列
      sheet.getRange(lastRow + 1, 2, newRows.length, 1).setNumberFormat("@"); // 日付列
    }

    return jsonResponse({ status: "ok", imported: newRows.length });
  } catch(err) {
    return jsonResponse({ status: "error", message: err.toString() });
  }
}

// ── 画像取得 ──
function getImages(yearMonth) {
  const sheet = getSheet(IMG_SHEET_NAME);
  const data = sheet.getDataRange().getValues();
  if (data.length < 2) return jsonResponse({ status: "ok", images: [] });
  const rows = data.slice(1).filter(r => r[0] !== "" && r[1] === yearMonth);
  const images = rows.map(row => ({
    id: String(row[0]),
    yearMonth: row[1],
    filename: row[2],
    url: row[3],
  }));
  return jsonResponse({ status: "ok", images });
}

// ── 画像保存 ──
function saveImage(body) {
  const sheet = getSheet(IMG_SHEET_NAME);
  const data = sheet.getDataRange().getValues();
  const count = data.slice(1).filter(r => r[0] !== "" && r[1] === body.yearMonth).length;
  if (count >= 10) return jsonResponse({ status: "error", message: "この月の画像は最大10枚までです" });
  const lastRow = sheet.getLastRow();
  sheet.getRange(lastRow + 1, 1, 1, 4).setValues([[
    body.id, body.yearMonth, body.filename, body.url
  ]]);
  return jsonResponse({ status: "ok" });
}

// ── 画像削除 ──
function deleteImage(id) {
  const sheet = getSheet(IMG_SHEET_NAME);
  const data = sheet.getDataRange().getValues();
  for (let i = 1; i < data.length; i++) {
    if (String(data[i][0]) === String(id)) {
      sheet.deleteRow(i + 1);
      return jsonResponse({ status: "ok" });
    }
  }
  return jsonResponse({ status: "error", message: "not found" });
}

// ── アクセスログ ──
function saveAccessLog(body) {
  const sheet = getSheet(LOG_SHEET_NAME);
  const now = Utilities.formatDate(new Date(), "Asia/Tokyo", "yyyy-MM-dd HH:mm:ss");
  sheet.appendRow([
    now,
    String(body.deviceId || ""),
    String(body.userAgent || "").substring(0, 100),
    String(body.screen || ""),
    String(body.language || ""),
    body.isNew ? "初回" : ""
  ]);

  const data = sheet.getDataRange().getValues();
  const rows = data.slice(1).filter(r => r[1] === body.deviceId);
  const isKnown = rows.length > 1;
  const knownDevices = [...new Set(data.slice(1).map(r => r[1]).filter(Boolean))];

  return jsonResponse({ status: "ok", isKnownDevice: isKnown, knownDevices });
}

// ── テスト用 ──
function testSync() {
  const [year, month] = ["2026", "06"].map(Number);
  const startDate = new Date(year, month - 1, 1);
  const endDate = new Date(year, month, 1);
  const cal = CalendarApp.getDefaultCalendar();
  Logger.log("カレンダー名: " + cal.getName());
  const gEvents = cal.getEvents(startDate, endDate);
  Logger.log("取得した予定数: " + gEvents.length);
  gEvents.forEach(ev => Logger.log(ev.getStartTime() + " / " + ev.getTitle()));
}

// ── レスポンス ──
function jsonResponse(obj) {
  return ContentService
    .createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}

// ====================================================================
// NocoDB 連携
// ====================================================================

const NOCODB_BASE_URL = "https://shif-pc-da970mab.tail052499.ts.net";
const NOCODB_TOKEN = "nc_pat_ermieoZvZBmK4D0wzFZzJxqirX2tvr9l9LQPCsD6";
const NOCODB_MEMBER_TABLE_ID = "m8vphzapdjk158t";
const NOCODB_EVENT_TABLE_ID = "mbe4zeu0cah6srv";

function checkMemberIds() {
  const url = `${NOCODB_BASE_URL}/api/v2/tables/${NOCODB_MEMBER_TABLE_ID}/records?limit=50`;
  const res = UrlFetchApp.fetch(url, {
    method: "get",
    headers: { "xc-token": NOCODB_TOKEN },
    muteHttpExceptions: true
  });
  Logger.log(res.getContentText());
}

const MEMBER_ID_MAP = {
  papa: 1,
  mama: 2,
  mari: 3,
  aki: 4,
  hiro: 5,
  iku: 6
};

function nocodbRequest(method, path, payload) {
  const options = {
    method: method,
    headers: {
      "xc-token": NOCODB_TOKEN,
      "Content-Type": "application/json"
    },
    muteHttpExceptions: true
  };
  if (payload !== undefined) options.payload = JSON.stringify(payload);
  const res = UrlFetchApp.fetch(`${NOCODB_BASE_URL}${path}`, options);
  const code = res.getResponseCode();
  if (code >= 300) {
    throw new Error(`NocoDB API error ${code}: ${res.getContentText()}`);
  }
  return res.getContentText() ? JSON.parse(res.getContentText()) : null;
}

function syncToNocoDB() {
  const sheet = getSheet(SHEET_NAME);
  const data = sheet.getDataRange().getValues();
  if (data.length < 2) {
    Logger.log("予定シートにデータがありません。同期をスキップします。");
    return;
  }
  const headers = data[0];
  const rows = data.slice(1).filter(r => r[0] !== "");
  const events = rows.map(row => {
    const obj = {};
    headers.forEach((h, i) => obj[h] = row[i]);
    try { obj.memberIds = JSON.parse(obj.memberids); } catch { obj.memberIds = []; }
    if (!Array.isArray(obj.memberIds)) obj.memberIds = [];
    return obj;
  });

  const listPath = `/api/v2/tables/${NOCODB_EVENT_TABLE_ID}/records?limit=1000`;
  const existing = nocodbRequest("get", listPath);
  const existingIds = (existing.list || []).map(r => r.Id);

  if (existingIds.length > 0) {
    const deletePayload = existingIds.map(id => ({ Id: id }));
    nocodbRequest("delete", `/api/v2/tables/${NOCODB_EVENT_TABLE_ID}/records`, deletePayload);
  }

  const createPayload = events.map(ev => {
    const isGcal = String(ev.id).startsWith("gcal_");
    const memo = [ev.categoryEmoji || "", ev.note || ""].filter(Boolean).join(" ");
    return {
      "タイトル": ev.title || "",
      "開始時間": ev.date || null,
      "終了日時": ev.date || null,
      "種別": "",
      "場所": "",
      "出典": isGcal ? "Googleカレンダー" : "朝日カレンダー",
      "メモ": memo
    };
  });

  if (createPayload.length === 0) {
    Logger.log("作成対象の予定がありません。");
    return;
  }

  const created = nocodbRequest("post", `/api/v2/tables/${NOCODB_EVENT_TABLE_ID}/records`, createPayload);

  created.forEach((createdRecord, idx) => {
    const ev = events[idx];
    const memberIds = (ev.memberIds || [])
      .map(code => MEMBER_ID_MAP[code])
      .filter(id => id !== undefined);

    Logger.log(`[link debug] event="${ev.title}" createdId=${createdRecord.Id} memberCodes=${JSON.stringify(ev.memberIds)} memberIds=${JSON.stringify(memberIds)}`);

    if (memberIds.length === 0) return;

    const linkPath = `/api/v2/tables/${NOCODB_EVENT_TABLE_ID}/links/csyxmyszs5b9euw/records/${createdRecord.Id}`;
    const linkPayload = memberIds.map(id => ({ Id: id }));
    const linkResult = nocodbRequest("post", linkPath, linkPayload);
    Logger.log(`[link debug] linkResult=${JSON.stringify(linkResult)}`);
  });

  Logger.log(`同期完了: ${createPayload.length}件の予定をNocoDBに反映しました。`);
}

function setupSyncTrigger() {
  ScriptApp.getProjectTriggers().forEach(t => {
    if (t.getHandlerFunction() === "syncToNocoDB") ScriptApp.deleteTrigger(t);
  });
  ScriptApp.newTrigger("syncToNocoDB")
    .timeBased()
    .atHour(1)
    .everyDays(1)
    .create();
  Logger.log("深夜1時の同期トリガーを設定しました。");
}

function checkEventFields() {
  const url = `${NOCODB_BASE_URL}/api/v2/meta/tables/${NOCODB_EVENT_TABLE_ID}`;
  const res = UrlFetchApp.fetch(url, {
    method: "get",
    headers: { "xc-token": NOCODB_TOKEN },
    muteHttpExceptions: true
  });
  const data = JSON.parse(res.getContentText());
  (data.columns || []).forEach(c => {
    Logger.log(`title="${c.title}" id=${c.id} type=${c.uidt}`);
  });
}
