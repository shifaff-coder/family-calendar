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
    try { obj.memberIds = JSON.parse(obj.memberIds); } catch { obj.memberIds = []; }
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
        return [id, r[1], r[2], JSON.stringify(ev.memberIds || []), ev.categoryEmoji || r[4], ev.note || r[5]];
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
      return [ev.id, date, ev.title, JSON.stringify(ev.memberIds), ev.categoryEmoji, ev.note || ""];
    });

  const allRows = [...manualRows, ...gcalRows];
  if (allRows.length > 0) {
    sheet.getRange(2, 1, allRows.length, 6).setValues(allRows);
    sheet.getRange(2, 2, allRows.length, 1).setNumberFormat("@");
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

    // 今月分のGoogleカレンダーIDセット
    const currentGcalIds = new Set(
      gEvents.map(ev => "gcal_" + ev.getId().replace(/[^a-zA-Z0-9]/g, "").substring(0, 30))
    );

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

      // 開始日・終了日を日本時間で取得
      const start = ev.getStartTime();
      const end = ev.getEndTime();

      // 終日イベントは終了日が exclusive なので1日引く、時間イベントはそのまま
      const isAllDay = ev.isAllDayEvent();
      const endAdj = isAllDay ? new Date(end.getTime() - 1) : end;

      const startStr = Utilities.formatDate(start, "Asia/Tokyo", "yyyy-MM-dd");
      const endStr   = Utilities.formatDate(endAdj, "Asia/Tokyo", "yyyy-MM-dd");

      // 開始日から終了日まで1日ずつ登録
      const cur = new Date(start);
      cur.setHours(0, 0, 0, 0);
      let dayIndex = 0;
      while (true) {
        const curStr = Utilities.formatDate(cur, "Asia/Tokyo", "yyyy-MM-dd");
        if (curStr > endStr) break;

        // 今月の範囲内のみ登録
        if (curStr.startsWith(yearMonth)) {
          const gcalId = baseId + "_d" + dayIndex;
          if (!existingGcalIds.has(gcalId)) {
            const dayLabel = startStr === endStr ? "" : "（" + (dayIndex + 1) + "日目）";
            newRows.push([gcalId, curStr, title + dayLabel, "[]", "📆", "Googleカレンダーより"]);
          }
        }

        cur.setDate(cur.getDate() + 1);
        dayIndex++;
        if (dayIndex > 31) break; // 無限ループ防止
      }
    });

    if (newRows.length > 0) {
      const lastRow = sheet.getLastRow();
      sheet.getRange(lastRow + 1, 1, newRows.length, 6).setValues(newRows);
      sheet.getRange(lastRow + 1, 2, newRows.length, 1).setNumberFormat("@");
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
