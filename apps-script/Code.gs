// 櫃台推薦紀錄 — Google 試算表後端（Apps Script 網頁應用程式）
// 部署：執行身分＝我，存取權＝任何人。實際保護靠下方 TOKEN（通行碼），前端每次請求都要帶。
// 請把 TOKEN 換成自己的長亂碼；這個值不要放上 GitHub。

const TOKEN = 'CHANGE-ME';
const SHEET_NAME = '紀錄';
const FIELDS = ['id', 'ts', 'date', 'time', 'wd', 'g', 'a', 'r', 'why', 'p'];
const HEADERS = ['id', '時間戳', '日期', '時間', '星期', '性別', '年齡', '推薦結果', '拒絕原因', '來店目的'];

function doGet() {
  return json({ ok: true, service: 'counter-log' });
}

function doPost(e) {
  let req;
  try {
    req = JSON.parse(e.postData.contents);
  } catch (_) {
    return json({ ok: false, error: 'bad_request' });
  }
  if (TOKEN === 'CHANGE-ME' || !req || req.token !== TOKEN) {
    return json({ ok: false, error: 'unauthorized' });
  }

  const lock = LockService.getScriptLock();
  lock.waitLock(10000);
  try {
    const sh = getSheet();
    if (req.action === 'list') {
      return json({ ok: true, rows: readAll(sh), sheetUrl: SpreadsheetApp.getActive().getUrl() });
    }
    if (req.action === 'add') {
      const row = req.row || {};
      if (!row.id) return json({ ok: false, error: 'bad_request' });
      if (findRow(sh, row.id)) return json({ ok: true, dup: true });
      sh.appendRow(FIELDS.map(f => clean(row[f])));
      return json({ ok: true });
    }
    if (req.action === 'delete') {
      const i = findRow(sh, req.id);
      if (i) sh.deleteRow(i);
      return json({ ok: true });
    }
    return json({ ok: false, error: 'unknown_action' });
  } finally {
    lock.releaseLock();
  }
}

function getSheet() {
  const ss = SpreadsheetApp.getActive();
  let sh = ss.getSheetByName(SHEET_NAME);
  if (!sh) {
    sh = ss.insertSheet(SHEET_NAME);
    sh.getRange('A:J').setNumberFormat('@'); // 全部存成文字，避免日期、「70+」被自動轉格式
    sh.appendRow(HEADERS);
    sh.setFrozenRows(1);
  }
  return sh;
}

function readAll(sh) {
  const last = sh.getLastRow();
  if (last < 2) return [];
  return sh.getRange(2, 1, last - 1, FIELDS.length).getValues().map(v => {
    const o = {};
    FIELDS.forEach((f, i) => { o[f] = String(v[i]); });
    return o;
  });
}

function findRow(sh, id) {
  const last = sh.getLastRow();
  if (last < 2 || !id) return 0;
  const ids = sh.getRange(2, 1, last - 1, 1).getValues();
  for (let i = ids.length - 1; i >= 0; i--) {
    if (String(ids[i][0]) === String(id)) return i + 2;
  }
  return 0;
}

// 擋掉會被試算表當成公式的開頭字元
function clean(v) {
  const s = v == null ? '' : String(v).slice(0, 100);
  return /^[=+\-@]/.test(s) ? "'" + s : s;
}

function json(o) {
  return ContentService.createTextOutput(JSON.stringify(o)).setMimeType(ContentService.MimeType.JSON);
}
