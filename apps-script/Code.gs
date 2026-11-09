/**
 * 孟子文言閱讀挑戰：成績後端（Google Apps Script）
 * 用途：把學生成績存到這份 Google 試算表，讓網頁能顯示全班累計成績。
 * 設定步驟請看 README.md。
 */

// ★ 請改成你自己的教師密碼（用來登入老師檢視、清除本班紀錄）
const TEACHER_PASSWORD = "change-me";

const SHEET_RECORDS = "成績";      // 每位學生每篇一列（程式使用，請勿手動改欄位）
const SHEET_LOG = "作答紀錄";       // 每次交卷一列（方便老師直接看）
const STORY_TITLES = {s1:"揠苗助長",s2:"一暴十寒",s3:"五十步笑百步",s4:"月攘一雞",s5:"齊人驕其妻妾",s6:"緣木求魚",s7:"一傅眾咻",s8:"校人烹魚",s9:"王顧左右而言他",s10:"孟母三遷"};
const STORY_TOTAL = {s1:3,s2:3,s3:3,s4:3,s5:3,s6:3,s7:3,s8:3,s9:3,s10:3};
const HEAD_REC = ["id","班級","座號","姓名","篇章","最高分","滿分","作答次數","最後作答","資料(JSON)"];
const HEAD_LOG = ["完成時間","班級","座號","姓名","篇章","第幾次","得分","滿分","用時(秒)","作答"];

function doGet() {
  return json_({ ok: true, message: "孟子文言閱讀挑戰 成績後端運作中" });
}

function doPost(e) {
  try {
    const req = JSON.parse(e.postData.contents || "{}");
    if (req.action === "list") {
      const t = isTeacher_(req.password);
      if (req.password && !t) return json_({ ok: false, error: "wrong password" });
      const recs = listRecords_();
      return json_({ ok: true, teacher: t, records: t ? recs : recs.map(publicView_) });
    }
    if (req.action === "save") return json_({ ok: true, record: saveRecord_(req.rec) });
    if (req.action === "clearClass") return json_({ ok: true, removed: clearClass_(req.cls, req.password) });
    return json_({ ok: false, error: "unknown action" });
  } catch (err) {
    return json_({ ok: false, error: String(err && err.message || err) });
  }
}

// 老師密碼是否正確（還沒改掉 change-me 時一律視為不正確）
function isTeacher_(pw) {
  return TEACHER_PASSWORD !== "change-me" && pw !== undefined && pw !== null && String(pw) === TEACHER_PASSWORD;
}

// 學生看得到的欄位：只有排行榜需要的資料，不含作答時間、作答內容
function publicView_(r) {
  return { id: r.id, cls: r.cls, seat: r.seat, name: r.name, story: r.story, total: r.total, best: r.best, attempts: r.attempts };
}

function json_(o) {
  return ContentService.createTextOutput(JSON.stringify(o)).setMimeType(ContentService.MimeType.JSON);
}

function sheet_(name, head) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sh = ss.getSheetByName(name);
  if (!sh) {
    sh = ss.insertSheet(name);
    sh.appendRow(head);
    sh.setFrozenRows(1);
    sh.getRange(1, 1, 1, head.length).setFontWeight("bold");
  }
  return sh;
}

function listRecords_() {
  const sh = sheet_(SHEET_RECORDS, HEAD_REC);
  const n = sh.getLastRow() - 1;
  if (n < 1) return [];
  return sh.getRange(2, 10, n, 1).getValues().map(r => { try { return JSON.parse(r[0]); } catch (e) { return null; } }).filter(Boolean);
}

function str_(v, max) { return String(v == null ? "" : v).slice(0, max); }
function ans_(a, n) { return Array.isArray(a) ? a.slice(0, n).map(x => (x === null || x === undefined) ? null : Math.max(0, Math.min(3, Number(x) | 0))) : []; }

function saveRecord_(rec) {
  if (!rec || !STORY_TITLES[rec.story]) throw new Error("bad record");
  const total = STORY_TOTAL[rec.story];
  const clamp = v => Math.max(0, Math.min(total, Number(v) || 0));
  const id = str_(rec.id, 120);
  if (!/^[A-Za-z0-9_\-.~:@+]+$/.test(id)) throw new Error("bad id");

  const lock = LockService.getScriptLock();
  lock.waitLock(20000);
  try {
    const sh = sheet_(SHEET_RECORDS, HEAD_REC);
    const n = sh.getLastRow() - 1;
    let row = -1, old = null;
    if (n > 0) {
      const ids = sh.getRange(2, 1, n, 1).getValues();
      for (let i = 0; i < ids.length; i++) if (ids[i][0] === id) { row = i + 2; break; }
      if (row > 0) { try { old = JSON.parse(sh.getRange(row, 10).getValue()); } catch (e) {} }
    }
    // 合併作答紀錄（以時間去重）
    const log = (old && Array.isArray(old.log)) ? old.log.slice() : [];
    const added = [];
    (Array.isArray(rec.log) ? rec.log : []).forEach(e => {
      const ts = Number(e && e.ts);
      if (ts && !log.some(x => x.ts === ts)) {
        const item = { ts: ts, dur: Math.max(0, Number(e.dur) || 0), sc: clamp(e.sc), ans: ans_(e.ans, total) };
        log.push(item); added.push(item);
      }
    });
    log.sort((a, b) => a.ts - b.ts);
    const keep = log.slice(-20);
    const merged = {
      id: id, cls: str_(rec.cls, 10), seat: Number(rec.seat) || 0, name: str_(rec.name, 12), story: rec.story, total: total,
      best: Math.max(clamp(rec.best), old ? clamp(old.best) : 0, ...keep.map(x => x.sc)),
      attempts: Math.max(Number(rec.attempts) || 1, old ? Number(old.attempts) || 0 : 0, log.length),
      firstAns: (old && Array.isArray(old.firstAns) && old.firstAns.length) ? old.firstAns : ans_(rec.firstAns, total),
      lastAns: ans_(rec.lastAns, total),
      ts: Math.max(Number(rec.ts) || 0, old ? Number(old.ts) || 0 : 0),
      firstTs: (old && old.firstTs) || (log[0] && log[0].ts) || Number(rec.ts) || Date.now(),
      log: keep
    };
    const values = [id, merged.cls, merged.seat, merged.name, STORY_TITLES[merged.story], merged.best, total, merged.attempts, new Date(merged.ts), JSON.stringify(merged)];
    if (row > 0) sh.getRange(row, 1, 1, values.length).setValues([values]);
    else sh.appendRow(values);

    if (added.length) {
      const lg = sheet_(SHEET_LOG, HEAD_LOG);
      added.forEach(e => {
        const nth = log.findIndex(x => x.ts === e.ts) + 1;
        lg.appendRow([new Date(e.ts), merged.cls, merged.seat, merged.name, STORY_TITLES[merged.story], nth, e.sc, total, e.dur, e.ans.map(a => a === null ? "-" : "ABCD"[a]).join("")]);
      });
    }
    return merged;
  } finally {
    lock.releaseLock();
  }
}

function clearClass_(cls, password) {
  if (!isTeacher_(password)) throw new Error("wrong password");
  const lock = LockService.getScriptLock();
  lock.waitLock(20000);
  try {
    const sh = sheet_(SHEET_RECORDS, HEAD_REC);
    const n = sh.getLastRow() - 1;
    if (n < 1) return 0;
    const cl = sh.getRange(2, 2, n, 1).getValues();
    let removed = 0;
    for (let i = cl.length - 1; i >= 0; i--) {
      if (String(cl[i][0]) === String(cls)) { sh.deleteRow(i + 2); removed++; }
    }
    return removed;
  } finally {
    lock.releaseLock();
  }
}
