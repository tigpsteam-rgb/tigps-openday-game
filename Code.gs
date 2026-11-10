/**
 * 你以為的數位世代：匿名作答統計（Google Apps Script 網頁應用程式）
 *
 * 只記錄：時間、隨機場次代碼、事件（start／finish）、玩法、身分、年齡區間、答對題數、題數、各題對錯。
 * 不記錄姓名、電話、Email、IP 或裝置資訊。
 *
 * 部署方式請見 README.md。
 */
const SHEET_NAME = 'log';
const HEADERS = ['time', 'sid', 'event', 'mode', 'player', 'role', 'age', 'score', 'total', 'answers'];
const ROLES = ['學生', '家長', '老師／教育工作者', '其他'];
const AGES = ['12歲以下', '13–15歲', '16–18歲', '19–24歲', '25–34歲', '35–44歲', '45–54歲', '55–64歲', '65歲以上'];
const MODES = ['quick', 'pair', 'full'];

function sheet_() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sh = ss.getSheetByName(SHEET_NAME);
  if (!sh) { sh = ss.insertSheet(SHEET_NAME); sh.appendRow(HEADERS); sh.setFrozenRows(1); }
  return sh;
}

// 寫入一筆紀錄（前端用 fetch POST，Content-Type: text/plain）
function doPost(e) {
  let body;
  try { body = JSON.parse(e.postData.contents); } catch (err) { return json_({ ok: false, error: 'bad json' }); }
  const rows = (Array.isArray(body) ? body : [body]).slice(0, 2).map(clean_).filter(Boolean);
  if (!rows.length) return json_({ ok: false, error: 'invalid' });
  const lock = LockService.getScriptLock();
  lock.waitLock(10000);
  try {
    const sh = sheet_();
    rows.forEach(r => sh.appendRow(r));
  } finally { lock.releaseLock(); }
  CacheService.getScriptCache().remove('stats');
  return json_({ ok: true });
}

function clean_(r) {
  if (!r || typeof r !== 'object') return null;
  const ev = r.event === 'finish' ? 'finish' : r.event === 'start' ? 'start' : null;
  if (!ev) return null;
  const sid = String(r.sid || '').replace(/[^a-z0-9]/gi, '').slice(0, 16);
  if (!sid) return null;
  const mode = MODES.indexOf(r.mode) >= 0 ? r.mode : '';
  const role = ROLES.indexOf(r.role) >= 0 ? r.role : '';
  const age = AGES.indexOf(r.age) >= 0 ? r.age : '';
  const player = ['', 'kid', 'adult'].indexOf(r.player) >= 0 ? r.player : '';
  let score = '', total = '', answers = '';
  if (ev === 'finish') {
    total = Math.max(0, Math.min(50, parseInt(r.total, 10) || 0));
    score = Math.max(0, Math.min(total, parseInt(r.score, 10) || 0));
    answers = String(r.answers || '').replace(/[^0-9:,]/g, '').slice(0, 200); // 例如 "4:1,6:0"
  }
  return [new Date(), sid, ev, mode, player, role, age, score, total, answers];
}

// 讀取統計（前端用 JSONP：?callback=xxx）
function doGet(e) {
  const cache = CacheService.getScriptCache();
  let txt = cache.get('stats');
  if (!txt) { txt = JSON.stringify(stats_()); cache.put('stats', txt, 20); }
  const cb = (e && e.parameter && e.parameter.callback || '').replace(/[^\w.$]/g, '');
  if (cb) return ContentService.createTextOutput(cb + '(' + txt + ');').setMimeType(ContentService.MimeType.JAVASCRIPT);
  return ContentService.createTextOutput(txt).setMimeType(ContentService.MimeType.JSON);
}

function stats_() {
  const values = sheet_().getDataRange().getValues().slice(1);
  const out = { starts: 0, finishes: 0, role: {}, age: {}, scoreBuckets: [0, 0, 0, 0, 0], pctList: [], avgPct: null, perQ: {}, updated: new Date().toISOString() };
  const startSids = {};
  let pctSum = 0;
  values.forEach(v => {
    const [time, sid, ev, mode, player, role, age, score, total, answers] = v;
    if (ev === 'start') {
      // 親子模式一個場次有兩位玩家，人次以玩家計
      const key = sid + '|' + player;
      if (!startSids[key]) { startSids[key] = 1; out.starts++; }
      return;
    }
    if (ev !== 'finish') return;
    out.finishes++;
    if (role) out.role[role] = (out.role[role] || 0) + 1;
    if (age) out.age[age] = (out.age[age] || 0) + 1;
    const t = Number(total) || 0, s = Number(score) || 0;
    if (t > 0) {
      const p = s / t;
      pctSum += p;
      out.pctList.push(Math.round(p * 100));
      out.scoreBuckets[Math.min(4, Math.floor(p * 5))]++;
    }
    String(answers || '').split(',').forEach(x => {
      const m = x.split(':'); if (m.length !== 2) return;
      const q = m[0], ok = m[1] === '1';
      const o = out.perQ[q] || (out.perQ[q] = { n: 0, ok: 0 });
      o.n++; if (ok) o.ok++;
    });
  });
  out.avgPct = out.finishes ? Math.round(pctSum / out.finishes * 1000) / 10 : null;
  out.pctList.sort((a, b) => a - b);
  return out;
}

function json_(o) {
  return ContentService.createTextOutput(JSON.stringify(o)).setMimeType(ContentService.MimeType.JSON);
}
