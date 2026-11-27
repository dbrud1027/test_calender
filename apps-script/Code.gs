/**
 * 주간 할 일 – 저장소 + 일정표 시트 가져오기
 *
 * - 일정표 시트는 읽기만 합니다. 이 파일에는 시트를 고치는 코드가 없고,
 *   appsscript.json에서 스프레드시트 권한을 '읽기 전용'으로 제한해 두었어요.
 * - 할 일·메모·기분은 이 스크립트의 저장 공간(스크립트 속성)에 주 단위로 보관돼요.
 * - 10분마다 시트에서 '김유경' 칸을 읽어 이번 주부터의 할 일을 채워 넣어요.
 */

const CONFIG = {
  SHEET_ID: '1JO43sl8ORcPNsqGHPz5fLSoxwgV3xhJpY6GP_jpnSm8', // [앱기획] 일일 일정표
  NAME: '김유경',                                          // 가져올 이름 칸
  SYNC_MINUTES: 10,                                        // 1, 5, 10, 15, 30 중 하나
  KEEP_WEEKS: 26                                           // 이보다 오래된 주는 정리
};

/* ---------- 처음 한 번 실행 ---------- */

/** 연결 키를 만들고, 10분마다 시트를 가져오는 자동 실행을 켜요. */
function setup() {
  const p = props_();
  if (!p.getProperty('SECRET')) p.setProperty('SECRET', Utilities.getUuid().replace(/-/g, ''));
  ScriptApp.getProjectTriggers()
    .filter(t => t.getHandlerFunction() === 'syncSheet')
    .forEach(t => ScriptApp.deleteTrigger(t));
  ScriptApp.newTrigger('syncSheet').timeBased().everyMinutes(CONFIG.SYNC_MINUTES).create();
  const info = syncSheet();
  Logger.log('설정 완료! 읽은 탭: ' + info.tab + ' · 반영한 주: ' + info.changed + '개');
  Logger.log('연결 키: ' + p.getProperty('SECRET'));
}

/** 연결 키를 다시 확인할 때 실행하세요. */
function showKey() {
  Logger.log('연결 키: ' + props_().getProperty('SECRET'));
}

/** 연결 키를 바꾸고 싶을 때 실행하세요. (모든 기기에서 새 키로 다시 연결해야 해요) */
function resetKey() {
  props_().setProperty('SECRET', Utilities.getUuid().replace(/-/g, ''));
  showKey();
}

/* ---------- 웹 앱 ---------- */

function doGet(e) { return handle_((e && e.parameter) || {}); }

function doPost(e) {
  let body = {};
  try { body = JSON.parse(e.postData.contents); } catch (err) {}
  return handle_(body);
}

function handle_(q) {
  const secret = props_().getProperty('SECRET');
  if (!secret || q.key !== secret) return out_({ ok: false, error: 'key' });
  const lock = LockService.getScriptLock();
  let locked = false;
  try {
    switch (q.action) {
      case 'ping':
        return out_({ ok: true, sync: lastSync_() });
      case 'week':
        return out_({ ok: true, doc: getDoc_(weekKey_(q.k)), sync: lastSync_() });
      case 'save': {
        const k = weekKey_(q.k);
        const data = q.data;
        if (!data || typeof data !== 'object' || Array.isArray(data)) return out_({ ok: false, error: 'data' });
        if (JSON.stringify(data).length > 150000) return out_({ ok: false, error: 'too_big' });
        lock.waitLock(20000); locked = true;
        data.updated = Date.now();
        putDoc_(k, data);
        return out_({ ok: true, updated: data.updated });
      }
      case 'sync': {
        const info = syncSheet();
        return out_({ ok: true, sync: info, doc: q.k ? getDoc_(weekKey_(q.k)) : null });
      }
      default:
        return out_({ ok: false, error: 'action' });
    }
  } catch (err) {
    return out_({ ok: false, error: String((err && err.message) || err) });
  } finally {
    if (locked) lock.releaseLock();
  }
}

function out_(o) {
  return ContentService.createTextOutput(JSON.stringify(o)).setMimeType(ContentService.MimeType.JSON);
}

/* ---------- 시트 가져오기 (읽기 전용) ---------- */

function syncSheet() {
  const lock = LockService.getScriptLock();
  lock.waitLock(25000);
  try {
    const ss = SpreadsheetApp.openById(CONFIG.SHEET_ID);
    const sh = pickTab_(ss);
    if (!sh) throw new Error('반기 탭(예: 2026_하반기)을 찾지 못했어요');
    const rows = sh.getDataRange().getDisplayValues();
    const weeks = parseSheet_(rows, new Date());
    let changed = 0;
    weeks.forEach((w, k) => {
      const st = normalize_(getDoc_(k));
      if (mergeWeek_(st, w)) { st.updated = Date.now(); putDoc_(k, st); changed++; }
    });
    prune_();
    const info = { tab: sh.getName(), weeks: weeks.size, changed: changed, at: new Date().toISOString() };
    props_().setProperty('LAST_SYNC', JSON.stringify(info));
    return info;
  } finally {
    lock.releaseLock();
  }
}

/** 오늘 날짜에 맞는 반기 탭(예: 2026_하반기). 아직 없으면 가장 최근 반기 탭. */
function pickTab_(ss) {
  const now = new Date();
  const want = now.getFullYear() + '_' + (now.getMonth() < 6 ? '상' : '하') + '반기';
  const tabs = ss.getSheets();
  const exact = tabs.find(s => s.getName().trim() === want);
  if (exact) return exact;
  const ranked = tabs.map(s => {
    const m = s.getName().trim().match(/^(\d{4})_(상|하)반기$/);
    return m ? { s: s, v: Number(m[1]) * 2 + (m[2] === '하' ? 1 : 0) } : null;
  }).filter(Boolean).sort((a, b) => b.v - a.v);
  return ranked.length ? ranked[0].s : null;
}

/* ---------- 시트 해석 (페이지 버전과 같은 규칙) ---------- */

const WD_ = ['월', '화', '수', '목', '금'];
const pad_ = n => String(n).padStart(2, '0');
function mondayOf_(d) { const x = new Date(d.getFullYear(), d.getMonth(), d.getDate()); x.setDate(x.getDate() - ((x.getDay() + 6) % 7)); return x; }
function keyOf_(m) { return m.getFullYear() + '-' + pad_(m.getMonth() + 1) + '-' + pad_(m.getDate()); }
function nearest_(mo, d, now) {
  let best = null;
  [now.getFullYear() - 1, now.getFullYear(), now.getFullYear() + 1].forEach(y => {
    const x = new Date(y, mo - 1, d);
    if (!best || Math.abs(x - now) < Math.abs(best - now)) best = x;
  });
  return best;
}

function parseSheet_(R, now) {
  const cell = (r, c) => String((R[r] || [])[c] == null ? '' : R[r][c]).trim();
  let wr = -1;
  for (let r = 0; r < Math.min(15, R.length); r++) {
    if (R[r].filter(x => String(x).trim() === '월').length >= 3) { wr = r; break; }
  }
  if (wr < 0) throw new Error('시트 모양이 예전과 달라서 읽지 못했어요 (요일 줄 없음)');
  const dr = wr + 1;
  let W = 0; for (let r = 0; r <= dr; r++) W = Math.max(W, (R[r] || []).length);
  let mr = -1;
  for (let r = 0; r < wr; r++) if (R[r].some(x => /^\d{1,2}월$/.test(String(x).trim()))) { mr = r; break; }
  let pc = R[dr].findIndex(x => String(x).trim() === '프로젝트'); if (pc < 1) pc = 2;

  const cols = []; let cur = null;
  for (let c = 0; c < W; c++) {
    if (mr >= 0) { const m = cell(mr, c).match(/^(\d{1,2})월$/); if (m) cur = Number(m[1]); }
    const w = cell(wr, c), v = cell(dr, c);
    if (WD_.indexOf(w) < 0 || !v) continue;
    let mo, d; const m = v.match(/^(\d{1,2})\/(\d{1,2})$/);
    if (m) { mo = Number(m[1]); d = Number(m[2]); }
    else if (/^\d{1,2}$/.test(v) && cur) { mo = cur; d = Number(v); }
    else continue;
    cols.push({ c: c, date: nearest_(mo, d, now) });
  }
  if (!cols.length) throw new Error('시트 모양이 예전과 달라서 읽지 못했어요 (날짜 없음)');

  const nr = R.findIndex(r => String(r[0]).trim() === CONFIG.NAME);
  if (nr < 0) throw new Error('시트에서 “' + CONFIG.NAME + '” 칸을 찾지 못했어요');
  let end = nr + 1; while (end < R.length && !cell(end, 0)) end++;
  let ir = -1; for (let r = nr - 1; r >= Math.max(0, nr - 3); r--) if (/^\d+$/.test(cell(r, 0))) { ir = r; break; }

  const items = c => {
    const res = [], cnt = {}; let p = '';
    for (let r = nr; r < end; r++) {
      const lab = cell(r, pc); if (lab) p = lab.slice(0, 100);
      String((R[r] || [])[c] || '').split(/\r?\n/).forEach(line => {
        const t = line.trim(); if (!t) return;
        const b = p + '|' + t; cnt[b] = (cnt[b] || 0) + 1;
        res.push({ p: p, t: t.slice(0, 500), k: (b + '|' + cnt[b]).slice(0, 600) });
      });
    }
    return res;
  };
  const leave = c => { const v = ir >= 0 ? cell(ir, c + 1) : ''; return v && v !== '-' && v !== '일정완료' ? v.slice(0, 20) : ''; };

  const from = mondayOf_(now), weeks = new Map();
  cols.forEach(o => {
    if (o.date < from) return;
    const i = (o.date.getDay() + 6) % 7; if (i > 4) return;
    const k = keyOf_(mondayOf_(o.date));
    if (!weeks.has(k)) weeks.set(k, { days: [[], [], [], [], []], leave: ['', '', '', '', ''] });
    const w = weeks.get(k); w.days[i] = items(o.c); w.leave[i] = leave(o.c);
  });
  return weeks;
}

/* ---------- 주 데이터 ---------- */

function blank_() {
  return { days: [[], [], [], [], [], [], []], memo: '', mood: ['', '', '', '', '', '', ''], leave: ['', '', '', '', '', '', ''], hide: [] };
}
function normalize_(o) {
  const s = blank_();
  if (!o || typeof o !== 'object') return s;
  if (Array.isArray(o.days)) for (let i = 0; i < 7; i++) if (Array.isArray(o.days[i])) s.days[i] = o.days[i].filter(t => t && typeof t.t === 'string');
  if (typeof o.memo === 'string') s.memo = o.memo;
  if (Array.isArray(o.mood)) for (let i = 0; i < 7; i++) s.mood[i] = typeof o.mood[i] === 'string' ? o.mood[i] : '';
  if (Array.isArray(o.leave)) for (let i = 0; i < 7; i++) s.leave[i] = typeof o.leave[i] === 'string' ? o.leave[i] : '';
  if (Array.isArray(o.hide)) s.hide = o.hide.filter(x => typeof x === 'string').slice(-300);
  return s;
}
function newId_() { return Date.now().toString(36) + Math.random().toString(36).slice(2, 7); }

/** 시트 내용을 한 주에 반영. 체크한 항목·직접 쓴 항목·지운 항목(hide)은 지켜요. */
function mergeWeek_(st, w) {
  const before = JSON.stringify([st.days, st.leave]);
  const hide = {}; st.hide.forEach(k => { hide[k] = true; });
  for (let i = 0; i < 5; i++) {
    const cur = st.days[i], imp = {};
    cur.forEach(x => { if (x.s && x.k) imp[x.k] = x; });
    const next = [];
    w.days[i].forEach(it => {
      if (hide[it.k]) return;
      const ex = imp[it.k];
      if (ex) { next.push(Object.assign({}, ex, { p: it.p })); delete imp[it.k]; }
      else next.push({ id: newId_(), t: it.t, d: false, s: true, k: it.k, p: it.p });
    });
    Object.keys(imp).forEach(k => { if (imp[k].d) next.push(imp[k]); });
    st.days[i] = next.concat(cur.filter(x => !x.s));
    st.leave[i] = w.leave[i] || '';
  }
  return JSON.stringify([st.days, st.leave]) !== before;
}

/* ---------- 저장 공간 (스크립트 속성, 값 하나당 9KB 제한이라 나눠 저장) ---------- */

const CHUNK_ = 2500; // 한글 3바이트 기준으로도 9KB 아래
function props_() { return PropertiesService.getScriptProperties(); }
function weekKey_(k) {
  if (!/^\d{4}-\d{2}-\d{2}$/.test(String(k))) throw new Error('잘못된 주 키');
  return String(k);
}
function putDoc_(k, obj) {
  const p = props_(), s = JSON.stringify(obj), parts = [];
  for (let i = 0; i < s.length;) {
    let j = Math.min(i + CHUNK_, s.length);
    const code = s.charCodeAt(j - 1);
    if (j < s.length && code >= 0xD800 && code <= 0xDBFF) j--; // 이모지가 반으로 잘리지 않게
    parts.push(s.slice(i, j)); i = j;
  }
  const old = Number(p.getProperty('W:' + k) || 0), set = {};
  parts.forEach((x, i) => { set['W:' + k + '#' + i] = x; });
  set['W:' + k] = String(parts.length);
  p.setProperties(set);
  for (let i = parts.length; i < old; i++) p.deleteProperty('W:' + k + '#' + i);
}
function getDoc_(k) {
  const p = props_(), n = Number(p.getProperty('W:' + k) || 0);
  if (!n) return null;
  let s = ''; for (let i = 0; i < n; i++) s += p.getProperty('W:' + k + '#' + i) || '';
  try { return JSON.parse(s); } catch (err) { return null; }
}
function prune_() {
  const p = props_(), all = p.getKeys();
  const limit = mondayOf_(new Date()); limit.setDate(limit.getDate() - 7 * CONFIG.KEEP_WEEKS);
  const cut = keyOf_(limit);
  all.forEach(key => {
    const m = key.match(/^W:(\d{4}-\d{2}-\d{2})(#\d+)?$/);
    if (m && m[1] < cut) p.deleteProperty(key);
  });
}
function lastSync_() {
  try { return JSON.parse(props_().getProperty('LAST_SYNC') || 'null'); } catch (err) { return null; }
}
