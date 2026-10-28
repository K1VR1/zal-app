/**
 * API для Telegram Mini App «Журнал зала».
 * Добавь этот файл в тот же проект Apps Script, где лежит основной скрипт.
 * Свойства скрипта (Настройки проекта → Свойства скрипта):
 *   BOT_TOKEN       — токен бота от @BotFather
 *   ALLOWED_USER_ID — твой Telegram id (приложение подскажет его при первом входе)
 */
const APP = {
  TODAY: 'Сегодня',
  LOG: 'Журнал',
  PROG: 'Прогресс',
  FIRST: 5,                       // первая строка упражнений на «Сегодня»
  START: [2026, 9, 12],           // понедельник первой недели (месяцы с нуля)
  GROUPS: { 1: 'Грудь и бицепс', 3: 'Ноги и трицепс', 5: 'Спина и предплечья' },
  MAX_AGE: 7 * 24 * 3600          // сколько секунд живёт подпись Telegram
};

function doGet() {
  return out_({ ok: true, app: 'zal' });
}

function doPost(e) {
  try {
    const req = JSON.parse((e && e.postData && e.postData.contents) || '{}');
    auth_(req.initData);
    switch (req.action) {
      case 'today':    return out_({ ok: true, data: apiToday_() });
      case 'save':     return out_({ ok: true, data: apiSaveMany_([req]) });
      case 'saveMany': return out_({ ok: true, data: apiSaveMany_(req.items || []) });
      case 'exercise': return out_({ ok: true, data: apiSetExercise_(req) });
      case 'progress': return out_({ ok: true, data: apiProgress_() });
      case 'history':  return out_({ ok: true, data: apiHistory_(req.name) });
      default:         return out_({ ok: false, error: 'Неизвестное действие' });
    }
  } catch (err) {
    return out_({ ok: false, error: String((err && err.message) || err) });
  }
}

function out_(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}

/* ---------- проверка, что запрос пришёл из Telegram от тебя ---------- */

function auth_(initData) {
  const props = PropertiesService.getScriptProperties();
  const token = props.getProperty('BOT_TOKEN');
  const allowed = props.getProperty('ALLOWED_USER_ID');
  if (!token) throw new Error('В свойствах скрипта не задан BOT_TOKEN');
  if (!initData) throw new Error('Открой приложение через Telegram');

  const params = {};
  initData.split('&').forEach(pair => {
    const i = pair.indexOf('=');
    if (i > 0) params[decodeURIComponent(pair.slice(0, i))] = decodeURIComponent(pair.slice(i + 1));
  });
  const hash = params.hash;
  delete params.hash;
  const check = Object.keys(params).sort().map(k => k + '=' + params[k]).join('\n');
  const secret = Utilities.computeHmacSha256Signature(
    Utilities.newBlob(token).getBytes(), Utilities.newBlob('WebAppData').getBytes());
  const sign = hex_(Utilities.computeHmacSha256Signature(Utilities.newBlob(check).getBytes(), secret));
  if (sign !== hash) throw new Error('Подпись Telegram не прошла проверку');

  const age = Date.now() / 1000 - Number(params.auth_date || 0);
  if (age > APP.MAX_AGE) throw new Error('Сессия устарела, перезапусти приложение');

  const user = JSON.parse(params.user || '{}');
  if (!allowed) throw new Error('Добавь в свойства скрипта ALLOWED_USER_ID = ' + user.id);
  if (String(user.id) !== String(allowed)) throw new Error('Нет доступа');
  return user;
}

function hex_(bytes) {
  return bytes.map(b => ('0' + (b & 0xff).toString(16)).slice(-2)).join('');
}

/* ---------- чтение таблицы ---------- */

function sheet_(name) { return SpreadsheetApp.getActive().getSheetByName(name); }
function tz_() { return SpreadsheetApp.getActive().getSpreadsheetTimeZone(); }
function ymd_(d) { return Utilities.formatDate(d, tz_(), 'yyyy-MM-dd'); }
function num_(v) { return v === '' || v === null ? null : Number(v); }
function startDate_() { return new Date(APP.START[0], APP.START[1], APP.START[2]); }
function dayOfWeek_(d) { return Number(Utilities.formatDate(d, tz_(), 'u')); } // 1 = пн

/* ---------- быстрое чтение: только нужные строки ---------- */

function dayKeyFn_() {
  const z = Utilities.formatDate(new Date(), tz_(), 'Z'); // например +0300
  const off = (z[0] === '-' ? -1 : 1) * (Number(z.slice(1, 3)) * 60 + Number(z.slice(3, 5))) * 60000;
  return d => Math.floor((d.getTime() + off) / 86400000);
}

function index_(sh) {
  const max = sh.getMaxRows();
  const col = sh.getRange(APP.FIRST, 1, max - APP.FIRST + 1, 1).getValues();
  const key = dayKeyFn_();
  const keys = [];
  for (let i = 0; i < col.length && col[i][0] instanceof Date; i++) keys.push(key(col[i][0]));
  return { keys: keys, key: key };
}

function block_(keys, k) {
  const s = keys.indexOf(k);
  if (s < 0) return null;
  let e = s;
  while (e + 1 < keys.length && keys[e + 1] === k) e++;
  return { s: s, e: e };
}

function isoFromKey_(k) { return new Date(k * 86400000).toISOString().slice(0, 10); }
function wdFromKey_(k) { const w = new Date(k * 86400000).getUTCDay(); return w === 0 ? 7 : w; }

function readSession_(sh, keys, k, full) {
  const b = block_(keys, k);
  if (!b) return null;
  const n = b.e - b.s + 1;
  const rng = sh.getRange(APP.FIRST + b.s, 1, n, 13);
  const disp = rng.getDisplayValues();
  const vals = full ? rng.getValues() : null;
  const w = wdFromKey_(k);
  const list = [];
  for (let i = 0; i < n; i++) {
    if (full) {
      const v = vals[i], sets = [];
      for (let j = 0; j < 4; j++) sets.push([num_(v[4 + 2 * j]), num_(v[5 + 2 * j])]);
      list.push({ row: APP.FIRST + b.s + i, name: v[1], plan: disp[i][2], last: v[1] ? disp[i][3] : '',
        sets: sets, tonnage: Number(v[12]) || 0, extra: i === n - 1 });
    } else if (disp[i][1]) {
      list.push({ name: disp[i][1], plan: disp[i][2], last: disp[i][3] });
    }
  }
  return { date: isoFromKey_(k), weekday: w, group: APP.GROUPS[w] || '', exercises: list };
}

function weekStats_(keys, todayKey) {
  const wd = wdFromKey_(todayKey);
  const from = todayKey - (wd - 1), to = from + 6;
  let s = -1, e = -1;
  keys.forEach((k, i) => { if (k >= from && k <= to) { if (s < 0) s = i; e = i; } });
  let done = {}, tonnage = 0;
  if (s >= 0) {
    // строки «Журнала» идут в том же порядке, что и на «Сегодня», со сдвигом на 3
    const rows = sheet_(APP.LOG).getRange(APP.FIRST + s - 3, 1, e - s + 1, 7).getValues();
    rows.forEach((r, i) => {
      if (r[3] === '') return;
      done[keys[s + i]] = true;
      tonnage += Number(r[6]) || 0;
    });
  }
  const prs = Number(String(sheet_(APP.PROG).getRange('J5').getDisplayValue()).replace(/\s/g, '')) || 0;
  return { done: Object.keys(done).length, total: 3, tonnage: Math.round(tonnage), prs: prs };
}

function options_() {
  const c = CacheService.getScriptCache();
  const hit = c.get('options');
  if (hit) return JSON.parse(hit);
  const list = sheet_(APP.PROG).getRange('C7:C27').getValues().map(r => r[0]).filter(String);
  c.put('options', JSON.stringify(list), 3600);
  return list;
}

/* ---------- действия ---------- */

function apiToday_() {
  const sh = sheet_(APP.TODAY);
  const ix = index_(sh);
  const todayKey = ix.key(new Date());
  const nextKey = ix.keys.find(k => k > todayKey);
  const today = readSession_(sh, ix.keys, todayKey, true);
  return {
    date: isoFromKey_(todayKey),
    weekday: wdFromKey_(todayKey),
    today: today,
    next: today || nextKey === undefined ? null : readSession_(sh, ix.keys, nextKey, false),
    week: weekStats_(ix.keys, todayKey),
    options: options_()
  };
}

function apiSaveMany_(items) {
  const sh = sheet_(APP.TODAY);
  const max = sh.getMaxRows();
  const clean = v => (v === null || v === undefined || v === '') ? '' : Number(String(v).replace(',', '.'));
  const lock = LockService.getScriptLock();
  lock.waitLock(10000);
  try {
    items.forEach(it => {
      const row = Number(it.row), set = Number(it.set);
      if (!(row >= APP.FIRST && row <= max)) throw new Error('Неверная строка');
      if (!(set >= 0 && set <= 3)) throw new Error('Неверный подход');
      const kg = clean(it.kg), reps = clean(it.reps);
      if ((kg !== '' && !(kg >= 0)) || (reps !== '' && !(reps >= 0))) throw new Error('Неверное число');
      sh.getRange(row, 5 + 2 * set, 1, 2).setValues([[kg, reps]]);
    });
  } finally {
    lock.releaseLock();
  }
  return { saved: items.length };
}

function apiSetExercise_(req) {
  const sh = sheet_(APP.TODAY);
  const row = Number(req.row);
  const name = String(req.name || '').trim().slice(0, 60);
  const two = sh.getRange(row, 1, 2, 1).getValues();
  const key = dayKeyFn_();
  const lastOfDay = two[0][0] instanceof Date && !(two[1][0] instanceof Date && key(two[1][0]) === key(two[0][0]));
  if (row < APP.FIRST || !lastOfDay) throw new Error('Менять можно только доп. упражнение');
  sh.getRange(row, 2).setValue(name);
  SpreadsheetApp.flush();
  return apiToday_();
}

function logRows_() {
  const sh = sheet_(APP.LOG);
  const last = sh.getLastRow();
  if (last < 2) return [];
  return sh.getRange(2, 1, last - 1, 9).getValues()
    .filter(r => r[0] instanceof Date && r[2] !== '');
}

function apiProgress_() {
  const p = sheet_(APP.PROG);
  const tiles = p.getRange('B5:L5').getDisplayValues()[0];
  const rows = p.getRange('B7:L27').getDisplayValues().filter(r => r[1]);
  return {
    tiles: { workouts: tiles[0], tonnage: tiles[2], week: tiles[5], prs: tiles[8] },
    exercises: rows.map(r => ({
      day: r[0], name: r[1], sets: r[2], reps: r[3], start: r[4], best: r[5],
      gain: r[6], gainPct: r[7], maxReps: r[8], count: r[9], last: r[10]
    }))
  };
}

function apiHistory_(name) {
  const log = logRows_().filter(r => r[3] !== '');
  const start = startDate_().getTime();
  const weeks = {};
  log.forEach(r => {
    const w = Math.floor((r[0].getTime() - start) / (7 * 86400000)) + 1;
    weeks[w] = (weeks[w] || 0) + (Number(r[6]) || 0);
  });
  const nowWeek = Math.floor((Date.now() - start) / (7 * 86400000)) + 1;
  const weekList = [];
  for (let w = 1; w <= Math.max(nowWeek, 1); w++) weekList.push({ week: w, tonnage: Math.round(weeks[w] || 0) });
  const names = [];
  log.forEach(r => { if (names.indexOf(r[2]) < 0) names.push(r[2]); });
  const points = name ? log.filter(r => r[2] === name).map(r => ({
    date: ymd_(r[0]), weight: Number(r[3]) || 0, reps: Number(r[5]) || 0, tonnage: Number(r[6]) || 0
  })) : [];
  return { weeks: weekList, names: names, points: points };
}
