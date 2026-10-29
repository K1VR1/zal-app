/**
 * API для Telegram Mini App «Журнал зала» + бот-напоминалка.
 * Добавь этот файл в тот же проект Apps Script, где лежит основной скрипт.
 * Свойства скрипта (Настройки проекта → Свойства скрипта):
 *   BOT_TOKEN       — токен бота от @BotFather
 *   ALLOWED_USER_ID — твой Telegram id
 * После замены кода: запусти setupReminders() и обнови развёртывание (новая версия).
 */
const APP = {
  TODAY: 'Сегодня',
  LOG: 'Журнал',
  PROG: 'Прогресс',
  WEIGHT: 'Вес',
  FIRST: 5,                       // первая строка упражнений на «Сегодня»
  START: [2026, 9, 12],           // понедельник первой недели (месяцы с нуля)
  GROUPS: { 1: 'Грудь и бицепс', 3: 'Ноги и трицепс', 5: 'Спина и предплечья' },
  DAYCODE: { 'Пн': 1, 'Ср': 3, 'Пт': 5 },
  BODYWEIGHT: ['Подтягивания', 'Пресс на брусьях'],
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
      case 'note':     return out_({ ok: true, data: apiNote_(req) });
      case 'move':     return out_({ ok: true, data: apiMove_(req) });
      case 'weight':   return out_({ ok: true, data: apiWeight_(req) });
      case 'progress': return out_({ ok: true, data: apiProgress_() });
      case 'history':  return out_({ ok: true, data: apiHistory_(req.name) });
      case 'week':     return out_({ ok: true, data: apiWeek_(Number(req.week)) });
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

/* ---------- общие помощники ---------- */

function sheet_(name) { return SpreadsheetApp.getActive().getSheetByName(name); }
function tz_() { return SpreadsheetApp.getActive().getSpreadsheetTimeZone(); }
function ymd_(d) { return Utilities.formatDate(d, tz_(), 'yyyy-MM-dd'); }
function num_(v) { return v === '' || v === null || v === undefined ? null : Number(v); }
function startDate_() { return new Date(APP.START[0], APP.START[1], APP.START[2]); }
function dayOfWeek_(d) { return Number(Utilities.formatDate(d, tz_(), 'u')); } // 1 = пн

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
function uniqKeys_(keys) { return keys.filter((k, i) => i === 0 || keys[i - 1] !== k); }

// справочник с «Прогресса»: список упражнений, их день и заметки (кэш на час)
function meta_() {
  const c = CacheService.getScriptCache();
  const hit = c.get('meta2');
  if (hit) return JSON.parse(hit);
  const rows = sheet_(APP.PROG).getRange('B7:M27').getValues();
  const m = { options: [], day: {}, notes: {} };
  rows.forEach(r => {
    const name = String(r[1] || '').trim();
    if (!name) return;
    m.options.push(name);
    m.day[name] = String(r[0] || '');
    if (String(r[11] || '').trim()) m.notes[name] = String(r[11]).trim();
  });
  c.put('meta2', JSON.stringify(m), 3600);
  return m;
}
function dropMeta_() { CacheService.getScriptCache().remove('meta2'); }

// день программы (1/3/5) по первому упражнению сессии — работает и для перенесённых тренировок
function planDay_(names, fallback) {
  const m = meta_();
  for (let i = 0; i < names.length; i++) {
    const d = APP.DAYCODE[m.day[names[i]]];
    if (d) return d;
  }
  return fallback;
}

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
  const pw = planDay_(list.map(x => x.name).filter(String), w);
  return { date: isoFromKey_(k), weekday: w, pw: pw, group: APP.GROUPS[pw] || '', exercises: list, start: b.s };
}

function hasData_(s) {
  return !!s && s.exercises.some(e => e.sets && e.sets.some(p => p[0] !== null || p[1] !== null));
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

// последние 3 записи по каждому упражнению сегодняшней тренировки (для «застоя»)
function recent_(startIdx, names) {
  const out = {};
  names.forEach(n => out[n] = []);
  const lastRow = APP.FIRST + startIdx - 3 - 1; // строка «Журнала» прямо перед сегодняшней тренировкой
  if (lastRow < 2) return out;
  const rows = sheet_(APP.LOG).getRange(2, 3, lastRow - 1, 4).getValues(); // C:F
  for (let i = rows.length - 1; i >= 0; i--) {
    const r = rows[i], list = out[r[0]];
    if (!list || list.length >= 3 || r[1] === '') continue;
    list.push({ kg: Number(r[1]) || 0, sets: Number(r[2]) || 0, reps: Number(r[3]) || 0 });
  }
  return out;
}

/* ---------- вес тела ---------- */

function weightSheet_() {
  let sh = sheet_(APP.WEIGHT);
  if (sh) return sh;
  sh = SpreadsheetApp.getActive().insertSheet(APP.WEIGHT);
  sh.getRange('A1:B1').setValues([['Дата', 'Вес, кг']]).setFontWeight('bold').setFontColor('#8E95A6');
  sh.getRange('A:B').setBackground('#121418');
  sh.getRange('A2:B').setFontColor('#F2F4F8');
  sh.getRange('A2:A').setNumberFormat('dd.mm.yyyy');
  sh.getRange('B2:B').setNumberFormat('0.0');
  sh.setFrozenRows(1);
  sh.setColumnWidths(1, 2, 120);
  return sh;
}

function weights_(limit) {
  const sh = sheet_(APP.WEIGHT);
  if (!sh) return [];
  const last = sh.getLastRow();
  if (last < 2) return [];
  const n = Math.min(limit || 400, last - 1);
  return sh.getRange(last - n + 1, 1, n, 2).getValues()
    .filter(r => r[0] instanceof Date && r[1] !== '')
    .map(r => ({ date: ymd_(r[0]), kg: Number(r[1]) }));
}

function addWeight_(kg) {
  kg = Number(String(kg).replace(',', '.'));
  if (!(kg >= 30 && kg <= 250)) throw new Error('Вес должен быть числом от 30 до 250');
  kg = Math.round(kg * 10) / 10;
  const sh = weightSheet_();
  const today = ymd_(new Date());
  const lock = LockService.getScriptLock();
  lock.waitLock(10000);
  try {
    const last = sh.getLastRow();
    if (last >= 2) {
      const prev = sh.getRange(last, 1).getValue();
      if (prev instanceof Date && ymd_(prev) === today) { sh.getRange(last, 2).setValue(kg); return kg; }
    }
    sh.getRange(last + 1, 1, 1, 2).setValues([[Utilities.parseDate(today, tz_(), 'yyyy-MM-dd'), kg]]);
  } finally {
    lock.releaseLock();
  }
  return kg;
}

/* ---------- действия приложения ---------- */

function apiToday_() {
  const sh = sheet_(APP.TODAY);
  const ix = index_(sh);
  const todayKey = ix.key(new Date());
  const nextKey = ix.keys.find(k => k > todayKey);
  const today = readSession_(sh, ix.keys, todayKey, true);
  const m = meta_();
  const w = weights_(1);
  const res = {
    date: isoFromKey_(todayKey),
    weekday: wdFromKey_(todayKey),
    today: today,
    next: today || nextKey === undefined ? null : readSession_(sh, ix.keys, nextKey, false),
    week: weekStats_(ix.keys, todayKey),
    options: m.options,
    notes: m.notes,
    bodyweight: w.length ? w[w.length - 1].kg : null,
    move: moveOptions_(sh, ix, todayKey, today)
  };
  if (today) {
    const names = today.exercises.map(e => e.name).filter(String);
    res.recent = recent_(today.start, names);
  }
  return res;
}

// что можно перенести: пропущенную тренировку на сегодня или сегодняшнюю на завтра
function moveOptions_(sh, ix, todayKey, today) {
  const u = uniqKeys_(ix.keys);
  if (today) {
    const after = u.find(k => k > todayKey);
    if (!hasData_(today) && (after === undefined || after > todayKey + 1)) {
      return { kind: 'tomorrow', from: isoFromKey_(todayKey), to: isoFromKey_(todayKey + 1) };
    }
    return null;
  }
  const prev = u.filter(k => k < todayKey).pop();
  if (prev === undefined || todayKey - prev > 3) return null;
  const ps = readSession_(sh, ix.keys, prev, true);
  if (hasData_(ps)) return null;
  return { kind: 'today', from: isoFromKey_(prev), to: isoFromKey_(todayKey), group: ps.group, weekday: ps.weekday };
}

function apiMove_(req) {
  const sh = sheet_(APP.TODAY);
  const lock = LockService.getScriptLock();
  lock.waitLock(10000);
  try {
    const ix = index_(sh);
    const to = Utilities.parseDate(String(req.to), tz_(), 'yyyy-MM-dd');
    const fromKey = ix.key(Utilities.parseDate(String(req.from), tz_(), 'yyyy-MM-dd'));
    const toKey = ix.key(to);
    const todayKey = ix.key(new Date());
    const b = block_(ix.keys, fromKey);
    if (!b) throw new Error('Тренировка не найдена');
    if (toKey < todayKey || toKey > todayKey + 1 || Math.abs(toKey - fromKey) > 3) throw new Error('Перенести можно только на сегодня или завтра');
    const u = uniqKeys_(ix.keys);
    const before = u.filter(k => k < fromKey).pop();
    const after = u.find(k => k > fromKey);
    if ((before !== undefined && toKey <= before) || (after !== undefined && toKey >= after)) {
      throw new Error('На этот день уже стоит другая тренировка');
    }
    if (hasData_(readSession_(sh, ix.keys, fromKey, true))) throw new Error('В этой тренировке уже есть подходы');
    const n = b.e - b.s + 1;
    sh.getRange(APP.FIRST + b.s, 1, n, 1).setValues(Array.from({ length: n }, () => [to]));
    SpreadsheetApp.flush();
  } finally {
    lock.releaseLock();
  }
  if (typeof showToday === 'function') { try { showToday(); } catch (e) {} }
  return apiToday_();
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

function apiNote_(req) {
  const name = String(req.name || '').trim();
  const text = String(req.text || '').trim().slice(0, 200);
  const p = sheet_(APP.PROG);
  const names = p.getRange('C7:C27').getValues().map(r => String(r[0]).trim());
  const i = names.indexOf(name);
  if (i < 0) throw new Error('Заметки можно оставлять только к упражнениям из списка');
  p.getRange(7 + i, 13).setValue(text);
  dropMeta_();
  return { name: name, text: text };
}

function apiWeight_(req) {
  const kg = addWeight_(req.kg);
  return { kg: kg, list: weights_(60) };
}

function logRows_() {
  const sh = sheet_(APP.LOG);
  const last = sh.getLastRow();
  if (last < 2) return [];
  return sh.getRange(2, 1, last - 1, 9).getValues()
    .filter(r => r[0] instanceof Date && r[2] !== '');
}

/* ---------- статистика, серии и достижения ---------- */

function weekNo_(d) { return Math.floor((d.getTime() - startDate_().getTime()) / (7 * 86400000)) + 1; }

function stats_(log) {
  log = log || logRows_();
  const done = log.filter(r => r[3] !== '');
  const dates = {}, perWeek = {};
  let tonnage = 0, maxLift = 0, prs = 0;
  done.forEach(r => {
    const iso = ymd_(r[0]);
    if (!dates[iso]) { dates[iso] = true; const w = weekNo_(r[0]); perWeek[w] = (perWeek[w] || 0) + 1; }
    tonnage += Number(r[6]) || 0;
    if (APP.BODYWEIGHT.indexOf(r[2]) < 0) maxLift = Math.max(maxLift, Number(r[3]) || 0);
    if (String(r[8]).indexOf('PR') === 0) prs++;
  });
  // серия: сколько недель подряд по 3 тренировки (текущая неделя считается, если уже 3/3)
  const nowWeek = weekNo_(new Date());
  let streak = 0, w = (perWeek[nowWeek] || 0) >= 3 ? nowWeek : nowWeek - 1;
  while (w >= 1 && (perWeek[w] || 0) >= 3) { streak++; w--; }
  let best = 0, run = 0;
  for (let i = 1; i <= nowWeek; i++) { run = (perWeek[i] || 0) >= 3 ? run + 1 : 0; best = Math.max(best, run); }
  return { workouts: Object.keys(dates).length, tonnage: Math.round(tonnage), maxLift: maxLift, prs: prs,
    streak: streak, bestStreak: best, perWeek: perWeek, thisWeek: perWeek[nowWeek] || 0 };
}

function achievements_(s) {
  const list = [
    ['w1', '🎯', 'Первый шаг', 'Первая тренировка', s.workouts, 1],
    ['w10', '💪', 'Десятка', '10 тренировок', s.workouts, 10],
    ['w25', '🏋️', 'Втянулся', '25 тренировок', s.workouts, 25],
    ['w50', '🦾', 'Полсотни', '50 тренировок', s.workouts, 50],
    ['w100', '👑', 'Сотня', '100 тренировок', s.workouts, 100],
    ['s1', '📅', 'Полная неделя', '3 тренировки за неделю', s.bestStreak, 1],
    ['s4', '🔥', 'Месяц без пропусков', '4 полные недели подряд', s.bestStreak, 4],
    ['s8', '☄️', 'Машина', '8 полных недель подряд', s.bestStreak, 8],
    ['t10', '🧱', '10 тонн', 'Тоннаж всего 10 т', s.tonnage, 10000],
    ['t50', '🚚', '50 тонн', 'Тоннаж всего 50 т', s.tonnage, 50000],
    ['t100', '🚂', '100 тонн', 'Тоннаж всего 100 т', s.tonnage, 100000],
    ['t250', '🚀', '250 тонн', 'Тоннаж всего 250 т', s.tonnage, 250000],
    ['p5', '🏆', 'Рекордсмен', '5 личных рекордов', s.prs, 5],
    ['p25', '🥇', 'Без потолка', '25 личных рекордов', s.prs, 25],
    ['k50', '🪨', 'Полтинник', 'Рабочий вес 50 кг', s.maxLift, 50],
    ['k100', '🗿', 'Сотка', 'Рабочий вес 100 кг', s.maxLift, 100]
  ];
  return list.map(a => ({ id: a[0], icon: a[1], title: a[2], desc: a[3],
    got: a[4] >= a[5], have: Math.min(a[4], a[5]), need: a[5] }));
}

function apiProgress_() {
  const p = sheet_(APP.PROG);
  const tiles = p.getRange('B5:L5').getDisplayValues()[0];
  const rows = p.getRange('B7:L27').getDisplayValues().filter(r => r[1]);
  const log = logRows_();
  const s = stats_(log);
  return {
    tiles: { workouts: tiles[0], tonnage: tiles[2], week: tiles[5], prs: tiles[8] },
    exercises: rows.map(r => ({
      day: r[0], name: r[1], sets: r[2], reps: r[3], start: r[4], best: r[5],
      gain: r[6], gainPct: r[7], maxReps: r[8], count: r[9], last: r[10]
    })),
    streak: s.streak, bestStreak: s.bestStreak, thisWeek: s.thisWeek,
    achievements: achievements_(s),
    weights: weights_(60),
    spark: spark_(log),
    weeks5: weeks5_(s),
    start: ymd_(startDate_()),
    nowWeek: weekNo_(new Date())
  };
}

// рабочий вес по последним тренировкам каждого упражнения (для мини-графиков)
function spark_(log) {
  const out = {};
  log.filter(r => r[3] !== '').forEach(r => { (out[r[2]] = out[r[2]] || []).push(Number(r[3]) || 0); });
  Object.keys(out).forEach(k => { out[k] = out[k].slice(-6); });
  return out;
}

function weeks5_(s) {
  const now = weekNo_(new Date()), list = [];
  for (let w = Math.max(1, now - 4); w <= now; w++) list.push({ week: w, n: s.perWeek[w] || 0, current: w === now });
  return list;
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
  // две последние тренировки
  const sess = {}, order = [];
  log.forEach(r => {
    const iso = ymd_(r[0]);
    if (!sess[iso]) { sess[iso] = { date: iso, t: 0, sets: 0, prs: 0, names: [] }; order.push(iso); }
    const x = sess[iso];
    x.t += Number(r[6]) || 0; x.sets += Number(r[4]) || 0; x.names.push(r[2]);
    if (String(r[8]).indexOf('PR') === 0) x.prs++;
  });
  const recent = order.slice(-2).reverse().map(iso => {
    const x = sess[iso], pw = planDay_(x.names, 0);
    return { date: iso, group: APP.GROUPS[pw] || '', pw: pw, tonnage: Math.round(x.t), sets: x.sets, prs: x.prs };
  });
  return { weeks: weekList, names: names, points: points, recent: recent, start: ymd_(startDate_()), nowWeek: nowWeek };
}

// подробности одной недели программы: дни, упражнения, подходы, рекорды
function apiWeek_(n) {
  if (!(n >= 1)) throw new Error('Неверная неделя');
  const sh = sheet_(APP.TODAY);
  const ix = index_(sh);
  const k0 = ix.key(startDate_()) + (n - 1) * 7, k1 = k0 + 6;
  let s = -1, e = -1;
  ix.keys.forEach((k, i) => { if (k >= k0 && k <= k1) { if (s < 0) s = i; e = i; } });
  const res = { week: n, from: isoFromKey_(k0), to: isoFromKey_(k1), tonnage: 0, done: 0, total: 0, sets: 0, prs: 0, delta: null, days: [] };
  if (s < 0) return res;
  // Журнал с начала и до конца недели: тоннаж, рекорды, прошлые тренировки той же группы
  const jr = sheet_(APP.LOG).getRange(2, 1, APP.FIRST + e - 3 - 1, 9).getValues();
  const rows = sh.getRange(APP.FIRST + s, 1, e - s + 1, 13).getValues();
  const blocks = [];
  for (let i = 0; i <= e; i++) {
    if (i === 0 || ix.keys[i] !== ix.keys[i - 1]) blocks.push({ k: ix.keys[i], s: i, e: i });
    else blocks[blocks.length - 1].e = i;
  }
  const tOf = b => { let t = 0; for (let i = b.s; i <= b.e; i++) t += Number(jr[i][6]) || 0; return t; };
  const pwOf = b => { const nm = []; for (let i = b.s; i <= b.e; i++) if (jr[i][2]) nm.push(jr[i][2]); return planDay_(nm, wdFromKey_(b.k)); };
  const hasOf = b => { for (let i = b.s; i <= b.e; i++) if (jr[i][3] !== '') return true; return false; };
  let prevWeekT = 0;
  blocks.forEach((b, bi) => {
    if (b.k >= k0 - 7 && b.k < k0) prevWeekT += tOf(b);
    if (b.k < k0) return;
    const pw = pwOf(b), has = hasOf(b);
    let prevT = null;
    for (let j = bi - 1; j >= 0; j--) if (pwOf(blocks[j]) === pw && hasOf(blocks[j])) { prevT = Math.round(tOf(blocks[j])); break; }
    const day = { date: isoFromKey_(b.k), weekday: wdFromKey_(b.k), pw: pw, group: APP.GROUPS[pw] || '', done: has,
      tonnage: Math.round(tOf(b)), prevT: prevT, sets: 0, prs: [], exercises: [] };
    for (let i = b.s; i <= b.e; i++) {
      const r = rows[i - s], j = jr[i];
      if (!r[1]) continue;
      const sets = [];
      for (let q = 0; q < 4; q++) { const kg = num_(r[4 + 2 * q]), rp = num_(r[5 + 2 * q]); if (kg !== null || rp !== null) sets.push([kg, rp]); }
      const pr = String(j[8]).indexOf('PR') === 0;
      if (pr) day.prs.push({ name: r[1], kg: Number(j[3]) || 0 });
      day.sets += sets.length;
      if (sets.length) day.exercises.push({ name: r[1], sets: sets, tonnage: Math.round(Number(r[12]) || 0), pr: pr });
    }
    res.days.push(day);
    res.total++;
    if (has) res.done++;
    res.tonnage += day.tonnage; res.sets += day.sets; res.prs += day.prs.length;
  });
  if (prevWeekT > 0) res.delta = Math.round((res.tonnage / prevWeekT - 1) * 100);
  return res;
}

/* =====================================================================
 * БОТ: напоминания, команды, запись подходов текстом, итог месяца.
 * Один раз запусти setupReminders() — он включит оба таймера:
 *   remind() раз в час: 9:00 план, 20:00 пинок, вс 20:00 итог недели, 1-го числа итог месяца;
 *   poll()   раз в минуту: читает твои сообщения боту и отвечает.
 * ===================================================================== */
const REMIND = {
  APP_URL: 'https://k1vr1.github.io/zal-app/',
  MORNING: 9,
  EVENING: 20,
  MONTHLY: 10
};
const MONTHS = ['Январь', 'Февраль', 'Март', 'Апрель', 'Май', 'Июнь', 'Июль', 'Август', 'Сентябрь', 'Октябрь', 'Ноябрь', 'Декабрь'];

function setupReminders() {
  stopReminders();
  ScriptApp.newTrigger('remind').timeBased().everyHours(1).create();
  ScriptApp.newTrigger('poll').timeBased().everyMinutes(1).create();
  tgApi_('deleteWebhook', {});
  tgApi_('setMyCommands', { commands: [
    { command: 'today', description: 'План на сегодня' },
    { command: 'week', description: 'Итог недели' },
    { command: 'month', description: 'Итог месяца картинкой' },
    { command: 'records', description: 'Мои рекорды' },
    { command: 'weight', description: 'Вес тела' },
    { command: 'help', description: 'Что умеет бот' }
  ] });
  weightSheet_();
  tg_('✅ Бот включён: напоминания, команды и запись подходов текстом.\n\n' + helpText_());
}

function stopReminders() {
  ScriptApp.getProjectTriggers()
    .filter(t => ['remind', 'poll'].indexOf(t.getHandlerFunction()) >= 0)
    .forEach(t => ScriptApp.deleteTrigger(t));
}

function remind() {
  const now = new Date();
  const hour = Number(Utilities.formatDate(now, tz_(), 'H'));
  const wd = dayOfWeek_(now);
  const day = Number(Utilities.formatDate(now, tz_(), 'd'));
  if (hour === REMIND.MORNING) once_('morning', morningText_, true);
  if (hour === REMIND.EVENING) once_('evening', eveningText_, true);
  if (wd === 7 && hour === REMIND.EVENING) once_('week', () => weekText_(false, true), false);
  if (day === 1 && hour === REMIND.MONTHLY) {
    const props = PropertiesService.getScriptProperties();
    const tag = Utilities.formatDate(now, tz_(), 'yyyy-MM');
    if (props.getProperty('sent_month') !== tag) {
      props.setProperty('sent_month', tag);
      const prev = new Date(now.getTime() - 5 * 86400000);
      sendMonth_(Number(Utilities.formatDate(prev, tz_(), 'yyyy')), Number(Utilities.formatDate(prev, tz_(), 'M')) - 1);
    }
  }
}

// отправляет не больше одного сообщения каждого вида в день
function once_(kind, build, button) {
  const props = PropertiesService.getScriptProperties();
  const key = 'sent_' + kind;
  const day = ymd_(new Date());
  if (props.getProperty(key) === day) return;
  props.setProperty(key, day);
  const text = build();
  if (text) tg_(text, button);
}

function sessionToday_() {
  const sh = sheet_(APP.TODAY);
  const ix = index_(sh);
  return { sh: sh, ix: ix, s: readSession_(sh, ix.keys, ix.key(new Date()), true) };
}

function planText_(s) {
  const lines = s.exercises.filter(e => e.name).map(e => {
    const last = e.last && e.last !== '—' && e.last.indexOf('←') < 0 ? ' — ' + e.last.trim().split(/\s{2,}/).join(' · ') : '';
    return '• <b>' + esc_(e.name) + '</b> ' + esc_(e.plan) + esc_(last);
  });
  return lines.join('\n');
}

function morningText_() {
  const s = sessionToday_().s;
  if (!s) return '';
  return '💪 <b>Сегодня: ' + esc_(s.group) + '</b>\n\n' + planText_(s) + '\n\nПодсказки по весу — в приложении 👇';
}

function eveningText_() {
  const t = sessionToday_();
  if (!t.s || hasData_(t.s)) return '';
  return '⏰ Сегодня по плану <b>' + esc_(t.s.group) + '</b>, а в журнале пока пусто.\n' +
    'Если уже позанимался — отметь подходы. Не успеваешь — тренировку можно перенести на завтра в приложении.';
}

function fmtT_(n) { return String(Math.round(n)).replace(/\B(?=(\d{3})+(?!\d))/g, ' '); }
function fmtN_(n) { return String(Math.round(n * 100) / 100).replace('.', ','); }

function weekText_(dry, askWeight) {
  const ix = index_(sheet_(APP.TODAY));
  const todayKey = ix.key(new Date());
  const cur = weekStats_(ix.keys, todayKey);
  const prev = weekStats_(ix.keys, todayKey - 7);
  const props = PropertiesService.getScriptProperties();
  const prsBefore = Number(props.getProperty('prs_last_week') || cur.prs);
  if (!dry) props.setProperty('prs_last_week', String(cur.prs));
  const newPrs = Math.max(0, cur.prs - prsBefore);
  const st = stats_();

  let delta = '';
  if (prev.tonnage > 0) {
    const pct = Math.round((cur.tonnage / prev.tonnage - 1) * 100);
    delta = ' (' + (pct > 0 ? '+' : '') + pct + '% к прошлой)';
  }
  const mark = cur.done >= 3 ? '🔥 Все три тренировки!' :
    cur.done === 0 ? (askWeight ? 'На этой неделе тренировок не было — с понедельника заново 💪' : 'Пока 0 из 3 — всё впереди 💪') :
    'Тренировок: ' + cur.done + ' из 3.';
  let text = '📊 <b>Итог недели</b>\n\n' + mark + '\n' +
    (cur.done ? 'Тоннаж: <b>' + fmtT_(cur.tonnage) + ' кг</b>' + delta + '\n' : '') +
    (newPrs ? 'Новых рекордов: <b>' + newPrs + '</b> 🏆\n' : '') +
    (st.streak >= 2 ? 'Серия: <b>' + st.streak + ' нед. подряд</b> без пропусков 🔥\n' : '');
  if (askWeight) text += '\n⚖️ Взвесься и пришли вес числом, например <code>71,5</code> — запишу в таблицу.';
  return text;
}

function recordsText_() {
  const rows = sheet_(APP.PROG).getRange('B7:L27').getDisplayValues().filter(r => r[1]);
  const names = { 'Пн': 'Понедельник', 'Ср': 'Среда', 'Пт': 'Пятница' };
  let text = '🏆 <b>Лучшие веса</b>\n';
  ['Пн', 'Ср', 'Пт'].forEach(d => {
    const list = rows.filter(r => r[0] === d);
    if (!list.length) return;
    text += '\n<b>' + names[d] + '</b>\n';
    list.forEach(r => {
      const best = r[5] ? r[5] + ' кг' : (r[8] ? r[8] + ' раз' : '—');
      text += '• ' + esc_(r[1]) + ' — <b>' + esc_(best) + '</b>' + (r[7] && r[7] !== '0%' ? ' (' + esc_(r[7]) + ')' : '') + '\n';
    });
  });
  return text;
}

function weightText_() {
  const list = weights_(10);
  if (!list.length) return '⚖️ Вес ещё не записан. Пришли число, например <code>71,5</code>.';
  const last = list[list.length - 1], first = list[0];
  const d = last.kg - first.kg;
  return '⚖️ Сейчас: <b>' + fmtN_(last.kg) + ' кг</b> (' + last.date.slice(8) + '.' + last.date.slice(5, 7) + ')' +
    (list.length > 1 ? '\nЗа последние ' + list.length + ' записей: ' + (d > 0 ? '+' : '') + fmtN_(d) + ' кг' : '') +
    '\n\nЧтобы записать новый — просто пришли число.';
}

function helpText_() {
  return 'Что я умею:\n' +
    '• /today — план на сегодня\n' +
    '• /week — итог недели, /month — итог месяца картинкой\n' +
    '• /records — лучшие веса\n' +
    '• пришли число, например <code>71,5</code> — запишу вес тела\n' +
    '• пришли <code>жим 40 10</code> — запишу подход в сегодняшнюю тренировку (<code>жим 40 10 3</code> — сразу 3 подхода), /undo — отменить';
}

/* ---------- приём сообщений (раз в минуту) ---------- */

function poll() {
  const lock = LockService.getScriptLock();
  if (!lock.tryLock(1000)) return;
  let updates = [];
  try {
    const props = PropertiesService.getScriptProperties();
    const offset = Number(props.getProperty('tg_offset') || 0);
    const res = tgApi_('getUpdates', { offset: offset, timeout: 0, allowed_updates: ['message'] });
    updates = (res && res.result) || [];
    if (updates.length) props.setProperty('tg_offset', String(updates[updates.length - 1].update_id + 1));
  } finally {
    lock.releaseLock();
  }
  const allowed = String(PropertiesService.getScriptProperties().getProperty('ALLOWED_USER_ID'));
  updates.forEach(u => {
    const m = u.message;
    if (!m || !m.text || !m.chat) return;
    if (String(m.chat.id) !== allowed) { tgTo_(m.chat.id, 'Это личный бот 🙂'); return; }
    try { tg_(handle_(m.text.trim()), true); } catch (err) { tg_('⚠️ ' + esc_(err.message || err)); }
  });
}

function handle_(text) {
  const cmd = text.toLowerCase().split(/[\s@]/)[0];
  if (cmd === '/start' || cmd === '/help') return helpText_();
  if (cmd === '/today') {
    const s = sessionToday_().s;
    return s ? '💪 <b>Сегодня: ' + esc_(s.group) + '</b>\n\n' + planText_(s) : '💤 Сегодня день отдыха.';
  }
  if (cmd === '/week') return weekText_(true, false);
  if (cmd === '/records') return recordsText_();
  if (cmd === '/weight') return weightText_();
  if (cmd === '/undo') return undo_();
  if (cmd === '/month') {
    const now = new Date();
    sendMonth_(Number(Utilities.formatDate(now, tz_(), 'yyyy')), Number(Utilities.formatDate(now, tz_(), 'M')) - 1);
    return '';
  }
  if (/^\d{2,3}([.,]\d{1,2})?$/.test(text)) {
    const kg = addWeight_(text);
    const list = weights_(2);
    const prev = list.length > 1 ? list[0].kg : null;
    return '⚖️ Записал: <b>' + fmtN_(kg) + ' кг</b>' +
      (prev !== null ? ' (' + (kg - prev > 0 ? '+' : '') + fmtN_(kg - prev) + ' с прошлого раза)' : '');
  }
  const set = parseSet_(text);
  if (set) return logSet_(set);
  return 'Не понял 🤔\n\n' + helpText_();
}

/* ---------- запись подхода текстом: «жим 40 10», «жим 40х10 3» ---------- */

function norm_(s) { return String(s).toLowerCase().replace(/ё/g, 'е').replace(/[^a-zа-я0-9 ]/g, ' ').replace(/\s+/g, ' ').trim(); }

function parseSet_(text) {
  const m = String(text).trim().match(/^(.*?[a-zа-яё][^\d]*?)\s+(\d+(?:[.,]\d+)?)\s*(?:кг)?\s*(?:[xх×*]|\s)\s*(\d+)(?:\s*(?:раз)?\s*(?:[xх×*]|\s)\s*(\d))?\s*$/i);
  if (!m) return null;
  return { q: m[1], kg: Number(m[2].replace(',', '.')), reps: Number(m[3]), times: m[4] ? Number(m[4]) : 1 };
}

function findExercise_(list, q) {
  const words = norm_(q).split(' ').filter(String);
  const score = e => {
    const nw = norm_(e.name).split(' ');
    if (!words.every(w => nw.some(x => x.indexOf(w) === 0))) return 0;
    return norm_(e.name).indexOf(norm_(q)) === 0 ? 2 : 1;
  };
  const hits = list.map(e => ({ e: e, s: score(e) })).filter(x => x.s > 0);
  if (!hits.length) return { none: true };
  const top = Math.max.apply(null, hits.map(x => x.s));
  const best = hits.filter(x => x.s === top);
  if (best.length > 1) {
    const exact = best.filter(x => norm_(x.e.name) === norm_(q));
    if (exact.length === 1) return { e: exact[0].e };
    // «жим» → самое короткое подходящее («Жим лежа»); остальные можно уточнить словом
    best.sort((a, b) => a.e.name.length - b.e.name.length);
    return { e: best[0].e, others: best.slice(1).map(x => x.e.name) };
  }
  return { e: best[0].e };
}

function logSet_(p) {
  const t = sessionToday_();
  if (!t.s) return '💤 Сегодня по плану отдых — записывать некуда. Если тренировка перенесена, перенеси её в приложении.';
  if (!(p.times >= 1 && p.times <= 4)) throw new Error('Подходов можно от 1 до 4');
  const list = t.s.exercises.filter(e => e.name);
  const f = findExercise_(list, p.q);
  if (f.none) return 'Не нашёл «' + esc_(p.q) + '» в сегодняшней тренировке.\nСегодня: ' + list.map(e => esc_(e.name)).join(', ');
  const e = f.e;
  const free = [];
  e.sets.forEach((s, i) => { if (s[0] === null && s[1] === null) free.push(i); });
  if (free.length < p.times) throw new Error('В «' + e.name + '» свободно подходов: ' + free.length + ' из 4');
  const items = free.slice(0, p.times).map(i => ({ row: e.row, set: i, kg: p.kg, reps: p.reps }));
  apiSaveMany_(items);
  PropertiesService.getScriptProperties().setProperty('last_logged', JSON.stringify({ name: e.name, items: items }));
  const nums = items.map(x => x.set + 1).join(', ');
  return '✅ <b>' + esc_(e.name) + '</b>: ' + (p.times > 1 ? 'подходы ' : 'подход ') + nums + ' — ' + fmtN_(p.kg) + '×' + p.reps +
    (f.others && f.others.length ? '\nНе то упражнение? /undo и напиши точнее: ' + f.others.map(esc_).join(', ') : '\nОшибся — /undo');
}

function undo_() {
  const props = PropertiesService.getScriptProperties();
  const last = JSON.parse(props.getProperty('last_logged') || 'null');
  if (!last) return 'Нечего отменять.';
  apiSaveMany_(last.items.map(x => ({ row: x.row, set: x.set, kg: '', reps: '' })));
  props.deleteProperty('last_logged');
  return '↩️ Отменил запись в «' + esc_(last.name) + '».';
}

/* ---------- итог месяца картинкой ---------- */

function monthData_(year, month) {
  const log = logRows_().filter(r => r[3] !== '');
  const ym = d => Utilities.formatDate(d, tz_(), 'yyyy-MM');
  const pad = n => ('0' + (n + 1)).slice(-2);
  const cur = year + '-' + pad(month);
  const prevTag = month === 0 ? (year - 1) + '-12' : year + '-' + pad(month - 1);
  const days = {}, order = [];
  let tonnage = 0, prs = 0, prevT = 0;
  log.forEach(r => {
    const tag = ym(r[0]);
    if (tag === prevTag) prevT += Number(r[6]) || 0;
    if (tag !== cur) return;
    const iso = ymd_(r[0]);
    if (!days[iso]) { days[iso] = { t: 0, names: [] }; order.push(iso); }
    days[iso].t += Number(r[6]) || 0;
    days[iso].names.push(r[2]);
    tonnage += Number(r[6]) || 0;
    if (String(r[8]).indexOf('PR') === 0) prs++;
  });
  return {
    title: MONTHS[month] + ' ' + year,
    workouts: order.length, tonnage: Math.round(tonnage), prs: prs,
    delta: prevT ? Math.round((tonnage / prevT - 1) * 100) : null,
    bars: order.map(iso => ({ label: iso.slice(8) + '.' + iso.slice(5, 7), t: Math.round(days[iso].t), pw: planDay_(days[iso].names, 0) }))
  };
}

function monthChart_(d) {
  const colors = { 1: '#C6FF3D', 3: '#FF8A3D', 5: '#3DD9FF', 0: '#8E95A6' };
  const sub = d.workouts + ' трен.  ·  ' + fmtN_(Math.round(d.tonnage / 100) / 10) + ' т  ·  рекордов: ' + d.prs +
    (d.delta !== null ? '  ·  ' + (d.delta > 0 ? '+' : '') + d.delta + '% к прошлому' : '');
  return "{type:'bar',data:{labels:" + JSON.stringify(d.bars.map(b => b.label)) +
    ",datasets:[{data:" + JSON.stringify(d.bars.map(b => b.t)) + ",backgroundColor:" + JSON.stringify(d.bars.map(b => colors[b.pw] || colors[0])) + ",borderWidth:0,maxBarThickness:64}]}," +
    "options:{layout:{padding:{left:16,right:24,top:8,bottom:8}},legend:{display:false}," +
    "title:{display:true,text:" + JSON.stringify([d.title.toUpperCase(), sub]) + ",fontColor:'#F2F4F8',fontSize:24,fontStyle:'bold',padding:24}," +
    "scales:{xAxes:[{gridLines:{display:false},ticks:{fontColor:'#8E95A6',fontSize:13}}]," +
    "yAxes:[{gridLines:{color:'#2E3340',zeroLineColor:'#2E3340'},ticks:{fontColor:'#8E95A6',beginAtZero:true,callback:(v)=>(v/1000)+' т'}}]}," +
    "plugins:{datalabels:{anchor:'end',align:'top',color:'#D9DEE8',font:{weight:'bold',size:12},formatter:(v)=>(v/1000).toFixed(1).replace('.',',')}}}}";
}

function sendMonth_(year, month) {
  const d = monthData_(year, month);
  if (!d.workouts) { tg_('📅 <b>' + d.title + '</b>: тренировок пока нет.'); return; }
  const caption = '📅 <b>Итог: ' + d.title + '</b>\n\nТренировок: <b>' + d.workouts + '</b>\nТоннаж: <b>' + fmtT_(d.tonnage) + ' кг</b>' +
    (d.delta !== null ? ' (' + (d.delta > 0 ? '+' : '') + d.delta + '% к прошлому месяцу)' : '') +
    '\nРекордов: <b>' + d.prs + '</b>' + (d.prs ? ' 🏆' : '');
  const img = UrlFetchApp.fetch('https://quickchart.io/chart', {
    method: 'post', contentType: 'application/json', muteHttpExceptions: true,
    payload: JSON.stringify({ version: '2', backgroundColor: '#121418', width: 800, height: 480, devicePixelRatio: 2, format: 'png', chart: monthChart_(d) })
  });
  if (img.getResponseCode() !== 200) { tg_(caption); return; }
  const props = PropertiesService.getScriptProperties();
  const res = UrlFetchApp.fetch('https://api.telegram.org/bot' + props.getProperty('BOT_TOKEN') + '/sendPhoto', {
    method: 'post', muteHttpExceptions: true,
    payload: { chat_id: props.getProperty('ALLOWED_USER_ID'), caption: caption, parse_mode: 'HTML', photo: img.getBlob().setName('month.png') }
  });
  if (res.getResponseCode() !== 200) tg_(caption);
}

/* ---------- Telegram ---------- */

function esc_(s) {
  return String(s == null ? '' : s).replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;');
}

function tgApi_(method, body) {
  const token = PropertiesService.getScriptProperties().getProperty('BOT_TOKEN');
  if (!token) throw new Error('Нужен BOT_TOKEN в свойствах скрипта');
  const res = UrlFetchApp.fetch('https://api.telegram.org/bot' + token + '/' + method, {
    method: 'post', contentType: 'application/json', payload: JSON.stringify(body), muteHttpExceptions: true
  });
  if (res.getResponseCode() !== 200) throw new Error('Telegram: ' + res.getContentText());
  return JSON.parse(res.getContentText());
}

function tgTo_(chat, text, button) {
  if (!text) return;
  const body = { chat_id: chat, text: text, parse_mode: 'HTML', disable_web_page_preview: true };
  if (button) body.reply_markup = { inline_keyboard: [[{ text: '🏋️ Открыть тренировку', web_app: { url: REMIND.APP_URL } }]] };
  tgApi_('sendMessage', body);
}

function tg_(text, button) {
  const chat = PropertiesService.getScriptProperties().getProperty('ALLOWED_USER_ID');
  if (!chat) throw new Error('Нужен ALLOWED_USER_ID в свойствах скрипта');
  tgTo_(chat, text, button);
}

// для проверки вручную: шлёт все сообщения прямо сейчас
function testReminders() {
  tg_(morningText_() || '(сегодня нет тренировки — утреннего сообщения нет)', true);
  tg_(eveningText_() || '(вечернего пинка нет: день отдыха или подходы уже записаны)', true);
  tg_(weekText_(true, true));
  const now = new Date();
  sendMonth_(Number(Utilities.formatDate(now, tz_(), 'yyyy')), Number(Utilities.formatDate(now, tz_(), 'M')) - 1);
}
