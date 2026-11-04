/**
 * Бэкенд мини-приложения «MY GYM» (версия 3: данные живут в Telegram).
 *
 * Что он делает:
 *   1. export — один раз отдаёт владельцу его старые тренировки из Google Таблицы (для переноса);
 *   2. sync   — принимает от приложения расписание пользователя (на 2 недели вперёд);
 *   3. remind — раз в час рассылает напоминания всем, у кого они включены, по их часовому поясу.
 *
 * Свойства скрипта (Настройки проекта → Свойства скрипта):
 *   BOT_TOKEN       — токен бота от @BotFather
 *   ALLOWED_USER_ID — твой Telegram id (только тебе доступен перенос из таблицы)
 *
 * После замены кода: запусти setupReminders() один раз и обнови развёртывание (новая версия).
 * Саму таблицу не удаляй — в ней живёт этот скрипт. Открывать её больше не нужно.
 */
const APP = {
  MAX_AGE: 7 * 24 * 3600,          // сколько секунд живёт подпись Telegram
  APP_URL: 'https://k1vr1.github.io/zal-app/',
  MAX_USERS: 400                    // защита от переполнения свойств скрипта
};

function doGet() {
  return out_({ ok: true, app: 'mygym', v: 3 });
}

function doPost(e) {
  try {
    const req = JSON.parse((e && e.postData && e.postData.contents) || '{}');
    const user = auth_(req.initData);
    switch (req.action) {
      case 'export': return out_({ ok: true, data: apiExport_(user) });
      case 'sync':   return out_({ ok: true, data: apiSync_(user, req) });
      case 'ping':   return out_({ ok: true, data: { id: user.id } });
      default:       return out_({ ok: false, error: 'Неизвестное действие' });
    }
  } catch (err) {
    return out_({ ok: false, error: String((err && err.message) || err) });
  }
}

function out_(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj)).setMimeType(ContentService.MimeType.JSON);
}

/* ---------- проверка, что запрос пришёл из Telegram ---------- */

function auth_(initData) {
  const token = PropertiesService.getScriptProperties().getProperty('BOT_TOKEN');
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
  const secret = Utilities.computeHmacSha256Signature(Utilities.newBlob(token).getBytes(), Utilities.newBlob('WebAppData').getBytes());
  const sign = hex_(Utilities.computeHmacSha256Signature(Utilities.newBlob(check).getBytes(), secret));
  if (sign !== hash) throw new Error('Подпись Telegram не прошла проверку');
  if (Date.now() / 1000 - Number(params.auth_date || 0) > APP.MAX_AGE) throw new Error('Сессия устарела, перезапусти приложение');
  const user = JSON.parse(params.user || '{}');
  if (!user.id) throw new Error('Нет пользователя');
  return user;
}

function hex_(bytes) { return bytes.map(b => ('0' + (b & 0xff).toString(16)).slice(-2)).join(''); }

/* ---------- 1. перенос старых данных из таблицы (только владельцу) ---------- */

function apiExport_(user) {
  const owner = PropertiesService.getScriptProperties().getProperty('ALLOWED_USER_ID');
  if (!owner || String(user.id) !== String(owner)) throw new Error('Нет доступа');
  const ss = SpreadsheetApp.getActive();
  const tz = ss.getSpreadsheetTimeZone();
  const ymd = d => Utilities.formatDate(d, tz, 'yyyy-MM-dd');
  const num = v => (v === '' || v === null || v === undefined) ? null : Number(v);

  // программа: день, упражнение, подходы, повторы, старт, ..., заметка
  const prog = ss.getSheetByName('Прогресс').getRange('B7:M27').getValues()
    .filter(r => String(r[1]).trim())
    .map(r => ({ day: String(r[0]), name: String(r[1]).trim(), sets: num(r[2]), reps: r[3] === 'макс' ? null : num(r[3]), start: num(r[4]), note: String(r[11] || '').trim() }));

  // тренировки: все строки «Сегодня», где есть хоть один подход
  const sh = ss.getSheetByName('Сегодня');
  const rows = sh.getRange(5, 1, sh.getMaxRows() - 4, 12).getValues();
  const sessions = {}, order = [];
  for (let i = 0; i < rows.length && rows[i][0] instanceof Date; i++) {
    const r = rows[i], name = String(r[1] || '').trim();
    const sets = [];
    for (let j = 0; j < 4; j++) {
      const kg = num(r[4 + 2 * j]), reps = num(r[5 + 2 * j]);
      sets.push(kg === null && reps === null ? null : [kg, reps]);
    }
    if (!name || !sets.some(Boolean)) continue;
    const d = ymd(r[0]);
    if (!sessions[d]) { sessions[d] = []; order.push(d); }
    sessions[d].push({ name: name, sets: sets });
  }

  const wsh = ss.getSheetByName('Вес');
  const weights = wsh && wsh.getLastRow() > 1
    ? wsh.getRange(2, 1, wsh.getLastRow() - 1, 2).getValues().filter(r => r[0] instanceof Date && r[1] !== '').map(r => [ymd(r[0]), Number(r[1])])
    : [];

  return { program: prog, sessions: order.map(d => ({ date: d, ex: sessions[d] })), weights: weights };
}

/* ---------- 2. расписание пользователя для напоминаний ---------- */
// req: { on, tz (минуты к UTC), m (час утра), e (час вечера), plan: { 'yyyy-MM-dd': { t, x: [названия] } }, done: ['yyyy-MM-dd'] }

function apiSync_(user, req) {
  const props = PropertiesService.getScriptProperties();
  const key = 'u_' + user.id;
  if (!req.on) { props.deleteProperty(key); return { on: false }; }
  const cur = JSON.parse(props.getProperty(key) || '{}');
  if (!cur.id) {
    const n = Object.keys(props.getProperties()).filter(k => k.indexOf('u_') === 0).length;
    if (n >= APP.MAX_USERS) throw new Error('Напоминания временно недоступны');
  }
  const clampH = h => Math.max(0, Math.min(23, Math.round(Number(h) || 0)));
  const plan = {};
  Object.keys(req.plan || {}).slice(0, 21).forEach(d => {
    if (!/^\d{4}-\d{2}-\d{2}$/.test(d)) return;
    const p = req.plan[d] || {};
    plan[d] = { t: String(p.t || 'Тренировка').slice(0, 40), x: (p.x || []).slice(0, 12).map(s => String(s).slice(0, 40)) };
  });
  const done = (req.done || []).filter(d => /^\d{4}-\d{2}-\d{2}$/.test(d)).slice(-5);
  const rec = {
    id: user.id, tz: Math.max(-720, Math.min(840, Math.round(Number(req.tz) || 0))),
    m: clampH(req.m == null ? 9 : req.m), e: clampH(req.e == null ? 20 : req.e),
    plan: plan, done: done, sm: cur.sm || '', se: cur.se || ''
  };
  props.setProperty(key, JSON.stringify(rec));
  return { on: true };
}

/* ---------- 3. рассылка напоминаний ---------- */

function setupReminders() {
  ScriptApp.getProjectTriggers()
    .filter(t => ['remind', 'poll'].indexOf(t.getHandlerFunction()) >= 0)
    .forEach(t => ScriptApp.deleteTrigger(t));
  ScriptApp.newTrigger('remind').timeBased().everyHours(1).create();
  // старые свойства от версии с таблицей больше не нужны
  const props = PropertiesService.getScriptProperties();
  ['sent_morning', 'sent_evening', 'sent_week', 'sent_month', 'prs_last_week', 'tg_offset', 'last_logged'].forEach(k => props.deleteProperty(k));
  try { tgApi_('deleteMyCommands', {}); } catch (e) {}
  const owner = props.getProperty('ALLOWED_USER_ID');
  if (owner) tgTo_(owner, '✅ Бот обновлён: напоминания теперь берут расписание из приложения. Открой приложение один раз, чтобы оно передало расписание.', true);
}

function localNow_(tzMin) {
  const d = new Date(Date.now() + tzMin * 60000);
  return { date: d.toISOString().slice(0, 10), hour: d.getUTCHours() };
}

function remind() {
  const props = PropertiesService.getScriptProperties();
  const all = props.getProperties();
  Object.keys(all).filter(k => k.indexOf('u_') === 0).forEach(k => {
    let u;
    try { u = JSON.parse(all[k]); } catch (e) { return; }
    const now = localNow_(u.tz || 0);
    const p = u.plan && u.plan[now.date];
    let changed = false;
    try {
      if (p && now.hour === u.m && u.sm !== now.date) {
        u.sm = now.date; changed = true;
        tgTo_(u.id, '💪 <b>Сегодня: ' + esc_(p.t) + '</b>\n\n' + p.x.map(x => '• ' + esc_(x)).join('\n') + '\n\nВеса и подсказки — в приложении 👇', true);
      }
      if (p && now.hour === u.e && u.se !== now.date) {
        u.se = now.date; changed = true;
        if ((u.done || []).indexOf(now.date) < 0) {
          tgTo_(u.id, '⏰ Сегодня по плану <b>' + esc_(p.t) + '</b>, а в журнале пока пусто.\nЕсли уже позанимался — отметь подходы. Не успеваешь — перенеси тренировку в приложении.', true);
        }
      }
    } catch (err) {
      // пользователь заблокировал бота — убираем его из рассылки
      if (/blocked|deactivated|chat not found/i.test(String(err.message))) { props.deleteProperty(k); return; }
    }
    if (changed) props.setProperty(k, JSON.stringify(u));
  });
}

/* ---------- Telegram ---------- */

function esc_(s) { return String(s == null ? '' : s).replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;'); }

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
  const body = { chat_id: chat, text: text, parse_mode: 'HTML', disable_web_page_preview: true };
  if (button) body.reply_markup = { inline_keyboard: [[{ text: '🏋️ Открыть тренировку', web_app: { url: APP.APP_URL } }]] };
  tgApi_('sendMessage', body);
}

// для проверки вручную: шлёт тебе утреннее напоминание по твоему расписанию на сегодня
function testReminders() {
  const props = PropertiesService.getScriptProperties();
  const owner = props.getProperty('ALLOWED_USER_ID');
  const u = JSON.parse(props.getProperty('u_' + owner) || 'null');
  if (!u) { tgTo_(owner, 'Расписание ещё не пришло: открой приложение и включи напоминания в настройках.', true); return; }
  const now = localNow_(u.tz || 0), p = u.plan[now.date];
  tgTo_(owner, p ? '💪 <b>Сегодня: ' + esc_(p.t) + '</b>\n\n' + p.x.map(x => '• ' + esc_(x)).join('\n') : 'Сегодня по расписанию отдых. Дней в расписании: ' + Object.keys(u.plan).length, true);
}
