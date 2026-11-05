/**
 * Сервер бота MY GYM на Cloudflare Workers.
 *
 * Что делает:
 *   POST /  { action: 'sync' }   — приложение присылает расписание пользователя на 2 недели вперёд;
 *   POST /  { action: 'test' }   — прислать пользователю пробное напоминание;
 *   POST /  { action: 'export' } — один раз отдать владельцу его старые тренировки (ключ export в KV);
 *   POST /  { action: 'report' } — прислать отчёт о тренировке: у всех по правилам, у владельца по желанию от ИИ;
 *   POST /  { action: 'backup' } — прислать пользователю файлы с его тренировками (на сервере не хранятся);
 *   POST /  { action: 'card' }   — прислать пользователю картинку с его рекордом, чтобы переслать другу;
 *   POST /  с заголовком секрета — сообщения боту от Telegram: отвечает на /start и /help;
 *   по расписанию (cron «0 * * * *», раз в час) — напоминания, итог недели и предупреждение о конце расписания
 *   каждому по его часовому поясу.
 *
 * Настройки воркера (Settings):
 *   Secret   BOT_TOKEN — токен бота от @BotFather
 *   Secret   ANTHROPIC_API_KEY — ключ Claude API для ИИ-отчёта (без него приходит отчёт по правилам)
 *   Secret   WEBHOOK_SECRET — случайная строка: так воркер отличает сообщения от Telegram (без неё бот не отвечает на /start)
 *   Text     OWNER_ID  — Telegram id владельца
 *   KV       KV        — хранилище расписаний (binding с именем KV)
 *   Cron     0 * * * *
 */
import Anthropic from '@anthropic-ai/sdk';

// Старые тренировки владельца для одноразового переноса из Google Таблицы лежат в KV (ключ export), не в репозитории.
const OWNER_EXPORT = null;

const APP_URL = 'https://k1vr1.github.io/zal-app/';
const WORKER_URL = 'https://mygym.oleg25092007.workers.dev/';
const MAX_AGE = 24 * 3600; // сколько секунд живёт подпись Telegram
const SYNC_PER_DAY = 40; // сколько раз в сутки один человек может менять расписание в KV
const AI_TIMEOUT = 25000; // воркеру дают около 30 секунд, если приложение уже закрыли
const AI_MAX_FACTS = 8000; // больше символов данных об одной тренировке не бывает
const BACKUP_MAX = 2000000; // предел для файла с копией, символов
const CARD_MAX = 4000000; // предел для картинки рекорда, символов base64

const CORS = { 'Access-Control-Allow-Origin': '*', 'Access-Control-Allow-Methods': 'GET, POST, OPTIONS', 'Access-Control-Allow-Headers': 'Content-Type' };
const json = obj => new Response(JSON.stringify(obj), { headers: Object.assign({ 'Content-Type': 'application/json; charset=utf-8' }, CORS) });

export default {
  async fetch(req, env, ctx) {
    if (req.method === 'OPTIONS') return new Response(null, { headers: CORS });
    if (req.method !== 'POST') return json({ ok: true, app: 'mygym', v: 5 });
    const hook = req.headers.get('X-Telegram-Bot-Api-Secret-Token');
    if (hook !== null) {
      if (!env.WEBHOOK_SECRET || hook !== env.WEBHOOK_SECRET) return new Response('Нет доступа', { status: 403 });
      try { await onUpdate(JSON.parse(await req.text() || '{}'), env); } catch (err) { console.error('Сообщение боту', String((err && err.message) || err)); }
      return new Response('ok'); // иначе Telegram будет присылать то же сообщение снова
    }
    try {
      const body = JSON.parse(await req.text() || '{}');
      const user = await auth(body.initData, env);
      const { success } = await env.LIMIT.limit({ key: String(user.id) });
      if (!success) throw new Error('Слишком много запросов, подожди минуту');
      if (body.action === 'sync') return json({ ok: true, data: await sync(user, body, env) });
      if (body.action === 'test') return json({ ok: true, data: await test(user, env) });
      if (body.action === 'report') {
        const job = report(user, body, env);
        ctx.waitUntil(job.catch(() => {})); // отчёт дойдёт, даже если приложение закроют раньше
        return json({ ok: true, data: await job });
      }
      if (body.action === 'card') return json({ ok: true, data: await card(user, body, env) });
      if (body.action === 'backup') return json({ ok: true, data: await backup(user, body, env) });
      if (body.action === 'export') {
        if (String(user.id) !== String(env.OWNER_ID)) throw new Error('Нет доступа');
        const data = OWNER_EXPORT || JSON.parse((await env.KV.get('export')) || 'null');
        if (!data) throw new Error('Данных для переноса нет');
        return json({ ok: true, data: data });
      }
      if (body.action === 'ping') return json({ ok: true, data: { id: user.id } });
      return json({ ok: false, error: 'Неизвестное действие' });
    } catch (err) {
      return json({ ok: false, error: String((err && err.message) || err) });
    }
  },
  async scheduled(event, env, ctx) {
    ctx.waitUntil(Promise.allSettled([remind(env), ensureWebhook(env)]));
  }
};

/* ---------- проверка, что запрос пришёл из Telegram ---------- */
const enc = new TextEncoder();
async function hmac(key, msg) {
  const k = await crypto.subtle.importKey('raw', typeof key === 'string' ? enc.encode(key) : key, { name: 'HMAC', hash: 'SHA-256' }, false, ['sign']);
  return new Uint8Array(await crypto.subtle.sign('HMAC', k, enc.encode(msg)));
}
// crypto.subtle.verify сравнивает подпись за постоянное время
async function signatureOk(secret, msg, hash) {
  if (!/^[0-9a-f]{64}$/.test(hash || '')) return false;
  const key = await crypto.subtle.importKey('raw', secret, { name: 'HMAC', hash: 'SHA-256' }, false, ['verify']);
  const sig = Uint8Array.from(hash.match(/../g), h => parseInt(h, 16));
  return crypto.subtle.verify('HMAC', key, sig, enc.encode(msg));
}
async function auth(initData, env) {
  if (!env.BOT_TOKEN) throw new Error('На сервере не задан BOT_TOKEN');
  if (!initData) throw new Error('Открой приложение через Telegram');
  const params = new URLSearchParams(initData);
  const hash = params.get('hash');
  params.delete('hash');
  const check = Array.from(params.keys()).sort().map(k => k + '=' + params.get(k)).join('\n');
  const secret = await hmac('WebAppData', env.BOT_TOKEN);
  if (!(await signatureOk(secret, check, hash))) throw new Error('Подпись Telegram не прошла проверку');
  if (Date.now() / 1000 - Number(params.get('auth_date') || 0) > MAX_AGE) throw new Error('Сессия устарела, перезапусти приложение');
  const user = JSON.parse(params.get('user') || '{}');
  if (!user.id) throw new Error('Нет пользователя');
  return user;
}

/* ---------- расписание пользователя ---------- */
const isDate = d => /^\d{4}-\d{2}-\d{2}$/.test(d);
async function sync(user, req, env) {
  const key = 'u:' + user.id;
  // счётчик записей хранится в самой записи: так лимит не тратит лишние записи в KV
  const old = JSON.parse((await env.KV.get(key)) || 'null');
  const today = new Date().toISOString().slice(0, 10);
  const count = old && old.wd === today ? old.wn || 0 : 0;
  if (count >= SYNC_PER_DAY) throw new Error('Слишком много изменений за сегодня, попробуй завтра');
  const hour = (h, def) => Math.max(0, Math.min(23, Math.round(Number(h == null ? def : h) || 0)));
  const plan = {};
  // выключенные напоминания — это пустой план: рассылка таких пропускает, а счётчик остаётся
  if (req.on) Object.keys(req.plan || {}).filter(isDate).slice(0, 21).forEach(d => {
    const p = req.plan[d] || {};
    plan[d] = { t: String(p.t || 'Тренировка').slice(0, 40), x: (p.x || []).slice(0, 12).map(s => String(s).slice(0, 40)) };
  });
  const rec = {
    id: user.id, tz: Math.max(-720, Math.min(840, Math.round(Number(req.tz) || 0))),
    m: hour(req.m, 9), e: hour(req.e, 20), plan: plan, done: (req.done || []).filter(isDate).slice(-5),
    wd: today, wn: count + 1, week: cleanWeek(req.week),
    warned: old && old.warned, wsent: old && old.wsent // чтобы обновление расписания не присылало их повторно
  };
  await env.KV.put(key, JSON.stringify(rec));
  return { on: !!req.on };
}

function cleanWeek(w) {
  if (!w || !isDate(w.from)) return null;
  const list = v => (Array.isArray(v) ? v : []).slice(0, 5).map(s => String(s).slice(0, 60));
  const num = v => Math.max(0, Math.round(Number(v) || 0));
  return { from: w.from, done: num(w.done), target: num(w.target), ton: num(w.ton), prs: list(w.prs), stuck: list(w.stuck) };
}

/* ---------- рассылка ---------- */
const addDays = (d, n) => new Date(Date.parse(d + 'T00:00:00Z') + n * 864e5).toISOString().slice(0, 10);
const isSunday = d => new Date(d + 'T00:00:00Z').getUTCDay() === 0;
const dm = d => d.slice(8, 10) + '.' + d.slice(5, 7);
function localNow(tz) {
  const d = new Date(Date.now() + tz * 60000);
  return { date: d.toISOString().slice(0, 10), hour: d.getUTCHours() };
}
const esc = s => String(s == null ? '' : s).replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;');
const morningText = p => '💪 <b>Сегодня: ' + esc(p.t) + '</b>\n\n' + p.x.map(x => '• ' + esc(x)).join('\n') + '\n\nВеса и подсказки — в приложении 👇';
const eveningText = p => '⏰ Сегодня по плану <b>' + esc(p.t) + '</b>, а в журнале пока пусто.\nЕсли уже позанимался — отметь подходы. Не успеваешь — перенеси тренировку в приложении.';

const endText = last => '⏳ Напоминания скоро закончатся: расписание есть только до ' + dm(last) + '.\nОткрой приложение, и оно обновится само.';
function weekText(w, monday) {
  const fmt = n => String(n).replace(/\B(?=(\d{3})+$)/g, ' ');
  if (!w || w.from !== monday || !w.done) return '📅 <b>Итог недели</b>\n\nНа этой неделе тренировок не было. Новая неделя — хороший повод начать 💪';
  let t = '📅 <b>Итог недели</b>\n\nТренировок: ' + w.done + (w.target ? ' из ' + w.target : '') + '\nТоннаж: ' + fmt(w.ton) + ' кг';
  if (w.prs.length) t += '\n\n🏆 <b>Рекорды</b>\n' + w.prs.map(x => '• ' + esc(x)).join('\n');
  if (w.stuck.length) t += '\n\n⚠️ <b>Вес стоит на месте</b>\n' + w.stuck.map(x => '• ' + esc(x)).join('\n');
  return t;
}

async function remindUser(env, name, u) {
  const days = Object.keys(u.plan || {}).sort();
  if (!days.length) return; // напоминания выключены
  const now = localNow(u.tz || 0), p = u.plan[now.date], last = days[days.length - 1], msgs = [];
  let changed = false;
  if (p && now.hour === u.m) msgs.push(morningText(p));
  else if (p && now.hour === u.e && (u.done || []).indexOf(now.date) < 0) msgs.push(eveningText(p));
  if (now.hour === u.m && last >= now.date && last <= addDays(now.date, 2) && u.warned !== last) { msgs.push(endText(last)); u.warned = last; changed = true; }
  if (now.hour === u.e && isSunday(now.date) && last >= now.date) {
    const monday = addDays(now.date, -6);
    if (u.wsent !== monday) { msgs.push(weekText(u.week, monday)); u.wsent = monday; changed = true; }
  }
  for (const text of msgs) {
    const r = await send(env, u.id, text);
    if (r.status === 403) { await env.KV.delete(name); return; } // бот заблокирован — убираем из рассылки
  }
  if (changed) await env.KV.put(name, JSON.stringify(u));
}

async function remind(env) {
  let cursor;
  do {
    const page = await env.KV.list({ prefix: 'u:', cursor: cursor });
    for (const k of page.keys) {
      try { // сбой на одном человеке не должен оставить остальных без напоминаний
        const u = JSON.parse((await env.KV.get(k.name)) || 'null');
        if (u) await remindUser(env, k.name, u);
      } catch (err) {
        console.error('Напоминание не отправлено', k.name, String((err && err.message) || err));
      }
    }
    cursor = page.list_complete ? null : page.cursor;
  } while (cursor);
}

async function test(user, env) {
  const u = JSON.parse((await env.KV.get('u:' + user.id)) || 'null');
  const now = localNow(u ? u.tz : 0), p = u && u.plan[now.date];
  const days = u ? Object.keys(u.plan).length : 0;
  const r = await send(env, user.id, p ? morningText(p) : '🔔 Так будут выглядеть напоминания. Сегодня по расписанию отдых' + (days ? ', тренировок впереди: ' + days + '.' : '.'));
  if (r.status !== 200) throw new Error(r.status === 403 ? 'Бот не может тебе написать — нажми «Старт» в чате с ботом' : 'Telegram: ' + r.status);
  return { sent: true };
}

/* ---------- отчёт о тренировке ---------- */
const lines = v => (Array.isArray(v) ? v : []).slice(0, 12).map(s => '• ' + esc(String(s).slice(0, 160))).join('\n');
function rulesText(r) {
  const part = (title, list) => (list ? '\n\n<b>' + title + '</b>\n' + list : '');
  return '📊 <b>' + esc(String(r.t || 'Тренировка').slice(0, 60)) + '</b>' + (r.head ? '\n' + esc(String(r.head).slice(0, 160)) : '') +
    part('✅ Хорошо', lines(r.good)) + part('⚠️ Плохо', lines(r.bad)) + part('👉 В следующий раз', lines(r.next));
}

const AI_SYSTEM = 'Ты тренер в журнале тренировок MY GYM. По данным одной силовой тренировки напиши короткий разбор на русском, обращаясь на «ты». ' +
  'Формат: три блока с заголовками «✅ Хорошо», «⚠️ Плохо», «👉 В следующий раз», в каждом 1–4 коротких пункта, начинающихся с «• ». Всего не больше 120 слов. ' +
  'Обычный текст без разметки и без вступления. Опирайся только на присланные цифры и ничего не выдумывай. Если сказать в блоке нечего, напиши об этом одной фразой. ' +
  'Рекомендации по весу бери из поля next, когда оно есть. Поле good, bad и next — выводы приложения, можно их уточнять, но не противоречить цифрам. Медицинских советов не давай. ' +
  'В данных kg — вес, reps — повторы, last — прошлая тренировка, recent — лучший вес за прошлые разы, начиная с самого свежего.';

async function aiText(facts, rules, env) {
  const data = JSON.stringify({ тренировка: facts, выводы: rules });
  if (data.length > AI_MAX_FACTS) throw new Error('Слишком большая тренировка');
  const client = new Anthropic({ apiKey: env.ANTHROPIC_API_KEY, maxRetries: 0, timeout: AI_TIMEOUT });
  const res = await client.messages.create({
    model: 'claude-haiku-5-5',
    max_tokens: 16000,
    output_config: { effort: 'low' },
    system: AI_SYSTEM,
    messages: [{ role: 'user', content: data }]
  });
  if (res.stop_reason !== 'end_turn') throw new Error('Ответ ИИ не завершён: ' + res.stop_reason);
  const text = res.content.filter(b => b.type === 'text').map(b => b.text).join('').trim();
  if (!text) throw new Error('Пустой ответ ИИ');
  return text;
}

async function report(user, req, env) {
  const r = req.report || {};
  let text = rulesText(r), ai = false, aiError = null;
  // ИИ-отчёт только владельцу: решает сервер, а не приложение
  if (req.ai && String(user.id) === String(env.OWNER_ID)) {
    try {
      if (!env.ANTHROPIC_API_KEY) throw new Error('На сервере не задан ANTHROPIC_API_KEY');
      const body = await aiText(req.facts || {}, { good: r.good, bad: r.bad, next: r.next }, env);
      text = '🤖 <b>' + esc(String(r.t || 'Тренировка').slice(0, 60)) + '</b>' + (r.head ? '\n' + esc(String(r.head).slice(0, 160)) : '') + '\n\n' + esc(body.slice(0, 3400));
      ai = true;
    } catch (err) {
      aiError = String((err && err.message) || err).slice(0, 200); // ИИ не ответил — уйдёт отчёт по правилам
    }
  }
  const s = await send(env, user.id, text);
  if (s.status !== 200) throw new Error(s.status === 403 ? 'Бот не может тебе написать — нажми «Старт» в чате с ботом' : 'Telegram: ' + s.status);
  return { sent: true, ai: ai, aiError: aiError };
}

/* ---------- карточка рекорда ---------- */
async function card(user, req, env) {
  const b64 = req.image;
  if (typeof b64 !== 'string' || !b64 || b64.length > CARD_MAX || !/^[A-Za-z0-9+/]+={0,2}$/.test(b64)) throw new Error('Не получилось подготовить картинку');
  const bytes = Uint8Array.from(atob(b64), c => c.charCodeAt(0));
  if (bytes[0] !== 0xFF || bytes[1] !== 0xD8 || bytes[2] !== 0xFF) throw new Error('Не получилось подготовить картинку'); // принимаем только JPEG
  const form = new FormData();
  form.append('chat_id', String(user.id));
  form.append('caption', '🏆 MY GYM — журнал тренировок в Telegram: @my_gyymm_bot');
  form.append('photo', new File([bytes], 'my-gym-rekord.jpg', { type: 'image/jpeg' }));
  const r = await fetch('https://api.telegram.org/bot' + env.BOT_TOKEN + '/sendPhoto', { method: 'POST', body: form });
  if (r.status !== 200) throw new Error(r.status === 403 ? 'Бот не может тебе написать — нажми «Старт» в чате с ботом' : 'Telegram: ' + r.status);
  return { sent: true };
}

/* ---------- копия данных ---------- */
async function backup(user, req, env) {
  const files = [['my-gym-trenirovki.csv', req.csv, 'text/csv', 'Таблица: открывается в Excel и Google Таблицах'], ['my-gym-kopiya.json', req.json, 'application/json', 'Полная копия: программа, тренировки, вес тела']];
  for (const f of files) if (typeof f[1] !== 'string' || !f[1] || f[1].length > BACKUP_MAX) throw new Error('Не получилось собрать копию');
  for (const f of files) {
    const form = new FormData();
    form.append('chat_id', String(user.id));
    form.append('caption', f[3]);
    form.append('document', new File([f[1]], f[0], { type: f[2] }));
    const r = await fetch('https://api.telegram.org/bot' + env.BOT_TOKEN + '/sendDocument', { method: 'POST', body: form });
    if (r.status !== 200) throw new Error(r.status === 403 ? 'Бот не может тебе написать — нажми «Старт» в чате с ботом' : 'Telegram: ' + r.status);
  }
  return { sent: true };
}

/* ---------- сообщения боту ---------- */
const WELCOME = '👋 <b>Привет! Это MY GYM — журнал тренировок.</b>\n\nСоставь программу, записывай подходы, а я подскажу вес, напомню о тренировке и пришлю отчёт после неё.\n\nЖми кнопку ниже 👇';
async function onUpdate(update, env) {
  const m = update.message;
  if (!m || !m.chat || m.chat.type !== 'private' || !/^\/(start|help)(\s|@|$)/.test(m.text || '')) return;
  const { success } = await env.LIMIT.limit({ key: String(m.chat.id) });
  if (success) await send(env, m.chat.id, WELCOME);
}
// Telegram должен узнать адрес воркера: делаем это один раз из cron, потому что токен есть только здесь
async function ensureWebhook(env) {
  if (!env.WEBHOOK_SECRET || (await env.KV.get('webhook')) === '1') return;
  const r = await fetch('https://api.telegram.org/bot' + env.BOT_TOKEN + '/setWebhook', {
    method: 'POST', headers: { 'Content-Type': 'application/json' },
    body: JSON.stringify({ url: WORKER_URL, secret_token: env.WEBHOOK_SECRET, allowed_updates: ['message'], drop_pending_updates: true })
  });
  if (r.status === 200) await env.KV.put('webhook', '1');
  else console.error('setWebhook', r.status);
}

function send(env, chat, text) {
  return fetch('https://api.telegram.org/bot' + env.BOT_TOKEN + '/sendMessage', {
    method: 'POST', headers: { 'Content-Type': 'application/json' },
    body: JSON.stringify({ chat_id: chat, text: text, parse_mode: 'HTML', disable_web_page_preview: true,
      reply_markup: { inline_keyboard: [[{ text: '🏋️ Открыть тренировку', web_app: { url: APP_URL } }]] } })
  });
}
