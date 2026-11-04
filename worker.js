/**
 * Сервер бота MY GYM на Cloudflare Workers.
 *
 * Что делает:
 *   POST /  { action: 'sync' }   — приложение присылает расписание пользователя на 2 недели вперёд;
 *   POST /  { action: 'test' }   — прислать пользователю пробное напоминание;
 *   POST /  { action: 'export' } — один раз отдать владельцу его старые тренировки (константа OWNER_EXPORT ниже);
 *   по расписанию (cron «0 * * * *», раз в час) — рассылает напоминания каждому по его часовому поясу.
 *
 * Настройки воркера (Settings):
 *   Secret   BOT_TOKEN — токен бота от @BotFather
 *   Text     OWNER_ID  — Telegram id владельца
 *   KV       KV        — хранилище расписаний (binding с именем KV)
 *   Cron     0 * * * *
 */

// Старые тренировки владельца для одноразового переноса из Google Таблицы.
// В общем репозитории здесь пусто; рабочая копия с данными — только в твоём аккаунте Cloudflare.
const OWNER_EXPORT = null;

const APP_URL = 'https://k1vr1.github.io/zal-app/';
const MAX_AGE = 7 * 24 * 3600; // сколько секунд живёт подпись Telegram

const CORS = { 'Access-Control-Allow-Origin': '*', 'Access-Control-Allow-Methods': 'GET, POST, OPTIONS', 'Access-Control-Allow-Headers': 'Content-Type' };
const json = obj => new Response(JSON.stringify(obj), { headers: Object.assign({ 'Content-Type': 'application/json; charset=utf-8' }, CORS) });

export default {
  async fetch(req, env) {
    if (req.method === 'OPTIONS') return new Response(null, { headers: CORS });
    if (req.method !== 'POST') return json({ ok: true, app: 'mygym', v: 4 });
    try {
      const body = JSON.parse(await req.text() || '{}');
      const user = await auth(body.initData, env);
      if (body.action === 'sync') return json({ ok: true, data: await sync(user, body, env) });
      if (body.action === 'test') return json({ ok: true, data: await test(user, env) });
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
    ctx.waitUntil(remind(env));
  }
};

/* ---------- проверка, что запрос пришёл из Telegram ---------- */
const enc = new TextEncoder();
async function hmac(key, msg) {
  const k = await crypto.subtle.importKey('raw', typeof key === 'string' ? enc.encode(key) : key, { name: 'HMAC', hash: 'SHA-256' }, false, ['sign']);
  return new Uint8Array(await crypto.subtle.sign('HMAC', k, enc.encode(msg)));
}
const hex = b => Array.from(b).map(x => x.toString(16).padStart(2, '0')).join('');
async function auth(initData, env) {
  if (!env.BOT_TOKEN) throw new Error('На сервере не задан BOT_TOKEN');
  if (!initData) throw new Error('Открой приложение через Telegram');
  const params = new URLSearchParams(initData);
  const hash = params.get('hash');
  params.delete('hash');
  const check = Array.from(params.keys()).sort().map(k => k + '=' + params.get(k)).join('\n');
  const secret = await hmac('WebAppData', env.BOT_TOKEN);
  if (hex(await hmac(secret, check)) !== hash) throw new Error('Подпись Telegram не прошла проверку');
  if (Date.now() / 1000 - Number(params.get('auth_date') || 0) > MAX_AGE) throw new Error('Сессия устарела, перезапусти приложение');
  const user = JSON.parse(params.get('user') || '{}');
  if (!user.id) throw new Error('Нет пользователя');
  return user;
}

/* ---------- расписание пользователя ---------- */
const isDate = d => /^\d{4}-\d{2}-\d{2}$/.test(d);
async function sync(user, req, env) {
  const key = 'u:' + user.id;
  if (!req.on) { await env.KV.delete(key); return { on: false }; }
  const hour = (h, def) => Math.max(0, Math.min(23, Math.round(Number(h == null ? def : h) || 0)));
  const plan = {};
  Object.keys(req.plan || {}).filter(isDate).slice(0, 21).forEach(d => {
    const p = req.plan[d] || {};
    plan[d] = { t: String(p.t || 'Тренировка').slice(0, 40), x: (p.x || []).slice(0, 12).map(s => String(s).slice(0, 40)) };
  });
  const rec = {
    id: user.id, tz: Math.max(-720, Math.min(840, Math.round(Number(req.tz) || 0))),
    m: hour(req.m, 9), e: hour(req.e, 20), plan: plan, done: (req.done || []).filter(isDate).slice(-5)
  };
  await env.KV.put(key, JSON.stringify(rec));
  return { on: true };
}

/* ---------- рассылка ---------- */
function localNow(tz) {
  const d = new Date(Date.now() + tz * 60000);
  return { date: d.toISOString().slice(0, 10), hour: d.getUTCHours() };
}
const esc = s => String(s == null ? '' : s).replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;');
const morningText = p => '💪 <b>Сегодня: ' + esc(p.t) + '</b>\n\n' + p.x.map(x => '• ' + esc(x)).join('\n') + '\n\nВеса и подсказки — в приложении 👇';
const eveningText = p => '⏰ Сегодня по плану <b>' + esc(p.t) + '</b>, а в журнале пока пусто.\nЕсли уже позанимался — отметь подходы. Не успеваешь — перенеси тренировку в приложении.';

async function remind(env) {
  let cursor;
  do {
    const page = await env.KV.list({ prefix: 'u:', cursor: cursor });
    for (const k of page.keys) {
      const u = JSON.parse((await env.KV.get(k.name)) || 'null');
      if (!u) continue;
      const now = localNow(u.tz || 0), p = u.plan && u.plan[now.date];
      if (!p) continue;
      let text = null;
      if (now.hour === u.m) text = morningText(p);
      else if (now.hour === u.e && (u.done || []).indexOf(now.date) < 0) text = eveningText(p);
      if (!text) continue;
      const r = await send(env, u.id, text);
      if (r.status === 403) await env.KV.delete(k.name); // бот заблокирован — убираем из рассылки
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

function send(env, chat, text) {
  return fetch('https://api.telegram.org/bot' + env.BOT_TOKEN + '/sendMessage', {
    method: 'POST', headers: { 'Content-Type': 'application/json' },
    body: JSON.stringify({ chat_id: chat, text: text, parse_mode: 'HTML', disable_web_page_preview: true,
      reply_markup: { inline_keyboard: [[{ text: '🏋️ Открыть тренировку', web_app: { url: APP_URL } }]] } })
  });
}
