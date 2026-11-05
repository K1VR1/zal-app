import { createHmac } from 'node:crypto';
// Проверки воркера с подставными Telegram и Claude: настоящие сервисы не вызываются.
const worker = (await import('../worker.js')).default;
const TOKEN='123:test', OWNER='947241771';
function initData(id){const p=new URLSearchParams({auth_date:String(Math.floor(Date.now()/1000)),user:JSON.stringify({id})});
 const check=[...p.keys()].sort().map(k=>k+'='+p.get(k)).join('\n');const secret=createHmac('sha256','WebAppData').update(TOKEN).digest();
 p.set('hash',createHmac('sha256',secret).update(check).digest('hex'));return p.toString();}
let calls; let aiMode='ok';
globalThis.fetch=async(u,o={})=>{u=String(u.url||u);
 if(u.includes('api.telegram.org')){const b=typeof o.body==='string'?JSON.parse(o.body):Object.fromEntries(o.body);calls.tg.push(b);return new Response('{}',{status:calls.tgStatus||200});}
 if(u.includes('api.anthropic.com')){const b=JSON.parse(o.body);const h=new Headers(o.headers);calls.ai.push({b,beta:h.get('anthropic-beta'),key:h.get('x-api-key')});
  if(aiMode==='err')return new Response(JSON.stringify({type:'error',error:{type:'overloaded_error',message:'busy'}}),{status:529,headers:{'content-type':'application/json'}});
  return new Response(JSON.stringify({id:'m',type:'message',role:'assistant',model:b.model,stop_reason:aiMode==='refusal'?'refusal':'end_turn',usage:{input_tokens:1,output_tokens:1},
   content:aiMode==='refusal'?[]:[{type:'text',text:'✅ Хорошо\n• <b>рекорд</b> & рост'}]}),{status:200,headers:{'content-type':'application/json'}});}
 throw new Error('unexpected '+u)};
const env=k=>({BOT_TOKEN:TOKEN,OWNER_ID:OWNER,KV:{},LIMIT:{limit:async()=>({success:true})},...(k?{ANTHROPIC_API_KEY:'sk-test'}:{})});
const rep={t:'Спина <1> · 06.11',head:'Упражнения: 5 из 7',good:['Жим: рекорд'],bad:[],next:['Жим: 40 кг ↑']};
async function run(name,id,body,e,want){calls={tg:[],ai:[],tgStatus:body._s};delete body._s;const waits=[];
 const r=await worker.fetch(new Request('https://x/',{method:'POST',body:JSON.stringify({action:'report',initData:initData(id),...body})}),e,{waitUntil:p=>waits.push(p)});
 const d=await r.json();await Promise.all(waits);const got={ok:d.ok,ai:d.data&&d.data.ai,err:(d.data&&d.data.aiError||d.error||'').slice(0,50),aiCalls:calls.ai.length,tg:calls.tg.length};
 const pass=Object.entries(want).every(([k,v])=>got[k]===v);console.log(pass?'PASS':'FAIL',name,JSON.stringify(got));if(!pass)process.exitCode=1;return calls;}
await run('обычный пользователь без ИИ',111,{report:rep},env(true),{ok:true,ai:false,aiCalls:0,tg:1});
const c1=await run('чужой просит ИИ — отказ молча, ИИ не вызван',111,{ai:true,report:rep,facts:{a:1}},env(true),{ok:true,ai:false,aiCalls:0,tg:1,err:''});
console.log('  text:',JSON.stringify(c1.tg[0].text));
await run('владелец без галочки',OWNER,{report:rep},env(true),{ok:true,ai:false,aiCalls:0,tg:1});
const c2=await run('владелец с ИИ',OWNER,{ai:true,report:rep,facts:{a:1}},env(true),{ok:true,ai:true,aiCalls:1,tg:1});
console.log('  req:',c2.ai[0].b.model,c2.ai[0].b.fallbacks,c2.ai[0].beta,JSON.stringify(c2.ai[0].b.output_config));console.log('  text:',JSON.stringify(c2.tg[0].text));
await run('владелец, нет ключа — правила',OWNER,{ai:true,report:rep,facts:{}},env(false),{ok:true,ai:false,aiCalls:0,tg:1,err:'На сервере не задан ANTHROPIC_API_KEY'});
aiMode='err';await run('ИИ упал — правила',OWNER,{ai:true,report:rep,facts:{}},env(true),{ok:true,ai:false,aiCalls:1,tg:1});
aiMode='refusal';await run('ИИ отказал — правила',OWNER,{ai:true,report:rep,facts:{}},env(true),{ok:true,ai:false,aiCalls:1,tg:1});
aiMode='ok';await run('слишком много данных — правила',OWNER,{ai:true,report:rep,facts:{x:'я'.repeat(9000)}},env(true),{ok:true,ai:false,aiCalls:0,tg:1});
await run('бот заблокирован',111,{report:rep,_s:403},env(true),{ok:false,tg:1});
const bad=await worker.fetch(new Request('https://x/',{method:'POST',body:JSON.stringify({action:'report',ai:true,initData:'auth_date=1&hash=00',report:rep})}),env(true),{waitUntil(){}});
console.log('поддельная подпись:',JSON.stringify(await bad.json()));
// --- новые проверки ---
calls={tg:[],ai:[]};
const ok=(n,c,x='')=>{console.log(c?'PASS':'FAIL',n,x);if(!c)process.exitCode=1};
const memKV=(init={})=>{const m=new Map(Object.entries(init));return {m,puts:0,get:async k=>m.get(k)??null,put(k,v){this.puts++;m.set(k,v);return Promise.resolve()},delete:async k=>{m.delete(k)},list:async({prefix})=>({keys:[...m.keys()].filter(k=>k.startsWith(prefix)).map(name=>({name})),list_complete:true})}};
const call=async(id,body,e,init)=>{const r=await worker.fetch(new Request('https://x/',{method:'POST',body:JSON.stringify({initData:init||initData(id),...body})}),e,{waitUntil(){}});return r.json()};
const kv=memKV();const E={...env(false),KV:kv};
const day=new Date().toISOString().slice(0,10);
let r=await call(5,{action:'sync',on:true,tz:180,m:9,e:20,plan:{[day]:{t:'A',x:['B']}},done:[]},E);ok('sync включает',r.ok&&r.data.on===true);
let u=JSON.parse(kv.m.get('u:5'));ok('план и счётчик',u.plan[day].t==='A'&&u.wn===1&&u.wd===day);
r=await call(5,{action:'sync',on:false},E);u=JSON.parse(kv.m.get('u:5'));ok('выключение: план пуст, счётчик цел',r.data.on===false&&Object.keys(u.plan).length===0&&u.wn===2);
for(let i=0;i<38;i++)await call(5,{action:'sync',on:true,plan:{}},E);
const puts=kv.puts;r=await call(5,{action:'sync',on:true,plan:{}},E);ok('41-я запись отклонена и не пишет в KV',!r.ok&&kv.puts===puts,r.error);
const y=new Date(Date.now()-864e5).toISOString().slice(0,10);kv.m.set('u:5',JSON.stringify({...JSON.parse(kv.m.get('u:5')),wd:y,wn:99}));r=await call(5,{action:'sync',on:true,plan:{}},E);ok('на новый день счётчик сбрасывается',r.ok&&JSON.parse(kv.m.get('u:5')).wn===1);
r=await call(5,{action:'test'},{...E});ok('пробное после пустого плана работает',r.ok,r.error||'');
r=await call(6,{action:'ping'},{...E,LIMIT:{limit:async()=>({success:false})}});ok('лимит запросов',!r.ok&&r.error.includes('подожди'),r.error);
function signed(f){const p=new URLSearchParams({user:JSON.stringify({id:7}),...f});const c=[...p.keys()].sort().map(k=>k+'='+p.get(k)).join('\n');const s=createHmac('sha256','WebAppData').update(TOKEN).digest();p.set('hash',createHmac('sha256',s).update(c).digest('hex'));return p.toString()}
const now=Math.floor(Date.now()/1000);
r=await call(0,{action:'ping'},E,signed({auth_date:now-23*3600}));ok('подпись 23 ч проходит',r.ok,r.error||'');
r=await call(0,{action:'ping'},E,signed({auth_date:now-25*3600}));ok('подпись 25 ч устарела',!r.ok&&r.error.includes('устарела'),r.error);
const good=signed({auth_date:now});const q=new URLSearchParams(good);q.set('hash',q.get('hash').replace(/.$/,c=>c==='0'?'1':'0'));
r=await call(0,{action:'ping'},E,q.toString());ok('изменённый символ подписи',!r.ok);
r=await call(0,{action:'ping'},E,'auth_date=1&user=%7B%22id%22%3A7%7D');ok('нет подписи',!r.ok);
r=await call(0,{action:'ping'},E,good.replace('hash=','hash=zz'));ok('мусор в подписи',!r.ok);
const far=new Date(Date.now()+7*864e5).toISOString().slice(0,10);
const today={tz:0,m:new Date().getUTCHours(),e:(new Date().getUTCHours()+5)%24,plan:{[day]:{t:'T',x:['y']},[far]:{t:'T',x:['y']}},done:[]};
const kv2=memKV({'u:1':JSON.stringify({id:1,...today}),'u:2':'{broken','u:3':JSON.stringify({id:3,...today}),'u:4':JSON.stringify({id:4,...today})});
const sent=[];const realFetch=globalThis.fetch;globalThis.fetch=async(u,o)=>{const b=JSON.parse(o.body);sent.push(b.chat_id);if(b.chat_id===3)throw new Error('network');if(b.chat_id===4)return new Response('{}',{status:403});return new Response('{}')};
const ce=console.error;let logged=0;console.error=()=>{logged++};let w;await worker.scheduled({}, {...E,KV:kv2},{waitUntil:p=>w=p});await w;console.error=ce;globalThis.fetch=realFetch;
ok('сбой одного не мешает остальным',sent.join()==='1,3,4'&&logged===2,sent.join()+' ошибок '+logged);ok('заблокировавший бота удалён',!kv2.m.has('u:4')&&kv2.m.has('u:1'));

// --- копия данных ---
calls={tg:[],ai:[]};
r=await call(8,{action:'backup',csv:'a;b',json:'{}'},E);ok('копия: два файла тому, кто просил',r.ok&&calls.tg.length===2&&calls.tg.every(b=>b.chat_id==='8'&&b.document.name),String(calls.tg.map(b=>b.document&&b.document.name)));
r=await call(8,{action:'backup',csv:'',json:'{}'},E);ok('копия: пустой файл отклонён',!r.ok);
r=await call(8,{action:'backup',csv:'x'.repeat(2000001),json:'{}'},E);ok('копия: слишком большой файл отклонён',!r.ok);
ok('копия не сохраняется на сервере',!kv.m.has('u:8'));
// --- сообщения боту ---
const hook=(hdr,update,e)=>worker.fetch(new Request('https://x/',{method:'POST',headers:hdr?{'X-Telegram-Bot-Api-Secret-Token':hdr}:{},body:JSON.stringify(update)}),e,{waitUntil(){}});
const HE={...E,WEBHOOK_SECRET:'s3cret'};const msg=(text,type='private')=>({message:{text,chat:{id:42,type},from:{id:42}}});
calls={tg:[],ai:[]};
let h=await hook('s3cret',msg('/start'),HE);ok('/start: приветствие с кнопкой',h.status===200&&calls.tg.length===1&&calls.tg[0].chat_id===42&&!!calls.tg[0].reply_markup,calls.tg[0]&&calls.tg[0].text.slice(0,20));
calls={tg:[],ai:[]};h=await hook('s3cret',msg('/start abc'),HE);ok('/start с параметром',calls.tg.length===1);
calls={tg:[],ai:[]};h=await hook('s3cret',msg('привет'),HE);ok('обычный текст: молчит, но 200',h.status===200&&calls.tg.length===0);
calls={tg:[],ai:[]};h=await hook('s3cret',msg('/start','group'),HE);ok('в группе не отвечает',calls.tg.length===0);
calls={tg:[],ai:[]};h=await hook('wrong',msg('/start'),HE);ok('чужой секрет: 403, без сообщения',h.status===403&&calls.tg.length===0);
calls={tg:[],ai:[]};h=await hook('x',msg('/start'),E);ok('секрет не задан: 403',h.status===403&&calls.tg.length===0);
calls={tg:[],ai:[]};h=await hook(null,msg('/start'),HE);ok('без заголовка: считается запросом приложения',calls.tg.length===0&&(await h.json()).ok===false);
// --- webhook ставится один раз ---
const kv3=memKV();calls={tg:[],ai:[]};let w3;await worker.scheduled({}, {...HE,KV:kv3},{waitUntil:p=>w3=p});await w3;
ok('webhook: адрес и секрет',calls.tg.length===1&&calls.tg[0].secret_token==='s3cret'&&calls.tg[0].url.startsWith('https://mygym.')&&kv3.m.get('webhook')==='1');
calls={tg:[],ai:[]};await worker.scheduled({}, {...HE,KV:kv3},{waitUntil:p=>w3=p});await w3;ok('webhook: второй раз не вызывается',calls.tg.length===0);
// --- итог недели и конец расписания: подставляем время (воскресенье 8 ноября 2026, 20:00 UTC) ---
const realNow=Date.now;const at=(iso)=>{Date.now=()=>Date.parse(iso)};
const runCron=async(k)=>{calls={tg:[],ai:[]};let p;await worker.scheduled({}, {...E,KV:k},{waitUntil:x=>p=x});await p;return calls.tg.map(b=>b.text)};
const mk=(o)=>memKV({'u:9':JSON.stringify({id:9,tz:0,m:9,e:20,done:[],plan:{'2026-11-20':{t:'T',x:[]}},...o})});
at('2026-11-08T20:00:00Z');
let k=mk({week:{from:'2026-11-02',done:3,target:3,ton:9660,prs:['Жим <b>: 50 кг'],stuck:['Тяга']}});let t=await runCron(k);
ok('итог недели в воскресенье',t.length===1&&t[0].includes('Итог недели')&&t[0].includes('3 из 3')&&t[0].includes('9 660')&&t[0].includes('Жим &lt;b&gt;')&&t[0].includes('Тяга'),t[0]&&t[0].slice(0,30));
t=await runCron(k);ok('итог недели не повторяется',t.length===0);
t=await runCron(mk({week:{from:'2026-10-26',done:3,target:3,ton:1,prs:[],stuck:[]}}));ok('старая неделя считается пустой',t.length===1&&t[0].includes('не было'));
at('2026-11-07T20:00:00Z');t=await runCron(mk({week:null}));ok('в субботу итога нет',t.length===0);
at('2026-11-08T19:00:00Z');t=await runCron(mk({week:null}));ok('в другой час итога нет',t.length===0);
at('2026-11-18T09:00:00Z');k=mk({});t=await runCron(k);ok('за 2 дня предупреждает о конце расписания',t.length===1&&t[0].includes('20.11'),t[0]&&t[0].slice(0,30));
t=await runCron(k);ok('предупреждение один раз',t.length===0);
at('2026-11-17T09:00:00Z');t=await runCron(mk({}));ok('за 3 дня ещё рано',t.length===0);
at('2026-11-21T09:00:00Z');t=await runCron(mk({}));ok('расписание кончилось: тишина',t.length===0);
at('2026-11-18T09:00:00Z');t=await runCron(mk({plan:{}}));ok('выключенным ничего',t.length===0);
at('2026-11-19T09:00:00Z');k=mk({plan:{'2026-11-19':{t:'Спина',x:['Тяга']},'2026-11-30':{t:'T',x:[]}}});t=await runCron(k);ok('обычное утреннее сообщение на месте',t.length===1&&t[0].includes('Спина'));
Date.now=realNow;
// --- обновление расписания не сбрасывает отметки ---
const kv4=memKV({'u:5':JSON.stringify({id:5,warned:'2026-11-20',wsent:'2026-11-02'})});r=await call(5,{action:'sync',on:true,plan:{},week:{from:'2026-11-02',done:'2',target:3,ton:100.6,prs:['a'.repeat(99)],stuck:'x'}},{...E,KV:kv4});u=JSON.parse(kv4.m.get('u:5'));
ok('sync хранит отметки и чистит неделю',r.ok&&u.warned==='2026-11-20'&&u.wsent==='2026-11-02'&&u.week.done===2&&u.week.ton===101&&u.week.prs[0].length===60&&u.week.stuck.length===0);
// --- карточка рекорда ---
const jpeg=Buffer.from([0xFF,0xD8,0xFF,0xE0,1,2,3]).toString('base64');
calls={tg:[],ai:[]};r=await call(8,{action:'card',image:jpeg},E);ok('карточка: картинка уходит тому, кто просил',r.ok&&calls.tg.length===1&&calls.tg[0].chat_id==='8'&&calls.tg[0].photo.size===7);
calls={tg:[],ai:[]};r=await call(8,{action:'card',image:Buffer.from('<svg>').toString('base64')},E);ok('карточка: не JPEG отклонён',!r.ok&&calls.tg.length===0);
calls={tg:[],ai:[]};r=await call(8,{action:'card',image:'a'.repeat(4000001)},E);ok('карточка: слишком большая отклонена',!r.ok&&calls.tg.length===0);
calls={tg:[],ai:[]};r=await call(8,{action:'card',image:'не base64!'},E);ok('карточка: мусор отклонён',!r.ok&&calls.tg.length===0);
calls={tg:[],ai:[]};r=await call(8,{action:'card'},E);ok('карточка: без картинки отклонена',!r.ok&&calls.tg.length===0);
if(!process.exitCode)console.log('Все проверки прошли');
