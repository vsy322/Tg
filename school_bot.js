// Школьный Telegram-бот на JavaScript (Node.js + Telegraf).
// Запуск из терминала VS Code:  npm install  ->  npm start

const { Telegraf, Markup } = require('telegraf');
const express = require('express');
const axios = require('axios');
const pptxgen = require('pptxgenjs');
const path = require('path');
const os = require('os');
const fs = require('fs');
const { execSync } = require('child_process');
const { HttpsProxyAgent } = require('https-proxy-agent');
const { SocksProxyAgent } = require('socks-proxy-agent');

// ВАЖНО: токен зашит в файл. Не публикуй этот код в интернете.
// После проверки отзови токен у @BotFather (/revoke) и создай новый.
// Локально токен вписан; на хостинге задай переменную BOT_TOKEN — она имеет приоритет.
const TOKEN = process.env.BOT_TOKEN;
if (!TOKEN) throw new Error('BOT_TOKEN is required');

// --- Поддержка прокси (если в браузере Telegram открывается, а здесь нет) ---
function detectProxy() {
  const env = process.env.HTTPS_PROXY || process.env.https_proxy ||
              process.env.HTTP_PROXY || process.env.http_proxy;
  if (env) return env;
  if (process.platform === 'win32') {
    try {
      const out = execSync(
        'reg query "HKCU\\Software\\Microsoft\\Windows\\CurrentVersion\\Internet Settings" /v ProxyServer',
        { encoding: 'utf8', stdio: ['ignore', 'pipe', 'ignore'] }
      );
      const m = out.match(/ProxyServer\s+REG_SZ\s+(\S+)/i);
      if (m && m[1] && !/^\s*$/.test(m[1])) return m[1];
    } catch (e) {}
  }
  return '';
}

// Happ использует локальный SOCKS-прокси 127.0.0.1:10808 (см. config.json).
// Если у тебя другой прокси — задай его переменной BOT_PROXY в run_bot.bat.
// На Windows используем локальный прокси Happ. На Linux-хостинге прокси не нужен.
const HAPP_PROXY = process.platform === 'win32' ? 'socks://127.0.0.1:10808' : '';
const PROXY = process.env.BOT_PROXY || HAPP_PROXY || detectProxy();
let agent;
if (PROXY) {
  agent = /^socks/i.test(PROXY) ? new SocksProxyAgent(PROXY) : new HttpsProxyAgent(PROXY);
  console.log('Использую прокси: ' + PROXY);
} else {
  console.log('Прокси не найден — Trying direct connection.');
}
const netOpts = agent ? { agent, httpsAgent: agent } : {};

const bot = new Telegraf(TOKEN, agent ? { telegram: { agent } } : {});
const WIKI = 'https://ru.wikipedia.org/w/api.php';
const WD_API = 'https://www.wikidata.org/w/api.php';
const DDG = 'https://api.duckduckgo.com/';
const SPARQL = 'https://query.wikidata.org/sparql';

// Википедия отклоняет запросы без нормального User-Agent — из-за этого был 403.
const UA = 'SchoolBot/1.0 (school project; t.me/schol_spg_bot)';
axios.defaults.headers.common['User-Agent'] = UA;
axios.defaults.headers.common['Accept'] = 'application/json';

// --- Защита от лимитов Википедии/Викиданных (ошибка 429) ---
// Render использует общий IP, поэтому запросы идут реже, повторяются при 429
// и кэшируются, чтобы повторный вопрос про того же человека не тратил лимит.
const apiCache = new Map();
let lastApiCall = 0;
let apiBusy = false;

async function apiGet(url, params) {
  const key = url + '|' + JSON.stringify(params || {});
  if (apiCache.has(key)) return apiCache.get(key);

  // Несколько вызовов подряд (поиск в Википедии + Викиданных) идут строго по очереди.
  while (apiBusy) await new Promise(r => setTimeout(r, 120));
  apiBusy = true;
  try {
    const gap = 400 - (Date.now() - lastApiCall);
    if (gap > 0) await new Promise(r => setTimeout(r, gap));

    let lastErr;
    for (let attempt = 0; attempt < 5; attempt++) {
      lastApiCall = Date.now();
      try {
        const { data } = await axios.get(url, { params, timeout: 30000, ...netOpts });
        apiCache.set(key, data);
        return data;
      } catch (e) {
        lastErr = e;
        const status = e.response?.status;
        const retryable = status === 429 || status === 500 || status === 502 ||
                          status === 503 || status === 504 || !status;
        if (!retryable || attempt === 4) break;
        const header = Number(e.response?.headers?.['retry-after']);
        const wait = Number.isFinite(header) && header > 0
          ? header * 1000
          : Math.min(2000 * Math.pow(2, attempt), 15000);
        await new Promise(r => setTimeout(r, wait));
      }
    }
    throw lastErr;
  } finally {
    apiBusy = false;
  }
}

const TEXT_BTN = '📝 Готовый текст';
const PRES_BTN = '📊 Презентация';

const mainMenu = Markup.keyboard([[TEXT_BTN, PRES_BTN]]).resize();
const removeKb = Markup.removeKeyboard();

const YES_WORDS = ['да', 'давай', 'давай же', 'ага', 'угу', 'конечно', 'ок', 'окей',
  'ладно', 'хочу', 'сделай', 'презентацию', 'презентация', 'yes', 'yeah', 'добре'];
const NO_WORDS = ['нет', 'не надо', 'не хочу', 'не сейчас', 'позже', 'стоп', 'отмена', 'no'];

const yesNoKeyboard = () => Markup.inlineKeyboard([
  [Markup.button.callback('Да, сделать презентацию', 'pres:yes'),
   Markup.button.callback('Нет, спасибо', 'pres:no')],
]);

const STYLES = { classic: 'Классический', modern: 'Современный', kids: 'Детский' };
const COLORS = {
  blue: ['Синий', '1F4E8C'],
  green: ['Зелёный', '226E4E'],
  red: ['Красный', '962D37'],
  violet: ['Фиолетовый', '5E378C'],
};

const styleKeyboard = () => Markup.inlineKeyboard(
  Object.entries(STYLES).map(([k, label]) => Markup.button.callback(label, `style:${k}`))
);
const colorKeyboard = () => Markup.inlineKeyboard(
  Object.entries(COLORS).map(([k, v]) => Markup.button.callback(v[0], `color:${k}`))
);

// Простейшие сессии в памяти
const sessions = new Map();
function session(ctx) {
  const id = ctx.chat.id;
  if (!sessions.has(id)) sessions.set(id, { step: 'menu' });
  return sessions.get(id);
}

// ---------- Открытые данные ----------

// --- Понимание разговорных запросов ("напиши про энштейна" и т.п.) ---
const STOP_WORDS = new Set([
  'напиши','написать','напишите','про','об','about','сделай','сделать','сделайте',
  'давай','дай','пришли','инфа','инфу','информация','информацию',
  'текст','доклад','доклада','презентация','презентацию','презентации',
  'биография','биографию','биографии','расскажи','рассказать','кто','такой','такая',
  'что','какой','какая','нужно','надо','помоги','найди','ищи',
  'поищи','сведения','данные','материал','материалы','тему','тема','для','класс',
  'класса','школе','школа','школь','человек','личность','жизнь','жизни'
]);

// Частые сокращения и опечатки -> полное имя.
const ALIASES = {
  'энштейн': 'альберт эйнштейн', 'энштейна': 'альберт эйнштейн', 'айнштейн': 'альберт эйнштейн',
  'пушкин': 'александр пушкин', 'пушкина': 'александр пушкин',
  'лермонтов': 'михаил лермонтов', 'толстой': 'лев толстой', 'чехов': 'антон чехов',
  'достоевский': 'фёдор достоевский', 'гоголь': 'николай гоголь',
  'черчилль': 'уинстон черчилль', 'наполеон': 'наполеон бонапарт',
  'сталин': 'иосиф сталин', 'ленин': 'владимир ленин', 'гитлер': 'адольф гитлер',
  'петр 1': 'пётр великий', 'петр великий': 'пётр великий', 'пётр 1': 'пётр великий',
  'екатерина 2': 'екатерина великая', 'екатерина великая': 'екатерина великая',
  'гагарин': 'юрий гагарин', 'королев': 'сергей королёв', 'менделеев': 'дмитрий менделеев',
  'ломоносов': 'михаил ломоносов', 'суворов': 'александр суворов', 'жуков': 'георгий жуков'
};

function cleanQuery(raw) {
  let s = String(raw || '').toLowerCase().replace(/ё/g, 'е')
    .replace(/[^a-zа-я0-9\s-]/gi, ' ').replace(/\s+/g, ' ').trim();
  let words = s.split(' ').filter(w => w.length > 1 && !STOP_WORDS.has(w));
  s = words.join(' ');
  return ALIASES[s] || s || String(raw || '').trim();
}

function lev(a, b) {
  a = String(a || ''); b = String(b || '');
  const prev = Array.from({ length: b.length + 1 }, (_, i) => i);
  for (let i = 1; i <= a.length; i++) {
    let last = prev[0]; prev[0] = i;
    for (let j = 1; j <= b.length; j++) {
      const old = prev[j];
      prev[j] = Math.min(prev[j] + 1, prev[j - 1] + 1, last + (a[i - 1] === b[j - 1] ? 0 : 1));
      last = old;
    }
  }
  return prev[b.length];
}

function sim(a, b) {
  a = String(a || '').toLowerCase(); b = String(b || '').toLowerCase();
  if (!a || !b) return 0;
  return 1 - lev(a, b) / Math.max(a.length, b.length);
}

function scoreCandidate(q, title, desc) {
  const t = String(title || '').toLowerCase().replace(/ё/g, 'е');
  const words = q.split(' ').filter(Boolean);
  let best = 0;
  for (const w of words) {
    if (t.includes(w)) best = Math.max(best, 0.92);
    for (const tw of t.split(' ')) best = Math.max(best, sim(w, tw) * 0.9);
    best = Math.max(best, sim(w, t) * 0.85);
  }
  best = Math.max(best, sim(q, t));
  if (desc && /человек|личность|персон/i.test(desc)) best += 0.04;
  return best;
}

async function wikiSearchRaw(name, limit = 8) {
  const data = await apiGet(WIKI, {
    action: 'query', list: 'search', srsearch: name,
    format: 'json', origin: '*', srlimit: String(limit),
  });
  return data.query?.search || [];
}

async function wdSearch(name) {
  const data = await apiGet(WD_API, {
    action: 'wbsearchentities', search: name, language: 'ru',
    format: 'json', origin: '*', limit: '10',
  });
  return data.search || [];
}

async function ddgEntity(name) {
  const data = await apiGet(DDG, {
    q: name, format: 'json', no_html: 1, no_redirect: 1,
  });
  return data.Heading ? { title: data.Heading, desc: data.Abstract || '' } : null;
}

async function findPerson(raw) {
  let q = cleanQuery(raw);
  const queries = [q];
  if (q !== String(raw).trim()) queries.push(String(raw).trim());

  let candidates = [];
  for (const query of queries) {
    const [wiki, wd] = await Promise.allSettled([wikiSearchRaw(query, 8), wdSearch(query)]);
    if (wiki.status === 'fulfilled') {
      for (const h of wiki.value) {
        candidates.push({
          title: h.title, pageid: h.pageid,
          desc: String(h.snippet || '').replace(/<[^>]+>/g, ''),
          score: scoreCandidate(q, h.title, h.snippet || ''), source: 'wikipedia',
        });
      }
    }
    if (wd.status === 'fulfilled') {
      for (const e of wd.value) {
        candidates.push({
          title: e.label, qid: e.id, desc: e.description || '',
          score: scoreCandidate(q, e.label + ' ' + (e.aliases || []).join(' '), e.description || ''),
          source: 'wikidata',
        });
      }
    }
    if (candidates.length) break;

    // Если Википедия/Викиданные ничего не нашли — пробуем «яндексоподобный» ответ.
    try {
      const dd = await ddgEntity(query);
      if (dd) {
        const more = await wikiSearchRaw(dd.title, 3);
        for (const h of more) {
          candidates.push({
            title: h.title, pageid: h.pageid,
            desc: dd.desc || String(h.snippet || '').replace(/<[^>]+>/g, ''),
            score: scoreCandidate(q, h.title, dd.desc || ''), source: 'duckduckgo',
          });
        }
      }
    } catch (e) {}
    if (candidates.length) break;
  }

  const seen = new Set();
  candidates = candidates.filter(c => {
    const k = String(c.title || '').toLowerCase().replace(/ё/g, 'е');
    if (!k || seen.has(k)) return false;
    seen.add(k); return true;
  }).sort((a, b) => b.score - a.score);

  return { query: q, best: candidates[0] || null, all: candidates.slice(0, 5) };
}

async function wikiArticle(pageid) {
  const data = await apiGet(WIKI, {
    action: 'query', prop: 'extracts|pageprops', pageids: String(pageid),
    explaintext: '1', format: 'json', origin: '*',
  });
  const page = data.query?.pages?.[String(pageid)] || {};
  return {
    title: page.title || '',
    text: page.extract || '',
    qid: page.pageprops?.wikibase_item || '',
  };
}

async function sparql(q) {
  const data = await apiGet(SPARQL, { query: q, format: 'json' });
  return data.results?.bindings || [];
}

const val = (row, key) => row?.[key]?.value || '';

async function collect(raw) {
  const found = await findPerson(raw);
  if (!found.best) return null;

  let best = found.best;
  let pageid = best.pageid;

  if (!pageid && best.qid) {
    try {
      const data = await apiGet(WD_API, {
        action: 'wbgetentities', ids: best.qid, props: 'sitelinks',
        sitefilter: 'ruwiki', format: 'json', origin: '*',
      });
      const ruTitle = data.entities?.[best.qid]?.sitelinks?.ruwiki?.title;
      if (ruTitle) {
        const wikiHits = await wikiSearchRaw(ruTitle, 3);
        pageid = wikiHits[0]?.pageid;
      }
    } catch (e) {}
    if (!pageid) {
      try {
        const wikiHits = await wikiSearchRaw(best.title, 3);
        pageid = wikiHits[0]?.pageid;
      } catch (e) {}
    }
  }
  if (!pageid) return null;

  let article;
  try {
    article = await wikiArticle(pageid);
  } catch (e) {
    // Если статья временно недоступна — всё равно сделаем доклад из найденного.
    article = { title: best.title, text: best.desc || '', qid: best.qid || '' };
  }
  const qid = article.qid || best.qid || '';
  let facts = {}, family = {};

  if (qid) {
    try {
      const fRows = await sparql(`
        SELECT ?dob ?dod ?bornLabel ?diedLabel ?countryLabel ?jobLabel WHERE {
          VALUES ?person { wd:${qid} }
          OPTIONAL { ?person wdt:P569 ?dob }
          OPTIONAL { ?person wdt:P570 ?dod }
          OPTIONAL { ?person wdt:P19 ?born }
          OPTIONAL { ?person wdt:P20 ?died }
          OPTIONAL { ?person wdt:P27 ?country }
          OPTIONAL { ?person wdt:P106 ?job }
          SERVICE wikibase:label { bd:serviceParam wikibase:language "ru,en". }
        } LIMIT 1`);
      const r = fRows[0] || {};
      facts = {
        birth: (val(r, 'dob') || '—').slice(0, 4),
        death: (val(r, 'dod') || '—').slice(0, 4),
        born: val(r, 'bornLabel') || '—',
        died: val(r, 'diedLabel') || '—',
        country: val(r, 'countryLabel') || '—',
        job: val(r, 'jobLabel') || '—',
      };
    } catch (e) {}

    try {
      const famRows = await sparql(`
        SELECT ?kind ?relLabel WHERE {
          VALUES ?person { wd:${qid} }
          VALUES (?p ?kind) {
            (wdt:P22 "Отец") (wdt:P25 "Мать") (wdt:P26 "Супруг(а)") (wdt:P40 "Дети")
          }
          ?person ?p ?rel .
          SERVICE wikibase:label { bd:serviceParam wikibase:language "ru,en". }
        } ORDER BY ?kind`);
      for (const row of famRows) {
        const k = val(row, 'kind');
        (family[k] ||= []).push(val(row, 'relLabel'));
      }
    } catch (e) {}
  }

  return {
    ...article, facts, family,
    query: found.query, matched: article.title, candidates: found.all,
    wikiUrl: 'https://ru.wikipedia.org/wiki/' + article.title.replace(/ /g, '_'),
    wdUrl: qid ? `https://www.wikidata.org/wiki/${qid}` : '',
  };
}

function sentences(text, limit = 6) {
  return String(text).replace(/\n/g, ' ')
    .split(/(?<=[.!?])\s+/)
    .map(s => s.trim())
    .filter(s => s.length > 40)
    .slice(0, limit);
}

// ---------- Текст доклада ----------

function makeReport(d, note) {
  const f = d.facts;
  const lines = [];
  lines.push(`Доклад на тему: ${d.title}`, '');
  lines.push('Введение');
  lines.push(`${d.title} — известная личность. Родился(ась) в ${f.birth} году в городе ${f.born}. Гражданство: ${f.country}. Род деятельности: ${f.job}.`);
  lines.push('', 'Основная часть');
  for (const s of sentences(d.text, 6)) lines.push(s);
  lines.push('', 'Семья');
  const fam = Object.entries(d.family);
  if (fam.length) for (const [k, v] of fam) lines.push(`${k}: ${v.join(', ')}`);
  else lines.push('Сведения о семье в открытых источниках отсутствуют.');
  lines.push('', 'Заключение');
  lines.push(`${d.title} оставил(а) заметный след. Его/её биография изучается в школе и может быть использована для самостоятельной работы.`);
  if (note) lines.push('', 'Дополнительно учесть: ' + note);
  lines.push('', 'Источники:', d.wikiUrl);
  if (d.wdUrl) lines.push(d.wdUrl);
  return lines.join('\n');
}

// ---------- Презентация ----------

async function makePptx(d, styleKey, colorKey, filePath) {
  const color = COLORS[colorKey][1];
  const pptx = new pptxgen();
  pptx.layout = 'LAYOUT_16x9';

  function addSlide(title, bullets, fontSize = 20) {
    const slide = pptx.addSlide();
    slide.background = { color };
    slide.addText(title, {
      x: 0.7, y: 0.5, w: 12, h: 1, fontSize: styleKey === 'kids' ? 38 : 34,
      bold: true, color: 'FFFFFF', align: 'center',
    });
    const text = bullets.map(b => '• ' + b).join('\n');
    slide.addText(text, {
      x: 0.9, y: 1.8, w: 11.6, h: 5.2,
      fontSize: styleKey === 'kids' ? fontSize + 2 : fontSize,
      color: 'FFFFFF', breakLine: false, paraSpaceAfter: 10,
    });
  }

  const f = d.facts;
  addSlide(d.title, ['Школьная презентация', `Стиль: ${STYLES[styleKey]}`, `Цвет: ${COLORS[colorKey][0]}`], 24);
  addSlide('Кто это?', [
    `Родился(ась): ${f.birth}`,
    `Место рождения: ${f.born}`,
    `Гражданство: ${f.country}`,
    `Деятельность: ${f.job}`,
  ]);
  addSlide('Биография', sentences(d.text, 4), 17);
  const fam = Object.entries(d.family).map(([k, v]) => `${k}: ${v.join(', ')}`);
  addSlide('Семья', fam.length ? fam : ['Сведений нет']);
  addSlide('География', [
    `Рождение: ${f.born}`, `Смерть: ${f.died}`, `Страна: ${f.country}`,
  ]);
  addSlide('Интересные факты', sentences(d.text, 3).reverse(), 17);
  addSlide('Вывод', [
    `${d.title} — важная историческая личность.`,
    'Его/её жизнь связана с ключевыми событиями эпохи.',
    'Изучение биографии помогает лучше понять историю.',
  ]);
  addSlide('Источники', [d.wikiUrl, d.wdUrl || 'Викиданные'].filter(Boolean), 16);

  await pptx.writeFile({ fileName: filePath });
}

// ---------- Хендлеры ----------

function matchesWord(text, words) {
  const low = String(text || '').toLowerCase().replace(/ё/g, 'е').trim();
  return words.some(w => low === w || low.startsWith(w + ' ') || low.includes(' ' + w + ' '));
}

async function askStyle(ctx, s) {
  s.step = 'pres_style';
  await ctx.reply('Выбери стиль презентации:', styleKeyboard());
}

async function startPresentationForLast(ctx, s) {
  if (s.lastName) {
    s.name = s.lastName;
    s.note = s.lastNote || '';
    await ctx.reply(`Делаю презентацию про: ${s.lastName}`);
    await askStyle(ctx, s);
  } else {
    s.step = 'pres_topic';
    await ctx.reply(
      'Напиши, про кого презентация и что должно быть на слайдах.\nНапример:\nАльберт Эйнштейн\nБиография, семья, открытия, вывод',
      removeKb
    );
  }
}

async function handleTextRequest(ctx, s, text) {
  const parts = text.split('\n').map(x => x.trim()).filter(Boolean);
  if (!parts.length) return ctx.reply('Напиши, про кого сделать работу.');
  await ctx.reply('Понял запрос. Ищу и сверяю по нескольким источникам...');
  try {
    const d = await collect(parts[0]);
    if (!d) {
      return ctx.reply(
        'Не смог понять, о ком речь. Напиши имя и фамилию полностью, ' +
        'например: Альберт Эйнштейн. Или просто: напиши про Эйнштейна.'
      );
    }
    if (d.matched && d.query !== d.matched.toLowerCase()) {
      await ctx.reply(`Похоже, ты имел(а) в виду: ${d.matched}`);
    }
    const note = parts.slice(1).join(' ');
    const report = makeReport(d, note);
    for (let i = 0; i < report.length; i += 3500) {
      await ctx.reply(report.slice(i, i + 3500));
    }
    s.lastName = d.matched || parts[0];
    s.lastNote = note;
    s.awaitingPres = true;
    s.step = 'menu';
    await ctx.reply('Хочешь также презентацию?', yesNoKeyboard());
  } catch (e) {
    console.error(e);
    const msg = String(e.message || '');
    if (/429/.test(msg)) {
      await ctx.reply('Википедия временно ограничила число запросов (429). ' +
        'Я уже повторил несколько раз — попробуй ещё раз через минуту.');
    } else {
      await ctx.reply('Ошибка: ' + msg);
    }
  }
}

bot.start(async (ctx) => {
  sessions.set(ctx.chat.id, { step: 'menu' });
  await ctx.reply(
    'Привет! Я помогу с школьной работой про человека.\n\nВыбери, что нужно сделать:',
    mainMenu
  );
});

bot.command('help', async (ctx) => {
  await ctx.reply(
    '📝 Готовый текст — напиши, про кого, получишь текст для вставки.\n' +
    '📊 Презентация — напиши тему, ответь на вопросы про дизайн и получи файл.\n\n' +
    '/start — начать заново'
  );
});

bot.command('cancel', async (ctx) => {
  sessions.set(ctx.chat.id, { step: 'menu' });
  await ctx.reply('Отменено. /start — начать заново.', mainMenu);
});

bot.on('text', async (ctx) => {
  const text = ctx.message.text.trim();
  const s = session(ctx);

  if (s.awaitingPres && s.step === 'menu') {
    if (matchesWord(text, YES_WORDS)) {
      s.awaitingPres = false;
      await startPresentationForLast(ctx, s);
      return;
    }
    if (matchesWord(text, NO_WORDS)) {
      s.awaitingPres = false;
      await ctx.reply('Хорошо. Если передумаешь — нажми 📊 Презентация.', mainMenu);
      return;
    }
    s.awaitingPres = false;
  }

  if (text === TEXT_BTN) {
    s.step = 'text_topic';
    await ctx.reply(
      'Напиши, про кого сделать текст, и что учесть.\nНапример:\nУинстон Черчилль\nДля 7 класса, про семью и войну',
      removeKb
    );
    return;
  }

  if (text === PRES_BTN) {
    s.awaitingPres = false;
    s.step = 'pres_topic';
    await ctx.reply(
      'Напиши, про кого презентация и что должно быть на слайдах.\nНапример:\nАльберт Эйнштейн\nБиография, семья, открытия, вывод',
      removeKb
    );
    return;
  }

  if (s.step === 'text_topic') {
    await handleTextRequest(ctx, s, text);
    return;
  }

  if (s.step === 'pres_topic') {
    const parts = text.split('\n').map(x => x.trim()).filter(Boolean);
    if (!parts.length) return ctx.reply('Напиши имя человека.');
    s.name = parts[0];
    s.note = parts.slice(1).join(' ');
    s.step = 'pres_style';
    await ctx.reply('Выбери стиль презентации:', styleKeyboard());
    return;
  }

  // Если человек просто написал «напиши про энштейна» — сразу понимаем это как запрос текста.
  if (s.step === 'menu') {
    await handleTextRequest(ctx, s, text);
    return;
  }

  await ctx.reply('Выбери кнопку или напиши /start.', mainMenu);
});

bot.action('pres:yes', async (ctx) => {
  const s = session(ctx);
  await ctx.answerCbQuery();
  await ctx.editMessageText('Отлично, делаю презентацию.');
  s.awaitingPres = false;
  await startPresentationForLast(ctx, s);
});

bot.action('pres:no', async (ctx) => {
  const s = session(ctx);
  await ctx.answerCbQuery();
  await ctx.editMessageText('Хорошо. Если передумаешь — нажми 📊 Презентация.');
  s.awaitingPres = false;
  s.step = 'menu';
  await ctx.reply('Что ещё сделать?', mainMenu);
});

bot.action(/^style:(.+)$/, async (ctx) => {
  const key = ctx.match[1];
  const s = session(ctx);
  s.style = key;
  s.step = 'pres_color';
  await ctx.answerCbQuery();
  await ctx.editMessageText(`Стиль: ${STYLES[key]}. Теперь выбери цвет:`);
  await ctx.reply('Цвет:', colorKeyboard());
});

bot.action(/^color:(.+)$/, async (ctx) => {
  const key = ctx.match[1];
  const s = session(ctx);
  await ctx.answerCbQuery();
  await ctx.editMessageText(`Цвет: ${COLORS[key][0]}. Делаю презентацию...`);
  try {
    const d = await collect(s.name);
    if (!d) return ctx.reply('Не нашёл человека в Википедии.');
    const filePath = path.join(os.tmpdir(), `presentation_${ctx.chat.id}.pptx`);
    await makePptx(d, s.style, key, filePath);
    await ctx.replyWithDocument(
      { source: fs.createReadStream(filePath), filename: `Презентация_${d.title}.pptx` },
      { caption: `Готовая презентация: ${d.title}\nСтиль: ${STYLES[s.style]}, цвет: ${COLORS[key][0]}` }
    );
    fs.unlinkSync(filePath);
    await ctx.reply('Ещё что-нибудь?', mainMenu);
    s.step = 'menu';
  } catch (e) {
    console.error(e);
    await ctx.reply('Не удалось сделать презентацию: ' + e.message);
    await ctx.reply('/start — вернуться в меню', mainMenu);
    s.step = 'menu';
  }
});

async function launchWithRetry() {
  for (let i = 1; i <= 12; i++) {
    try {
      await bot.launch();
      console.log('Бот запущен (polling)');
      return;
    } catch (e) {
      console.log(`Попытка ${i}/12 не удалась: ${e.message}. Повтор через 5 сек...`);
      await new Promise(r => setTimeout(r, 5000));
    }
  }
  console.error('Не удалось подключиться к Telegram. Проверь, что Happ подключён и прокси 10808 работает.');
  process.exit(1);
}

// --- Облако: webhook вместо polling ---
const PORT = process.env.PORT || 3000;
const WEBHOOK_URL = process.env.RENDER_EXTERNAL_URL
  ? process.env.RENDER_EXTERNAL_URL.replace(/\/$/, '') + '/webhook'
  : (process.env.WEBHOOK_URL || '');

async function start() {
  if (WEBHOOK_URL) {
    const app = express();
    app.use(express.json({ limit: '1mb' }));
    app.get('/', (req, res) => res.send('Bot is alive'));
    app.post('/webhook', bot.webhookCallback('/webhook'));

    await bot.telegram.setWebhook(WEBHOOK_URL);
    app.listen(PORT, () => {
      console.log('Webhook server started on port ' + PORT);
      console.log('Webhook URL: ' + WEBHOOK_URL);
    });
    return;
  }
  await launchWithRetry();
}

start();

process.once('SIGINT', () => bot.stop('SIGINT'));
process.once('SIGTERM', () => bot.stop('SIGTERM'));


}

// Happ использует локальный SOCKS-прокси 127.0.0.1:10808 (см. config.json).
// Если у тебя другой прокси — задай его переменной BOT_PROXY в run_bot.bat.
// На Windows используем локальный прокси Happ. На Linux-хостинге прокси не нужен.
const HAPP_PROXY = process.platform === 'win32' ? 'socks://127.0.0.1:10808' : '';
const PROXY = process.env.BOT_PROXY || HAPP_PROXY || detectProxy();
let agent;
if (PROXY) {
  agent = /^socks/i.test(PROXY) ? new SocksProxyAgent(PROXY) : new HttpsProxyAgent(PROXY);
  console.log('Использую прокси: ' + PROXY);
} else {
  console.log('Прокси не найден — Trying direct connection.');
}
const netOpts = agent ? { agent, httpsAgent: agent } : {};

const bot = new Telegraf(TOKEN, agent ? { telegram: { agent } } : {});
const WIKI = 'https://ru.wikipedia.org/w/api.php';
const WD_API = 'https://www.wikidata.org/w/api.php';
const DDG = 'https://api.duckduckgo.com/';
const SPARQL = 'https://query.wikidata.org/sparql';

// Википедия отклоняет запросы без нормального User-Agent — из-за этого был 403.
const UA = 'SchoolBot/1.0 (school project; t.me/schol_spg_bot)';
axios.defaults.headers.common['User-Agent'] = UA;
axios.defaults.headers.common['Accept'] = 'application/json';

const TEXT_BTN = '📝 Готовый текст';
const PRES_BTN = '📊 Презентация';

const mainMenu = Markup.keyboard([[TEXT_BTN, PRES_BTN]]).resize();
const removeKb = Markup.removeKeyboard();

const YES_WORDS = ['да', 'давай', 'давай же', 'ага', 'угу', 'конечно', 'ок', 'окей',
  'ладно', 'хочу', 'сделай', 'презентацию', 'презентация', 'yes', 'yeah', 'добре'];
const NO_WORDS = ['нет', 'не надо', 'не хочу', 'не сейчас', 'позже', 'стоп', 'отмена', 'no'];

const yesNoKeyboard = () => Markup.inlineKeyboard([
  [Markup.button.callback('Да, сделать презентацию', 'pres:yes'),
   Markup.button.callback('Нет, спасибо', 'pres:no')],
]);

const STYLES = { classic: 'Классический', modern: 'Современный', kids: 'Детский' };
const COLORS = {
  blue: ['Синий', '1F4E8C'],
  green: ['Зелёный', '226E4E'],
  red: ['Красный', '962D37'],
  violet: ['Фиолетовый', '5E378C'],
};

const styleKeyboard = () => Markup.inlineKeyboard(
  Object.entries(STYLES).map(([k, label]) => Markup.button.callback(label, `style:${k}`))
);
const colorKeyboard = () => Markup.inlineKeyboard(
  Object.entries(COLORS).map(([k, v]) => Markup.button.callback(v[0], `color:${k}`))
);

// Простейшие сессии в памяти
const sessions = new Map();
function session(ctx) {
  const id = ctx.chat.id;
  if (!sessions.has(id)) sessions.set(id, { step: 'menu' });
  return sessions.get(id);
}

// ---------- Открытые данные ----------

// --- Понимание разговорных запросов ("напиши про энштейна" и т.п.) ---
const STOP_WORDS = new Set([
  'напиши','написать','напишите','про','об','about','сделай','сделать','сделайте',
  'давай','дай','пришли','инфа','инфу','информация','информацию',
  'текст','доклад','доклада','презентация','презентацию','презентации',
  'биография','биографию','биографии','расскажи','рассказать','кто','такой','такая',
  'что','какой','какая','нужно','надо','помоги','найди','ищи',
  'поищи','сведения','данные','материал','материалы','тему','тема','для','класс',
  'класса','школе','школа','школь','человек','личность','жизнь','жизни'
]);

// Частые сокращения и опечатки -> полное имя.
const ALIASES = {
  'энштейн': 'альберт эйнштейн', 'энштейна': 'альберт эйнштейн', 'айнштейн': 'альберт эйнштейн',
  'пушкин': 'александр пушкин', 'пушкина': 'александр пушкин',
  'лермонтов': 'михаил лермонтов', 'толстой': 'лев толстой', 'чехов': 'антон чехов',
  'достоевский': 'фёдор достоевский', 'гоголь': 'николай гоголь',
  'черчилль': 'уинстон черчилль', 'наполеон': 'наполеон бонапарт',
  'сталин': 'иосиф сталин', 'ленин': 'владимир ленин', 'гитлер': 'адольф гитлер',
  'петр 1': 'пётр великий', 'петр великий': 'пётр великий', 'пётр 1': 'пётр великий',
  'екатерина 2': 'екатерина великая', 'екатерина великая': 'екатерина великая',
  'гагарин': 'юрий гагарин', 'королев': 'сергей королёв', 'менделеев': 'дмитрий менделеев',
  'ломоносов': 'михаил ломоносов', 'суворов': 'александр суворов', 'жуков': 'георгий жуков'
};

function cleanQuery(raw) {
  let s = String(raw || '').toLowerCase().replace(/ё/g, 'е')
    .replace(/[^a-zа-я0-9\s-]/gi, ' ').replace(/\s+/g, ' ').trim();
  let words = s.split(' ').filter(w => w.length > 1 && !STOP_WORDS.has(w));
  s = words.join(' ');
  return ALIASES[s] || s || String(raw || '').trim();
}

function lev(a, b) {
  a = String(a || ''); b = String(b || '');
  const prev = Array.from({ length: b.length + 1 }, (_, i) => i);
  for (let i = 1; i <= a.length; i++) {
    let last = prev[0]; prev[0] = i;
    for (let j = 1; j <= b.length; j++) {
      const old = prev[j];
      prev[j] = Math.min(prev[j] + 1, prev[j - 1] + 1, last + (a[i - 1] === b[j - 1] ? 0 : 1));
      last = old;
    }
  }
  return prev[b.length];
}

function sim(a, b) {
  a = String(a || '').toLowerCase(); b = String(b || '').toLowerCase();
  if (!a || !b) return 0;
  return 1 - lev(a, b) / Math.max(a.length, b.length);
}

function scoreCandidate(q, title, desc) {
  const t = String(title || '').toLowerCase().replace(/ё/g, 'е');
  const words = q.split(' ').filter(Boolean);
  let best = 0;
  for (const w of words) {
    if (t.includes(w)) best = Math.max(best, 0.92);
    for (const tw of t.split(' ')) best = Math.max(best, sim(w, tw) * 0.9);
    best = Math.max(best, sim(w, t) * 0.85);
  }
  best = Math.max(best, sim(q, t));
  if (desc && /человек|личность|персон/i.test(desc)) best += 0.04;
  return best;
}

async function wikiSearchRaw(name, limit = 8) {
  const url = WIKI + '?' + new URLSearchParams({
    action: 'query', list: 'search', srsearch: name,
    format: 'json', origin: '*', srlimit: String(limit),
  });
  const { data } = await axios.get(url, { timeout: 25000, ...netOpts });
  return data.query?.search || [];
}

async function wdSearch(name) {
  const url = WD_API + '?' + new URLSearchParams({
    action: 'wbsearchentities', search: name, language: 'ru',
    format: 'json', origin: '*', limit: '10',
  });
  const { data } = await axios.get(url, { timeout: 25000, ...netOpts });
  return data.search || [];
}

async function ddgEntity(name) {
  const { data } = await axios.get(DDG, {
    params: { q: name, format: 'json', no_html: 1, no_redirect: 1 },
    timeout: 20000, ...netOpts,
  });
  return data.Heading ? { title: data.Heading, desc: data.Abstract || '' } : null;
}

async function findPerson(raw) {
  let q = cleanQuery(raw);
  const queries = [q];
  if (q !== String(raw).trim()) queries.push(String(raw).trim());

  let candidates = [];
  for (const query of queries) {
    const [wiki, wd] = await Promise.allSettled([wikiSearchRaw(query, 8), wdSearch(query)]);
    if (wiki.status === 'fulfilled') {
      for (const h of wiki.value) {
        candidates.push({
          title: h.title, pageid: h.pageid,
          desc: String(h.snippet || '').replace(/<[^>]+>/g, ''),
          score: scoreCandidate(q, h.title, h.snippet || ''), source: 'wikipedia',
        });
      }
    }
    if (wd.status === 'fulfilled') {
      for (const e of wd.value) {
        candidates.push({
          title: e.label, qid: e.id, desc: e.description || '',
          score: scoreCandidate(q, e.label + ' ' + (e.aliases || []).join(' '), e.description || ''),
          source: 'wikidata',
        });
      }
    }
    if (candidates.length) break;

    // Если Википедия/Викиданные ничего не нашли — пробуем «яндексоподобный» ответ.
    try {
      const dd = await ddgEntity(query);
      if (dd) {
        const more = await wikiSearchRaw(dd.title, 3);
        for (const h of more) {
          candidates.push({
            title: h.title, pageid: h.pageid,
            desc: dd.desc || String(h.snippet || '').replace(/<[^>]+>/g, ''),
            score: scoreCandidate(q, h.title, dd.desc || ''), source: 'duckduckgo',
          });
        }
      }
    } catch (e) {}
    if (candidates.length) break;
  }

  const seen = new Set();
  candidates = candidates.filter(c => {
    const k = String(c.title || '').toLowerCase().replace(/ё/g, 'е');
    if (!k || seen.has(k)) return false;
    seen.add(k); return true;
  }).sort((a, b) => b.score - a.score);

  return { query: q, best: candidates[0] || null, all: candidates.slice(0, 5) };
}

async function wikiArticle(pageid) {
  const url = WIKI + '?' + new URLSearchParams({
    action: 'query', prop: 'extracts|pageprops', pageids: String(pageid),
    explaintext: '1', format: 'json', origin: '*',
  });
  const { data } = await axios.get(url, { timeout: 25000, ...netOpts });
  const page = data.query?.pages?.[String(pageid)] || {};
  return {
    title: page.title || '',
    text: page.extract || '',
    qid: page.pageprops?.wikibase_item || '',
  };
}

async function sparql(q) {
  const { data } = await axios.get(SPARQL, {
    params: { query: q, format: 'json' }, timeout: 25000, ...netOpts,
  });
  return data.results?.bindings || [];
}

const val = (row, key) => row?.[key]?.value || '';

async function collect(raw) {
  const found = await findPerson(raw);
  if (!found.best) return null;

  let best = found.best;
  let pageid = best.pageid;

  if (!pageid && best.qid) {
    try {
      const url = WD_API + '?' + new URLSearchParams({
        action: 'wbgetentities', ids: best.qid, props: 'sitelinks',
        sitefilter: 'ruwiki', format: 'json', origin: '*',
      });
      const { data } = await axios.get(url, { timeout: 25000, ...netOpts });
      const ruTitle = data.entities?.[best.qid]?.sitelinks?.ruwiki?.title;
      if (ruTitle) {
        const wikiHits = await wikiSearchRaw(ruTitle, 3);
        pageid = wikiHits[0]?.pageid;
      }
    } catch (e) {}
    if (!pageid) {
      try {
        const wikiHits = await wikiSearchRaw(best.title, 3);
        pageid = wikiHits[0]?.pageid;
      } catch (e) {}
    }
  }
  if (!pageid) return null;

  const article = await wikiArticle(pageid);
  const qid = article.qid || best.qid || '';
  let facts = {}, family = {};

  if (qid) {
    const fRows = await sparql(`
      SELECT ?dob ?dod ?bornLabel ?diedLabel ?countryLabel ?jobLabel WHERE {
        VALUES ?person { wd:${qid} }
        OPTIONAL { ?person wdt:P569 ?dob }
        OPTIONAL { ?person wdt:P570 ?dod }
        OPTIONAL { ?person wdt:P19 ?born }
        OPTIONAL { ?person wdt:P20 ?died }
        OPTIONAL { ?person wdt:P27 ?country }
        OPTIONAL { ?person wdt:P106 ?job }
        SERVICE wikibase:label { bd:serviceParam wikibase:language "ru,en". }
      } LIMIT 1`);
    const r = fRows[0] || {};
    facts = {
      birth: (val(r, 'dob') || '—').slice(0, 4),
      death: (val(r, 'dod') || '—').slice(0, 4),
      born: val(r, 'bornLabel') || '—',
      died: val(r, 'diedLabel') || '—',
      country: val(r, 'countryLabel') || '—',
      job: val(r, 'jobLabel') || '—',
    };

    const famRows = await sparql(`
      SELECT ?kind ?relLabel WHERE {
        VALUES ?person { wd:${qid} }
        VALUES (?p ?kind) {
          (wdt:P22 "Отец") (wdt:P25 "Мать") (wdt:P26 "Супруг(а)") (wdt:P40 "Дети")
        }
        ?person ?p ?rel .
        SERVICE wikibase:label { bd:serviceParam wikibase:language "ru,en". }
      } ORDER BY ?kind`);
    for (const row of famRows) {
      const k = val(row, 'kind');
      (family[k] ||= []).push(val(row, 'relLabel'));
    }
  }

  return {
    ...article, facts, family,
    query: found.query, matched: article.title, candidates: found.all,
    wikiUrl: 'https://ru.wikipedia.org/wiki/' + article.title.replace(/ /g, '_'),
    wdUrl: qid ? `https://www.wikidata.org/wiki/${qid}` : '',
  };
}

function sentences(text, limit = 6) {
  return String(text).replace(/\n/g, ' ')
    .split(/(?<=[.!?])\s+/)
    .map(s => s.trim())
    .filter(s => s.length > 40)
    .slice(0, limit);
}

// ---------- Текст доклада ----------

function makeReport(d, note) {
  const f = d.facts;
  const lines = [];
  lines.push(`Доклад на тему: ${d.title}`, '');
  lines.push('Введение');
  lines.push(`${d.title} — известная личность. Родился(ась) в ${f.birth} году в городе ${f.born}. Гражданство: ${f.country}. Род деятельности: ${f.job}.`);
  lines.push('', 'Основная часть');
  for (const s of sentences(d.text, 6)) lines.push(s);
  lines.push('', 'Семья');
  const fam = Object.entries(d.family);
  if (fam.length) for (const [k, v] of fam) lines.push(`${k}: ${v.join(', ')}`);
  else lines.push('Сведения о семье в открытых источниках отсутствуют.');
  lines.push('', 'Заключение');
  lines.push(`${d.title} оставил(а) заметный след. Его/её биография изучается в школе и может быть использована для самостоятельной работы.`);
  if (note) lines.push('', 'Дополнительно учесть: ' + note);
  lines.push('', 'Источники:', d.wikiUrl);
  if (d.wdUrl) lines.push(d.wdUrl);
  return lines.join('\n');
}

// ---------- Презентация ----------

async function makePptx(d, styleKey, colorKey, filePath) {
  const color = COLORS[colorKey][1];
  const pptx = new pptxgen();
  pptx.layout = 'LAYOUT_16x9';

  function addSlide(title, bullets, fontSize = 20) {
    const slide = pptx.addSlide();
    slide.background = { color };
    slide.addText(title, {
      x: 0.7, y: 0.5, w: 12, h: 1, fontSize: styleKey === 'kids' ? 38 : 34,
      bold: true, color: 'FFFFFF', align: 'center',
    });
    const text = bullets.map(b => '• ' + b).join('\n');
    slide.addText(text, {
      x: 0.9, y: 1.8, w: 11.6, h: 5.2,
      fontSize: styleKey === 'kids' ? fontSize + 2 : fontSize,
      color: 'FFFFFF', breakLine: false, paraSpaceAfter: 10,
    });
  }

  const f = d.facts;
  addSlide(d.title, ['Школьная презентация', `Стиль: ${STYLES[styleKey]}`, `Цвет: ${COLORS[colorKey][0]}`], 24);
  addSlide('Кто это?', [
    `Родился(ась): ${f.birth}`,
    `Место рождения: ${f.born}`,
    `Гражданство: ${f.country}`,
    `Деятельность: ${f.job}`,
  ]);
  addSlide('Биография', sentences(d.text, 4), 17);
  const fam = Object.entries(d.family).map(([k, v]) => `${k}: ${v.join(', ')}`);
  addSlide('Семья', fam.length ? fam : ['Сведений нет']);
  addSlide('География', [
    `Рождение: ${f.born}`, `Смерть: ${f.died}`, `Страна: ${f.country}`,
  ]);
  addSlide('Интересные факты', sentences(d.text, 3).reverse(), 17);
  addSlide('Вывод', [
    `${d.title} — важная историческая личность.`,
    'Его/её жизнь связана с ключевыми событиями эпохи.',
    'Изучение биографии помогает лучше понять историю.',
  ]);
  addSlide('Источники', [d.wikiUrl, d.wdUrl || 'Викиданные'].filter(Boolean), 16);

  await pptx.writeFile({ fileName: filePath });
}

// ---------- Хендлеры ----------

function matchesWord(text, words) {
  const low = String(text || '').toLowerCase().replace(/ё/g, 'е').trim();
  return words.some(w => low === w || low.startsWith(w + ' ') || low.includes(' ' + w + ' '));
}

async function askStyle(ctx, s) {
  s.step = 'pres_style';
  await ctx.reply('Выбери стиль презентации:', styleKeyboard());
}

async function startPresentationForLast(ctx, s) {
  if (s.lastName) {
    s.name = s.lastName;
    s.note = s.lastNote || '';
    await ctx.reply(`Делаю презентацию про: ${s.lastName}`);
    await askStyle(ctx, s);
  } else {
    s.step = 'pres_topic';
    await ctx.reply(
      'Напиши, про кого презентация и что должно быть на слайдах.\nНапример:\nАльберт Эйнштейн\nБиография, семья, открытия, вывод',
      removeKb
    );
  }
}

async function handleTextRequest(ctx, s, text) {
  const parts = text.split('\n').map(x => x.trim()).filter(Boolean);
  if (!parts.length) return ctx.reply('Напиши, про кого сделать работу.');
  await ctx.reply('Понял запрос. Ищу и сверяю по нескольким источникам...');
  try {
    const d = await collect(parts[0]);
    if (!d) {
      return ctx.reply(
        'Не смог понять, о ком речь. Напиши имя и фамилию полностью, ' +
        'например: Альберт Эйнштейн. Или просто: напиши про Эйнштейна.'
      );
    }
    if (d.matched && d.query !== d.matched.toLowerCase()) {
      await ctx.reply(`Похоже, ты имел(а) в виду: ${d.matched}`);
    }
    const note = parts.slice(1).join(' ');
    const report = makeReport(d, note);
    for (let i = 0; i < report.length; i += 3500) {
      await ctx.reply(report.slice(i, i + 3500));
    }
    s.lastName = d.matched || parts[0];
    s.lastNote = note;
    s.awaitingPres = true;
    s.step = 'menu';
    await ctx.reply('Хочешь также презентацию?', yesNoKeyboard());
  } catch (e) {
    console.error(e);
    await ctx.reply('Ошибка: ' + e.message);
  }
}

bot.start(async (ctx) => {
  sessions.set(ctx.chat.id, { step: 'menu' });
  await ctx.reply(
    'Привет! Я помогу с школьной работой про человека.\n\nВыбери, что нужно сделать:',
    mainMenu
  );
});

bot.command('help', async (ctx) => {
  await ctx.reply(
    '📝 Готовый текст — напиши, про кого, получишь текст для вставки.\n' +
    '📊 Презентация — напиши тему, ответь на вопросы про дизайн и получи файл.\n\n' +
    '/start — начать заново'
  );
});

bot.command('cancel', async (ctx) => {
  sessions.set(ctx.chat.id, { step: 'menu' });
  await ctx.reply('Отменено. /start — начать заново.', mainMenu);
});

bot.on('text', async (ctx) => {
  const text = ctx.message.text.trim();
  const s = session(ctx);

  if (s.awaitingPres && s.step === 'menu') {
    if (matchesWord(text, YES_WORDS)) {
      s.awaitingPres = false;
      await startPresentationForLast(ctx, s);
      return;
    }
    if (matchesWord(text, NO_WORDS)) {
      s.awaitingPres = false;
      await ctx.reply('Хорошо. Если передумаешь — нажми 📊 Презентация.', mainMenu);
      return;
    }
    s.awaitingPres = false;
  }

  if (text === TEXT_BTN) {
    s.step = 'text_topic';
    await ctx.reply(
      'Напиши, про кого сделать текст, и что учесть.\nНапример:\nУинстон Черчилль\nДля 7 класса, про семью и войну',
      removeKb
    );
    return;
  }

  if (text === PRES_BTN) {
    s.awaitingPres = false;
    s.step = 'pres_topic';
    await ctx.reply(
      'Напиши, про кого презентация и что должно быть на слайдах.\nНапример:\nАльберт Эйнштейн\nБиография, семья, открытия, вывод',
      removeKb
    );
    return;
  }

  if (s.step === 'text_topic') {
    await handleTextRequest(ctx, s, text);
    return;
  }

  if (s.step === 'pres_topic') {
    const parts = text.split('\n').map(x => x.trim()).filter(Boolean);
    if (!parts.length) return ctx.reply('Напиши имя человека.');
    s.name = parts[0];
    s.note = parts.slice(1).join(' ');
    s.step = 'pres_style';
    await ctx.reply('Выбери стиль презентации:', styleKeyboard());
    return;
  }

  // Если человек просто написал «напиши про энштейна» — сразу понимаем это как запрос текста.
  if (s.step === 'menu') {
    await handleTextRequest(ctx, s, text);
    return;
  }

  await ctx.reply('Выбери кнопку или напиши /start.', mainMenu);
});

bot.action('pres:yes', async (ctx) => {
  const s = session(ctx);
  await ctx.answerCbQuery();
  await ctx.editMessageText('Отлично, делаю презентацию.');
  s.awaitingPres = false;
  await startPresentationForLast(ctx, s);
});

bot.action('pres:no', async (ctx) => {
  const s = session(ctx);
  await ctx.answerCbQuery();
  await ctx.editMessageText('Хорошо. Если передумаешь — нажми 📊 Презентация.');
  s.awaitingPres = false;
  s.step = 'menu';
  await ctx.reply('Что ещё сделать?', mainMenu);
});

bot.action(/^style:(.+)$/, async (ctx) => {
  const key = ctx.match[1];
  const s = session(ctx);
  s.style = key;
  s.step = 'pres_color';
  await ctx.answerCbQuery();
  await ctx.editMessageText(`Стиль: ${STYLES[key]}. Теперь выбери цвет:`);
  await ctx.reply('Цвет:', colorKeyboard());
});

bot.action(/^color:(.+)$/, async (ctx) => {
  const key = ctx.match[1];
  const s = session(ctx);
  await ctx.answerCbQuery();
  await ctx.editMessageText(`Цвет: ${COLORS[key][0]}. Делаю презентацию...`);
  try {
    const d = await collect(s.name);
    if (!d) return ctx.reply('Не нашёл человека в Википедии.');
    const filePath = path.join(os.tmpdir(), `presentation_${ctx.chat.id}.pptx`);
    await makePptx(d, s.style, key, filePath);
    await ctx.replyWithDocument(
      { source: fs.createReadStream(filePath), filename: `Презентация_${d.title}.pptx` },
      { caption: `Готовая презентация: ${d.title}\nСтиль: ${STYLES[s.style]}, цвет: ${COLORS[key][0]}` }
    );
    fs.unlinkSync(filePath);
    await ctx.reply('Ещё что-нибудь?', mainMenu);
    s.step = 'menu';
  } catch (e) {
    console.error(e);
    await ctx.reply('Не удалось сделать презентацию: ' + e.message);
    await ctx.reply('/start — вернуться в меню', mainMenu);
    s.step = 'menu';
  }
});

async function launchWithRetry() {
  for (let i = 1; i <= 12; i++) {
    try {
      await bot.launch();
      console.log('Бот запущен (polling)');
      return;
    } catch (e) {
      console.log(`Попытка ${i}/12 не удалась: ${e.message}. Повтор через 5 сек...`);
      await new Promise(r => setTimeout(r, 5000));
    }
  }
  console.error('Не удалось подключиться к Telegram. Проверь, что Happ подключён и прокси 10808 работает.');
  process.exit(1);
}

// --- Облако: webhook вместо polling ---
const PORT = process.env.PORT || 3000;
const WEBHOOK_URL = process.env.RENDER_EXTERNAL_URL
  ? process.env.RENDER_EXTERNAL_URL.replace(/\/$/, '') + '/webhook'
  : (process.env.WEBHOOK_URL || '');

async function start() {
  if (WEBHOOK_URL) {
    const app = express();
    app.use(express.json({ limit: '1mb' }));
    app.get('/', (req, res) => res.send('Bot is alive'));
    app.post('/webhook', bot.webhookCallback('/webhook'));

    await bot.telegram.setWebhook(WEBHOOK_URL);
    app.listen(PORT, () => {
      console.log('Webhook server started on port ' + PORT);
      console.log('Webhook URL: ' + WEBHOOK_URL);
    });
    return;
  }
  await launchWithRetry();
}

start();

process.once('SIGINT', () => bot.stop('SIGINT'));
process.once('SIGTERM', () => bot.stop('SIGTERM'));

