/* Excel 公式练习区 · 页面 */
(function () {
  'use strict';
  const X = window.XL, B = window.PRACTICE_BANK, C = window.PRACTICE_CORE;
  const $ = (s, r = document) => r.querySelector(s);
  const $$ = (s, r = document) => [...r.querySelectorAll(s)];
  function h(tag, attrs, ...kids) {
    const el = document.createElement(tag);
    if (attrs) for (const [k, v] of Object.entries(attrs)) {
      if (v == null || v === false) continue;
      if (k === 'class') el.className = v;
      else if (k.startsWith('on')) el.addEventListener(k.slice(2), v);
      else if (k === 'html') el.innerHTML = v;
      else el.setAttribute(k, v === true ? '' : v);
    }
    for (const k of kids.flat(Infinity)) { if (k == null || k === false) continue; el.append(k instanceof Node ? k : document.createTextNode(String(k))); }
    return el;
  }
  const esc = s => String(s).replace(/[&<>"']/g, c => ({ '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;', "'": '&#39;' }[c]));
  const rich = (tag, s, attrs) => h(tag, { ...(attrs || {}), html: esc(s).replace(/«(.+?)»/g, '<code>$1</code>') });
  const pad = n => String(n).padStart(2, '0');
  const dayKey = d => `${d.getFullYear()}-${pad(d.getMonth() + 1)}-${pad(d.getDate())}`;
  const TODAY = dayKey(new Date());
  const addDays = (key, n) => { const [y, m, d] = key.split('-').map(Number); return dayKey(new Date(y, m - 1, d + n)); };
  const WEEK = ['周日', '周一', '周二', '周三', '周四', '周五', '周六'];
  const isWeekend = [0, 6].includes(new Date().getDay());
  function hash(s) { let x = 2166136261; for (let i = 0; i < s.length; i++) { x ^= s.charCodeAt(i); x = Math.imul(x, 16777619); } return (x >>> 0).toString(36); }

  /* ---------- 存储（只在这台电脑的这个浏览器里） ---------- */
  const LS = 'excel-practice-v1';
  const S = { items: {}, daily: {}, size: 3 };
  try { Object.assign(S, JSON.parse(localStorage.getItem(LS) || '{}')); } catch (e) { /* 浏览器不允许存储 */ }
  let canStore = true;
  function save() { try { localStorage.setItem(LS, JSON.stringify(S)); } catch (e) { canStore = false; } }
  const INTERVAL = [1, 2, 4, 7, 15, 30];
  function record(id, ok, patch) {
    const st = { box: 0, ok: 0, bad: 0, ...(S.items[id] || {}), ...(patch || {}) };
    const peeked = st.peek === TODAY;
    if (ok && !peeked) { st.box = Math.min((st.box || 0) + 1, 5); st.ok++; st.due = addDays(TODAY, INTERVAL[st.box]); }
    else { if (!ok) st.bad++; st.box = 0; st.due = addDays(TODAY, 1); }
    st.lastOk = ok && !peeked; st.last = Date.now(); st.doneOn = TODAY;
    S.items[id] = st; save(); refreshChrome();
  }
  function status(id) {
    const st = S.items[id];
    if (!st || !st.last) return st && st.peek ? 'peek' : 'new';
    if (st.doneOn === TODAY) return st.lastOk ? 'done' : 'redo';
    if (st.due && st.due <= TODAY) return 'due';
    return st.box >= 3 ? 'master' : 'learning';
  }
  const ST = { new: ['没做过', 'st-new'], peek: ['看过答案', 'st-peek'], done: ['今天做对了', 'st-done'], redo: ['明天再练', 'st-redo'], due: ['该复习了', 'st-due'], master: ['已掌握', 'st-done'], learning: ['练习中', 'st-new'] };
  const pill = id => { const [t, c] = ST[status(id)]; return h('span', { class: 'pill ' + c, 'data-pill': id }, t); };

  /* ---------- 题目 ---------- */
  const ITEMS = []; const BY = {};
  const add = it => { if (BY[it.id]) return; BY[it.id] = it; ITEMS.push(it); };
  B.SCENARIOS.forEach(sc => sc.exercises.forEach((ex, i) => add({ id: ex.id, kind: 'grid', area: sc.area, link: sc.link, title: ex.title, fns: ex.fns, sc, ex, order: ITEMS.length, challenge: !!ex.challenge })));
  B.RECALL.forEach(r => add({ ...r, kind: 'recall', title: r.q, order: ITEMS.length }));
  const KIND = { grid: '表格题', recall: '回忆题', cloze: '填空题' };
  let SITE = { ok: null, fnCount: {}, sections: 0 };
  const FN_RE = /\b([A-Za-z][A-Za-z0-9]*(?:\.[A-Za-z]+)*)\s*\(/g;
  const NOT_FN = new Set(['IN', 'BY']);
  function fnsOf(text) { const out = []; let m; FN_RE.lastIndex = 0; while ((m = FN_RE.exec(text))) { const f = m[1].toUpperCase(); if (!NOT_FN.has(f) && !out.includes(f) && /[A-Z]/.test(f)) out.push(f); } return out; }
  async function loadSite() {
    try {
      const res = await fetch('index.html', { cache: 'no-cache' });
      if (!res.ok) throw new Error(res.status);
      const doc = new DOMParser().parseFromString(await res.text(), 'text/html');
      let n = 0; const seenText = new Set();
      $$('main section', doc).forEach(sec => {
        const title = (sec.querySelector('h2') || {}).textContent || '';
        const id = sec.id || '';
        const scen = /场景\s*(\d+)/.exec(title);
        const area = scen ? '场景 ' + scen[1] : '通用';
        const link = id ? 'index.html#' + id : 'index.html';
        $$('.formula', sec).forEach(f => {
          const text = f.textContent.trim();
          const fns = fnsOf(text);
          fns.forEach(fn => (SITE.fnCount[fn] = (SITE.fnCount[fn] || 0) + 1));
          if (!fns.length || seenText.has(text) || !/[(]/.test(text) || text.length < 12) return;
          seenText.add(text);
          fns.slice(0, 4).forEach(fn => add({ id: 'cz-' + hash(text + '|' + fn), kind: 'cloze', area, link, title: `${title.replace(/^场景\s*\d+：/, '')}：补全 ${fns.length > 1 ? '其中一个' : ''}函数`, fns: [fn], text, fn, from: title, order: ITEMS.length }));
          n++;
        });
        $$('.quiz-item', sec).forEach(q => {
          const qs = (q.querySelector('strong') || {}).textContent || ''; const a = (q.querySelector('.answer') || {}).textContent || '';
          if (!qs || !a) return;
          const fns = fnsOf(a);
          fns.forEach(fn => (SITE.fnCount[fn] = (SITE.fnCount[fn] || 0) + 1));
          add({ id: 'qz-' + hash(qs), kind: 'recall', area: '主动自测', link: 'index.html#quiz', title: qs, q: qs, a: [a], fns, order: ITEMS.length, fromSite: true });
        });
      });
      SITE.ok = true; SITE.sections = n;
    } catch (e) { SITE.ok = false; }
  }

  /* ---------- 今日题目 ---------- */
  function fnPractice() { const c = {}; for (const it of ITEMS) { const st = S.items[it.id]; if (!st || !st.last) continue; for (const f of it.fns || []) c[f] = (c[f] || 0) + 1; } return c; }
  function buildSet(n, exclude) {
    exclude = new Set(exclude || []);
    const pick = [];
    const due = ITEMS.filter(it => !exclude.has(it.id) && S.items[it.id] && S.items[it.id].due && S.items[it.id].due <= TODAY && S.items[it.id].doneOn !== TODAY)
      .sort((a, b) => S.items[a.id].due.localeCompare(S.items[b.id].due));
    for (const it of due) { if (pick.length >= Math.max(1, n - 1)) break; pick.push(it.id); }
    const seen = fnPractice();
    let fresh = ITEMS.filter(it => !exclude.has(it.id) && !pick.includes(it.id) && !(S.items[it.id] && S.items[it.id].last));
    const order = ['grid', 'recall', 'cloze'];
    let ti = 0;
    while (pick.length < n && fresh.length) {
      const want = order[ti++ % order.length];
      let pool = fresh.filter(it => it.kind === want); if (!pool.length) pool = fresh;
      const score = it => (it.challenge ? 100 : 0) + Math.min(...(it.fns && it.fns.length ? it.fns : ['-']).map(f => seen[f] || 0)) * 10;
      pool.sort((a, b) => score(a) - score(b) || (a.kind === 'grid' ? a.order - b.order : hash(TODAY + a.id).localeCompare(hash(TODAY + b.id))));
      const it = pool[0]; pick.push(it.id); fresh = fresh.filter(x => x !== it);
      (it.fns || []).forEach(f => (seen[f] = (seen[f] || 0) + 1));
    }
    for (const it of due) { if (pick.length >= n) break; if (!pick.includes(it.id)) pick.push(it.id); }
    return pick;
  }
  function todaySet() { return (S.daily[TODAY] || []).filter(id => BY[id]); }
  function ensureToday() {
    if (!isWeekend && !S.daily[TODAY]) { S.daily[TODAY] = buildSet(S.size || 3); const keys = Object.keys(S.daily).sort(); while (keys.length > 60) delete S.daily[keys.shift()]; save(); }
  }

  /* ---------- 表格 ---------- */
  let K = 0;
  function sheetMeta(sc, name) { return sc.sheets.find(s => s.name === name); }
  function renderGrid(it, key) {
    const { sc, ex } = it;
    const tg = C.parseTarget(ex.target);
    const exSheet = C.sheetOf(sc, ex);
    const T = C.truth(sc);
    const exp = T.cellsByEx[ex.id];
    const box = h('div', { class: 'grid-box' });
    const tabs = h('div', { class: 'tabs', role: 'tablist' });
    const wrap = h('div', { class: 'tbl-wrap' });
    const info = h('p', { class: 'grid-info', id: key + '-info' }, '点任意格子，看它的地址和内容。黄色格子是你要写公式的地方。');
    let current = exSheet;
    const targetSet = new Set(exp.map(c => c.addr));
    const paint = () => {
      $$('.tab', tabs).forEach(t => t.classList.toggle('on', t.dataset.sheet === current));
      wrap.textContent = '';
      const meta = sheetMeta(sc, current); const book = T.book; const size = book.size(current);
      let rows = size.rows, cols = size.cols;
      if (current === exSheet) { rows = Math.max(rows, tg.r2, ...exp.map(c => c.row)); cols = Math.max(cols, tg.c2, ...exp.map(c => c.col)); }
      rows = Math.min(rows, 40); cols = Math.min(cols, 16);
      const table = h('table', { class: 'grid' });
      table.append(h('thead', null, h('tr', null, h('th', { class: 'corner' }), Array.from({ length: cols }, (_, i) => h('th', null, X.numToCol(i + 1))))));
      const tb = h('tbody');
      for (let r = 1; r <= rows; r++) {
        const tr = h('tr', null, h('th', { class: 'rowh' }, r));
        for (let c = 1; c <= cols; c++) {
          const addr = X.numToCol(c) + r;
          const isT = current === exSheet && targetSet.has(addr);
          const v = isT ? null : (meta.cells[addr] !== undefined ? meta.cells[addr] : null);
          const fmt = (meta.fmt || {})[X.numToCol(c)];
          const cls = ['cell']; if (r === 1) cls.push('hd'); if (isT) cls.push('tgt'); if (typeof v === 'number') cls.push('num');
          if (isT && ex.spill && addr !== tg.first) cls.push('spill');
          tr.append(h('td', { class: cls.join(' '), 'data-addr': addr, onclick: () => cellInfo(it, key, current, addr) }, v === null ? '' : X.show(v, fmt)));
        }
        tb.append(tr);
      }
      table.append(tb); wrap.append(table);
      paintResult(it, key);
    };
    sc.sheets.forEach(s => tabs.append(h('button', { type: 'button', class: 'tab', 'data-sheet': s.name, role: 'tab', onclick: () => { current = s.name; paint(); } }, s.name)));
    box.append(tabs, wrap, info);
    box._paint = paint; box._current = () => current; box._exSheet = exSheet;
    setTimeout(paint, 0);
    return box;
  }
  const RESULTS = {};
  function paintResult(it, key) {
    const card = document.getElementById(key); if (!card) return;
    const box = $('.grid-box', card); if (!box || box._current() !== box._exSheet) return;
    const res = RESULTS[key];
    const fmt = it.ex.fmt || (sheetMeta(it.sc, box._exSheet).fmt || {});
    $$('td.tgt', box).forEach(td => { td.classList.remove('ok', 'bad', 'err'); td.textContent = ''; });
    $$('td.extra', box).forEach(td => { td.classList.remove('extra', 'bad'); td.textContent = ''; });
    if (!res || res.error) return;
    for (const c of res.cells) {
      const td = $(`td[data-addr="${c.addr}"]`, box); if (!td) continue;
      const f = typeof fmt === 'string' ? fmt : fmt[c.addr.replace(/\d+/, '')];
      td.textContent = X.show(c.v, f);
      td.classList.add(c.ok ? 'ok' : 'bad'); if (X.isErr(c.v)) td.classList.add('err'); if (c.extra) td.classList.add('extra');
    }
  }
  function cellInfo(it, key, sheet, addr) {
    const info = document.getElementById(key + '-info'); if (!info) return; info.textContent = '';
    const exSheet = C.sheetOf(it.sc, it.ex);
    const res = RESULTS[key];
    const c = res && !res.error ? res.cells.find(x => x.addr === addr) : null;
    if (sheet === exSheet && c) {
      const f = it.ex.fmt || (sheetMeta(it.sc, exSheet).fmt || {})[addr.replace(/\d+/, '')];
      if (res.literal) { info.append(h('strong', null, addr), '：没有以 = 开头，所以显示的是文字本身。'); return; }
      if (it.ex.spill) info.append(h('strong', null, addr), '：由 ', h('code', null, C.parseTarget(it.ex.target).first), ' 的公式溢出得到 → ');
      else { let ft = ''; try { ft = '=' + X.shiftText(res.body, c.dr || 0, c.dc || 0); } catch (e) { ft = ''; } info.append(h('strong', null, addr), ' 实际执行 ', h('code', null, ft), ' → '); }
      info.append(h('span', { class: c.ok ? 'ok' : 'bad' }, X.show(c.v, f) || '（空）'));
      if (!c.ok) info.append(h('span', { class: 'muted' }, `，应该是 ${X.show(c.exp, f) || '（空）'}`));
      return;
    }
    const meta = sheetMeta(it.sc, sheet); const v = meta.cells[addr];
    const f = (meta.fmt || {})[addr.replace(/\d+/, '')];
    const ref = sheet === exSheet ? addr : (/^[A-Za-z_㐀-鿿][\w.㐀-鿿]*$/.test(sheet) && !/[（）()]/.test(sheet) ? `${sheet}!${addr}` : `'${sheet.replace(/'/g, "''")}'!${addr}`);
    info.append(h('strong', null, ref), v === undefined || v === '' ? '：空格子' : `：${X.show(v, f)}${typeof v === 'number' ? (f ? '（日期/时间，本质是数字）' : '（数字）') : '（文字）'}`);
  }

  /* ---------- 题卡 ---------- */
  function card(it) {
    const key = 'k' + (++K);
    const head = h('header', { class: 'ex-h' },
      h('span', { class: 'kind kind-' + it.kind }, KIND[it.kind]),
      h('span', { class: 'area' }, it.area),
      h('h3', null, it.kind === 'grid' ? it.title : it.kind === 'cloze' ? it.title : '回忆一下'),
      pill(it.id));
    const el = h('article', { class: 'ex' + (it.challenge ? ' chal' : ''), id: key, 'data-item': it.id }, head);
    if (it.kind === 'grid') gridCard(it, key, el);
    else if (it.kind === 'recall') recallCard(it, key, el);
    else clozeCard(it, key, el);
    return el;
  }
  const put = (el, ...kids) => kids.flat(Infinity).forEach(k => { if (k != null && k !== false) el.append(k); });
  function linkTo(it) { return it.link ? h('a', { class: 'more', href: it.link }, '看经验库讲解 →') : null; }

  function gridCard(it, key, el) {
    const { sc, ex } = it;
    const tg = C.parseTarget(ex.target);
    const sheetNote = ex.sheet ? `（在「${ex.sheet}」表里）` : '';
    const where = ex.spill ? `在 «${tg.first}» 写一个公式${sheetNote}，结果会自动往下溢出。`
      : tg.cells.length === 1 ? `在 «${tg.first}» 这一个格子里写公式${sheetNote}。`
        : (tg.c2 > tg.c1 && tg.r2 > tg.r1) ? `在 «${tg.first}» 写公式${sheetNote}，网站会往下、往右填满到 «${tg.last}»。`
          : `在 «${tg.first}» 写公式${sheetNote}，网站会像表格一样自动往下填充到 «${tg.last}»。`;
    const st = S.items[it.id] || {};
    const inp = h('input', { id: key + '-f', class: 'fin', type: 'text', spellcheck: 'false', autocomplete: 'off', autocapitalize: 'off', placeholder: '=', 'aria-label': '公式', onkeydown: e => { if (e.key === 'Enter') { e.preventDefault(); run(); } } });
    if (st.formula) inp.value = st.formula;
    let hints = 0;
    const hintList = h('ol', { class: 'hints' });
    const hintBtn = h('button', { class: 'btn small', type: 'button', onclick: () => { hints = Math.min(hints + 1, ex.hints.length); paintHints(); } });
    const paintHints = () => { hintList.textContent = ''; ex.hints.slice(0, hints).forEach((t, i) => hintList.append(h('li', null, h('span', { class: 'hint-n' }, `提示 ${i + 1}`), rich('span', t)))); hintBtn.textContent = hints >= ex.hints.length ? '提示已全部给出' : hints ? `再给一个提示（${hints}/${ex.hints.length}）` : '给我一个提示'; hintBtn.disabled = hints >= ex.hints.length; };
    paintHints();
    const fb = h('div', { class: 'fb', 'aria-live': 'polite' });
    const ans = h('div', { class: 'ans', hidden: true });
    const confirm = confirmBox(() => { peek(it.id); showAns(); });
    function showAns() {
      ans.textContent = '';
      ans.append(
        h('div', { class: 'ans-f' }, h('span', { class: 'lbl' }, '参考写法'), h('code', null, ex.answer), copyBtn(ex.answer)),
        h('table', { class: 'steps' }, h('tbody', null, ex.steps.map(s => h('tr', null, h('th', null, h('code', null, s[0])), h('td', null, s[1]))))),
        ex.alt ? h('div', { class: 'alts' }, h('p', { class: 'lbl' }, '其他写法'), ex.alt.map(a => h('p', null, h('code', null, a[0]), h('span', { class: 'muted' }, a[1])))) : null,
        linkTo(it));
      ans.hidden = false;
    }
    function run() {
      const f = inp.value;
      const res = C.grade(sc, ex, f); RESULTS[key] = res;
      paintResult(it, key); paintFb(res, true);
      if (!f.trim() || res.error && !f.trim()) return;
      record(it.id, res.solved, { formula: f.trim() });
    }
    function paintFb(res, fresh) {
      fb.textContent = '';
      if (res.error) { fb.append(h('div', { class: 'msg bad' }, h('strong', null, '公式写法有问题'), h('p', null, res.error))); return; }
      if (res.solved) { fb.append(h('div', { class: 'msg ok' }, h('strong', null, `全部正确，${res.total} / ${res.total}`), h('p', null, fresh ? (S.items[it.id] && S.items[it.id].peek === TODAY ? '做对了。因为今天看过答案，这题明天还会再出现一次。' : '这个套路你已经会了，过几天会再来复习一次。') : '这是你上次写的答案。'), h('button', { class: 'btn small', type: 'button', onclick: showAns }, '对比参考写法'))); return; }
      fb.append(h('div', { class: 'msg warn' }, h('strong', null, res.total > 1 ? `${res.ok} / ${res.total} 格正确` : '结果不对'), h('ul', null, res.tips.map(t => h('li', null, t)))));
    }
    put(el,
      h('p', { class: 'ask' }, ex.ask),
      h('p', { class: 'story muted' }, `${sc.title}｜${sc.story}`),
      renderGrid(it, key),
      rich('p', where, { class: 'where muted' }),
      h('div', { class: 'fbar' }, h('label', { class: 'fx', for: key + '-f' }, h('span', { class: 'cellname' }, tg.first), h('i', null, 'fx')), inp, h('button', { class: 'btn primary', type: 'button', onclick: run }, '运行检查')),
      fb,
      h('div', { class: 'tools' }, hintBtn, h('button', { class: 'btn small', type: 'button', onclick: () => (status(it.id) === 'done' || status(it.id) === 'master' ? showAns() : (confirm.hidden = false)) }, '看答案和拆解'), linkTo(it)),
      confirm, hintList, ans);
    if (st.formula) setTimeout(() => { const res = C.grade(sc, ex, st.formula); RESULTS[key] = res; paintResult(it, key); paintFb(res, false); }, 0);
  }
  function confirmBox(onYes) {
    const box = h('div', { class: 'confirm', hidden: true },
      h('p', null, '确定现在看答案吗？看过答案的题明天会再出现一次，自己写对才算掌握。'),
      h('div', { class: 'row' },
        h('button', { class: 'btn primary small', type: 'button', onclick: () => { box.hidden = true; onYes(); } }, '看答案'),
        h('button', { class: 'btn small', type: 'button', onclick: () => (box.hidden = true) }, '再想想')));
    return box;
  }
  function peek(id) { const st = S.items[id] || { box: 0, ok: 0, bad: 0 }; if (S.items[id] && S.items[id].doneOn === TODAY && S.items[id].lastOk) return; st.peek = TODAY; S.items[id] = st; save(); refreshChrome(); }
  function copyBtn(text) {
    return h('button', { class: 'btn small', type: 'button', onclick: async e => { const b = e.currentTarget; try { await navigator.clipboard.writeText(text); b.textContent = '已复制'; setTimeout(() => (b.textContent = '复制'), 1400); } catch (_) { const r = document.createRange(); r.selectNodeContents(b.previousElementSibling); const s = getSelection(); s.removeAllRanges(); s.addRange(r); } } }, '复制');
  }
  function recallCard(it, key, el) {
    const inp = h('textarea', { id: key + '-f', class: 'rin', rows: '2', spellcheck: 'false', placeholder: '先在脑子里想，再把公式或答案写下来' });
    const out = h('div', { class: 'fb' });
    const reveal = (mine) => {
      out.textContent = '';
      const exact = mine !== null ? C.checkRecall(it, inp.value) : false;
      if (exact) { out.append(h('div', { class: 'msg ok' }, h('strong', null, '完全一致'), h('p', null, h('code', null, it.a[0])))); record(it.id, true); return; }
      out.append(h('div', { class: 'msg warn' },
        h('strong', null, mine === null ? '参考答案' : '对照参考答案'),
        h('p', null, h('code', null, it.a[0])),
        it.note ? h('p', { class: 'muted' }, it.note) : null,
        mine === null ? h('p', { class: 'muted' }, '这题记作没答上，明天会再出现。') :
          h('div', { class: 'row' }, h('p', { class: 'muted' }, '意思一样也算对，你自己判断：'),
            h('button', { class: 'btn small primary', type: 'button', onclick: () => { record(it.id, true); out.textContent = ''; out.append(h('div', { class: 'msg ok' }, h('strong', null, '记作答对'), h('p', null, h('code', null, it.a[0])))); } }, '我答对了'),
            h('button', { class: 'btn small', type: 'button', onclick: () => { record(it.id, false); out.textContent = ''; out.append(h('div', { class: 'msg warn' }, h('strong', null, '记作没答对，明天再练'), h('p', null, h('code', null, it.a[0])))); } }, '没答对'))));
      if (mine === null) record(it.id, false);
    };
    put(el,
      h('p', { class: 'ask' }, it.q),
      it.fromSite ? h('p', { class: 'story muted' }, '来自经验库「主动遮挡自测」') : null,
      inp,
      h('div', { class: 'tools' },
        h('button', { class: 'btn primary', type: 'button', onclick: () => { if (!inp.value.trim()) { inp.focus(); return; } reveal(inp.value); } }, '对答案'),
        h('button', { class: 'btn small quiet', type: 'button', onclick: () => reveal(null) }, '想不起来，看答案'),
        linkTo(it)),
      out);
  }
  function clozeCard(it, key, el) {
    const idx = it.text.toUpperCase().search(new RegExp('\\b' + it.fn.replace(/\./g, '\\.') + '\\s*\\(', 'i'));
    const pre = h('pre', { class: 'cloze' });
    if (idx >= 0) pre.append(it.text.slice(0, idx), h('mark', null, '＿'.repeat(Math.max(3, Math.min(it.fn.length, 8)))), it.text.slice(idx + it.fn.length));
    else pre.append(it.text);
    const inp = h('input', { id: key + '-f', class: 'fin solo', type: 'text', spellcheck: 'false', autocomplete: 'off', autocapitalize: 'characters', placeholder: '函数名', onkeydown: e => { if (e.key === 'Enter') check(); } });
    const out = h('div', { class: 'fb' });
    const FR = { SI: 'IF', SIERREUR: 'IFERROR', 'SI.NON.DISP': 'IFNA', ESTNUM: 'ISNUMBER', 'NB.SI': 'COUNTIF', 'NB.SI.ENS': 'COUNTIFS', RECHERCHEV: 'VLOOKUP', RECHERCHEX: 'XLOOKUP', EQUIV: 'MATCH', SUPPRESPACE: 'TRIM', CNUM: 'VALUE', SOMMEPROD: 'SUMPRODUCT', FILTRE: 'FILTER', ENT: 'INT', JOURSEM: 'WEEKDAY', TEXTE: 'TEXT', SUBSTITUE: 'SUBSTITUTE', 'JOINDRE.TEXTE': 'TEXTJOIN', LIGNES: 'ROWS', SOMME: 'SUM', MINUSCULE: 'LOWER', TEMPS: 'TIME' };
    function check() {
      const v = inp.value.trim().toUpperCase().replace(/\(.*$/, ''); if (!v) { inp.focus(); return; }
      const ok = v === it.fn || FR[v] === it.fn;
      out.textContent = '';
      out.append(h('div', { class: 'msg ' + (ok ? 'ok' : 'warn') }, h('strong', null, ok ? '答对了' : `不对，这里是 ${it.fn}`), h('p', null, ok ? '回到原场景，看看它在整个公式里负责哪一步。' : `你写的是 ${v}。去经验库里看看 ${it.fn} 在这里负责哪一步。`)));
      record(it.id, ok);
    }
    put(el,
      h('p', { class: 'ask' }, `下面是经验库「${it.from}」里的公式，空着的地方是哪个函数？`),
      pre,
      h('div', { class: 'tools' }, inp, h('button', { class: 'btn primary', type: 'button', onclick: check }, '检查'), h('button', { class: 'btn small quiet', type: 'button', onclick: () => { inp.value = ''; out.textContent = ''; out.append(h('div', { class: 'msg warn' }, h('strong', null, `答案：${it.fn}`))); peek(it.id); record(it.id, false); } }, '看答案'), linkTo(it)),
      out);
  }

  /* ---------- 视图 ---------- */
  const main = () => $('#main');
  function viewToday(m) {
    if (SITE.ok === null && !S.daily[TODAY]) { m.append(h('p', { class: 'empty' }, '正在准备今天的题…')); return; }
    ensureToday();
    const set = todaySet();
    const d = new Date();
    const done = set.filter(id => S.items[id] && S.items[id].doneOn === TODAY).length;
    m.append(h('header', { class: 'hero' },
      h('p', { class: 'eyebrow' }, `${d.getMonth() + 1}月${d.getDate()}日 ${WEEK[d.getDay()]}`),
      h('h1', null, isWeekend && !set.length ? '周末不推送新题' : '今日练习'),
      h('p', { class: 'lede' }, isWeekend && !set.length ? '工作日每天早上会自动准备好新的一组题。周末可以休息，也可以复习一下到期的题。' : `今天 ${set.length} 道题，完成 ${done} 道。题目来自经验库里出现过的公式：做错或看过答案的题明天会再出现，做对的题会隔几天再来复习。`),
      h('div', { class: 'row' },
        h('span', { class: 'muted' }, '每天题量'),
        [3, 5].map(n => h('button', { type: 'button', class: 'chip' + ((S.size || 3) === n ? ' on' : ''), 'aria-pressed': (S.size || 3) === n ? 'true' : 'false', onclick: () => { S.size = n; if (S.daily[TODAY] && S.daily[TODAY].length < n) S.daily[TODAY] = S.daily[TODAY].concat(buildSet(n - S.daily[TODAY].length, S.daily[TODAY])); save(); render(); } }, `${n} 道`)),
        h('button', { type: 'button', class: 'btn small', onclick: () => { S.daily[TODAY] = (S.daily[TODAY] || []).concat(buildSet(3, S.daily[TODAY] || [])); save(); render(); } }, isWeekend && !set.length ? '想练就来 3 道' : '再来 3 道'))));
    if (!canStore) m.append(h('p', { class: 'note' }, '这个浏览器不允许保存数据，刷新后进度会丢失。'));
    if (SITE.ok === false) m.append(h('p', { class: 'note' }, '现在是直接双击打开的本地文件，读不到经验库页面，所以暂时没有「填空题」和自测回忆题。用 GitHub Pages 的网址打开就有了。'));
    if (!set.length) {
      const due = ITEMS.filter(it => status(it.id) === 'due');
      if (due.length) { m.append(h('h2', { class: 'sec-h' }, `到期该复习的题（${due.length}）`)); m.append(h('div', { class: 'list' }, due.map(row))); }
      return;
    }
    const all = done === set.length;
    if (all) m.append(h('div', { class: 'msg ok wide' }, h('strong', null, '今天的题做完了'), h('p', null, '明天早上会有新的一组。想多练可以点「再来 3 道」。')));
    m.append(h('div', { class: 'cards' }, set.map(id => card(BY[id]))));
  }
  function row(it) {
    const st = S.items[it.id] || {};
    return h('a', { class: 'row-i', href: '#ex=' + encodeURIComponent(it.id) },
      h('span', { class: 'kind kind-' + it.kind }, KIND[it.kind]),
      h('span', { class: 'row-t' }, h('strong', null, it.title), h('span', { class: 'muted' }, `${it.area} · ${(it.fns || []).slice(0, 4).join(' + ') || '基础运算'}`)),
      h('span', { class: 'muted small' }, st.last ? `对 ${st.ok || 0} · 错 ${st.bad || 0}` : ''),
      pill(it.id));
  }
  let BANK = { area: '全部', kind: '全部', fn: '', q: '' };
  function viewBank(m) {
    const areas = ['全部', ...new Set(ITEMS.map(i => i.area))];
    m.append(h('header', { class: 'page-h' }, h('h1', null, '题库'), h('p', { class: 'lede' }, `一共 ${ITEMS.length} 道：表格题 ${ITEMS.filter(i => i.kind === 'grid').length} 道（在网页里写公式、自动判对错），回忆题 ${ITEMS.filter(i => i.kind === 'recall').length} 道，填空题 ${ITEMS.filter(i => i.kind === 'cloze').length} 道。填空题和自测回忆题是从经验库页面自动生成的，经验库新增公式后会自动多出来。`)));
    const list = h('div', { class: 'list' });
    const paint = () => {
      list.textContent = '';
      const q = BANK.q.trim().toLowerCase();
      const res = ITEMS.filter(i => (BANK.area === '全部' || i.area === BANK.area) && (BANK.kind === '全部' || KIND[i.kind] === BANK.kind) && (!BANK.fn || (i.fns || []).includes(BANK.fn)) && (!q || [i.title, i.area, (i.fns || []).join(' '), i.text || '', i.q || ''].join(' ').toLowerCase().includes(q)));
      if (!res.length) { list.append(h('p', { class: 'empty' }, '没有符合条件的题。')); return; }
      res.forEach(i => list.append(row(i)));
    };
    m.append(h('div', { class: 'filters' },
      h('input', { type: 'search', class: 'search', placeholder: '搜标题、函数…', value: BANK.q, oninput: e => { BANK.q = e.target.value; paint(); } }),
      h('div', { class: 'chips' }, ['全部', ...Object.values(KIND)].map(k => h('button', { type: 'button', class: 'chip' + (BANK.kind === k ? ' on' : ''), onclick: () => { BANK.kind = k; render(); } }, k))),
      h('div', { class: 'chips' }, areas.map(a => h('button', { type: 'button', class: 'chip' + (BANK.area === a ? ' on' : ''), onclick: () => { BANK.area = a; render(); } }, a))),
      BANK.fn ? h('p', null, '只看函数：', h('code', null, BANK.fn), ' ', h('button', { type: 'button', class: 'link', onclick: () => { BANK.fn = ''; render(); } }, '清除')) : null), list);
    paint();
  }
  function viewCoverage(m) {
    const fns = {};
    for (const it of ITEMS) for (const f of it.fns || []) { fns[f] = fns[f] || { grid: 0, other: 0, master: 0, tried: 0 }; fns[f][it.kind === 'grid' ? 'grid' : 'other']++; const st = S.items[it.id]; if (st && st.last) fns[f].tried++; if (st && st.box >= 3) fns[f].master++; }
    for (const f of Object.keys(SITE.fnCount)) fns[f] = fns[f] || { grid: 0, other: 0, master: 0, tried: 0 };
    const list = Object.entries(fns).sort((a, b) => (SITE.fnCount[b[0]] || 0) - (SITE.fnCount[a[0]] || 0) || a[0].localeCompare(b[0]));
    const noGrid = list.filter(([, v]) => !v.grid).map(([k]) => k);
    m.append(h('header', { class: 'page-h' }, h('h1', null, '函数覆盖'), h('p', { class: 'lede' }, '经验库里出现过的每个函数，都有对应的练习。按在经验库里出现的次数排序。'),
      noGrid.length ? h('p', { class: 'note' }, `这些函数暂时只有回忆题或填空题，还没有表格题：${noGrid.join('、')}。需要的话可以让 Claude 补充。`) : null));
    const tb = h('tbody');
    for (const [f, v] of list) {
      tb.append(h('tr', null,
        h('td', null, h('code', null, f)),
        h('td', { class: 'num' }, SITE.fnCount[f] || '–'),
        h('td', { class: 'num' }, v.grid || '–'),
        h('td', { class: 'num' }, v.other || '–'),
        h('td', { class: 'num' }, `${v.master} / ${v.grid + v.other}`),
        h('td', null, h('button', { type: 'button', class: 'btn small', onclick: () => { BANK = { area: '全部', kind: '全部', fn: f, q: '' }; location.hash = 'bank'; } }, '练这个'))));
    }
    m.append(h('div', { class: 'tbl-wrap card' }, h('table', { class: 'cov' }, h('thead', null, h('tr', null, h('th', null, '函数'), h('th', null, '经验库出现'), h('th', null, '表格题'), h('th', null, '回忆/填空'), h('th', null, '已掌握'), h('th', null, ''))), tb)));
  }
  function viewReview(m) {
    const groups = [['due', '该复习了'], ['redo', '今天做错 / 看过答案'], ['peek', '看过答案还没做'], ['learning', '练习中'], ['master', '已掌握']];
    m.append(h('header', { class: 'page-h' }, h('h1', null, '复习本'), h('p', { class: 'lede' }, '做对的题按 1、2、4、7、15、30 天的间隔回来复习；做错或看过答案的题，第二天会回来。')));
    let any = false;
    for (const [k, t] of groups) {
      const items = ITEMS.filter(it => status(it.id) === k); if (!items.length) continue; any = true;
      m.append(h('h2', { class: 'sec-h' }, `${t}（${items.length}）`), h('div', { class: 'list' }, items.map(row)));
    }
    if (!any) m.append(h('p', { class: 'empty' }, '还没有做过题。先去「今日练习」做第一组吧。'));
  }
  function viewBackup(m) {
    const n = Object.values(S.items).filter(s => s.last).length;
    const file = h('input', { type: 'file', accept: '.json,application/json', id: 'imp', class: 'file' });
    const msg = h('p', { class: 'muted' });
    file.addEventListener('change', async () => {
      const f = file.files[0]; if (!f) return;
      try { const j = JSON.parse(await f.text()); if (!j || typeof j.items !== 'object') throw new Error('格式不对'); S.items = { ...S.items, ...j.items }; S.daily = { ...(j.daily || {}), ...S.daily }; save(); msg.textContent = `已导入 ${Object.keys(j.items).length} 道题的记录。`; refreshChrome(); }
      catch (e) { msg.textContent = '这个文件不是练习区导出的进度文件。'; }
    });
    const reset = h('div', { class: 'confirm', hidden: true }, h('p', null, '确定清空所有练习记录吗？清空后无法恢复，建议先导出备份。'),
      h('div', { class: 'row' }, h('button', { class: 'btn small danger', type: 'button', onclick: () => { S.items = {}; S.daily = {}; save(); render(); } }, '清空'), h('button', { class: 'btn small', type: 'button', onclick: () => (reset.hidden = true) }, '取消')));
    m.append(h('header', { class: 'page-h' }, h('h1', null, '备份进度'), h('p', { class: 'lede' }, `练习记录保存在这台电脑的这个浏览器里（目前 ${n} 道题有记录）。换电脑、换浏览器或清理浏览器数据之前，先导出一份。`)),
      h('section', { class: 'card' },
        h('h2', null, '导出'), h('p', { class: 'muted' }, '下载一个 .json 文件，里面是每道题的对错次数和复习日期。'),
        h('div', null, h('button', { class: 'btn primary', type: 'button', onclick: () => { const blob = new Blob([JSON.stringify({ items: S.items, daily: S.daily, exportedAt: new Date().toISOString() }, null, 1)], { type: 'application/json' }); const a = h('a', { href: URL.createObjectURL(blob), download: `excel练习进度-${TODAY}.json` }); document.body.append(a); a.click(); a.remove(); } }, '导出进度文件'))),
      h('section', { class: 'card' }, h('h2', null, '导入'), h('p', { class: 'muted' }, '在新电脑上选择之前导出的文件，记录会合并进来。'), h('label', { class: 'btn', for: 'imp' }, '选择进度文件'), file, msg),
      h('section', { class: 'card' }, h('h2', null, '清空'), h('div', null, h('button', { class: 'btn small danger', type: 'button', onclick: () => (reset.hidden = false) }, '清空所有记录')), reset));
  }
  function viewOne(m, id) {
    const it = BY[id];
    m.append(h('p', null, h('a', { class: 'more', href: '#bank' }, '← 回到题库')));
    if (!it) { m.append(h('p', { class: 'empty' }, SITE.ok === null ? '正在读取经验库…' : '没有找到这道题，可能经验库里的公式改过了。')); return; }
    m.append(card(it));
  }

  /* ---------- 外框 ---------- */
  function route() { const t = decodeURIComponent((location.hash || '').slice(1)); if (t.startsWith('ex=')) return ['ex', t.slice(3)]; return [['today', 'bank', 'coverage', 'review', 'backup'].includes(t) ? t : 'today']; }
  function refreshChrome() {
    const [v] = route();
    $$('#nav a').forEach(a => a.classList.toggle('on', a.dataset.v === v || (v === 'ex' && a.dataset.v === 'bank')));
    const set = todaySet(); const left = set.filter(id => !(S.items[id] && S.items[id].doneOn === TODAY)).length;
    const b = $('#nav-today-n'); if (b) b.textContent = set.length ? (left ? `${left} 道待做` : '已完成') : '';
    const due = ITEMS.filter(it => status(it.id) === 'due').length;
    const r = $('#nav-review-n'); if (r) r.textContent = due ? String(due) : '';
    $$('[data-pill]').forEach(p => { const [t, c] = ST[status(p.dataset.pill)]; p.textContent = t; p.className = 'pill ' + c; });
  }
  function render() {
    const m = main(); m.textContent = '';
    const [v, arg] = route();
    if (v === 'bank') viewBank(m); else if (v === 'coverage') viewCoverage(m); else if (v === 'review') viewReview(m); else if (v === 'backup') viewBackup(m); else if (v === 'ex') viewOne(m, arg); else viewToday(m);
    refreshChrome();
  }
  window.addEventListener('hashchange', () => { render(); window.scrollTo(0, 0); });
  render();
  loadSite().then(render);
})();
