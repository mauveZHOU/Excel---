/* Excel 公式练习区 · 公式引擎
 * 支持：单元格/区域/整列/整行引用、跨表引用 '表名'!A1、往下/往右填充时的 $ 锁定、
 * 数组运算（ARRAYFORMULA / FILTER / UNIQUE / SUMPRODUCT …）、LET、LAMBDA、MAP、日期时间函数。
 * 法语版函数名和分号分隔也能识别。
 */
(function (G) {
  'use strict';
  class XErr { constructor(code) { this.code = code; } }
  const ERR = c => new XErr(c);
  const isErr = v => v instanceof XErr;
  class FE extends Error {}
  class Arr {
    constructor(rows) { this.rows = rows; }
    get h() { return this.rows.length; }
    get w() { return this.rows[0] ? this.rows[0].length : 0; }
    flat() { return this.rows.flat(); }
    static col(list) { return new Arr(list.map(v => [v])); }
  }
  class Lambda { constructor(params, body, scope) { this.params = params; this.body = body; this.scope = scope; } }
  const isArr = v => v instanceof Arr;

  function colToNum(c) { c = c.toUpperCase(); let n = 0; for (const ch of c) n = n * 26 + ch.charCodeAt(0) - 64; return n; }
  function numToCol(n) { let s = ''; while (n > 0) { const m = (n - 1) % 26; s = String.fromCharCode(65 + m) + s; n = Math.floor((n - 1) / 26); } return s; }

  const CN_PUNCT = { '，': ',', '（': '(', '）': ')', '“': '"', '”': '"', '：': ':', '；': ';', '＄': '$', '＞': '>', '＜': '<', '＝': '=', '＂': '"', '‘': "'", '’': "'" };
  const FR_ALIAS = {
    'SI': 'IF', 'SIERREUR': 'IFERROR', 'SI.NON.DISP': 'IFNA', 'ESTNUM': 'ISNUMBER', 'ESTTEXTE': 'ISTEXT', 'ESTVIDE': 'ISBLANK',
    'ESTERREUR': 'ISERROR', 'ESTNA': 'ISNA', 'ET': 'AND', 'OU': 'OR', 'NON': 'NOT', 'ARRONDI': 'ROUND', 'ARRONDI.SUP': 'ROUNDUP',
    'ARRONDI.INF': 'ROUNDDOWN', 'ENT': 'INT', 'SOMME': 'SUM', 'MOYENNE': 'AVERAGE', 'NB': 'COUNT', 'NBVAL': 'COUNTA',
    'NB.VIDE': 'COUNTBLANK', 'NB.SI': 'COUNTIF', 'NB.SI.ENS': 'COUNTIFS', 'SOMME.SI': 'SUMIF', 'SOMME.SI.ENS': 'SUMIFS',
    'MOYENNE.SI': 'AVERAGEIF', 'RECHERCHEV': 'VLOOKUP', 'RECHERCHEX': 'XLOOKUP', 'EQUIV': 'MATCH', 'SUPPRESPACE': 'TRIM',
    'CNUM': 'VALUE', 'NBCAR': 'LEN', 'GAUCHE': 'LEFT', 'DROITE': 'RIGHT', 'STXT': 'MID', 'MAJUSCULE': 'UPPER',
    'MINUSCULE': 'LOWER', 'CONCATENER': 'CONCATENATE', 'SUBSTITUE': 'SUBSTITUTE', 'TROUVE': 'FIND', 'CHERCHE': 'SEARCH',
    'TEXTE': 'TEXT', 'JOINDRE.TEXTE': 'TEXTJOIN', 'FILTRE': 'FILTER', 'UNIQUE': 'UNIQUE', 'LIGNES': 'ROWS', 'COLONNES': 'COLUMNS',
    'SOMMEPROD': 'SUMPRODUCT', 'INDIRECT': 'INDIRECT', 'JOURSEM': 'WEEKDAY', 'SERIE.JOUR.OUVRE.INTL': 'WORKDAY.INTL',
    'NB.JOURS.OUVRES.INTL': 'NETWORKDAYS.INTL', 'TEMPS': 'TIME', 'DATE': 'DATE', 'MOD': 'MOD', 'MAX': 'MAX', 'MIN': 'MIN'
  };
  const SHEET_NAME = /^[A-Za-z_㐀-鿿][A-Za-z0-9_.㐀-鿿]*(?=!)/;

  /* ---------- 词法 ---------- */
  function tokenize(src) {
    const t = []; let i = 0;
    const after = len => { let k = i + len; while (src[k] === ' ' || src[k] === '\n' || src[k] === '\t') k++; return src[k]; };
    const prev = () => t[t.length - 1];
    let sheet = null, sheetStart = 0;
    const take = (tok, len) => { if (sheet) { tok.sheet = sheet; tok.s = sheetStart; sheet = null; } t.push(tok); i += len; };
    while (i < src.length) {
      const c = src[i];
      if (c === ' ' || c === '\t' || c === '\n' || c === '\r' || c === ' ') { i++; continue; }
      // 跨表引用：'表名'!A1 或 表名!A1
      if (c === "'") {
        let j = i + 1, name = '';
        for (;;) {
          if (j >= src.length) throw new FE("表名两边的单引号 ' 没有配对。跨表引用要写成 '表名'!A1。");
          if (src[j] === "'") { if (src[j + 1] === "'") { name += "'"; j += 2; continue; } break; }
          name += src[j++];
        }
        if (src[j + 1] !== '!') throw new FE(`'${name}' 后面要紧跟感叹号 !，例如 '${name}'!A1。`);
        sheet = name; sheetStart = i; i = j + 2; continue;
      }
      const sm = SHEET_NAME.exec(src.slice(i));
      if (sm) { sheet = sm[0]; sheetStart = i; i += sm[0].length + 1; continue; }
      if (/[㐀-鿿]/.test(c)) { const m = /^[㐀-鿿\w]+/.exec(src.slice(i)); throw new FE(`「${m[0]}」是文字。公式里的文字要用英文双引号括起来，写成 "${m[0]}"。如果它是表名，后面要加感叹号，例如 ${m[0]}!A1。`); }
      if (CN_PUNCT[c]) throw new FE(`「${c}」是中文标点。公式里的逗号、括号、引号、冒号都要用英文半角符号，这里应该写「${CN_PUNCT[c]}」。`);
      if (c === '"') {
        let j = i + 1, v = '';
        for (;;) {
          if (j >= src.length) throw new FE('有一个英文双引号 " 没有配对。文字两边都要有引号，例如 "一致"。');
          if (src[j] === '"') { if (src[j + 1] === '"') { v += '"'; j += 2; continue; } break; }
          v += src[j++];
        }
        take({ t: 'str', v, s: i, e: j + 1 }, j + 1 - i); continue;
      }
      // 整行引用 $3:$3 / 1:27
      const p = prev();
      const rr = /^(\$?)(\d+):(\$?)(\d+)(?![\d.A-Za-z])/.exec(src.slice(i));
      if (rr && (sheet || !p || (p.t === 'op' && !')}%'.includes(p.v)))) {
        take({ t: 'rows', r1: +rr[2], r2: +rr[4], a1: !!rr[1], a2: !!rr[3], s: i, e: i + rr[0].length }, rr[0].length); continue;
      }
      if (/[0-9.]/.test(c)) {
        const m = /^(\d+\.?\d*|\.\d+)(e[+-]?\d+)?/i.exec(src.slice(i));
        if (m) { take({ t: 'num', v: parseFloat(m[0]), s: i, e: i + m[0].length }, m[0].length); continue; }
      }
      if (c === '$' || /[A-Za-z_]/.test(c)) {
        const rest = src.slice(i);
        const rm = /^(\$?)([A-Za-z]{1,3})(\$?)(\d+)(?![A-Za-z0-9_.])/.exec(rest);
        if (rm && after(rm[0].length) !== '(') {
          take({ t: 'ref', col: colToNum(rm[2]), row: +rm[4], ac: !!rm[1], ar: !!rm[3], s: i, e: i + rm[0].length }, rm[0].length); continue;
        }
        const cm = /^(\$?)([A-Za-z]{1,3})(?![A-Za-z0-9_.(])/.exec(rest);
        if (cm && (after(cm[0].length) === ':' || (p && p.t === 'op' && p.v === ':'))) {
          take({ t: 'col', col: colToNum(cm[2]), ac: !!cm[1], s: i, e: i + cm[0].length }, cm[0].length); continue;
        }
        const wm = /^[A-Za-z_][A-Za-z0-9_.]*/.exec(rest);
        if (wm) {
          const w = wm[0].toUpperCase();
          if (after(wm[0].length) === '(') { take({ t: 'fn', v: FR_ALIAS[w] || w, raw: wm[0], s: i, e: i + wm[0].length }, wm[0].length); continue; }
          if (w === 'TRUE' || w === 'FALSE' || w === 'VRAI' || w === 'FAUX') { take({ t: 'bool', v: w === 'TRUE' || w === 'VRAI', s: i, e: i + wm[0].length }, wm[0].length); continue; }
          take({ t: 'name', v: wm[0], s: i, e: i + wm[0].length }, wm[0].length); continue;
        }
        throw new FE(`看不懂「${c}」。`);
      }
      const two = src.slice(i, i + 2);
      if (two === '<=' || two === '>=' || two === '<>') { take({ t: 'op', v: two, s: i, e: i + 2 }, 2); continue; }
      if ('+-*/^&=<>(),:%;{}'.includes(c)) { take({ t: 'op', v: c, s: i, e: i + 1 }, 1); continue; }
      throw new FE(`字符「${c}」不能出现在公式里。`);
    }
    if (sheet) throw new FE(`表名 ${sheet}! 后面少了单元格地址，例如 ${sheet}!A1。`);
    return t;
  }

  /* ---------- 语法 ---------- */
  function parse(src) {
    const toks = tokenize(src); let p = 0;
    const isOp = v => toks[p] && toks[p].t === 'op' && toks[p].v === v;
    const expect = v => { if (!isOp(v)) throw new FE(v === ')' ? '括号没有配对：少了一个右括号 )。' : v === '}' ? '大括号没有配对：少了 }。' : `这里应该是「${v}」。`); p++; };
    const CMP = ['=', '<>', '<', '>', '<=', '>='];
    const expr = () => cmp();
    function cmp() { let l = cat(); while (toks[p] && toks[p].t === 'op' && CMP.includes(toks[p].v)) { const op = toks[p++].v; l = { t: 'bin', op, l, r: cat() }; } return l; }
    function cat() { let l = add(); while (isOp('&')) { p++; l = { t: 'bin', op: '&', l, r: add() }; } return l; }
    function add() { let l = mul(); while (isOp('+') || isOp('-')) { const op = toks[p++].v; l = { t: 'bin', op, l, r: mul() }; } return l; }
    function mul() { let l = pw(); while (isOp('*') || isOp('/')) { const op = toks[p++].v; l = { t: 'bin', op, l, r: pw() }; } return l; }
    function pw() { let l = un(); while (isOp('^')) { p++; l = { t: 'bin', op: '^', l, r: un() }; } return l; }
    function un() { if (isOp('-')) { p++; return { t: 'neg', e: un() }; } if (isOp('+')) { p++; return un(); } return post(); }
    function post() { let e = prim(); while (isOp('%')) { p++; e = { t: 'pct', e }; } return e; }
    function prim() {
      const k = toks[p];
      if (!k) throw new FE('公式还没写完，结尾少了内容。');
      if (k.t === 'num' || k.t === 'str' || k.t === 'bool') { p++; return { t: k.t, v: k.v }; }
      if (k.t === 'name') { p++; return { t: 'name', v: k.v }; }
      if (k.t === 'rows') { p++; return { t: 'rows', k }; }
      if (k.t === 'ref') {
        p++;
        if (isOp(':')) {
          p++; const b = toks[p];
          if (b && b.t === 'ref') { p++; return { t: 'range', a: k, b, sheet: k.sheet }; }
          if (b && b.t === 'col') { p++; return { t: 'range', a: k, b: { col: b.col, ac: b.ac, row: null, ar: true }, sheet: k.sheet }; }
          throw new FE('冒号 : 后面应该是单元格，例如 D6:D15。');
        }
        return { t: 'ref', r: k, sheet: k.sheet };
      }
      if (k.t === 'col') { p++; expect(':'); const b = toks[p]; if (!b || b.t !== 'col') throw new FE('整列引用要写成 E:E 这样的形式。'); p++; return { t: 'colrange', a: k, b, sheet: k.sheet }; }
      if (k.t === 'fn') {
        p++; expect('(');
        const args = [];
        if (isOp(')')) { p++; return { t: 'call', name: k.v, raw: k.raw, args }; }
        for (;;) {
          if (isOp(',') || isOp(';') || isOp(')')) args.push({ t: 'empty' }); else args.push(expr());
          if (isOp(',') || isOp(';')) { p++; continue; }
          if (isOp(')')) { p++; break; }
          if (!toks[p]) throw new FE(`函数 ${k.raw} 少了右括号 )。`);
          if (isOp(':')) throw new FE('这个网站暂不支持「函数结果:函数结果」这种区域写法。');
          throw new FE(`函数 ${k.raw} 的参数之间要用英文逗号 , 隔开。`);
        }
        return { t: 'call', name: k.v, raw: k.raw, args };
      }
      if (k.t === 'op' && k.v === '{') {
        p++; const rows = [[]];
        for (;;) {
          rows[rows.length - 1].push(expr());
          if (isOp(',')) { p++; continue; }
          if (isOp(';')) { p++; rows.push([]); continue; }
          expect('}'); break;
        }
        return { t: 'arr', rows };
      }
      if (k.t === 'op' && k.v === '(') { p++; const e = expr(); expect(')'); return e; }
      if (k.t === 'op' && k.v === ')') throw new FE('多了一个右括号 )，或者括号里是空的。');
      if (k.t === 'op' && (k.v === ',' || k.v === ';')) throw new FE('逗号只能出现在函数的括号里，用来分隔参数。');
      throw new FE(`「${k.v}」出现的位置不对。`);
    }
    const ast = expr();
    if (p < toks.length) {
      const k = toks[p];
      if (k.t === 'op' && k.v === ')') throw new FE('多了一个右括号 )。');
      throw new FE('公式中间好像缺了运算符或逗号，检查一下有没有多写或漏写的内容。');
    }
    return { ast, toks };
  }

  /* ---------- 值 ---------- */
  const sc = v => (isArr(v) ? (v.h === 1 && v.w === 1 ? v.rows[0][0] : ERR('#VALUE!')) : v);
  function fmtNum(v) { if (Number.isInteger(v)) return String(v); return String(parseFloat(v.toPrecision(12))); }
  const NUMRE = /^[-+]?(\d+\.?\d*|\.\d+)(e[+-]?\d+)?$/i;
  function num(v) {
    if (isErr(v)) return v;
    if (v === null || v === undefined) return 0;
    if (typeof v === 'boolean') return v ? 1 : 0;
    if (typeof v === 'number') return v;
    if (v instanceof Lambda) return ERR('#VALUE!');
    const s = String(v).trim();
    if (s !== '' && NUMRE.test(s)) return parseFloat(s);
    const d = parseDateText(s); if (d !== null) return d;
    return ERR('#VALUE!');
  }
  function str(v) { if (v === null || v === undefined) return ''; if (typeof v === 'boolean') return v ? 'TRUE' : 'FALSE'; if (typeof v === 'number') return fmtNum(v); return String(v); }
  function truth(v) {
    if (isErr(v)) return v;
    if (v === null || v === undefined) return false;
    if (typeof v === 'boolean') return v;
    if (typeof v === 'number') return v !== 0;
    const s = String(v).toUpperCase();
    if (s === 'TRUE') return true; if (s === 'FALSE') return false;
    return ERR('#VALUE!');
  }
  const rank = v => (typeof v === 'number' ? 0 : typeof v === 'string' ? 1 : 2);
  function cmpVals(a, b) {
    if (a === null) a = typeof b === 'string' ? '' : typeof b === 'boolean' ? false : 0;
    if (b === null) b = typeof a === 'string' ? '' : typeof a === 'boolean' ? false : 0;
    const ra = rank(a), rb = rank(b);
    if (ra !== rb) return ra - rb;
    if (ra === 0) return Math.abs(a - b) < 1e-9 ? 0 : a < b ? -1 : 1;
    if (ra === 1) { const x = a.toLowerCase(), y = b.toLowerCase(); return x === y ? 0 : x < y ? -1 : 1; }
    return (a ? 1 : 0) - (b ? 1 : 0);
  }

  /* ---------- 日期 ---------- */
  const EPOCH = Date.UTC(1899, 11, 30);
  const toSerial = (y, m, d) => Math.round((Date.UTC(y, m - 1, d) - EPOCH) / 864e5);
  const fromSerial = n => { const dt = new Date(EPOCH + Math.floor(n) * 864e5); return { y: dt.getUTCFullYear(), m: dt.getUTCMonth() + 1, d: dt.getUTCDate(), wd: dt.getUTCDay() }; };
  function parseDateText(s) {
    let m = /^(\d{4})[-/.](\d{1,2})[-/.](\d{1,2})(?:[ T](\d{1,2}):(\d{2})(?::(\d{2}))?)?$/.exec(s);
    let y, mo, d, hh = 0, mi = 0, ss = 0;
    if (m) { [y, mo, d] = [+m[1], +m[2], +m[3]]; hh = +(m[4] || 0); mi = +(m[5] || 0); ss = +(m[6] || 0); }
    else {
      m = /^(\d{1,2})[/.](\d{1,2})[/.](\d{4})(?: (\d{1,2}):(\d{2})(?::(\d{2}))?)?$/.exec(s);
      if (m) { [d, mo, y] = [+m[1], +m[2], +m[3]]; hh = +(m[4] || 0); mi = +(m[5] || 0); ss = +(m[6] || 0); }
      else {
        m = /^(\d{1,2}):(\d{2})(?::(\d{2}))?$/.exec(s);
        if (m) return (+m[1] * 3600 + +m[2] * 60 + +(m[3] || 0)) / 86400;
        return null;
      }
    }
    if (mo < 1 || mo > 12 || d < 1 || d > 31) return null;
    return toSerial(y, mo, d) + (hh * 3600 + mi * 60 + ss) / 86400;
  }
  const pad = (n, w = 2) => String(n).padStart(w, '0');
  function fmtDate(n, kind) {
    if (typeof n !== 'number') return null;
    const { y, m, d } = fromSerial(n);
    const secs = Math.round((n - Math.floor(n)) * 86400);
    const hh = Math.floor(secs / 3600) % 24, mi = Math.floor(secs / 60) % 60;
    if (kind === 'date') return `${y}-${pad(m)}-${pad(d)}`;
    if (kind === 'time') return `${pad(hh)}:${pad(mi)}`;
    if (kind === 'datetime') return `${y}-${pad(m)}-${pad(d)} ${pad(hh)}:${pad(mi)}`;
    return null;
  }
  function weekendMask(w) {
    if (w === null || w === undefined) return '0000011';
    if (typeof w === 'string' && /^[01]{7}$/.test(w)) return w;
    const codes = { 1: '0000011', 2: '1000001', 3: '1100000', 4: '0110000', 5: '0011000', 6: '0001100', 7: '0000110', 11: '0000001', 12: '1000000', 13: '0100000', 14: '0010000', 15: '0001000', 16: '0000100', 17: '0000010' };
    const n = num(w); if (!isErr(n) && codes[n]) return codes[n];
    return null;
  }
  const isoDay = n => (fromSerial(n).wd + 6) % 7; // 周一=0 … 周日=6

  /* ---------- 广播 ---------- */
  function shape(vals) { let h = 1, w = 1; for (const v of vals) if (isArr(v)) { h = Math.max(h, v.h); w = Math.max(w, v.w); } return [h, w]; }
  const at = (v, r, c) => { if (!isArr(v)) return v; const rr = v.h === 1 ? 0 : r, cc = v.w === 1 ? 0 : c; return v.rows[rr] && v.rows[rr][cc] !== undefined ? v.rows[rr][cc] : ERR('#N/A'); };
  function lift(vals, fn) {
    if (!vals.some(isArr)) return fn(vals);
    const [h, w] = shape(vals); const rows = [];
    for (let r = 0; r < h; r++) { const row = []; for (let c = 0; c < w; c++) { let v = fn(vals.map(x => at(x, r, c))); if (isArr(v)) v = sc(v); row.push(v); } rows.push(row); }
    return new Arr(rows);
  }
  function bin(op, a, b) {
    return lift([a, b], ([a, b]) => {
      if (isErr(a)) return a; if (isErr(b)) return b;
      if (op === '&') return str(a) + str(b);
      if ('+-*/^'.includes(op)) {
        const x = num(a), y = num(b);
        if (isErr(x)) return x; if (isErr(y)) return y;
        if (op === '+') return x + y; if (op === '-') return x - y; if (op === '*') return x * y;
        if (op === '/') return y === 0 ? ERR('#DIV/0!') : x / y;
        return Math.pow(x, y);
      }
      const c = cmpVals(a, b);
      return op === '=' ? c === 0 : op === '<>' ? c !== 0 : op === '<' ? c < 0 : op === '>' ? c > 0 : op === '<=' ? c <= 0 : c >= 0;
    });
  }

  /* ---------- 条件 ---------- */
  function makeCrit(c) {
    c = sc(c);
    if (isErr(c)) return () => false;
    if (typeof c === 'number') return v => typeof v === 'number' && Math.abs(v - c) < 1e-9;
    if (typeof c === 'boolean') return v => v === c;
    let s = c === null ? '' : String(c), op = '=';
    const m = /^(<=|>=|<>|=|<|>)/.exec(s);
    if (m) { op = m[1]; s = s.slice(op.length); }
    const n = s.trim() !== '' && !isNaN(+s) ? +s : null;
    if (n !== null) {
      return v => {
        if (typeof v !== 'number') { if (typeof v === 'string' && NUMRE.test(v.trim())) v = +v; else return op === '<>'; }
        const d = Math.abs(v - n) < 1e-9 ? 0 : v < n ? -1 : 1;
        return op === '=' ? d === 0 : op === '<>' ? d !== 0 : op === '<' ? d < 0 : op === '>' ? d > 0 : op === '<=' ? d <= 0 : d >= 0;
      };
    }
    if (s === '') { if (op === '=') return v => v === null || v === ''; if (op === '<>') return v => !(v === null || v === ''); }
    if (op === '=' || op === '<>') {
      const re = new RegExp('^' + s.replace(/[.+^${}()|[\]\\]/g, '\\$&').replace(/\*/g, '.*').replace(/\?/g, '.') + '$', 'i');
      return op === '=' ? (v => typeof v === 'string' && re.test(v)) : (v => !(typeof v === 'string' && re.test(v)));
    }
    return v => { if (typeof v !== 'string') return false; const d = cmpVals(v, s); return op === '<' ? d < 0 : op === '>' ? d > 0 : op === '<=' ? d <= 0 : d >= 0; };
  }

  /* ---------- 函数库 ---------- */
  const FNS = {};
  function def(name, min, max, fn, syntax, liftIdx) {
    FNS[name] = { min, max, syntax, fn: liftIdx ? (a, ctx) => {
      const idx = liftIdx === 'all' ? a.map((_, i) => i) : liftIdx;
      if (!idx.some(i => isArr(a[i]))) return fn(a, ctx);
      const [h, w] = shape(idx.map(i => a[i])); const rows = [];
      for (let r = 0; r < h; r++) { const row = []; for (let c = 0; c < w; c++) { const b = a.slice(); idx.forEach(i => { if (i < b.length) b[i] = at(a[i], r, c); }); let v = fn(b, ctx); if (isArr(v)) v = sc(v); row.push(v); } rows.push(row); }
      return new Arr(rows);
    } : fn };
  }
  const val = x => (x === null ? 0 : x);
  const asArr = v => (isArr(v) ? v : new Arr([[v]]));
  const needArr = (fn, i, x, eg) => { if (!isArr(x)) throw new FE(`${fn} 的第 ${i} 个参数要是一个区域，例如 ${eg}。`); return x; };
  function numsOf(args) {
    const out = [];
    for (const a of args) {
      if (isArr(a)) { for (const v of a.flat()) { if (isErr(v)) return v; if (typeof v === 'number') out.push(v); } }
      else { if (a === null) continue; const n = num(a); if (isErr(n)) return n; out.push(n); }
    }
    return out;
  }
  function boolsOf(args) {
    const out = [];
    for (const a of args) {
      if (isArr(a)) { for (const v of a.flat()) { if (isErr(v)) return v; if (typeof v === 'boolean' || typeof v === 'number') out.push(!!v); } }
      else { const b = truth(a); if (isErr(b)) return b; out.push(b); }
    }
    return out;
  }
  const lookEq = (v, look) => {
    if (typeof look === 'string' && typeof v === 'string') return v.toLowerCase() === look.toLowerCase();
    if (typeof look === 'number' && typeof v === 'number') return Math.abs(v - look) < 1e-9;
    if (typeof look === 'boolean') return v === look;
    return false;
  };
  function approxIndex(list, look) {
    let best = -1;
    for (let i = 0; i < list.length; i++) { const v = list[i]; if (v === null || rank(v) !== rank(look)) continue; if (cmpVals(v, look) <= 0) best = i; else break; }
    return best;
  }
  const vec = a => (a.h === 1 ? a.rows[0] : a.rows.map(r => r[0]));

  def('IF', 2, 3, a => { const c = truth(a[0]); if (isErr(c)) return c; return c ? val(a[1]) : (a.length > 2 ? val(a[2]) : false); }, 'IF(条件, 成立时显示, 不成立时显示)', 'all');
  def('IFERROR', 2, 2, a => (isErr(a[0]) ? val(a[1]) : val(a[0])), 'IFERROR(原公式, 出错时显示)', 'all');
  def('IFNA', 2, 2, a => (isErr(a[0]) && a[0].code === '#N/A' ? val(a[1]) : val(a[0])), 'IFNA(原公式, 找不到时显示)', 'all');
  def('ISNUMBER', 1, 1, a => typeof a[0] === 'number', 'ISNUMBER(值)', 'all');
  def('ISTEXT', 1, 1, a => typeof a[0] === 'string', 'ISTEXT(值)', 'all');
  def('ISBLANK', 1, 1, a => a[0] === null, 'ISBLANK(单元格)', 'all');
  def('ISERROR', 1, 1, a => isErr(a[0]), 'ISERROR(值)', 'all');
  def('ISNA', 1, 1, a => isErr(a[0]) && a[0].code === '#N/A', 'ISNA(值)', 'all');
  def('N', 1, 1, a => (typeof a[0] === 'number' ? a[0] : typeof a[0] === 'boolean' ? (a[0] ? 1 : 0) : isErr(a[0]) ? a[0] : 0), 'N(值)', 'all');
  def('AND', 1, 30, a => { const b = boolsOf(a); return isErr(b) ? b : b.every(Boolean); }, 'AND(条件1, 条件2, …)');
  def('OR', 1, 30, a => { const b = boolsOf(a); return isErr(b) ? b : b.some(Boolean); }, 'OR(条件1, 条件2, …)');
  def('NOT', 1, 1, a => { const b = truth(a[0]); return isErr(b) ? b : !b; }, 'NOT(条件)', 'all');
  const N1 = (name, f, syn) => def(name, 1, 1, a => { const n = num(a[0]); return isErr(n) ? n : f(n); }, syn, 'all');
  N1('ABS', Math.abs, 'ABS(数字)');
  N1('INT', Math.floor, 'INT(数字)');
  def('MOD', 2, 2, a => { const n = num(a[0]), d = num(a[1]); if (isErr(n)) return n; if (isErr(d)) return d; if (d === 0) return ERR('#DIV/0!'); const r = n - d * Math.floor(n / d); return Math.abs(r) < 1e-12 ? 0 : r; }, 'MOD(数字, 除数)', 'all');
  const roundTo = (n, d, mode) => { const f = Math.pow(10, d), x = n * f; const r = mode === 'up' ? Math.sign(x) * Math.ceil(Math.abs(x) - 1e-9) : mode === 'down' ? Math.sign(x) * Math.floor(Math.abs(x) + 1e-9) : Math.sign(x) * Math.round(Math.abs(x) + 1e-9); return r / f; };
  for (const [nm, mode] of [['ROUND', 'n'], ['ROUNDUP', 'up'], ['ROUNDDOWN', 'down']]) def(nm, 1, 2, a => { const n = num(a[0]), d = a.length > 1 ? num(a[1]) : 0; if (isErr(n)) return n; if (isErr(d)) return d; return roundTo(n, d, mode); }, `${nm}(数字, 保留几位小数)`, 'all');
  def('SUM', 1, 30, a => { const n = numsOf(a); return isErr(n) ? n : n.reduce((x, y) => x + y, 0); }, 'SUM(区域)');
  def('AVERAGE', 1, 30, a => { const n = numsOf(a); if (isErr(n)) return n; return n.length ? n.reduce((x, y) => x + y, 0) / n.length : ERR('#DIV/0!'); }, 'AVERAGE(区域)');
  def('MAX', 1, 30, a => { const n = numsOf(a); return isErr(n) ? n : n.length ? Math.max(...n) : 0; }, 'MAX(数字1, 数字2 …)');
  def('MIN', 1, 30, a => { const n = numsOf(a); return isErr(n) ? n : n.length ? Math.min(...n) : 0; }, 'MIN(数字1, 数字2 …)');
  def('COUNT', 1, 30, a => { let c = 0; for (const x of a) { if (isArr(x)) c += x.flat().filter(v => typeof v === 'number').length; else if (typeof x === 'number') c++; } return c; }, 'COUNT(区域)');
  def('COUNTA', 1, 30, a => { let c = 0; for (const x of a) { if (isArr(x)) c += x.flat().filter(v => v !== null && v !== '').length; else if (x !== null) c++; } return c; }, 'COUNTA(区域)');
  def('COUNTBLANK', 1, 1, a => needArr('COUNTBLANK', 1, a[0], 'A2:A10').flat().filter(v => v === null || v === '').length, 'COUNTBLANK(区域)');
  def('COUNTIF', 2, 2, a => { const r = needArr('COUNTIF', 1, a[0], '$D$2:$D$10').flat(); return lift([a[1]], ([c]) => { const f = makeCrit(c); return r.filter(f).length; }); }, 'COUNTIF(在哪个范围数, 数什么)');
  def('COUNTIFS', 2, 30, a => {
    if (a.length % 2) throw new FE('COUNTIFS 的参数要成对出现：区域1, 条件1, 区域2, 条件2 …');
    const pairs = []; for (let i = 0; i < a.length; i += 2) pairs.push([needArr('COUNTIFS', i + 1, a[i], 'A2:A10').flat(), makeCrit(a[i + 1])]);
    const n = pairs[0][0].length; if (pairs.some(p => p[0].length !== n)) return ERR('#VALUE!');
    let c = 0; for (let i = 0; i < n; i++) if (pairs.every(([r, f]) => f(r[i]))) c++; return c;
  }, 'COUNTIFS(区域1, 条件1, 区域2, 条件2, …)');
  def('SUMIF', 2, 3, a => {
    const r = needArr('SUMIF', 1, a[0], 'A2:A10').flat(); const f = makeCrit(a[1]);
    const s = a.length > 2 ? needArr('SUMIF', 3, a[2], 'B2:B10').flat() : r;
    let t = 0; r.forEach((v, i) => { if (f(v) && typeof s[i] === 'number') t += s[i]; }); return t;
  }, 'SUMIF(条件区域, 条件, 求和区域)');
  def('SUMIFS', 3, 31, a => {
    const s = needArr('SUMIFS', 1, a[0], 'B2:B10').flat();
    if ((a.length - 1) % 2) throw new FE('SUMIFS 的参数：求和区域, 条件区域1, 条件1, 条件区域2, 条件2 …');
    const pairs = []; for (let i = 1; i < a.length; i += 2) pairs.push([needArr('SUMIFS', i + 1, a[i], 'A2:A10').flat(), makeCrit(a[i + 1])]);
    if (pairs.some(p => p[0].length !== s.length)) return ERR('#VALUE!');
    let t = 0; for (let i = 0; i < s.length; i++) if (pairs.every(([r, f]) => f(r[i])) && typeof s[i] === 'number') t += s[i]; return t;
  }, 'SUMIFS(求和区域, 条件区域1, 条件1, …)');
  def('SUMPRODUCT', 1, 30, a => {
    const arrs = a.map(asArr); const h = arrs[0].h, w = arrs[0].w;
    if (arrs.some(x => x.h !== h || x.w !== w)) return ERR('#VALUE!');
    let t = 0;
    for (let r = 0; r < h; r++) for (let c = 0; c < w; c++) {
      let prod = 1;
      for (const x of arrs) { const v = x.rows[r][c]; if (isErr(v)) return v; prod *= typeof v === 'number' ? v : 0; }
      t += prod;
    }
    return t;
  }, 'SUMPRODUCT(数组1, 数组2, …)');
  def('VLOOKUP', 3, 4, a => {
    const tb = needArr('VLOOKUP', 2, a[1], '$F$2:$G$10');
    return lift([a[0]], ([look]) => {
      if (isErr(look)) return look;
      const ci = num(sc(a[2])); if (isErr(ci)) return ci;
      if (ci < 1) return ERR('#VALUE!'); if (ci > tb.w) return ERR('#REF!');
      let exact = false;
      if (a.length > 3) { const v = sc(a[3]); if (v === null) exact = true; else { const b = truth(v); if (isErr(b)) return b; exact = !b; } }
      const first = tb.rows.map(r => r[0]);
      const idx = exact ? first.findIndex(v => lookEq(v, look)) : approxIndex(first, look);
      return idx < 0 ? ERR('#N/A') : tb.rows[idx][Math.floor(ci) - 1];
    });
  }, 'VLOOKUP(找谁, 在哪个区域找, 带回第几列, 0)');
  def('MATCH', 2, 3, a => {
    const r = needArr('MATCH', 2, a[1], '$F$2:$F$10');
    if (r.h > 1 && r.w > 1) return ERR('#N/A');
    const list = vec(r);
    const type = a.length > 2 ? num(sc(a[2])) : 1; if (isErr(type)) return type;
    return lift([a[0]], ([look]) => {
      if (isErr(look)) return look;
      let idx;
      if (type === 0) idx = list.findIndex(v => (typeof look === 'string' && typeof v === 'string') ? makeCrit(look)(v) : lookEq(v, look));
      else if (type > 0) idx = approxIndex(list, look);
      else { idx = -1; for (let i = 0; i < list.length; i++) { const v = list[i]; if (v === null || rank(v) !== rank(look)) continue; if (cmpVals(v, look) >= 0) idx = i; else break; } }
      return idx < 0 ? ERR('#N/A') : idx + 1;
    });
  }, 'MATCH(找谁, 在哪一列找, 0)');
  def('INDEX', 2, 3, a => {
    const r = needArr('INDEX', 1, a[0], '$G$2:$G$10');
    return lift([a[1], a[2] === undefined ? null : a[2]], ([ri0, ci0]) => {
      let ri = num(ri0); if (isErr(ri)) return ri;
      let ci = a.length > 2 && ci0 !== null ? num(ci0) : null; if (isErr(ci)) return ci;
      if (ci === null) { if (r.w === 1) ci = 1; else if (r.h === 1) { ci = ri; ri = 1; } else ci = 1; }
      ri = Math.floor(ri); ci = Math.floor(ci);
      if (ri < 1 || ci < 1 || ri > r.h || ci > r.w) return ERR('#REF!');
      return r.rows[ri - 1][ci - 1];
    });
  }, 'INDEX(区域, 第几行, [第几列])');
  def('XLOOKUP', 3, 6, a => {
    const la = vec(needArr('XLOOKUP', 2, a[1], '$F$2:$F$10'));
    const ra = needArr('XLOOKUP', 3, a[2], '$G$2:$G$10');
    const rv = vec(ra);
    if (la.length !== rv.length) return ERR('#VALUE!');
    return lift([a[0]], ([look]) => {
      if (isErr(look)) return look;
      const idx = la.findIndex(v => lookEq(v, look));
      if (idx < 0) return a.length > 3 && a[3] !== null ? sc(a[3]) : ERR('#N/A');
      return rv[idx];
    });
  }, 'XLOOKUP(找谁, 在哪列找, 带回哪列, 找不到时显示)');
  def('FILTER', 2, 30, a => {
    const src = asArr(a[0]);
    const conds = []; let ifEmpty;
    for (let i = 1; i < a.length; i++) {
      const c = a[i];
      if (!isArr(c) && i === 2 && a.length === 3 && typeof c !== 'boolean' && typeof c !== 'number') { ifEmpty = c; continue; }
      conds.push(asArr(c));
    }
    const byRow = conds.every(c => c.h === src.h && (c.w === 1 || src.h > 1));
    const keep = [];
    if (byRow || src.w === 1) {
      for (let r = 0; r < src.h; r++) {
        let ok = true;
        for (const c of conds) { const v = c.rows[r] ? c.rows[r][0] : null; if (isErr(v)) return v; const b = truth(v); if (isErr(b) || !b) { ok = false; break; } }
        if (ok) keep.push(src.rows[r]);
      }
      if (!keep.length) return ifEmpty !== undefined ? ifEmpty : ERR('#N/A');
      return new Arr(keep);
    }
    const cols = [];
    for (let c = 0; c < src.w; c++) { if (conds.every(cd => truth(cd.rows[0][c]) === true)) cols.push(c); }
    if (!cols.length) return ifEmpty !== undefined ? ifEmpty : ERR('#N/A');
    return new Arr(src.rows.map(r => cols.map(c => r[c])));
  }, 'FILTER(要返回的区域, 条件1, 条件2, …)');
  def('UNIQUE', 1, 3, a => {
    const src = asArr(a[0]); const seen = new Set(); const out = [];
    for (const row of src.rows) { const key = JSON.stringify(row.map(v => (typeof v === 'string' ? v.toLowerCase() : v))); if (!seen.has(key)) { seen.add(key); out.push(row); } }
    return new Arr(out);
  }, 'UNIQUE(区域)');
  def('ROWS', 1, 1, a => (isErr(a[0]) ? a[0] : asArr(a[0]).h), 'ROWS(区域)');
  def('COLUMNS', 1, 1, a => (isErr(a[0]) ? a[0] : asArr(a[0]).w), 'COLUMNS(区域)');
  def('ARRAYFORMULA', 1, 1, a => a[0], 'ARRAYFORMULA(公式)');
  def('TEXTJOIN', 3, 60, a => {
    const d = str(sc(a[0])); const ign = truth(sc(a[1])); const parts = [];
    for (const x of a.slice(2)) { for (const v of (isArr(x) ? x.flat() : [x])) { if (isErr(v)) return v; if (ign === true && (v === null || v === '')) continue; parts.push(str(v)); } }
    return parts.join(d);
  }, 'TEXTJOIN(分隔符, 是否忽略空白, 文字1, …)');
  const S1 = (name, f, syn) => def(name, 1, 1, a => (isErr(a[0]) ? a[0] : f(str(a[0]))), syn, 'all');
  S1('TRIM', s => s.replace(/ +/g, ' ').trim(), 'TRIM(文字)');
  S1('LOWER', s => s.toLowerCase(), 'LOWER(文字)');
  S1('UPPER', s => s.toUpperCase(), 'UPPER(文字)');
  S1('LEN', s => s.length, 'LEN(文字)');
  def('LEFT', 1, 2, a => { if (isErr(a[0])) return a[0]; const n = a.length > 1 ? num(a[1]) : 1; return isErr(n) ? n : str(a[0]).slice(0, n); }, 'LEFT(文字, 取几个字)', 'all');
  def('RIGHT', 1, 2, a => { if (isErr(a[0])) return a[0]; const n = a.length > 1 ? num(a[1]) : 1; if (isErr(n)) return n; return n <= 0 ? '' : str(a[0]).slice(-n); }, 'RIGHT(文字, 取几个字)', 'all');
  def('MID', 3, 3, a => { if (isErr(a[0])) return a[0]; const s0 = num(a[1]), n = num(a[2]); if (isErr(s0)) return s0; if (isErr(n)) return n; if (s0 < 1) return ERR('#VALUE!'); return str(a[0]).substr(s0 - 1, n); }, 'MID(文字, 从第几个开始, 取几个)', 'all');
  def('EXACT', 2, 2, a => { if (isErr(a[0])) return a[0]; if (isErr(a[1])) return a[1]; return str(a[0]) === str(a[1]); }, 'EXACT(文字1, 文字2)', 'all');
  def('CONCATENATE', 1, 30, a => { let s = ''; for (const v of a) { if (isErr(v)) return v; s += str(v); } return s; }, 'CONCATENATE(文字1, 文字2, …)', 'all');
  def('CONCAT', 1, 30, a => { let s = ''; for (const x of a) for (const v of (isArr(x) ? x.flat() : [x])) { if (isErr(v)) return v; s += str(v); } return s; }, 'CONCAT(文字或区域, …)');
  def('SUBSTITUTE', 3, 3, a => { for (const x of a) if (isErr(x)) return x; const so = str(a[1]); return so === '' ? str(a[0]) : str(a[0]).split(so).join(str(a[2])); }, 'SUBSTITUTE(文字, 旧内容, 新内容)', 'all');
  def('FIND', 2, 3, a => { const i = str(a[1]).indexOf(str(a[0])); return i < 0 ? ERR('#VALUE!') : i + 1; }, 'FIND(找什么, 在哪段文字里找)', [0, 1]);
  def('SEARCH', 2, 3, a => { const i = str(a[1]).toLowerCase().indexOf(str(a[0]).toLowerCase()); return i < 0 ? ERR('#VALUE!') : i + 1; }, 'SEARCH(找什么, 在哪段文字里找)', [0, 1]);
  def('VALUE', 1, 1, a => { if (isErr(a[0])) return a[0]; if (typeof a[0] === 'number') return a[0]; const s = str(a[0]).trim(); if (NUMRE.test(s)) return parseFloat(s); const d = parseDateText(s); return d === null ? ERR('#VALUE!') : d; }, 'VALUE(文本)', 'all');
  def('TEXT', 2, 2, a => {
    const v = num(a[0]), f = str(a[1]); if (isErr(v)) return v;
    if (/[yd]|h|s/i.test(f.replace(/"[^"]*"/g, ''))) {
      const { y, m, d } = fromSerial(v); const secs = Math.round((v - Math.floor(v)) * 86400);
      const hh = Math.floor(secs / 3600) % 24, mi = Math.floor(secs / 60) % 60, ss = secs % 60;
      let out = '', i = 0, lastH = false;
      while (i < f.length) {
        const rest = f.slice(i);
        let m2;
        if ((m2 = /^yyyy/i.exec(rest))) { out += y; i += 4; continue; }
        if ((m2 = /^yy/i.exec(rest))) { out += pad(y % 100); i += 2; continue; }
        if ((m2 = /^dd/i.exec(rest))) { out += pad(d); i += 2; lastH = false; continue; }
        if ((m2 = /^d/i.exec(rest))) { out += d; i += 1; lastH = false; continue; }
        if ((m2 = /^hh/i.exec(rest))) { out += pad(hh); i += 2; lastH = true; continue; }
        if ((m2 = /^h/i.exec(rest))) { out += hh; i += 1; lastH = true; continue; }
        if ((m2 = /^ss/i.exec(rest))) { out += pad(ss); i += 2; continue; }
        if ((m2 = /^mm/i.exec(rest))) { out += lastH ? pad(mi) : pad(m); i += 2; lastH = false; continue; }
        if ((m2 = /^m/i.exec(rest))) { out += lastH ? mi : m; i += 1; lastH = false; continue; }
        if (f[i] === '"') { const j = f.indexOf('"', i + 1); out += f.slice(i + 1, j < 0 ? f.length : j); i = j < 0 ? f.length : j + 1; continue; }
        out += f[i]; i++;
      }
      return out;
    }
    const pct = f.includes('%'); const intPart = f.split('.')[0].replace(/[^0]/g, '').length; const dec = (f.split('.')[1] || '').replace(/[^0#]/g, '').length;
    let x = pct ? v * 100 : v; let s = Math.abs(x).toFixed(dec);
    let [ip, dp] = s.split('.'); if (ip.length < intPart) ip = ip.padStart(intPart, '0');
    if (f.includes(',')) ip = ip.replace(/\B(?=(\d{3})+(?!\d))/g, ',');
    s = (x < 0 ? '-' : '') + ip + (dp ? '.' + dp : '');
    return pct ? s + '%' : s;
  }, 'TEXT(数字或日期, "格式")', [0]);
  def('DATE', 3, 3, a => { const [y, m, d] = a.map(num); for (const x of [y, m, d]) if (isErr(x)) return x; return Math.round((Date.UTC(y, m - 1, d) - EPOCH) / 864e5); }, 'DATE(年, 月, 日)', 'all');
  def('TIME', 3, 3, a => { const [h, m, s] = a.map(num); for (const x of [h, m, s]) if (isErr(x)) return x; const t = (h * 3600 + m * 60 + s) / 86400; return t - Math.floor(t); }, 'TIME(时, 分, 秒)', 'all');
  def('WEEKDAY', 1, 2, a => { const n = num(a[0]); if (isErr(n)) return n; const ty = a.length > 1 && a[1] !== null ? num(a[1]) : 1; const wd = fromSerial(n).wd; if (ty === 2) return wd === 0 ? 7 : wd; if (ty === 3) return (wd + 6) % 7; return wd + 1; }, 'WEEKDAY(日期, 2)', 'all');
  function holidaySet(h) { const s = new Set(); if (h === undefined || h === null) return s; for (const v of (isArr(h) ? h.flat() : [h])) if (typeof v === 'number') s.add(Math.floor(v)); return s; }
  def('WORKDAY.INTL', 2, 4, a => {
    let d = num(sc(a[0])), n = num(sc(a[1])); if (isErr(d)) return d; if (isErr(n)) return n;
    const mask = weekendMask(a.length > 2 ? sc(a[2]) : null); if (!mask) return ERR('#VALUE!');
    const hs = holidaySet(a[3]); d = Math.floor(d); n = Math.trunc(n);
    const step = n >= 0 ? 1 : -1; let left = Math.abs(n);
    while (left > 0) { d += step; if (mask[isoDay(d)] === '0' && !hs.has(d)) left--; }
    return d;
  }, 'WORKDAY.INTL(开始日期, 往后几个工作日, "0000011", 节假日)');
  def('NETWORKDAYS.INTL', 2, 4, a => {
    let s = num(sc(a[0])), e = num(sc(a[1])); if (isErr(s)) return s; if (isErr(e)) return e;
    const mask = weekendMask(a.length > 2 ? sc(a[2]) : null); if (!mask) return ERR('#VALUE!');
    const hs = holidaySet(a[3]); s = Math.floor(s); e = Math.floor(e);
    const sign = e >= s ? 1 : -1; let c = 0;
    for (let d = Math.min(s, e); d <= Math.max(s, e); d++) if (mask[isoDay(d)] === '0' && !hs.has(d)) c++;
    return c * sign;
  }, 'NETWORKDAYS.INTL(开始日期, 结束日期, "0000011", 节假日)');
  def('WORKDAY', 2, 3, a => FNS['WORKDAY.INTL'].fn([a[0], a[1], '0000011', a[2]]), 'WORKDAY(开始日期, 天数, 节假日)');
  def('NETWORKDAYS', 2, 3, a => FNS['NETWORKDAYS.INTL'].fn([a[0], a[1], '0000011', a[2]]), 'NETWORKDAYS(开始日期, 结束日期, 节假日)');
  def('INDIRECT', 1, 2, (a, ctx) => {
    const t = str(sc(a[0]));
    let parsed; try { parsed = parse(t); } catch (e) { return ERR('#REF!'); }
    if (!['ref', 'range', 'colrange', 'rows'].includes(parsed.ast.t)) return ERR('#REF!');
    if ((parsed.ast.sheet || null) && !ctx.book.has(parsed.ast.sheet)) return ERR('#REF!');
    return ctx.evalNode(parsed.ast, { dr: 0, dc: 0 });
  }, 'INDIRECT("地址文字")');

  /* ---------- 求值 ---------- */
  // book: { sheets: Map(name -> Map(addr -> value)), main, has(name), get(sheet, col, row), size(sheet) }
  function evaluate(ast, book, opt) {
    const dr = opt.dr || 0, dc = opt.dc || 0;
    const sheetOf = n => n || book.main;
    const R = (r, lock) => ((r.ar || lock) ? r.row : r.row + dr);
    const C = (r, lock) => ((r.ac || lock) ? r.col : r.col + dc);
    function cellsOf(sheet, c1, r1, c2, r2) {
      if (!book.has(sheet)) throw new FE(`没有叫「${sheet}」的工作表。检查一下表名，或者看看下方的表格标签。`);
      if (r1 < 1 || c1 < 1) return ERR('#REF!');
      const rows = [];
      for (let r = r1; r <= r2; r++) { const row = []; for (let c = c1; c <= c2; c++) row.push(book.get(sheet, c, r)); rows.push(row); }
      return new Arr(rows);
    }
    function ev(n, scope, o) {
      const lockR = o && o.lockRanges;
      const d0 = o && o.abs;
      const rr = (k, lock) => (d0 ? k.row : R(k, lock)), cc = (k, lock) => (d0 ? k.col : C(k, lock));
      switch (n.t) {
        case 'num': case 'str': case 'bool': return n.v;
        case 'empty': return null;
        case 'name': {
          for (let s = scope; s; s = s.parent) if (Object.prototype.hasOwnProperty.call(s.vars, n.v.toLowerCase())) return s.vars[n.v.toLowerCase()];
          throw new FE(`不认识「${n.v}」。如果它是文字，要加英文双引号："${n.v}"；如果是 LET 里定义的名字，检查一下拼写。`);
        }
        case 'ref': {
          const r = rr(n.r, false), c = cc(n.r, false);
          if (r < 1 || c < 1) return ERR('#REF!');
          const sh = sheetOf(n.sheet); if (!book.has(sh)) throw new FE(`没有叫「${sh}」的工作表。`);
          return book.get(sh, c, r);
        }
        case 'range': {
          const sh = sheetOf(n.sheet);
          const r1 = rr(n.a, lockR), c1 = cc(n.a, lockR), c2 = cc(n.b, lockR);
          const r2 = n.b.row === null ? Math.max(book.size(sh).rows, r1) : rr(n.b, lockR);
          return cellsOf(sh, Math.min(c1, c2), Math.min(r1, r2), Math.max(c1, c2), Math.max(r1, r2));
        }
        case 'colrange': { const sh = sheetOf(n.sheet); return cellsOf(sh, Math.min(n.a.col, n.b.col), 1, Math.max(n.a.col, n.b.col), Math.max(1, book.size(sh).rows)); }
        case 'rows': { const sh = sheetOf(n.k.sheet); return cellsOf(sh, 1, Math.min(n.k.r1, n.k.r2), Math.max(1, book.size(sh).cols), Math.max(n.k.r1, n.k.r2)); }
        case 'arr': {
          const blocks = n.rows.map(row => row.map(e => asArr(ev(e, scope, o))));
          const outRows = [];
          for (const row of blocks) {
            const h = Math.max(...row.map(b => b.h));
            for (let r = 0; r < h; r++) outRows.push(row.flatMap(b => (b.rows[r] || new Array(b.w).fill(ERR('#N/A')))));
          }
          return new Arr(outRows);
        }
        case 'neg': return lift([ev(n.e, scope, o)], ([v]) => { const x = num(v); return isErr(x) ? x : -x; });
        case 'pct': return lift([ev(n.e, scope, o)], ([v]) => { const x = num(v); return isErr(x) ? x : x / 100; });
        case 'bin': return bin(n.op, ev(n.l, scope, o), ev(n.r, scope, o));
        case 'call': {
          if (n.name === 'LET') {
            if (n.args.length < 3 || n.args.length % 2 === 0) throw new FE('LET 的写法：LET(名字1, 值1, 名字2, 值2, …, 最后要返回的计算)。参数个数应该是单数。');
            const s = { vars: {}, parent: scope };
            for (let i = 0; i < n.args.length - 1; i += 2) {
              const nm = n.args[i]; if (nm.t !== 'name') throw new FE('LET 里第 1、3、5… 个参数要是名字（不加引号），例如 LET(t, VALUE(A2), …)。');
              s.vars[nm.v.toLowerCase()] = ev(n.args[i + 1], s, o);
            }
            return ev(n.args[n.args.length - 1], s, o);
          }
          if (n.name === 'LAMBDA') {
            const params = n.args.slice(0, -1).map(x => { if (x.t !== 'name') throw new FE('LAMBDA 前面的参数要是名字，例如 LAMBDA(city, client, …)。'); return x.v.toLowerCase(); });
            return new Lambda(params, n.args[n.args.length - 1], scope);
          }
          if (n.name === 'MAP') {
            if (n.args.length < 2) throw new FE('MAP 的写法：MAP(数组1, …, LAMBDA(值1, …, 计算))。');
            const arrs = n.args.slice(0, -1).map(x => asArr(ev(x, scope, o)));
            const fnv = ev(n.args[n.args.length - 1], scope, o);
            if (!(fnv instanceof Lambda)) throw new FE('MAP 的最后一个参数要是 LAMBDA(…)。');
            if (fnv.params.length !== arrs.length) throw new FE(`MAP 输入了 ${arrs.length} 个数组，LAMBDA 就要定义 ${arrs.length} 个名字（现在是 ${fnv.params.length} 个）。`);
            const h = arrs[0].h, w = arrs[0].w; const rows = [];
            for (let r = 0; r < h; r++) { const row = []; for (let c = 0; c < w; c++) { const s = { vars: {}, parent: fnv.scope }; fnv.params.forEach((p, i) => (s.vars[p] = at(arrs[i], r, c))); let v = ev(fnv.body, s, o); if (isArr(v)) v = v.h === 1 && v.w === 1 ? v.rows[0][0] : ERR('#VALUE!'); row.push(v === null ? '' : v); } rows.push(row); }
            return new Arr(rows);
          }
          if (n.name === 'QUERY') throw new FE('QUERY 在这个网站里还不能运行（它需要单独的查询语言）。QUERY 相关的题目是「回忆题」，在题库里可以找到。');
          const f = FNS[n.name];
          if (!f) throw new FE(`不认识函数「${n.raw}」。检查一下拼写；如果拼写没问题，可能是这个网站还没收录它。`);
          if (n.args.length < f.min || n.args.length > f.max) throw new FE(`${n.raw} 的参数个数不对。它的写法是：${f.syntax}`);
          const args = n.args.map(x => ev(x, scope, o));
          if (args.some(x => x instanceof Lambda)) throw new FE('LAMBDA 只能放在 MAP 这类函数里使用。');
          return f.fn(args, { book, evalNode: (node, oo) => ev(node, scope, { ...o, abs: true }) });
        }
      }
      return ERR('#VALUE!');
    }
    return ev(ast, null, { lockRanges: opt.lockRanges });
  }
  function finalize(v) { if (isArr(v)) return v; if (v === null || v === undefined) return 0; if (v instanceof Lambda) return ERR('#VALUE!'); return v; }

  function makeBook(sheets, main) {
    const maps = new Map(); const sizes = {};
    for (const s of sheets) {
      const m = new Map(s.cells); maps.set(s.name, m);
    }
    const book = {
      sheets: maps, main: main || sheets[0].name,
      has: n => maps.has(n),
      get: (sh, c, r) => { const v = maps.get(sh).get(numToCol(c) + r); return v === undefined ? null : v; },
      set: (sh, addr, v) => { maps.get(sh).set(addr, v); delete sizes[sh]; },
      del: (sh, addr) => { maps.get(sh).delete(addr); delete sizes[sh]; },
      size: sh => {
        if (sizes[sh]) return sizes[sh];
        let rows = 1, cols = 1;
        for (const k of maps.get(sh).keys()) { const m = /^([A-Z]+)(\d+)$/.exec(k); if (m) { rows = Math.max(rows, +m[2]); cols = Math.max(cols, colToNum(m[1])); } }
        return (sizes[sh] = { rows, cols });
      },
      clone: () => makeBook([...maps].map(([name, cells]) => ({ name, cells })), main || sheets[0].name)
    };
    return book;
  }

  function shiftText(src, dr, dc) {
    const toks = tokenize(src); let out = '', last = 0;
    for (const k of toks) {
      if (k.t === 'ref' || k.t === 'col') {
        const start = k.e - (k.t === 'ref' ? (`${k.ac ? '$' : ''}${numToCol(k.col)}${k.ar ? '$' : ''}${k.row}`).length : (`${k.ac ? '$' : ''}${numToCol(k.col)}`).length);
        const col = k.ac ? k.col : k.col + (dc || 0);
        const txt = k.t === 'ref' ? `${k.ac ? '$' : ''}${numToCol(col)}${k.ar ? '$' : ''}${k.ar ? k.row : k.row + dr}` : `${k.ac ? '$' : ''}${numToCol(col)}`;
        out += src.slice(last, start) + txt; last = k.e;
      }
    }
    return out + src.slice(last);
  }
  function fnNamesIn(src) { try { return [...new Set(tokenize(src).filter(t => t.t === 'fn').map(t => t.v))]; } catch (e) { return []; } }
  function show(v, fmt) {
    if (isErr(v)) return v.code;
    if (typeof v === 'boolean') return v ? 'TRUE' : 'FALSE';
    if (typeof v === 'number') { const d = fmt && fmtDate(v, fmt); if (d) return d; if (fmt === 'hours') return fmtNum(Math.round(v * 100) / 100); return fmtNum(v); }
    if (v instanceof Lambda) return '#VALUE!';
    return v === null || v === undefined ? '' : String(v);
  }
  function sameVal(a, b) {
    if (isErr(a) || isErr(b)) return isErr(a) && isErr(b) && a.code === b.code;
    if (typeof a === 'number' && typeof b === 'number') return Math.abs(a - b) < 1e-6;
    if ((a === null || a === '') && (b === null || b === '')) return true;
    return a === b;
  }
  const D = s => parseDateText(s);

  G.XL = { XErr, ERR, isErr, FE, Arr, isArr, parse, tokenize, evaluate, finalize, makeBook, shiftText, fnNamesIn, show, sameVal, colToNum, numToCol, D, FNS, fmtDate };
})(typeof window !== 'undefined' ? window : globalThis);
