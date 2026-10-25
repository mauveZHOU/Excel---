/* Excel 公式练习区 · 判题 */
(function (G) {
  'use strict';
  const X = G.XL;
  const { isErr, isArr, FE } = X;

  function parseTarget(t) {
    const m = /^([A-Z]+)(\d+)(?::([A-Z]+)(\d+))?$/.exec(t);
    const c1 = X.colToNum(m[1]), r1 = +m[2], c2 = m[3] ? X.colToNum(m[3]) : c1, r2 = m[4] ? +m[4] : r1;
    const cells = [];
    for (let r = r1; r <= r2; r++) for (let c = c1; c <= c2; c++) cells.push({ addr: X.numToCol(c) + r, col: c, row: r, dr: r - r1, dc: c - c1 });
    return { c1, r1, c2, r2, cells, first: X.numToCol(c1) + r1, last: X.numToCol(c2) + r2 };
  }
  function baseBook(sc) { return X.makeBook(sc.sheets.map(s => ({ name: s.name, cells: new Map(Object.entries(s.cells)) })), sc.sheets[0].name); }
  const sheetOf = (sc, ex) => ex.sheet || sc.sheets[0].name;

  function spillCells(anchor, v) {
    const a = isArr(v) ? v : new X.Arr([[v]]); const out = [];
    a.rows.forEach((row, i) => row.forEach((x, j) => out.push({ addr: X.numToCol(anchor.col + j) + (anchor.row + i), col: anchor.col + j, row: anchor.row + i, v: x === null ? '' : x })));
    return out;
  }
  const TRUTH = new Map();
  function truth(sc) {
    if (TRUTH.has(sc.id)) return TRUTH.get(sc.id);
    const book = baseBook(sc); const cellsByEx = {};
    for (const ex of sc.exercises) {
      const sh = sheetOf(sc, ex); book.main = sh;
      const tg = parseTarget(ex.target);
      const { ast } = X.parse(ex.answer.replace(/^=/, ''));
      if (ex.spill) {
        const v = X.finalize(X.evaluate(ast, book, { dr: 0, dc: 0 }));
        const cells = spillCells(tg.cells[0], v);
        cells.forEach(c => book.set(sh, c.addr, c.v));
        cellsByEx[ex.id] = cells.map(c => ({ ...c, exp: c.v }));
      } else {
        cellsByEx[ex.id] = tg.cells.map(c => {
          let v = X.finalize(X.evaluate(ast, book, { dr: c.dr, dc: c.dc }));
          if (isArr(v)) v = v.h === 1 && v.w === 1 ? v.rows[0][0] : X.ERR('#VALUE!');
          book.set(sh, c.addr, v);
          return { ...c, exp: v };
        });
      }
    }
    book.main = sc.sheets[0].name;
    const t = { book, cellsByEx };
    TRUTH.set(sc.id, t);
    return t;
  }

  function grade(sc, ex, raw) {
    const out = { cells: [], ok: 0, total: 0, error: null, tips: [], solved: false, body: '' };
    const f = (raw || '').trim();
    const T = truth(sc); const expCells = T.cellsByEx[ex.id];
    out.total = expCells.length;
    if (!f) { out.error = '先在上面输入公式，再点「运行检查」。'; return out; }
    const sh = sheetOf(sc, ex);
    const book = T.book.clone(); book.main = sh;
    expCells.forEach(c => book.del(sh, c.addr));
    const tg = parseTarget(ex.target);
    if (!/^[=＝]/.test(f)) {
      out.literal = true;
      out.cells = expCells.map(c => ({ ...c, v: f, ok: false }));
      out.tips.push('公式要以 = 开头。不写 = 的话，表格会把你输入的内容当成普通文字原样显示。');
      return out;
    }
    const body = f.replace(/^[=＝]/, ''); out.body = body;
    let parsed;
    try { parsed = X.parse(body); } catch (e) { if (e instanceof FE) { out.error = e.message; return out; } throw e; }
    const acc = (v, exp) => { const a = ex.accept ? ex.accept(v, exp) : null; return a === null || a === undefined ? X.sameVal(v, exp) : a; };
    let spilledArr = false;
    const run = lock => {
      if (ex.spill) {
        const v = X.finalize(X.evaluate(parsed.ast, book, { dr: 0, dc: 0, lockRanges: lock }));
        const got = spillCells(tg.cells[0], v); const map = new Map(got.map(c => [c.addr, c.v]));
        const res = expCells.map(c => { const g = map.has(c.addr) ? map.get(c.addr) : null; return { ...c, v: g, ok: acc(g === null ? '' : g, c.exp) }; });
        const extra = got.filter(c => !expCells.some(e => e.addr === c.addr) && c.v !== '');
        extra.forEach(c => res.push({ ...c, exp: '', ok: false, extra: true }));
        return res;
      }
      return expCells.map(c => {
        let v = X.finalize(X.evaluate(parsed.ast, book, { dr: c.dr, dc: c.dc, lockRanges: lock }));
        if (isArr(v)) { if (v.h === 1 && v.w === 1) v = v.rows[0][0]; else { spilledArr = true; v = X.ERR('#VALUE!'); } }
        return { ...c, v, ok: acc(v, c.exp) };
      });
    };
    try { out.cells = run(false); } catch (e) { if (e instanceof FE) { out.error = e.message; return out; } throw e; }
    out.total = out.cells.length;
    out.ok = out.cells.filter(c => c.ok).length;
    const used = X.fnNamesIn(body);
    if (out.ok === out.total) {
      if (ex.require && !ex.require.every(n => used.includes(n))) out.tips.push(`结果是对的，不过这道题要练 ${ex.require.join(' + ')}，换这个写法再试一次。`);
      else out.solved = true;
      return out;
    }
    const d = ex.diag && ex.diag(f); if (d) out.tips.push(d);
    if (spilledArr) out.tips.push('结果是一整组值（数组），但这个格子只需要一个值。检查一下是不是少了 SUM、ROWS、TEXTJOIN 这类把一组值收成一个结果的函数，或者范围写成了整列。');
    if (ex.spill && out.cells.some(c => c.extra)) out.tips.push('溢出的结果比答案多，检查一下筛选条件或者有没有去重。');
    const vs = out.cells.map(c => c.v);
    if (out.cells.length > 1 && !ex.spill) {
      try {
        const L = run(true);
        if (L.every(c => c.ok)) {
          const tk = parsed.toks; let eg = '';
          for (let i = 0; i + 2 < tk.length; i++) if (tk[i].t === 'ref' && tk[i + 1].v === ':' && tk[i + 2].t === 'ref' && (!tk[i].ar || !tk[i + 2].ar)) { eg = `$${X.numToCol(tk[i].col)}$${tk[i].row}:$${X.numToCol(tk[i + 2].col)}$${tk[i + 2].row}`; break; }
          out.tips.push(`往下填充后，查找范围跟着一起跑了。把固定不动的范围加上 $${eg ? '，写成 ' + eg : ''}。`);
        } else if (out.cells[0].ok && !out.cells.some(c => !c.ok && isErr(c.v)) && ((tg.c2 > tg.c1 && tg.r2 > tg.r1) || parsed.toks.some((k, i) => k.t === 'ref' && !k.ar && (parsed.toks[i + 1] && parsed.toks[i + 1].v === ':' || (parsed.toks[i - 1] && parsed.toks[i - 1].v === ':'))))) {
          out.tips.push(tg.c2 > tg.c1 && tg.r2 > tg.r1 ? '左上角第一格是对的，往下或往右填充后有的格子不对。想一想：往右拖时哪一列不能变（锁列 $A3），往下拖时哪一行不能变（锁行 C$2）？' : '第一行是对的，往下填充后有的行不对。检查一下：哪些引用要跟着每一行变，哪些应该用 $ 锁住？');
        }
      } catch (e) { /* 忽略 */ }
    }
    if (vs.some(v => isErr(v) && v.code === '#N/A') && !out.cells.some(c => isErr(c.exp))) out.tips.push('#N/A 的意思是「找不到」。可以用 IFNA 或 IFERROR 给找不到的情况指定一个显示内容。');
    if (vs.some(v => isErr(v) && v.code === '#VALUE!') && !spilledArr) out.tips.push('#VALUE! 通常是拿文字做了数学运算，或者文本转不成数字。先判断、清洗再计算。');
    if (vs.some(v => isErr(v) && v.code === '#REF!')) out.tips.push('#REF! 表示引用无效：表名、地址写错了，或者 INDIRECT 拼出来的地址不对。');
    if (vs.some(v => typeof v === 'boolean') && out.cells.some(c => typeof c.exp === 'string')) out.tips.push('结果是 TRUE / FALSE。题目要显示文字，用 IF 把它翻译成文字。');
    if (vs.some(v => typeof v === 'number') && out.cells.every(c => typeof c.exp === 'string') && !ex.fmt) out.tips.push('结果是数字，题目要的是文字。用 IF 或 TEXT 把它变成要求的样子。');
    const wt = out.cells.find(c => !c.ok && typeof c.v === 'string' && typeof c.exp === 'string' && c.exp !== '');
    if (wt && !out.tips.length) out.tips.push(`显示的文字要和题目要求完全一样。${wt.addr} 应该显示「${X.show(wt.exp)}」。`);
    if (!out.tips.length) out.tips.push('对照一下标了 ✗ 的格子。点一下格子，可以看到它实际执行的公式。');
    return out;
  }

  /* 回忆题：去掉空格、统一大小写和标点后比较 */
  function normFormula(s) {
    let out = '', q = false;
    s = String(s).replace(/[“”＂]/g, '"').replace(/[‘’]/g, "'").replace(/；/g, ';').replace(/，/g, ',').replace(/（/g, '(').replace(/）/g, ')').trim().replace(/^=/, '');
    for (const ch of s) { if (ch === '"') q = !q; if (!q && /\s/.test(ch)) continue; out += q ? ch : ch.toUpperCase(); }
    return out.replace(/;/g, ',');
  }
  function checkRecall(item, raw) { const n = normFormula(raw); if (!n) return null; return item.a.some(a => normFormula(a) === n); }

  G.PRACTICE_CORE = { parseTarget, truth, grade, checkRecall, normFormula, sheetOf };
})(typeof window !== 'undefined' ? window : globalThis);
