/* Excel 公式练习区 · 题库
 * 新增场景时：在 SCENARIOS 里加一个场景（表格数据 + 练习题），或在 RECALL 里加回忆题。
 * 经验库页面（index.html）里的「原公式」和「主动遮挡自测」会被自动读取，生成填空题和回忆题，不用手动录入。
 */
(function (G) {
  'use strict';
  const D = G.XL.D;
  // T(二维数组) → 从 A1 开始的单元格；null 表示空格子
  function T(rows, startRow) {
    const m = {}; const r0 = startRow || 1;
    rows.forEach((row, i) => row.forEach((v, j) => { if (v !== null && v !== undefined && v !== '') m[G.XL.numToCol(j + 1) + (r0 + i)] = v; }));
    return m;
  }
  const H = ['2026-11-01', '2026-11-11', '2026-12-25', '2027-01-01'].map(D);
  const holidaySheet = { name: '法国节假日', cells: T([['法国节假日'], ...H.map(d => [d])]), fmt: { A: 'date' } };
  const mapSheet = { name: '映射表', cells: T([
    ['客户名称', 'BI客户名', '下单城市'],
    ['GOODCANG', 'GOODCANG-FR', 'LISSES'],
    ['GOODCANG', 'GOODCANG FR2', 'LISSES'],
    ['GOODCANG', 'GOODCANG-FR', 'EVRY'],
    ['CIRRO', 'CIRRO-Temu-BT-PU', 'LISSES'],
    ['CIRRO', 'CIRRO-PU', 'ROISSY']
  ]) };
  const biRows = [
    ['周次', 'BI客户', '城市'],
    ['W34', 'GOODCANG-FR', 'LISSES'], ['W34', 'GOODCANG FR2', 'LISSES'], ['W34', 'GOODCANG-FR', 'EVRY'], ['W34', 'GOODCANG-FR', 'ROISSY'],
    ['W35', 'GOODCANG-FR', 'LISSES'], ['W34', 'CIRRO-Temu-BT-PU', 'LISSES'], ['W34', 'CIRRO-Temu-BT-PU', 'LISSES'], ['W34', 'CIRRO-PU', 'ROISSY'],
    ['W34', 'CIRRO-PU', 'LISSES'], ['W34', 'OTHER', 'LISSES']
  ];

  const SCENARIOS = [
    /* ---------- 基础：两列对比 ---------- */
    {
      id: 'base1', area: '基础', title: '同行对比', sub: '系统件数 vs 司机回传件数',
      story: '每天收工后，司机会回传每个客户的揽收件数。逐个客户核对：和系统里的件数一样吗？差了几件？',
      sheets: [{ name: '件数核对', cells: { ...T([
        ['客户', '系统件数', '司机回传', '是否一致', '差异件数', '异常判断'],
        ['Client A', 120, 120], ['Client B', 85, 83], ['Client C', 200, 200], ['Client D', 46, 46], ['Client E', 150, 152],
        ['Client F', 33, 33], ['Client G', 98, 90], ['Client H', 60, 60], ['Client I', 175, 175], ['Client J', 12, 15]
      ]), A13: '不一致的客户数' } }],
      exercises: [
        { id: 'b-1', title: '两个件数是否一致', target: 'D2:D11', fns: ['IF'],
          ask: '在 D 列判断每个客户的两个件数是否一致：一致显示「一致」，不一致显示「不一致」。',
          hints: ['这是「逐行判断」：每一行只看自己这一行的 B 和 C。', '判断两个数相等可以直接写 «B2=C2»，结果是 TRUE 或 FALSE。', '用 IF 翻译成文字：«IF(条件, 成立时显示, 不成立时显示)»，文字加英文双引号。'],
          answer: '=IF(B2=C2,"一致","不一致")',
          steps: [['B2=C2', '比较这一行的两个件数，得到 TRUE / FALSE'], ['IF( 条件 , "一致" , "不一致" )', '把 TRUE / FALSE 翻译成文字'], ['往下填充', 'B2、C2 没有 $，到下一行会自动变成 B3、C3']] },
        { id: 'b-2', title: '差了几件', target: 'E2:E11', fns: [],
          ask: '在 E 列算出差异件数：司机回传 − 系统件数。',
          hints: ['不需要函数，直接做减法。', '回传 − 系统，也就是 C 减 B。'],
          answer: '=C2-B2', steps: [['C2-B2', '正数 = 司机多报，负数 = 少报']] },
        { id: 'b-3', title: '统计不一致的客户数', target: 'B13', fns: ['COUNTIF'],
          ask: '在 B13 统计 D 列一共有几个「不一致」。（D 列按标准答案的结果计算。）',
          hints: ['在 D2:D11 里数一数「不一致」出现了几次。', '«COUNTIF(在哪个范围数, 数什么)»'],
          answer: '=COUNTIF(D2:D11,"不一致")', steps: [['COUNTIF(D2:D11,"不一致")', '在 D 列里数「不一致」的个数']] },
        { id: 'b-4', title: '差异超过 2 件才算异常', target: 'F2:F11', fns: ['IF', 'ABS'], challenge: true,
          ask: '差异超过 2 件（不管多报还是少报）显示「异常」，否则显示「正常」。',
          hints: ['差异 «C2-B2» 可能是负数。', '«ABS()» 去掉负号。', '«ABS(C2-B2)>2» 当成 IF 的条件。'],
          answer: '=IF(ABS(C2-B2)>2,"异常","正常")',
          steps: [['ABS(C2-B2)', '只看差了多少，不管正负'], ['IF( …>2 , "异常" , "正常" )', '超过 2 件就是异常']],
          alt: [['=IF(OR(C2-B2>2,C2-B2<-2),"异常","正常")', '用 OR 连接两种情况']] }
      ]
    },
    {
      id: 'base2', area: '基础', title: '查找有无', sub: '系统下单 vs 司机扫描',
      story: '系统里有 10 个单号，司机扫描了 9 个，顺序也不一样。哪些单漏扫了？司机扫的单系统里都有吗？',
      sheets: [{ name: '扫描核对', cells: T([
        ['系统单号', '司机是否扫描', null, '司机扫描单号', '系统里是否有'],
        ['GF2026101', null, null, 'GF2026105'], ['GF2026102', null, null, 'GF2026101'], ['GF2026103', null, null, 'GF2026110'],
        ['GF2026104', null, null, 'GF2026103'], ['GF2026105', null, null, 'GF2026108'], ['GF2026106', null, null, 'GF2026102'],
        ['GF2026107', null, null, 'GF2026199'], ['GF2026108', null, null, 'GF2026106'], ['GF2026109', null, null, 'GF2026109'], ['GF2026110']
      ]) }],
      exercises: [
        { id: 'b-5', title: '找出漏扫的单', target: 'B2:B11', fns: ['COUNTIF', 'IF'],
          ask: '在 B 列判断系统里的每个单号司机有没有扫到：扫到显示「已扫描」，没扫到显示「漏扫」。',
          hints: ['换个问法：A2 这个单号，在 D 列里出现了几次？', '«COUNTIF($D$2:$D$10, A2)»，次数大于 0 就是扫到了。', '查找范围要用 $ 锁住，A2 不锁。'],
          answer: '=IF(COUNTIF($D$2:$D$10,A2)>0,"已扫描","漏扫")',
          steps: [['COUNTIF($D$2:$D$10,A2)', '在扫描列里数 A2 出现几次'], ['IF( …>0 , "已扫描" , "漏扫" )', '次数翻译成文字']] },
        { id: 'b-6', title: '找出系统里没有的单', target: 'E2:E10', fns: ['COUNTIF', 'IF'],
          ask: '反过来：E 列判断司机扫的每个单号系统里有没有。有显示「有」，没有显示「系统无此单」。',
          hints: ['同一个套路，方向反过来。', '在 «$A$2:$A$11» 里数 D2。'],
          answer: '=IF(COUNTIF($A$2:$A$11,D2)>0,"有","系统无此单")',
          steps: [['COUNTIF($A$2:$A$11,D2)', '拿扫描单号去系统列里数'], ['口诀', '两列找有无 = COUNTIF 数次数 + IF 翻译']] }
      ]
    },
    {
      id: 'base3', area: '基础', title: '带值对比', sub: '系统重量 vs 仓库称重',
      story: '按单号把仓库称重拉到左边，再算差了多少公斤。仓库表顺序不同，还少了一单。',
      sheets: [{ name: '重量核对', cells: T([
        ['单号', '系统重量kg', '仓库称重kg', '差异kg', null, '仓库单号', '仓库称重kg', null, 'INDEX+MATCH'],
        ['GF2026101', 12.5, null, null, null, 'GF2026103', 20.3], ['GF2026102', 8.0, null, null, null, 'GF2026110', 18.9],
        ['GF2026103', 20.3, null, null, null, 'GF2026101', 12.5], ['GF2026104', 5.6, null, null, null, 'GF2026106', 9.8],
        ['GF2026105', 15.0, null, null, null, 'GF2026102', 8.4], ['GF2026106', 9.8, null, null, null, 'GF2026109', 11.1],
        ['GF2026107', 30.0, null, null, null, 'GF2026105', 15.0], ['GF2026108', 7.2, null, null, null, 'GF2026107', 31.2],
        ['GF2026109', 11.1, null, null, null, 'GF2026108', 7.2], ['GF2026110', 18.4]
      ]) }],
      exercises: [
        { id: 'b-7', title: '把仓库称重拉过来', target: 'C2:C11', fns: ['VLOOKUP', 'IFERROR'],
          ask: '在 C 列按单号把仓库称重拉过来。仓库没有这一单的，显示「仓库无记录」。',
          hints: ['拿 A2 去仓库表 F 列找，带回同一行 G 列。', '«VLOOKUP(找谁, $F$2:$G$10, 2, 0)»', '找不到会报 #N/A，外面套 «IFERROR(…, "仓库无记录")»。'],
          answer: '=IFERROR(VLOOKUP(A2,$F$2:$G$10,2,0),"仓库无记录")',
          steps: [['VLOOKUP(A2,$F$2:$G$10,2,0)', '在 F 列找 A2，带回区域第 2 列；0 = 精确匹配'], ['IFERROR( … , "仓库无记录")', '找不到时显示指定文字']],
          alt: [['=XLOOKUP(A2,$F$2:$F$10,$G$2:$G$10,"仓库无记录")', 'XLOOKUP 一个函数就包含兜底']],
          diag: f => (/VLOOKUP/i.test(f) && !/,(0|FALSE)\)/i.test(f.replace(/\s/g, '')) ? 'VLOOKUP 的第 4 个参数要写 0（精确匹配），否则会带回别的单的重量。' : null) },
        { id: 'b-8', title: '算出重量差', target: 'D2:D11', fns: ['IF', 'ISNUMBER'],
          ask: '在 D 列算「仓库称重 − 系统重量」。C 列是「仓库无记录」的那一行显示「—」或任意文字。',
          hints: ['直接 «C2-B2» 会在「仓库无记录」那行报 #VALUE!。', '先判断 «ISNUMBER(C2)»。', '«IF(ISNUMBER(C2), C2-B2, "—")»'],
          answer: '=IF(ISNUMBER(C2),C2-B2,"—")',
          steps: [['ISNUMBER(C2)', '是数字才相减'], ['"—"', '不是数字就显示短横线']],
          alt: [['=IFERROR(C2-B2,"—")', '先减，出错再兜底']],
          accept: (v, exp) => (typeof exp === 'string' ? (typeof v === 'string' && v.length > 0) : null) },
        { id: 'b-9', title: '用 INDEX + MATCH 再查一遍', target: 'I2:I11', fns: ['INDEX', 'MATCH', 'IFERROR'], challenge: true, require: ['INDEX', 'MATCH'],
          ask: '在 I 列做出和 C 列一样的效果，这次用 INDEX + MATCH（不限查找方向）。',
          hints: ['«MATCH(A2, $F$2:$F$10, 0)» 找到第几个。', '«INDEX($G$2:$G$10, 第几个)» 取出重量。', '外面再套 IFERROR。'],
          answer: '=IFERROR(INDEX($G$2:$G$10,MATCH(A2,$F$2:$F$10,0)),"仓库无记录")',
          steps: [['MATCH(A2,$F$2:$F$10,0)', '定位'], ['INDEX($G$2:$G$10, …)', '取值'], ['IFERROR', '兜底']] }
      ]
    },

    /* ---------- 场景 1：揽收时间 ---------- */
    {
      id: 's1', area: '场景 1', link: 'index.html#scenario1', title: '揽收时间', sub: '下单时间 → 当天揽收还是顺延',
      story: '下单时间是系统导出的文本。截单时间默认 11:00，个别客户单独设置。周末和法国节假日不揽收。把场景 1 的大公式拆成小步骤，一步一步做出来。',
      sheets: [
        { name: '订单', fmt: { C: 'date', D: 'time', G: 'time', H: 'date', I: 'date' }, cells: T([
          ['客户', '下单时间', '下单日期', '下单时刻', '周末?', '节假日?', '截单时间', '下一个工作日', '揽收日期'],
          ['Client A', '2026-10-23 10:30'], ['Client B ', '2026-10-23 15:10'], ['Client C', '2026-10-24 09:00'],
          ['Client A', '2026-10-30 11:45'], ['Client D', '2026-11-11 08:20'], ['Client B', '2026-11-10 16:30']
        ]) },
        holidaySheet,
        { name: '揽收客户截单时间设置', fmt: { B: 'time' }, cells: T([['客户', '截单时间'], ['Client B', D('16:00')], ['Client D', D('09:30')]]) }
      ],
      exercises: [
        { id: 's1-1', title: '从下单时间里拿出日期', target: 'C2:C7', fmt: 'date', fns: ['VALUE', 'INT'],
          ask: 'B 列是文本格式的下单时间。在 C 列得到「哪一天」（不要几点几分）。',
          hints: ['B 列是文本，先用 «VALUE(B2)» 变成真正的日期时间数字。', '日期时间是一个数：整数部分是日期，小数部分是时间。', '«INT()» 取整数部分。'],
          answer: '=INT(VALUE(B2))',
          steps: [['VALUE(B2)', '验钞机：文本 → 可以计算的日期时间'], ['INT( … )', '拿日期：只留整数部分']] },
        { id: 's1-2', title: '拿出几点几分', target: 'D2:D7', fmt: 'time', fns: ['VALUE', 'MOD'],
          ask: '在 D 列得到下单的「几点几分」（不要日期）。',
          hints: ['和上一题是一对：这次要小数部分。', '«MOD(数字, 1)» 就是除以 1 的余数，也就是小数部分。'],
          answer: '=MOD(VALUE(B2),1)',
          steps: [['MOD( t , 1 )', '拿时间：只留小数部分'], ['口诀', 'INT 拿日期，MOD 拿钟表']] },
        { id: 's1-3', title: '是不是周末', target: 'E2:E7', fns: ['IF', 'WEEKDAY'],
          ask: 'E 列：C 列的日期是周六或周日显示「周末」，否则显示「工作日」。',
          hints: ['«WEEKDAY(日期, 2)» 把周一记为 1 … 周日记为 7。', '大于 5 就是周六、周日。'],
          answer: '=IF(WEEKDAY(C2,2)>5,"周末","工作日")',
          steps: [['WEEKDAY(C2,2)', '第 2 个参数写 2：周一 = 1 … 周日 = 7'], ['>5', '6 和 7 是周末']] },
        { id: 's1-4', title: '是不是法国节假日', target: 'F2:F7', fns: ['IF', 'COUNTIF'],
          ask: 'F 列：C 列的日期在「法国节假日」表里显示「节假日」，否则显示「否」。',
          hints: ['「在不在名单里」= 在名单里数一数出现几次。', "«COUNTIF('法国节假日'!$A$2:$A$5, C2)»，跨表引用要写表名。"],
          answer: "=IF(COUNTIF('法国节假日'!$A$2:$A$5,C2)>0,\"节假日\",\"否\")",
          steps: [["COUNTIF('法国节假日'!$A$2:$A$5,C2)", '在节假日表里数这个日期'], ['>0', '出现过就是节假日']] },
        { id: 's1-5', title: '查客户的截单时间', target: 'G2:G7', fmt: 'time', fns: ['IFNA', 'XLOOKUP', 'TRIM', 'TIME'],
          ask: 'G 列：去「揽收客户截单时间设置」表查这个客户的截单时间；没有单独设置的，默认 11:00。注意 A3 的客户名后面多了一个空格。',
          hints: ['查表用 «XLOOKUP(找谁, 在哪列找, 带回哪列)»。', '客户名先 «TRIM(A2)» 去掉空格再去找。', '查不到会报 #N/A，用 «IFNA(…, TIME(11,0,0))» 给默认值。'],
          answer: "=IFNA(XLOOKUP(TRIM(A2),'揽收客户截单时间设置'!$A$2:$A$3,'揽收客户截单时间设置'!$B$2:$B$3),TIME(11,0,0))",
          diag: f => (!/TRIM/i.test(f) ? 'A3 的客户名「Client B 」后面多了一个空格，直接查会查不到。先用 TRIM(A2) 去掉空格再查。' : null),
          steps: [['TRIM(A2)', '去掉客户名首尾空格'], ['XLOOKUP( … )', '在设置表里查截单时间'], ['IFNA( … , TIME(11,0,0))', '查不到就用默认 11:00']] },
        { id: 's1-6', title: '下一个法国工作日', target: 'H2:H7', fmt: 'date', fns: ['WORKDAY.INTL'],
          ask: 'H 列：从 C 列的日期出发，找下一个法国工作日（跳过周六周日和节假日）。',
          hints: ['«WORKDAY.INTL(开始日期, 往后几个工作日, 周末规则, 节假日)»', '周末规则 «"0000011"»：从周一数到周日，1 = 休息。'],
          answer: "=WORKDAY.INTL(C2,1,\"0000011\",'法国节假日'!$A$2:$A$5)",
          steps: [['C2 , 1', '从这天出发，往后找 1 个工作日'], ['"0000011"', '周六、周日休息'], ["'法国节假日'!$A$2:$A$5", '节假日也跳过']] },
        { id: 's1-7', title: '完整版：算出揽收日期', target: 'I2:I7', fmt: 'date', fns: ['IF', 'LET', 'VALUE', 'INT', 'MOD', 'XLOOKUP', 'IFNA', 'WEEKDAY', 'COUNTIF', 'WORKDAY.INTL'], challenge: true,
          ask: 'I 列：不超过截单时间、且当天是工作日 → 当天揽收；否则 → 下一个工作日。可以引用 C~H 列，也可以挑战用 LET 一个公式写完（就是经验库里的原公式）。',
          hints: ['先判断时间：«D2<=G2» 有没有赶上截单。', '赶上了还要看当天能不能揽：«OR(E2="周末", F2="节假日")»。', '口诀：空值先挡住，时间拆两半；客户找截单，晚了往后算；当天若放假，也要往后转。'],
          answer: "=IF(B2=\"\",\"\",LET(t,VALUE(B2),d,INT(t),h,'法国节假日'!$A$2:$A,cutoff,IFNA(XLOOKUP(TRIM(A2),'揽收客户截单时间设置'!$A:$A,'揽收客户截单时间设置'!$B:$B),TIME(11,0,0)),IF(MOD(t,1)<=cutoff,IF(OR(WEEKDAY(d,2)>5,COUNTIF(h,d)>0),WORKDAY.INTL(d,1,\"0000011\",h),d),WORKDAY.INTL(d,1,\"0000011\",h))))",
          steps: [['LET(t, …, d, …, h, …, cutoff, …)', '先把中间结果起好名字'], ['MOD(t,1)<=cutoff', '赶上截单了吗'], ['OR(WEEKDAY(d,2)>5, COUNTIF(h,d)>0)', '当天是周末或节假日吗'], ['WORKDAY.INTL(d,1,"0000011",h)', '往后找下一个工作日']],
          alt: [['=IF(AND(D2<=G2,E2="工作日",F2="否"),C2,H2)', '引用前面几列的拆分版，更好读']] }
      ]
    },

    /* ---------- 场景 2：CIFR 编号 ---------- */
    {
      id: 's2', area: '场景 2', link: 'index.html#scenario2', title: 'CIFR 编号', sub: '前缀 + 城市缩写 + 日期 + 运输商代码 + 序号',
      story: '每一行是一次揽收：运输商、日期、城市。编号 = CIFR + 城市缩写 + 日期(ddmmyyyy) + 运输商代码 + 同组序号（同运输商、同日期、同城市第几次）。',
      sheets: [
        { name: '揽收明细', fmt: { B: 'date' }, cells: T([
          ['运输商', '揽收日期', '城市', '日期片段', '城市缩写', '运输商代码', '序号', '完整编号'],
          ['Colis Express', D('2026-10-26'), 'LISSES'], ['Colis Express', D('2026-10-26'), 'LISSES'], ['Rapid Trans', D('2026-10-26'), 'ROUEN'],
          ['Colis Express', D('2026-10-27'), 'LISSES'], ['Colis Express', D('2026-10-26'), 'ROUEN'], ['Rapid Trans', D('2026-10-26'), 'ROUEN']
        ]) },
        { name: '城市缩写（数据源）', cells: T([['缩写', '城市'], ['LIS', 'LISSES'], ['ROU', 'ROUEN'], ['CDG', 'ROISSY']]) },
        { name: '运输商信息（数据源）', cells: T([['运输商', '代码'], ['Colis Express', 'CE'], ['Rapid Trans', 'RT']]) }
      ],
      exercises: [
        { id: 's2-1', title: '日期变成编号片段', target: 'D2:D7', fns: ['TEXT'],
          ask: 'D 列：把 B 列的日期变成 26102026 这样的 8 位片段。',
          hints: ['给日期「穿上固定格式的衣服」用 TEXT。', '格式写成 «"ddmmyyyy"»：日日月月年年年年。'],
          answer: '=TEXT(B2,"ddmmyyyy")', steps: [['TEXT(B2,"ddmmyyyy")', 'dd 日、mm 月、yyyy 年']] },
        { id: 's2-2', title: '查城市缩写', target: 'E2:E7', fns: ['IFERROR', 'INDEX', 'MATCH', 'TRIM'],
          ask: 'E 列：去「城市缩写（数据源）」表查城市对应的缩写。注意缩写在城市的左边。',
          hints: ['要找的列（缩写）在查找列（城市）左边，VLOOKUP 做不到，用 INDEX + MATCH。', "«MATCH(TRIM(C2),'城市缩写（数据源）'!B:B,0)» 找城市在第几行。", "«INDEX('城市缩写（数据源）'!A:A, 第几行)» 取缩写，外面套 IFERROR。"],
          answer: "=IFERROR(INDEX('城市缩写（数据源）'!A:A,MATCH(TRIM(C2),'城市缩写（数据源）'!B:B,0)),\"\")",
          steps: [['MATCH(TRIM(C2), 城市列, 0)', '城市在第几行'], ['INDEX(缩写列, 第几行)', '取出同一行的缩写'], ['IFERROR( … ,"")', '查不到就留空']] },
        { id: 's2-3', title: '查运输商代码', target: 'F2:F7', fns: ['IFERROR', 'VLOOKUP', 'TRIM'],
          ask: 'F 列：去「运输商信息（数据源）」表查运输商代码。',
          hints: ['这次代码在运输商右边，VLOOKUP 就可以。', "«VLOOKUP(TRIM(A2),'运输商信息（数据源）'!A:B,2,0)»"],
          answer: "=IFERROR(VLOOKUP(TRIM(A2),'运输商信息（数据源）'!A:B,2,0),\"\")",
          steps: [['VLOOKUP(TRIM(A2), A:B, 2, 0)', '在 A 列找运输商，带回第 2 列'], ['IFERROR', '查不到留空']] },
        { id: 's2-4', title: '同组第几次（补零）', target: 'G2:G7', fns: ['TEXT', 'COUNTIFS'],
          ask: 'G 列：同运输商、同日期、同城市，这是第几次？显示成 01、02。',
          hints: ['从第一行数到当前行：范围起点锁住，终点不锁，例如 «A$2:$A2»。', '«COUNTIFS(A$2:$A2,A2, B$2:$B2,B2, C$2:$C2,C2)»', '外面 «TEXT(…, "00")» 补零。'],
          answer: '=TEXT(COUNTIFS(A$2:$A2,A2,B$2:$B2,B2,C$2:$C2,C2),"00")',
          steps: [['A$2:$A2', '起点固定在第 2 行，终点跟着当前行往下长'], ['COUNTIFS( 三组条件 )', '三个条件都相同才算同一组'], ['TEXT( … ,"00")', '1 显示成 01']] },
        { id: 's2-5', title: '拼出完整编号', target: 'H2:H7', fns: ['IF', 'TEXT', 'COUNTIFS', 'INDEX', 'MATCH', 'VLOOKUP', 'IFERROR'], challenge: true,
          ask: 'H 列：编号 = "CIFR" & 城市缩写 & 日期片段 & 运输商代码 & 序号。可以引用 D~G 列，也可以一个公式写完。',
          hints: ['多段文字用 «&» 粘起来。', '顺序：CIFR → 城市缩写 → 日期 → 运输商代码 → 序号。'],
          answer: "=IF(A2=\"\",\"\",\"CIFR\"&IFERROR(INDEX('城市缩写（数据源）'!A:A,MATCH(TRIM(C2),'城市缩写（数据源）'!B:B,0)),\"\")&IFERROR(TEXT(B2,\"ddmmyyyy\"),\"\")&IFERROR(VLOOKUP(TRIM(A2),'运输商信息（数据源）'!A:B,2,0),\"\")&TEXT(COUNTIFS(A$2:$A2,A2,B$2:$B2,B2,C$2:$C2,C2),\"00\"))",
          steps: [['"CIFR" & …', '固定前缀'], ['& 城市缩写 & 日期 & 代码 & 序号', '文本胶水依次粘起来']],
          alt: [['="CIFR"&E2&D2&F2&G2', '引用拆分好的几列']] }
      ]
    },

    /* ---------- 场景 3：时效 ---------- */
    {
      id: 's3', area: '场景 3', link: 'index.html#scenario3', title: '下单到揽收时效', sub: '剔除周末和法国节假日',
      story: '计算下单到揽收用了多少小时，周末和节假日不算。先分别数自然日、工作日，再相减得到非工作日。',
      sheets: [
        { name: '时效', fmt: { A: 'datetime', B: 'datetime', F: 'hours', G: 'hours' }, cells: T([
          ['下单时间', '揽收时间', '自然日', '工作日', '非工作日', '时效(小时)', '完整版'],
          [D('2026-10-23 10:00'), D('2026-10-23 16:00')], [D('2026-10-23 15:00'), D('2026-10-26 10:00')],
          [D('2026-10-30 09:00'), D('2026-11-02 11:00')], [D('2026-11-10 14:00'), D('2026-11-12 09:00')],
          [D('2026-10-24 10:00'), D('2026-10-26 15:00')]
        ]) },
        holidaySheet
      ],
      exercises: [
        { id: 's3-1', title: '一共跨了几个自然日', target: 'C2:C6', fns: ['INT'],
          ask: 'C 列：从下单日到揽收日，一共几天（两头都算）。',
          hints: ['只看日期部分：«INT(B2)»、«INT(A2)»。', '两头都算要 +1。'],
          answer: '=INT(B2)-INT(A2)+1', steps: [['INT(B2)-INT(A2)', '相差几天'], ['+1', '起止当天都算']] },
        { id: 's3-2', title: '其中有几个工作日', target: 'D2:D6', fns: ['NETWORKDAYS.INTL', 'INT'],
          ask: 'D 列：这段日期里有几个法国工作日（两头都算）。',
          hints: ['数工作日用 NETWORKDAYS.INTL：一个「数日子」，WORKDAY.INTL 是「找日子」。', "«NETWORKDAYS.INTL(INT(A2), INT(B2), \"0000011\", '法国节假日'!$A$2:$A$5)»"],
          answer: "=NETWORKDAYS.INTL(INT(A2),INT(B2),\"0000011\",'法国节假日'!$A$2:$A$5)",
          steps: [['INT(A2) , INT(B2)', '开始和结束日期'], ['"0000011"', '周六周日休息'], ['节假日', '也要扣掉']] },
        { id: 's3-3', title: '非工作日天数', target: 'E2:E6', fns: ['NETWORKDAYS.INTL'],
          ask: 'E 列：非工作日有几天？',
          hints: ['自然日 − 工作日 = 非工作日。'],
          answer: '=C2-D2', steps: [['C2-D2', '自然日 − 工作日']] },
        { id: 's3-4', title: '剔除后的时效小时数', target: 'F2:F6', fmt: 'hours', fns: ['MAX'],
          ask: 'F 列：总小时数减去非工作日的小时数，最少为 0。',
          hints: ['日期时间相减得到的是「天」，乘 24 变成小时：«(B2-A2)*24»。', '每个非工作日扣 24 小时：«E2*24»。', '«MAX(0, …)» 保证不出现负数。'],
          answer: '=MAX(0,(B2-A2)*24-E2*24)', steps: [['(B2-A2)*24', '总小时数'], ['- E2*24', '扣掉非工作日'], ['MAX(0, …)', '最少是 0']] },
        { id: 's3-5', title: '完整版：下单落在周末时先顺延', target: 'G2:G6', fmt: 'hours', fns: ['IF', 'OR', 'LET', 'INT', 'WEEKDAY', 'COUNTIF', 'WORKDAY.INTL', 'NETWORKDAYS.INTL', 'MAX'], challenge: true,
          ask: 'G 列：经验库原公式的做法——如果下单或揽收落在周末/节假日，先把它顺延到下一个工作日 0 点再算。看看第 6 行和 F 列结果有什么不同。',
          hints: ['用 LET 定义 h、oDate、pDate。', 'Orderin：下单日是周末或节假日 → «WORKDAY.INTL(oDate,1,"0000011",h)»，否则用原时间。', '最后 «MAX(0, raw - nonwork*24)»。'],
          answer: "=IF(OR(A2=\"\",B2=\"\"),\"\",LET(h,'法国节假日'!$A$2:$A,oDate,INT(A2),pDate,INT(B2),Orderin,IF(OR(WEEKDAY(oDate,2)>5,COUNTIF(h,oDate)>0),WORKDAY.INTL(oDate,1,\"0000011\",h),A2),Pickup,IF(OR(WEEKDAY(pDate,2)>5,COUNTIF(h,pDate)>0),WORKDAY.INTL(pDate,1,\"0000011\",h),B2),raw,(Pickup-Orderin)*24,nonwork,(INT(Pickup)-INT(Orderin)+1)-NETWORKDAYS.INTL(INT(Orderin),INT(Pickup),\"0000011\",h),MAX(0,raw-nonwork*24)))",
          steps: [['Orderin / Pickup', '落在非工作日的先顺延'], ['raw', '顺延后的总小时'], ['nonwork', '中间夹着的非工作日'], ['MAX(0, raw-nonwork*24)', '扣掉后的时效']] }
      ]
    },

    /* ---------- 场景 4：同名聚合 ---------- */
    {
      id: 's4', area: '场景 4', link: 'index.html#scenario4', title: '同名客户聚合', sub: '客户清单 + 城市清单，三道闸门计数',
      story: '一个客户在 BI 里可能有好几个名字。先从映射表生成「BI 客户名清单」和「城市清单」，再在 BI 数据里数：周次对、客户在清单里、城市在清单里，三个条件都满足才算。',
      sheets: [
        { name: '周报', cells: T([['统计周次', null, '客户', 'BI名个数', '在清单?', 'BI客户名清单', '票数'], ['W34', null, 'GOODCANG'], [null, null, 'CIRRO']]) },
        mapSheet,
        { name: 'BI数据源', cells: T(biRows) }
      ],
      exercises: [
        { id: 's4-1', title: '这个客户有几个 BI 名字', target: 'D2:D3', fns: ['ROWS', 'UNIQUE', 'FILTER'],
          ask: 'D 列：在映射表里，这个客户对应几个不同的 BI 客户名？',
          hints: ["先筛：«FILTER('映射表'!$B$2:$B$6, '映射表'!$A$2:$A$6=C2)»。", '再去重：«UNIQUE(…)»。', '最后数行数：«ROWS(…)»。'],
          answer: "=ROWS(UNIQUE(FILTER('映射表'!$B$2:$B$6,'映射表'!$A$2:$A$6=C2)))",
          steps: [['FILTER(B列, A列=C2)', '筛出这个客户的所有 BI 名'], ['UNIQUE', '去重'], ['ROWS', '数有几个']] },
        { id: 's4-2', title: '列出 BI 客户名清单', target: 'F2', spill: true, fns: ['UNIQUE', 'FILTER'],
          ask: 'F2：写一个公式，把 GOODCANG（C2）对应的所有 BI 客户名去重后列出来，结果会自动往下溢出。',
          hints: ['和上一题一样先 FILTER 再 UNIQUE，只是不用 ROWS。'],
          answer: "=UNIQUE(FILTER('映射表'!$B$2:$B$6,'映射表'!$A$2:$A$6=C2))",
          steps: [['UNIQUE(FILTER( … ))', '做清单：先筛选，再去重']] },
        { id: 's4-3', title: '判断是否在清单里', sheet: 'BI数据源', target: 'D2:D11', fns: ['ISNUMBER', 'MATCH'],
          ask: '在「BI数据源」表的 D 列：这一行的 BI 客户名，在不在映射表的 B 列里？显示 TRUE / FALSE。',
          hints: ["«MATCH(B2, '映射表'!$B$2:$B$6, 0)» 找到返回位置号，找不到报错。", '«ISNUMBER(…)»：是数字 = 找到了。'],
          answer: "=ISNUMBER(MATCH(B2,'映射表'!$B$2:$B$6,0))",
          steps: [['MATCH(…,0)', '查名单，返回位置号或错误'], ['ISNUMBER', '有位置号就是 TRUE']] },
        { id: 's4-4', title: '三道闸门计数', target: 'G2:G3', fns: ['LET', 'UNIQUE', 'FILTER', 'SUMPRODUCT', 'N', 'ISNUMBER', 'MATCH'], challenge: true,
          ask: 'G 列：统计 A2 这一周，这个客户在 BI 数据源里的票数。三个条件都满足才算：周次相同、BI 客户名在清单里、城市在清单里。',
          hints: ['用 LET 定义 keys（BI 名清单）和 addr（城市清单），都是 UNIQUE(FILTER(…))。', '每个条件写成 TRUE/FALSE 的一整列，再用 «N()» 变成 1/0。', '«SUMPRODUCT(N(条件1), N(条件2), N(条件3))»：三个都是 1 才计 1。'],
          answer: "=LET(k,TRIM(C2),keys,UNIQUE(FILTER('映射表'!$B$2:$B$6,'映射表'!$A$2:$A$6=k)),addr,UNIQUE(FILTER('映射表'!$C$2:$C$6,'映射表'!$A$2:$A$6=k)),IF(k=\"\",\"\",SUMPRODUCT(N('BI数据源'!$A$2:$A$11=$A$2),N(ISNUMBER(MATCH('BI数据源'!$B$2:$B$11,keys,0))),N(ISNUMBER(MATCH('BI数据源'!$C$2:$C$11,addr,0))))))",
          steps: [['keys / addr', '两张清单'], ['N(… = $A$2)', '闸门 1：周次'], ['N(ISNUMBER(MATCH(客户, keys, 0)))', '闸门 2：客户在清单'], ['N(ISNUMBER(MATCH(城市, addr, 0)))', '闸门 3：城市在清单'], ['SUMPRODUCT', '三个 1 相乘再加总']] }
      ]
    },

    /* ---------- 场景 5：二维取数 ---------- */
    {
      id: 's5', area: '场景 5', link: 'index.html#scenario5', title: '二维取数', sub: '原表横向日期块 → 明细表',
      story: '原表按日期横向分块：第 2 行是日期（写在每个块的第一列），第 3 行是块内的指标名。明细表要按「日期 + 客户 + 指标」把数取出来。',
      sheets: [
        { name: '明细', fmt: { A: 'date' }, cells: T([
          ['日期', '客户', '下单票数', '揽收票数', '客户行号', '日期起始列'],
          [D('2026-10-26'), 'Client A'], [D('2026-10-26'), 'Client C'], [D('2026-10-27'), 'Client B'], [D('2026-10-27'), 'Client A']
        ]) },
        { name: '原表', fmt: { C: 'date', F: 'date' }, cells: T([
          ['原表', null, null, null, null, null, null, null],
          [null, null, D('2026-10-26'), null, null, D('2026-10-27')],
          [null, '客户', '下单票数', '揽收票数', '揽收率', '下单票数', '揽收票数', '揽收率'],
          [null, 'Client A', 120, 118, 0.98, 95, 95, 1],
          [null, 'Client B', 60, 54, 0.9, 72, 70, 0.97],
          [null, 'Client C', 33, 33, 1, 41, 39, 0.95]
        ]) }
      ],
      exercises: [
        { id: 's5-1', title: '客户在原表第几行', target: 'E2:E5', fns: ['MATCH'],
          ask: 'E 列：B 列的客户在原表 B 列的第几行？',
          hints: ['«MATCH(找谁, 在哪一列找, 0)» 返回第几个。', "在整列 «'原表'!$B:$B» 里找，返回的就是行号。"],
          answer: "=MATCH($B2,'原表'!$B:$B,0)", steps: [["MATCH($B2,'原表'!$B:$B,0)", '整列里找，位置号 = 行号']] },
        { id: 's5-2', title: '日期块从第几列开始', target: 'F2:F5', fns: ['MATCH'],
          ask: 'F 列：A 列的日期，在原表第 2 行的第几列？',
          hints: ["在整行 «'原表'!$2:$2» 里找日期。"],
          answer: "=MATCH($A2,'原表'!$2:$2,0)", steps: [["MATCH($A2,'原表'!$2:$2,0)", '整行里找，位置号 = 列号']] },
        { id: 's5-3', title: '行号 + 列号交叉取数', target: 'C2:D5', fns: ['INDEX', 'MATCH'], challenge: true,
          ask: 'C2:D5：按日期、客户和这一列的指标名（C1、D1）从原表取数。在 C2 写好后，网站会往下、往右填满整个区域，注意 $ 怎么锁。',
          hints: ["«INDEX('原表'!$1:$6, 行号, 列号)»。", '列号 = 日期起始列 + 指标在块内的位置 − 1。', "块内位置：«MATCH(C$1,'原表'!$C$3:$E$3,0)»。客户、日期锁列（$B2、$A2），指标名锁行（C$1）。"],
          answer: "=INDEX('原表'!$1:$6,MATCH($B2,'原表'!$B:$B,0),MATCH($A2,'原表'!$2:$2,0)+MATCH(C$1,'原表'!$C$3:$E$3,0)-1)",
          steps: [['MATCH($B2, 客户列, 0)', '行号'], ['MATCH($A2, 日期行, 0)', '日期块起点'], ['MATCH(C$1, 指标行, 0)', '块内第几个'], ['起点 + 块内位置 − 1', '起点本身就是第 1 个指标，所以减 1'], ['$B2 / C$1', '往右拖客户不能变，往下拖指标名不能变']] }
      ]
    },

    /* ---------- 场景 6：条件拼接 ---------- */
    {
      id: 's6', area: '场景 6', link: 'index.html#scenario6', title: '时效梯队', sub: 'COUNTIFS 条件拼接 + 混合锁定',
      story: '左边是客户，上面是时效天数。每个格子 = 这个客户、这个时效的票数。一个公式往下往右填满整张表，关键是 $A3 和 C$2。',
      sheets: [
        { name: '梯队', cells: T([['阈值(天)', 2, null, null, null, null], ['客户', null, 1, 2, 3, '超过阈值'], ['Client A'], ['Client B'], ['Client C']]) },
        { name: 'BI签入数据', cells: T([
          ['客户', '时效天数'], ['Client A', 1], ['Client A', 1], ['Client A', 2], ['Client A', 3], ['Client A', 4],
          ['Client B', 2], ['Client B', 2], ['Client B', 3], ['Client C', 1], ['Client C', 3], ['Client C', 3], ['Client C', 5]
        ]) }
      ],
      exercises: [
        { id: 's6-1', title: '客户 × 时效的票数', target: 'C3:E5', fns: ['COUNTIFS'],
          ask: 'C3:E5：每个格子统计「这个客户、时效等于这一列表头」的票数。在 C3 写一个公式，往下往右都要对。',
          hints: ["«COUNTIFS('BI签入数据'!$A:$A, 客户, 'BI签入数据'!$B:$B, 条件)»", '条件写成 «"="&C$2»：比较符号放引号里，单元格用 & 接上。', '客户只锁列 «$A3»，表头只锁行 «C$2»。'],
          answer: "=COUNTIFS('BI签入数据'!$A:$A,$A3,'BI签入数据'!$B:$B,\"=\"&C$2)",
          steps: [['$A3', '往右拖还是看 A 列客户，往下拖换客户'], ['"="&C$2', '等于当前列表头；往下拖表头行不变'], ['COUNTIFS', '两个条件同时满足才计数']] },
        { id: 's6-2', title: '超过阈值的票数', target: 'F3:F5', fns: ['COUNTIFS'],
          ask: 'F 列：时效「大于 B1 阈值」的票数。阈值以后可能会改，所以要引用 B1，不要直接写 2。',
          hints: ['大于某个单元格：«">"&$B$1»。'],
          answer: "=COUNTIFS('BI签入数据'!$A:$A,$A3,'BI签入数据'!$B:$B,\">\"&$B$1)",
          steps: [['">"&$B$1', '拼出「>2」这个条件，B1 改了结果自动变']] }
      ]
    },

    /* ---------- 场景 7 + 8：动态引用 ---------- */
    {
      id: 's7', area: '场景 7–8', link: 'index.html#scenario7', title: '动态引用工作表', sub: 'SUBSTITUTE + INDIRECT，统计唯一值和逗号金额',
      story: '每个客户一张工作表（A 列周次、B 列运单号、C 列金额，金额是 12,5 这种逗号小数）。汇总表 A 列写表名，用 INDIRECT 去对应的表里取数。',
      sheets: [
        { name: '汇总', cells: T([['表名', 'W43运单', 'W44运单', 'W43金额', 'W44金额', null, '安全表名'], ['客户A'], ["L'OREAL"]]) },
        { name: '客户A', cells: T([['周次', '运单号', '金额', '金额(数字)'], ['W43', 'P001', '12,5'], ['W43', 'P001', '3,5'], ['W43', 'P002', '8'], ['W44', 'P003', '10,25'], ['W44', null, '1,75']]) },
        { name: "L'OREAL", cells: T([['周次', '运单号', '金额'], ['W43', 'L01', '20,5'], ['W44', 'L02', '5,5'], ['W44', 'L03', 'abc'], ['W44', 'L02', '4']]) }
      ],
      exercises: [
        { id: 's7-0', title: '逗号小数转数字', sheet: '客户A', target: 'D2:D6', fns: ['VALUE', 'SUBSTITUTE'],
          ask: '在「客户A」表 D 列：把 C 列的 12,5 变成能计算的数字 12.5。',
          hints: ['直接 «VALUE(C2)» 会报错，因为逗号不被当成小数点。', '先 «SUBSTITUTE(C2, ",", ".")» 把逗号换成点，再 VALUE。'],
          diag: f => (!/SUBSTITUTE/i.test(f) ? '12,5 里的逗号不会被当成小数点，所以 VALUE 转不了。先用 SUBSTITUTE 把逗号换成点。' : null),
          answer: '=VALUE(SUBSTITUTE(C2,",","."))', steps: [['SUBSTITUTE(C2,",",".")', '12,5 → 12.5（还是文本）'], ['VALUE( … )', '文本 → 数字']] },
        { id: 's7-1', title: '把表名处理安全', target: 'G2:G3', fns: ['SUBSTITUTE'],
          ask: "G 列：拼出可以放进引用里的表名：两边加单引号，表名里自带的单引号要写成两个。L'OREAL 应该变成 'L''OREAL'。",
          hints: ['«SUBSTITUTE(A2, "\'", "\'\'")» 把一个单引号换成两个。', '两边再用 & 接上单引号：«"\'"&…&"\'"»。'],
          answer: "=\"'\"&SUBSTITUTE(A2,\"'\",\"''\")&\"'\"",
          steps: [['SUBSTITUTE(A2,"\'","\'\'")', "L'OREAL → L''OREAL"], ['"\'" & … & "\'"', '两边包上单引号']] },
        { id: 's7-2', title: '某周有几个不同运单', target: 'B2:C3', fns: ['LET', 'SUBSTITUTE', 'INDIRECT', 'IFERROR', 'ROWS', 'UNIQUE', 'FILTER'], challenge: true,
          ask: 'B2:C3：去 A 列那张表里，数「周次 = 表头周」的不同运单号有几个（空运单不算）。表头写的是 W43运单，周次取前 3 个字。',
          hints: ['表头周次：«LEFT(B$1,3)»。', "y = «INDIRECT(sh&\"!$B:$B\")»，x = «INDIRECT(sh&\"!$A:$A\")»。", '«ROWS(UNIQUE(FILTER(y,(x=周次)*(y<>""))))»，外面 IFERROR(…,0)。'],
          answer: "=LET(sh,\"'\"&SUBSTITUTE($A2,\"'\",\"''\")&\"'\",y,INDIRECT(sh&\"!$B:$B\"),x,INDIRECT(sh&\"!$A:$A\"),IFERROR(ROWS(UNIQUE(FILTER(y,(x=LEFT(B$1,3))*(y<>\"\")))),0))",
          steps: [['sh', '安全表名'], ['INDIRECT(sh&"!$B:$B")', '文字地址 → 真实引用'], ['(x=周次)*(y<>"")', '两个条件相乘 = 同时满足'], ['ROWS(UNIQUE(FILTER(…)))', '筛选、去重、数个数']] },
        { id: 's7-3', title: '某周金额合计（逗号小数）', target: 'D2:E3', fns: ['LET', 'INDIRECT', 'SUM', 'FILTER', 'ARRAYFORMULA', 'IFERROR', 'VALUE', 'SUBSTITUTE'], challenge: true,
          ask: 'D2:E3：去 A 列那张表里，把「周次 = 表头周」的金额加起来。金额是逗号小数，转不了数字的当 0。',
          hints: ['整列清洗：«ARRAYFORMULA(IFERROR(VALUE(SUBSTITUTE(n,",",".")),0))»。', '«SUM(FILTER(清洗后的金额, x=周次))»。'],
          answer: "=LET(sh,\"'\"&SUBSTITUTE($A2,\"'\",\"''\")&\"'\",n,INDIRECT(sh&\"!$C:$C\"),x,INDIRECT(sh&\"!$A:$A\"),IFERROR(SUM(FILTER(ARRAYFORMULA(IFERROR(VALUE(SUBSTITUTE(n,\",\",\".\")),0)),x=LEFT(D$1,3))),0))",
          steps: [['SUBSTITUTE → VALUE → IFERROR', '清洗链路：换小数点、转数字、坏数据当 0'], ['FILTER(…, x=周次)', '只留这一周'], ['SUM', '加总']] }
      ]
    },

    /* ---------- 场景 9：组合键 + MAP ---------- */
    {
      id: 's9', area: '场景 9', link: 'index.html#scenario9', title: '客户城市分析', sub: '组合键、FILTER + TEXTJOIN、MAP + LAMBDA',
      story: '同一个城市可能有好几个 BI 客户，客户和城市必须成对匹配。场景 4 的三道闸门会把「CIRRO-PU + LISSES」也算进来（这一对在映射表里不存在），用组合键就能避免。',
      sheets: [
        { name: '城市分析', cells: T([
          ['城市', 'BI客户', '映射客户', 'MAP版', null, '客户', '票数(组合键)'],
          ['LISSES', 'CIRRO-Temu-BT-PU', null, null, null, 'GOODCANG'], ['LISSES', 'GOODCANG-FR', null, null, null, 'CIRRO'],
          ['ROISSY', 'CIRRO-PU'], ['LISSES', 'CIRRO-PU'], ['EVRY', 'goodcang-fr ']
        ]) },
        mapSheet,
        { name: 'BI数据源', cells: T(biRows.map((r, i) => (i === 0 ? [...r, '组合键'] : r))) }
      ],
      exercises: [
        { id: 's9-1', title: '做组合键', sheet: 'BI数据源', target: 'D2:D11', fns: ['LOWER', 'TRIM'],
          ask: '在「BI数据源」表 D 列：把客户和城市拼成一个组合键，统一小写、去掉空格，中间用 ♦ 隔开，例如 goodcang-fr♦lisses。',
          hints: ['«LOWER(TRIM(B2))» 统一格式。', '用 «&"♦"&» 把两段接起来。'],
          answer: '=LOWER(TRIM(B2))&"♦"&LOWER(TRIM(C2))', steps: [['LOWER(TRIM(…))', '去空格、统一小写'], ['&"♦"&', '用不会出现在名字里的符号隔开']] },
        { id: 's9-2', title: '按城市 + 客户找映射客户', target: 'C2:C6', fns: ['IFNA', 'TEXTJOIN', 'UNIQUE', 'FILTER', 'LOWER', 'TRIM'],
          ask: 'C 列：去映射表里找「BI 客户名和城市都对得上」的客户名称，有多个用「 / 」连起来，一个都没有显示 NA。注意第 6 行大小写和空格都不规范。',
          hints: ['FILTER 可以写两个条件：客户相同、城市相同，必须在同一行同时满足。', '两边都用 «LOWER(TRIM(…))» 再比较。', '«IFNA(TEXTJOIN(" / ",TRUE,UNIQUE(FILTER(…))),"NA")»'],
          answer: "=IFNA(TEXTJOIN(\" / \",TRUE,UNIQUE(FILTER('映射表'!$A$2:$A$6,LOWER(TRIM('映射表'!$B$2:$B$6))=LOWER(TRIM(B2)),LOWER(TRIM('映射表'!$C$2:$C$6))=LOWER(TRIM(A2))))),\"NA\")",
          steps: [['FILTER(客户名称列, 条件1, 条件2)', '两道筛子同时满足'], ['UNIQUE', '同名只留一次'], ['TEXTJOIN(" / ",TRUE, …)', '多个结果都展示'], ['IFNA( … ,"NA")', '找不到时明确报警']] },
        { id: 's9-3', title: '用 MAP 一个公式生成整列', target: 'D2', spill: true, fns: ['MAP', 'LAMBDA', 'IF', 'IFNA', 'TEXTJOIN', 'UNIQUE', 'FILTER', 'LOWER', 'TRIM'], challenge: true, require: ['MAP', 'LAMBDA'],
          ask: 'D2：用 MAP + LAMBDA 写一个公式，同时读取 A2:A6 和 B2:B6，逐行做和 C 列一样的查找，结果自动溢出到 D6。',
          hints: ['«MAP(A2:A6, B2:B6, LAMBDA(city, client, …))»', 'LAMBDA 里把上一题公式的 A2 换成 city、B2 换成 client。', '空行挡住：«IF(city="","", …)»。'],
          answer: "=MAP(A2:A6,B2:B6,LAMBDA(city,client,IF(city=\"\",\"\",IFNA(TEXTJOIN(\" / \",TRUE,UNIQUE(FILTER('映射表'!$A$2:$A$6,LOWER(TRIM('映射表'!$B$2:$B$6))=LOWER(TRIM(client)),LOWER(TRIM('映射表'!$C$2:$C$6))=LOWER(TRIM(city))))),\"NA\"))))",
          steps: [['MAP(A2:A6, B2:B6, …)', '流水线：每一行送进加工台'], ['LAMBDA(city, client, …)', '给当前行的两个值起名字'], ['IF(city="","", …)', '空行不算']] },
        { id: 's9-4', title: '用组合键重新计票', target: 'G2:G3', fns: ['LET', 'UNIQUE', 'FILTER', 'LOWER', 'SUMPRODUCT', 'N', 'ISNUMBER', 'MATCH'], challenge: true,
          ask: 'G 列：统计 W34 这个客户的票数，但这次「客户 + 城市」必须作为一对出现在映射表里。对比一下：CIRRO 应该比场景 4 少 1 票。',
          hints: ['映射表的组合键清单：«UNIQUE(FILTER(LOWER(B列)&"♦"&LOWER(C列), A列=客户))»。', 'BI 的组合键：«LOWER(客户列)&"♦"&LOWER(城市列)»。', '«SUMPRODUCT(N(周次="W34"), N(ISNUMBER(MATCH(BI组合键, 清单, 0))))»'],
          answer: "=LET(k,TRIM(F2),keys,UNIQUE(FILTER(LOWER('映射表'!$B$2:$B$6)&\"♦\"&LOWER('映射表'!$C$2:$C$6),'映射表'!$A$2:$A$6=k)),SUMPRODUCT(N('BI数据源'!$A$2:$A$11=\"W34\"),N(ISNUMBER(MATCH(LOWER('BI数据源'!$B$2:$B$11)&\"♦\"&LOWER('BI数据源'!$C$2:$C$11),keys,0)))))",
          steps: [['keys', '映射表里合法的「客户♦城市」组合'], ['MATCH(BI组合键, keys, 0)', '这一对存在吗'], ['SUMPRODUCT(N(…), N(…))', '周次对 + 组合存在才计数']] }
      ]
    }
  ];

  // 回忆题：引擎不能运行的（QUERY 等）或适合口头回忆的
  const RECALL = [
    { id: 'r-q1', area: '场景 9', link: 'index.html#scenario9', fns: ['QUERY'], q: 'FILTER 已经筛出了 {城市, BI客户} 两列明细，要按城市 + BI 客户分组计数，QUERY 的查询语句怎么写？', a: ['select Col1,Col2,count(Col1) group by Col1,Col2'] },
    { id: 'r-q2', area: '场景 9', link: 'index.html#scenario9', fns: ['QUERY'], q: 'QUERY 的输入是 FILTER 的结果（没有标题行），第 3 个参数「标题行数」写几？', a: ['0'] },
    { id: 'r-q3', area: '场景 9', link: 'index.html#scenario9', fns: ['QUERY'], q: '不想要 QUERY 自动生成的标题（城市、客户、count），查询语句最后加什么？', a: ["label Col1 '',Col2 '',count(Col1) ''"] },
    { id: 'r-q4', area: '场景 9', link: 'index.html#scenario9', fns: ['QUERY'], q: 'select 里出现、但没有被 count / sum 处理的列，必须同时写在哪个子句里？', a: ['group by'] },
    { id: 'r-q5', area: '场景 9', link: 'index.html#scenario9', fns: ['QUERY'], q: '在 QUERY({D4:D100, F4:F100}, …) 里，城市（D 列）叫 Col 几？', a: ['Col1'] },
    { id: 'r-q6', area: '场景 9', link: 'index.html#scenario9', fns: ['QUERY'], q: '只按 Col1（城市）分组，会出现什么问题？', a: ['同一城市的不同 BI 客户会被合并成一行'] },
    { id: 'r-i1', area: '场景 7', link: 'index.html#scenario7', fns: ['INDIRECT', 'SUBSTITUTE'], q: "没有 INDIRECT 时，\"'客户A'!$Y:$Y\" 只是什么？", a: ['一段文字'] },
    { id: 'r-f1', area: '通用', fns: ['IFNA', 'IFERROR'], q: 'IFNA 和 IFERROR 怎么选？', a: ['查找类只处理找不到时用 IFNA；要接住各种报错才用 IFERROR'] },
    { id: 'r-f2', area: '通用', fns: ['WORKDAY.INTL', 'NETWORKDAYS.INTL'], q: 'WORKDAY.INTL 和 NETWORKDAYS.INTL 的区别，用一句话说？', a: ['WORKDAY.INTL 找日子，NETWORKDAYS.INTL 数日子'] },
    { id: 'r-f3', area: '通用', fns: ['MAP', 'LAMBDA'], q: 'MAP 输入了两列，LAMBDA 要定义几个名字？', a: ['2', '两个'] },
    { id: 'r-f4', area: '通用', fns: ['ARRAYFORMULA', 'MAP'], q: '什么情况下用 MAP + LAMBDA，而不是每行写一个公式往下拖？', a: ['行数会变（比如 QUERY 结果行数随周次变化），要一个公式自动生成整列'] }
  ];

  G.PRACTICE_BANK = { SCENARIOS, RECALL };
})(typeof window !== 'undefined' ? window : globalThis);
