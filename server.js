// 中山市香山小学 · 信息科组 电脑室调课系统 —— 零依赖 Node 后端
// 数据模型：2 个电脑室 × 周一~周五 × 上午1~4 / 下午1~2 × 20 周，每周从基础模板复制。
// 持久化：data.json（串行写，避免并发损坏）。

const http = require('http');
const fs = require('fs');
const path = require('path');
const { URL } = require('url');
const XLSX = require('xlsx');

const ROOT = __dirname;
const PUBLIC = path.join(ROOT, 'public');
const DATA_FILE = path.join(ROOT, 'data.json');
const PORT = process.env.PORT || 3000;

const DAYS = ['周一', '周二', '周三', '周四', '周五'];
const PERIODS = ['上午1', '上午2', '上午3', '上午4', '下午1', '下午2'];
const LABS = [{ id: 1, name: '电脑室1' }, { id: 2, name: '电脑室2' }];
const WEEKS = 20;
const TERM = '中山市香山小学 2026学年第一学期';
// 学期第1周起始日（周一）。锚点：2026-10-13 为第7周 → 第1周周一 = 2026-08-31。
// 每学年可在「导入基础课表」里改，自动周次随之更新。
const TERM_START_DEFAULT = '2026-08-31';

// 基础课表（按学校提供的两张课表照片录入；每学年换一份文件即可复用）
function C(cls, tea) { return { cls: cls, tea: tea }; }
const BASE_SCHEDULE = {
  // ===== 电脑室1（邝晓璇 / 郑嘉乐）=====
  1: {
    周一: { 上午1: null, 上午2: C('五年级(10)', '邝晓璇'), 上午3: C('四年级(12)', '邝晓璇'), 上午4: C('五年级(3)', '邝晓璇'), 下午1: null, 下午2: C('五年级(7)', '邝晓璇') },
    周二: { 上午1: null, 上午2: C('五年级(6)', '邝晓璇'), 上午3: C('五年级(1)', '邝晓璇'), 上午4: C('三年级(5)', '郑嘉乐'), 下午1: C('三年级(2)', '郑嘉乐'), 下午2: C('四年级(9)', '邝晓璇') },
    周三: { 上午1: C('五年级(5)', '邝晓璇'), 上午2: C('四年级(13)', '郑嘉乐'), 上午3: C('五年级(2)', '邝晓璇'), 上午4: C('三年级(3)', '郑嘉乐'), 下午1: C('三年级(1)', '郑嘉乐'), 下午2: C('四年级(11)', '邝晓璇') },
    周四: { 上午1: null, 上午2: C('五年级(9)', '邝晓璇'), 上午3: C('五年级(8)', '邝晓璇'), 上午4: C('三年级(4)', '郑嘉乐'), 下午1: C('三年级(6)', '郑嘉乐'), 下午2: C('四年级(8)', '邝晓璇') },
    周五: { 上午1: null, 上午2: null, 上午3: C('五年级(4)', '邝晓璇'), 上午4: C('四年级(10)', '邝晓璇'), 下午1: null, 下午2: null }
  },
  // ===== 电脑室2（张琳 / 郑嘉乐）=====
  2: {
    周一: { 上午1: null, 上午2: C('四年级(15)', '郑嘉乐'), 上午3: C('六年级(6)', '张琳'), 上午4: C('三年级(9)', '郑嘉乐'), 下午1: C('六年级(2)', '张琳'), 下午2: C('四年级(2)', '张琳') },
    周二: { 上午1: C('四年级(6)', '张琳'), 上午2: C('四年级(3)', '张琳'), 上午3: null, 上午4: null, 下午1: C('四年级(5)', '张琳'), 下午2: C('六年级(3)', '张琳') },
    周三: { 上午1: C('六年级(7)', '张琳'), 上午2: C('六年级(5)', '张琳'), 上午3: C('四年级(16)', '郑嘉乐'), 上午4: C('六年级(1)', '张琳'), 下午1: C('六年级(4)', '张琳'), 下午2: C('三年级(8)', '郑嘉乐') },
    周四: { 上午1: null, 上午2: null, 上午3: C('四年级(7)', '张琳'), 上午4: null, 下午1: C('六年级(8)', '张琳'), 下午2: C('四年级(1)', '张琳') },
    周五: { 上午1: C('四年级(4)', '张琳'), 上午2: null, 上午3: C('四年级(14)', '郑嘉乐'), 上午4: C('三年级(7)', '郑嘉乐'), 下午1: null, 下午2: null }
  }
};

function buildScheduleFromBase(base) {
  const sched = {};
  for (let w = 1; w <= WEEKS; w++) {
    sched[w] = {};
    for (const lab of LABS) {
      sched[w][lab.id] = {};
      for (const d of DAYS) {
        sched[w][lab.id][d] = {};
        for (const p of PERIODS) {
          const c = base[lab.id] && base[lab.id][d] && base[lab.id][d][p];
          sched[w][lab.id][d][p] = c ? { cls: c.cls, tea: c.tea } : null;
        }
      }
    }
  }
  return sched;
}

let db;
function loadDb() {
  if (fs.existsSync(DATA_FILE)) {
    try { db = JSON.parse(fs.readFileSync(DATA_FILE, 'utf8')); }
    catch (e) { console.error('数据文件损坏，重建', e); }
  }
  if (!db || !db.schedule || !db.schedule[1] || !db.periods || db.periods[0] !== PERIODS[0]) {
    db = {
      term: TERM,
      termStart: TERM_START_DEFAULT,
      labs: LABS, days: DAYS, periods: PERIODS, weeks: WEEKS,
      base: BASE_SCHEDULE,
      schedule: buildScheduleFromBase(BASE_SCHEDULE),
      log: []
    };
    saveDbSync();
  }
  if (!db.termStart) db.termStart = TERM_START_DEFAULT;
}

function parseDate(s) {
  if (!s) return null;
  const m = /^(\d{4})-(\d{1,2})-(\d{1,2})$/.exec(s);
  if (!m) return null;
  return new Date(Number(m[1]), Number(m[2]) - 1, Number(m[3]));
}
function atMid(d) { return new Date(d.getFullYear(), d.getMonth(), d.getDate()); }
// 自动当前周：以 termStart（周一）为第1周起点，按自然周推算，落在 1~WEEKS
function currentWeek() {
  const start = atMid(parseDate(db.termStart) || parseDate(TERM_START_DEFAULT));
  const days = Math.round((atMid(new Date()) - start) / 86400000);
  let w = Math.floor(days / 7) + 1;
  if (w < 1) w = 1;
  if (w > WEEKS) w = WEEKS;
  return w;
}
// 某周的日期范围（周一~周日）
function weekRange(week) {
  const start = atMid(parseDate(db.termStart) || parseDate(TERM_START_DEFAULT));
  const mon = new Date(start.getTime() + (week - 1) * 7 * 86400000);
  const sun = new Date(mon.getTime() + 6 * 86400000);
  const fmt = d => (d.getMonth() + 1) + '月' + d.getDate() + '日';
  return { mon: fmt(mon), sun: fmt(sun) };
}

let writeChain = Promise.resolve();
function saveDb() {
  writeChain = writeChain.then(() => new Promise((res, rej) => {
    fs.writeFile(DATA_FILE, JSON.stringify(db), (err) => err ? rej(err) : res());
  }));
  return writeChain;
}
function saveDbSync() { fs.writeFileSync(DATA_FILE, JSON.stringify(db)); }

function labName(id) { const l = LABS.find(x => x.id === id); return l ? l.name : ('电脑室' + id); }

// 取某周某格内容（addr 含 week）
function cellAt(addr) {
  const w = db.schedule[addr.week];
  return w ? w[addr.lab][addr.day][addr.period] : null;
}
function descAddr(a) { return '第' + a.week + '周 ' + labName(a.lab) + ' ' + a.day + a.period; }
function sameAddr(a, b) { return a.lab === b.lab && a.day === b.day && a.period === b.period && a.week === b.week; }
// 把操作映射到指定周（同周批量用：a/b 的 week 都置为 w）
function withWeek(op, w) {
  return Object.assign({}, op, {
    a: Object.assign({}, op.a, { week: w }),
    b: op.b ? Object.assign({}, op.b, { week: w }) : null
  });
}

// 单步校验（基于 op.a.week / op.b.week 访问具体周；支持跨周）
function checkOp(op) {
  if (op.type === 'swap') {
    const A = cellAt(op.a), B = cellAt(op.b);
    if (!A) return { ok: false, error: descAddr(op.a) + ' 无课，无法作为互换源' };
    if (!B) return { ok: false, error: descAddr(op.b) + ' 无课（互换需两节都有课）' };
    if (sameAddr(op.a, op.b)) return { ok: false, error: '不能与自己互换' };
    return { ok: true };
  }
  if (op.type === 'move') {
    const A = cellAt(op.a), B = cellAt(op.b);
    if (!A) return { ok: false, error: descAddr(op.a) + ' 无课，无法作为移动源' };
    if (B) return { ok: false, error: descAddr(op.b) + ' 已有课，不能移入（请改为互换）' };
    return { ok: true };
  }
  if (op.type === 'edit') {
    const A = cellAt(op.a);
    if (!A) return { ok: false, error: descAddr(op.a) + ' 无课，无法编辑' };
    if (!op.tea || !op.tea.trim()) return { ok: false, error: '教师不能为空' };
    return { ok: true };
  }
  return { ok: false, error: '未知操作类型' };
}

// 在指定 schedule 上应用（跨周/同周统一：swap/move 用各自的 week，edit 用 a.week）
function applyOpOnSchedule(schedule, op) {
  if (op.type === 'swap') {
    const wA = schedule[op.a.week], wB = schedule[op.b.week];
    const A = wA[op.a.lab][op.a.day][op.a.period];
    const B = wB[op.b.lab][op.b.day][op.b.period];
    wB[op.b.lab][op.b.day][op.b.period] = A;
    wA[op.a.lab][op.a.day][op.a.period] = B;
  } else if (op.type === 'move') {
    const wA = schedule[op.a.week], wB = schedule[op.b.week];
    const A = wA[op.a.lab][op.a.day][op.a.period];
    wB[op.b.lab][op.b.day][op.b.period] = A;
    wA[op.a.lab][op.a.day][op.a.period] = null;
  } else if (op.type === 'edit') {
    const A = schedule[op.a.week][op.a.lab][op.a.day][op.a.period];
    if (op.cls !== undefined && op.cls.trim() !== '') A.cls = op.cls.trim();
    A.tea = op.tea.trim();
  }
}
function cloneWeek(w) { return JSON.parse(JSON.stringify(db.schedule[w])); }

// 教师冲突：同一位老师在同一「星期+节次」被排到两个电脑室（跨室才会发生）
function teacherConflicts(sw) {
  const out = [];
  for (const d of DAYS) {
    for (const p of PERIODS) {
      const a = sw[1][d][p], b = sw[2][d][p];
      if (a && b && a.tea && b.tea && a.tea === b.tea) {
        out.push({
          key: d + '|' + p + '|' + a.tea,
          msg: d + p + '：' + a.tea + ' 老师同时在两间电脑室有课（' + a.cls + ' / ' + b.cls + '）'
        });
      }
    }
  }
  return out;
}
// 模拟操作后的新增教师冲突（cross=true 时只涉及 op 真实的两个周；否则对每个目标周用 withWeek）
function newTeacherConflicts(op, weeks, cross) {
  const before = new Set();
  for (const w of weeks) teacherConflicts(db.schedule[w]).forEach(x => before.add(w + '|' + x.key));
  const sim = {};
  for (const w of weeks) sim[w] = cloneWeek(w);
  if (cross) applyOpOnSchedule(sim, op);
  else for (const w of weeks) applyOpOnSchedule(sim, withWeek(op, w));
  const out = [];
  for (const w of weeks) teacherConflicts(sim[w]).forEach(x => {
    if (!before.has(w + '|' + x.key)) out.push({ key: w + '|' + x.key, msg: '第' + w + '周 ' + x.msg });
  });
  return out;
}

function rangeLabel(weeks) {
  if (weeks.length === 1) return '第' + weeks[0] + '周';
  const sorted = [...weeks].sort((x, y) => x - y);
  const min = sorted[0], max = sorted[sorted.length - 1];
  const consecutive = sorted.every((v, i) => v === min + i);
  if (consecutive) return '第' + min + '~' + max + '周';
  return '第' + sorted.join('、') + '周';
}

function parseCell(raw) {
  const s = (raw == null ? '' : String(raw)).trim();
  if (s === '' || s === '空' || s === '-' || s === '—') return null;
  const parts = s.split(/[／/]/);
  return { cls: parts[0].trim(), tea: (parts[1] || '').trim() };
}
function gridToBase(rows) {
  const grid = {};
  DAYS.forEach(d => grid[d] = {});
  if (!rows || !rows.length) return grid;
  let start = 0;
  if (rows[0] && Array.isArray(rows[0]) && rows[0].some(c => typeof c === 'string' && c.includes('周一'))) start = 1;
  for (let i = 0; i < PERIODS.length; i++) {
    const row = rows[start + i];
    if (!row) continue;
    for (let j = 0; j < DAYS.length; j++) {
      grid[DAYS[j]][PERIODS[i]] = parseCell(row[j]);
    }
  }
  return grid;
}
function countCells(grid) { let n = 0; DAYS.forEach(d => PERIODS.forEach(p => { if (grid[d][p]) n++; })); return n; }

function detailOf(op, cross) {
  if (cross) {
    if (op.type === 'swap') return descAddr(op.a) + ' ↔ ' + descAddr(op.b) + '（互换）';
    if (op.type === 'move') return descAddr(op.a) + ' → ' + descAddr(op.b) + '（移走，原格留空）';
    if (op.type === 'edit') return descAddr(op.a) + ' 教师改为 ' + op.tea;
  }
  const a = op.a;
  if (op.type === 'swap') {
    const A = cellAt(op.a);
    return labName(a.lab) + ' ' + a.day + a.period + ' ↔ ' + labName(op.b.lab) + ' ' + op.b.day + op.b.period +
      '（' + (A ? A.cls : '空') + ' 互换）';
  }
  if (op.type === 'move') {
    const A = cellAt(op.a);
    return labName(a.lab) + ' ' + a.day + a.period + ' → ' + labName(op.b.lab) + ' ' + op.b.day + op.b.period +
      '（' + (A ? A.cls : '') + ' 移走，原格留空）';
  }
  if (op.type === 'edit') {
    const A = cellAt(op.a);
    return labName(a.lab) + ' ' + a.day + a.period + ' 教师改为 ' + op.tea + '（班级 ' + (A ? A.cls : '') + '）';
  }
  return '';
}

// ---------- HTTP ----------
function sendJson(res, code, obj) {
  const s = JSON.stringify(obj);
  res.writeHead(code, { 'Content-Type': 'application/json; charset=utf-8' });
  res.end(s);
}
function serveFile(res, file, type) {
  fs.readFile(file, (err, data) => {
    if (err) { res.writeHead(404); res.end('not found'); return; }
    res.writeHead(200, {
      'Content-Type': type,
      'Cache-Control': 'no-cache, no-store, must-revalidate',
      'Pragma': 'no-cache',
      'Expires': '0'
    });
    res.end(data);
  });
}
function readBody(req, cb) {
  let body = '';
  req.on('data', c => body += c);
  req.on('end', () => {
    try { cb(JSON.parse(body || '{}')); }
    catch (e) { cb({}); }
  });
}

const server = http.createServer((req, res) => {
  const u = new URL(req.url, 'http://localhost');

  if (req.method === 'GET' && (u.pathname === '/' || u.pathname === '/index.html')) {
    return serveFile(res, path.join(PUBLIC, 'index.html'), 'text/html; charset=utf-8');
  }

  if (u.pathname === '/api/state') {
    const week = Math.min(WEEKS, Math.max(1, parseInt(u.searchParams.get('week') || '1', 10)));
    const cw = currentWeek();
    return sendJson(res, 200, {
      term: db.term || TERM,
      termStart: db.termStart,
      currentWeek: cw,
      labs: db.labs, days: db.days, periods: db.periods, weeks: db.weeks,
      week, schedule: db.schedule[week], weekRange: weekRange(week),
      log: db.log.slice(-80).reverse()
    });
  }

  if (req.method === 'POST' && u.pathname === '/api/op') {
    return readBody(req, (op) => {
      if (!op.who || !op.who.trim()) return sendJson(res, 400, { error: '请先输入操作人姓名' });
      if (!op.a || !op.a.lab || !op.a.day || !op.a.period) return sendJson(res, 400, { error: '未选择第一节课' });
      if ((op.type === 'swap' || op.type === 'move') && (!op.b || !op.b.lab || !op.b.day || !op.b.period))
        return sendJson(res, 400, { error: '未选择第二节/目标格' });
      const aw = Number(op.a.week) || 1;
      const bw = Number((op.b && op.b.week) || aw);
      op.a.week = aw;
      if (op.b) op.b.week = bw;
      // 跨周：两格来自不同周，仅在这两个具体周之间操作（weeks 不生效）
      const cross = (op.type === 'swap' || op.type === 'move') && aw !== bw;

      let weeks;
      if (cross) weeks = [aw, bw].filter(w => w >= 1 && w <= db.weeks);
      else weeks = (op.weeks && op.weeks.length ? op.weeks : [aw]).map(Number).filter(w => w >= 1 && w <= db.weeks);
      if (!weeks.length) return sendJson(res, 400, { error: '未指定有效周次' });

      // 先校验：跨周整体校验一次；同周对每个目标周校验
      if (cross) {
        const r = checkOp(op);
        if (!r.ok) return sendJson(res, 400, { error: r.error });
      } else {
        for (const w of weeks) {
          const r = checkOp(withWeek(op, w));
          if (!r.ok) return sendJson(res, 400, { error: '第' + w + '周：' + r.error });
        }
      }
      // 再校验：不能造成「同一位老师同一时段被排到两个电脑室」
      const conf = newTeacherConflicts(op, weeks, cross);
      if (conf.length) {
        return sendJson(res, 400, { error: '教师冲突：' + conf[0].msg + '。请改选其他时段。' });
      }

      // 跨周：op 在两具体周之间只应用一次；同周：对每个目标周用 withWeek 应用
      if (cross) applyOpOnSchedule(db.schedule, op);
      else for (const w of weeks) applyOpOnSchedule(db.schedule, withWeek(op, w));

      const rangeTxt = cross ? ('第' + aw + '周↔第' + bw + '周') : rangeLabel(weeks);
      db.log.push({
        ts: new Date().toISOString(), who: op.who.trim(),
        type: op.type, weeks, range: rangeTxt, detail: detailOf(op, cross)
      });
      saveDb().then(() => sendJson(res, 200, { ok: true }))
        .catch(e => sendJson(res, 500, { error: '保存失败：' + e.message }));
      return;
    });
  }

  if (req.method === 'POST' && u.pathname === '/api/import') {
    return readBody(req, (payload) => {
      const base = payload.base;
      if (!base || !base[1] || !base[2]) return sendJson(res, 400, { error: '基础课表格式不正确' });
      db.base = base;
      if (payload.term) db.term = payload.term;
      if (payload.termStart && parseDate(payload.termStart)) db.termStart = payload.termStart;
      db.schedule = buildScheduleFromBase(base);
      db.log.push({ ts: new Date().toISOString(), who: payload.who || '管理员', type: 'import', weeks: [], range: '全部20周', detail: '导入基础课表，已重置全部周次' });
      saveDb().then(() => sendJson(res, 200, { ok: true }))
        .catch(e => sendJson(res, 500, { error: '保存失败：' + e.message }));
      return;
    });
  }

  if (req.method === 'POST' && u.pathname === '/api/import-file') {
    return readBody(req, (payload) => {
      try {
        const base = { 1: null, 2: null };
        for (const f of (payload.files || [])) {
          const lab = f.lab === 2 ? 2 : 1;
          const buf = Buffer.from(f.data || '', 'base64');
          let rows;
          const name = (f.name || '').toLowerCase();
          if (name.endsWith('.csv') || name.endsWith('.txt')) {
            const text = buf.toString('utf8');
            rows = text.split(/\r?\n/).map(l => l.split(/[,，]/).map(s => s.trim()));
          } else {
            const wb = XLSX.read(buf, { type: 'buffer' });
            const ws = wb.Sheets[wb.SheetNames[0]];
            rows = XLSX.utils.sheet_to_json(ws, { header: 1, defval: '', raw: false });
          }
          base[lab] = gridToBase(rows);
        }
        if (!base[1] || !base[2]) return sendJson(res, 400, { error: '请提供电脑室1和电脑室2两个文件' });
        db.base = base;
        if (payload.termStart && parseDate(payload.termStart)) db.termStart = payload.termStart;
        db.schedule = buildScheduleFromBase(base);
        db.log.push({ ts: new Date().toISOString(), who: payload.who || '管理员', type: 'import', weeks: [], range: '全部20周', detail: '导入基础课表文件，已重置全部周次' });
        saveDb().then(() => sendJson(res, 200, { ok: true, cells1: countCells(base[1]), cells2: countCells(base[2]) }))
          .catch(e => sendJson(res, 500, { error: '保存失败：' + e.message }));
      } catch (e) {
        return sendJson(res, 400, { error: '解析失败：' + e.message });
      }
    });
  }

  if (req.method === 'GET' && u.pathname === '/api/export-log') {
    const termTxt = db.term || '未设置学期';
    const aoa = [['调课记录导出 — ' + termTxt], []];
    aoa.push(['时间', '操作人', '变动内容', '范围', '类型']);
    const typeTxt = { swap: '互换', move: '调位', edit: '改教师', import: '导入' };
    (db.log || []).forEach(e => {
      aoa.push([
        (e.ts || '').replace('T', ' ').slice(0, 16),
        e.who || '',
        e.detail || '',
        e.range || '',
        typeTxt[e.type] || e.type || ''
      ]);
    });
    const ws = XLSX.utils.aoa_to_sheet(aoa);
    const wb = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(wb, ws, '调课记录');
    const buf = XLSX.write(wb, { type: 'buffer', bookType: 'xlsx' });
    res.writeHead(200, {
      'Content-Type': 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
      'Cache-Control': 'no-cache, no-store, must-revalidate',
      'Content-Disposition': (/MicroMessenger/i.test(req.headers['user-agent'] || '') ? 'inline' : 'attachment') + '; filename="tiaoke-jilu.xlsx"'
    });
    res.end(buf);
    return;
  }

  if (req.method === 'POST' && u.pathname === '/api/settings') {
    return readBody(req, (payload) => {
      if (payload.termStart && parseDate(payload.termStart)) db.termStart = payload.termStart;
      if (payload.term && payload.term.trim()) db.term = payload.term.trim();
      saveDb().then(() => sendJson(res, 200, { ok: true, termStart: db.termStart, currentWeek: currentWeek() }))
        .catch(e => sendJson(res, 500, { error: '保存失败：' + e.message }));
    });
  }

  res.writeHead(404); res.end('not found');
});

loadDb();
server.listen(PORT, '0.0.0.0', () => console.log('电脑室调课系统已启动: http://0.0.0.0:' + PORT));
