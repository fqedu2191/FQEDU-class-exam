/******************************************************
 * 本週挑戰｜老師後台
 *
 * 這是「獨立」的 Apps Script 專案，跟學生作答用的後端分開，
 * 不會改到學生那支程式；兩者讀寫同一份試算表。
 *
 * 資料規則：
 *  - Responses 是唯一的「正本」（每次作答的完整紀錄）。
 *  - Scores 是排行榜 API 用的副本。後台每次修改 Responses 後，
 *    都會依 Responses 重新產生 Scores，讓兩邊永遠一致。
 ******************************************************/

/* ========== 設定區 ========== */
const SPREADSHEET_ID = '1SPyCYVfxBrzOnB8RRtK7VOmZI8IXtr89b0feTf9Sd-I';
const SHEET_ANS = 'AnswerKey';
const SHEET_RES = 'Responses';
const SHEET_SCO = 'Scores';
const POINTS_PER_CORRECT = 1;   // 與學生後端相同：每個市場答對 1 分

const RES_HEADERS = ['Timestamp','Group','QNO','Story','STOCK','BOND','FX','COM','RE','Score'];
const SCO_HEADERS = ['Timestamp','Group','QNO','Score'];
const MARKETS = ['STOCK','BOND','FX','COM','RE'];

const ANS_ALIASES = {
  QNO: ['QNO','題號','QUESTIONNO','Q_NO','NO'],
  QUESTIONTEXT: ['QUESTIONTEXT','QUESTIONTEST','題目','題幹','QUESTION','QUESTION_TEXT'],
  STOCK: ['STOCK','股市'],
  BOND:  ['BOND','債市','債券','債卷'],
  FX:    ['FX','外匯','DXY','美元指數'],
  COM:   ['COM','商品','COMMODITY','原物料'],
  RE:    ['RE','房地產','REAL ESTATE','REAL_ESTATE'],
};

/* ========== 網頁入口 ========== */
function doGet() {
  return HtmlService.createHtmlOutputFromFile('Dashboard')
    .setTitle('本週挑戰｜老師後台')
    .addMetaTag('viewport', 'width=device-width, initial-scale=1')
    // 允許嵌入 Google 協作平台
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

/* ========== 前端可呼叫的功能 ========== */

/** 讀取全部資料 */
function getDashboardData() {
  return buildPayload_(ss_());
}

/** 修改一筆作答的組別、題號、分數、備註 */
function updateResponse(ref, changes) {
  return withLock_(function () {
    const ss = ss_();
    const res = assertConsistent_(ss);
    const r = findRow_(res, ref);

    const group = S_(changes.group);
    const qno = normQno_(changes.qno);
    const score = Number(changes.score);
    if (!group) throw new Error('組別不可空白');
    if (!qno) throw new Error('題號不可空白');
    if (S_(changes.score) === '' || !isFinite(score)) throw new Error('分數必須是數字');

    const sh = res.sh, c = res.cols;
    sh.getRange(r.row, c.Group + 1).setValue(group);
    sh.getRange(r.row, c.QNO + 1).setValue(qno);
    sh.getRange(r.row, c.Score + 1).setValue(score);
    sh.getRange(r.row, c.Story + 1).setValue(S_(changes.note));
    SpreadsheetApp.flush();

    rebuildScores_(ss);
    return buildPayload_(ss);
  });
}

/** 刪除一筆作答 */
function deleteResponse(ref) {
  return withLock_(function () {
    const ss = ss_();
    const res = assertConsistent_(ss);
    const r = findRow_(res, ref);
    res.sh.deleteRow(r.row);
    SpreadsheetApp.flush();

    rebuildScores_(ss);
    return buildPayload_(ss);
  });
}

/** 手動補登一筆成績（五市場留空，備註預設「手動補登」） */
function addManualEntry(entry) {
  return withLock_(function () {
    const ss = ss_();
    const res = assertConsistent_(ss);

    const group = S_(entry.group);
    const qno = normQno_(entry.qno);
    const score = Number(entry.score);
    if (!group) throw new Error('組別不可空白');
    if (!qno) throw new Error('題號不可空白');
    if (S_(entry.score) === '' || !isFinite(score)) throw new Error('分數必須是數字');

    const sh = res.sh || ensureResponsesSheet_(ss);
    const c = res.cols || resCols_(RES_HEADERS);
    const width = Math.max(sh.getLastColumn(), RES_HEADERS.length);
    const row = new Array(width).fill('');
    row[c.Timestamp] = new Date();
    row[c.Group] = group;
    row[c.QNO] = qno;
    row[c.Story] = S_(entry.note) || '手動補登';
    row[c.Score] = score;
    sh.getRange(sh.getLastRow() + 1, 1, 1, width).setValues([row]);
    SpreadsheetApp.flush();

    rebuildScores_(ss);
    return buildPayload_(ss);
  });
}

/**
 * 依目前 AnswerKey 的正解，重新批改所有作答。
 * 五市場沒填完整的紀錄（例如手動補登）與找不到題號的紀錄會略過。
 * 有變動時會先備份一份 Responses 分頁。
 */
function regradeAll() {
  return withLock_(function () {
    const ss = ss_();
    const res = assertConsistent_(ss);
    const keyMap = answerKeyMap_(ss);
    const sh = res.sh;
    if (!sh || res.rows.length === 0) {
      return { data: buildPayload_(ss), changed: 0, skipped: 0, samples: [] };
    }

    const lastRow = sh.getLastRow();
    const scoreRange = sh.getRange(2, res.cols.Score + 1, lastRow - 1, 1);
    const scores = scoreRange.getValues();

    let changed = 0, skipped = 0;
    const samples = [];
    res.rows.forEach(function (r) {
      const key = keyMap[r.qno];
      if (!key || MARKETS.some(function (m) { return !r[m]; })) { skipped++; return; }
      const s = gradeOne_(key.ans, r);
      if (s !== r.score) {
        scores[r.row - 2][0] = s;
        changed++;
        if (samples.length < 30) samples.push({ group: r.group, qno: r.qno, from: r.score, to: s });
      }
    });

    if (changed > 0) {
      backupSheet_(ss, sh, 'Responses');
      scoreRange.setValues(scores);
      SpreadsheetApp.flush();
      rebuildScores_(ss);
    }
    return { data: buildPayload_(ss), changed: changed, skipped: skipped, samples: samples };
  });
}

/**
 * 以 Responses 重新產生 Scores。
 * 用在兩邊不一致的時候（例如有人直接在試算表改過 Scores）。會先備份舊的 Scores。
 */
function rebuildScores() {
  return withLock_(function () {
    const ss = ss_();
    const sco = ss.getSheetByName(SHEET_SCO);
    if (sco && sco.getLastRow() > 1) backupSheet_(ss, sco, 'Scores');
    rebuildScores_(ss);
    return buildPayload_(ss);
  });
}

/* ========== 讀取資料 ========== */

function buildPayload_(ss) {
  const res = readResponses_(ss);
  const sco = readScores_(ss);
  return {
    generatedAt: Date.now(),
    sheetUrl: ss.getUrl(),
    sheetName: ss.getName(),
    pointsPerCorrect: POINTS_PER_CORRECT,
    responses: res.rows.map(function (r) {
      const o = {};
      for (const k in r) if (k !== 'sh') o[k] = r[k];
      return o;
    }),
    answerKey: readAnswerKey_(ss),
    consistency: checkConsistency_(res.rows, sco.rows),
  };
}

function readResponses_(ss) {
  const sh = ss.getSheetByName(SHEET_RES);
  if (!sh) return { sh: null, cols: null, rows: [] };
  const v = sh.getDataRange().getValues();
  if (v.length === 0) return { sh: sh, cols: resCols_(RES_HEADERS), rows: [] };

  const cols = resCols_(v[0]);
  const rows = [];
  for (let i = 1; i < v.length; i++) {
    const r = v[i];
    const group = S_(r[cols.Group]);
    const qno = normQno_(r[cols.QNO]);
    if (!group && !qno) continue;
    const o = {
      row: i + 1,
      ts: toMs_(r[cols.Timestamp]),
      group: group,
      qno: qno,
      note: S_(r[cols.Story]),
      score: toScore_(r[cols.Score]),
    };
    MARKETS.forEach(function (m) { o[m] = normSign_(r[cols[m]]); });
    rows.push(o);
  }
  return { sh: sh, cols: cols, rows: rows };
}

function readScores_(ss) {
  const sh = ss.getSheetByName(SHEET_SCO);
  if (!sh) return { rows: [] };
  const v = sh.getDataRange().getValues();
  if (v.length < 2) return { rows: [] };
  const h = v[0];
  const iG = colIndex_(h, 'Group', 1), iQ = colIndex_(h, 'QNO', 2), iS = colIndex_(h, 'Score', 3);
  const rows = [];
  for (let i = 1; i < v.length; i++) {
    const group = S_(v[i][iG]);
    const qno = normQno_(v[i][iQ]);
    if (!group && !qno) continue;
    rows.push({ group: group, qno: qno, score: toScore_(v[i][iS]) });
  }
  return { rows: rows };
}

function readAnswerKey_(ss) {
  const sh = ss.getSheetByName(SHEET_ANS);
  if (!sh) return [];
  const v = sh.getDataRange().getValues();
  if (v.length < 2) return [];

  const aliasSet = {};
  Object.keys(ANS_ALIASES).forEach(function (k) {
    ANS_ALIASES[k].forEach(function (a) { aliasSet[UP_(a)] = true; });
  });
  let hr = 0;
  for (let r = 0; r < Math.min(v.length, 20); r++) {
    if (v[r].some(function (c) { return aliasSet[UP_(c)]; })) { hr = r; break; }
  }
  const head = v[hr].map(UP_);
  const idx = {};
  Object.keys(ANS_ALIASES).forEach(function (k) {
    idx[k] = -1;
    for (const a of ANS_ALIASES[k]) {
      const i = head.indexOf(UP_(a));
      if (i !== -1) { idx[k] = i; break; }
    }
  });

  const out = [];
  for (let r = hr + 1; r < v.length; r++) {
    const row = v[r];
    const qno = idx.QNO >= 0 ? normQno_(row[idx.QNO]) : '';
    if (!qno) continue;
    const ans = {};
    MARKETS.forEach(function (m) { ans[m] = idx[m] >= 0 ? normSign_(row[idx[m]]) : ''; });
    out.push({ qno: qno, text: idx.QUESTIONTEXT >= 0 ? S_(row[idx.QUESTIONTEXT]) : '', ans: ans });
  }
  return out;
}

function answerKeyMap_(ss) {
  const map = {};
  readAnswerKey_(ss).forEach(function (k) { if (!map[k.qno]) map[k.qno] = k; });
  return map;
}

/* ========== 一致性與重建 ========== */

/** 比對 Responses 與 Scores 的（組別, 題號, 分數）筆數是否相同 */
function checkConsistency_(resRows, scoRows) {
  const key = function (r) { return r.group + '|' + r.qno + '|' + (r.score == null ? '' : r.score); };
  const m = {};
  resRows.forEach(function (r) { m[key(r)] = (m[key(r)] || 0) + 1; });
  scoRows.forEach(function (r) { m[key(r)] = (m[key(r)] || 0) - 1; });
  const diffs = [];
  Object.keys(m).forEach(function (k) {
    if (m[k] !== 0) {
      const p = k.split('|');
      diffs.push({ group: p[0], qno: p[1], score: p[2], delta: m[k] });
    }
  });
  return {
    ok: diffs.length === 0,
    resCount: resRows.length,
    scoCount: scoRows.length,
    diffTotal: diffs.length,
    diffs: diffs.slice(0, 30),
  };
}

function assertConsistent_(ss) {
  const res = readResponses_(ss);
  const c = checkConsistency_(res.rows, readScores_(ss).rows);
  if (!c.ok) {
    throw new Error('Scores 與 Responses 目前不一致（' + c.diffTotal +
      ' 項差異）。為避免覆蓋資料，請先按「以 Responses 重建 Scores」再修改。');
  }
  return res;
}

function rebuildScores_(ss) {
  const res = readResponses_(ss);
  if (!res.sh) throw new Error('找不到 Responses 工作表，已停止重建以免清空 Scores');

  let sh = ss.getSheetByName(SHEET_SCO);
  if (!sh) sh = ss.insertSheet(SHEET_SCO);

  const out = res.rows.map(function (r) {
    return [r.ts != null ? new Date(r.ts) : '', r.group, r.qno, r.score == null ? '' : r.score];
  });
  const last = sh.getLastRow();
  if (last > 1) sh.getRange(2, 1, last - 1, Math.max(sh.getLastColumn(), 4)).clearContent();
  sh.getRange(1, 1, 1, 4).setValues([SCO_HEADERS]);
  if (out.length) sh.getRange(2, 1, out.length, 4).setValues(out);
  SpreadsheetApp.flush();
}

/* ========== 工具 ========== */

function ss_() { return SpreadsheetApp.openById(SPREADSHEET_ID); }

function S_(v) { return String(v == null ? '' : v).trim(); }
function UP_(v) { return S_(v).toUpperCase(); }

function normQno_(v) {
  const s = S_(v);
  const n = Number(s);
  return s !== '' && isFinite(n) ? String(n) : s;
}

function toScore_(v) {
  if (v === '' || v == null) return null;
  const n = Number(String(v).replace(/[, ]/g, ''));
  return isFinite(n) ? n : null;
}

function toMs_(v) {
  if (v instanceof Date) return v.getTime();
  if (v === '' || v == null) return null;
  const t = new Date(v).getTime();
  return isFinite(t) ? t : null;
}

function normSign_(v) {
  const t = String(v == null ? '' : v).trim().toLowerCase();
  const pos = ['+','＋','1','正','正面','up','long','多','多頭','pos','positive'];
  const neg = ['-','－','-1','負','負面','down','short','空','空頭','neg','negative'];
  if (pos.indexOf(t) !== -1) return '+';
  if (neg.indexOf(t) !== -1) return '-';
  return '';
}

function gradeOne_(ans, r) {
  let s = 0;
  MARKETS.forEach(function (m) { if (ans[m] && r[m] && ans[m] === r[m]) s += POINTS_PER_CORRECT; });
  return s;
}

function colIndex_(headerRow, name, fallback) {
  const i = headerRow.map(UP_).indexOf(UP_(name));
  return i >= 0 ? i : fallback;
}

function resCols_(headerRow) {
  const c = {};
  RES_HEADERS.forEach(function (h, i) { c[h] = colIndex_(headerRow, h, i); });
  return c;
}

function ensureResponsesSheet_(ss) {
  let sh = ss.getSheetByName(SHEET_RES);
  if (!sh) sh = ss.insertSheet(SHEET_RES);
  if (sh.getLastRow() === 0) sh.appendRow(RES_HEADERS);
  return sh;
}

/** 找到要修改的那一列，並確認內容沒被別人改過 */
function findRow_(res, ref) {
  const same = function (r) {
    return r.group === S_(ref.group) && r.qno === normQno_(ref.qno) && r.ts === (ref.ts == null ? null : ref.ts);
  };
  const direct = res.rows.filter(function (r) { return r.row === ref.row && same(r); });
  if (direct.length) return direct[0];
  const found = res.rows.filter(same);
  if (found.length === 1) return found[0];
  throw new Error('找不到這筆作答（資料可能剛被修改過），請按「重新整理」後再試。');
}

function backupSheet_(ss, sh, label) {
  const stamp = Utilities.formatDate(new Date(), 'Asia/Taipei', 'yyyyMMdd-HHmmss');
  sh.copyTo(ss).setName('備份_' + label + '_' + stamp);
}

function withLock_(fn) {
  const lock = LockService.getScriptLock();
  if (!lock.tryLock(20000)) throw new Error('系統忙碌中，請稍後再試');
  try { return fn(); } finally { lock.releaseLock(); }
}
