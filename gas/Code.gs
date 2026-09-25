/**
 * 学生スタッフ｜当日タスクボード — Google Apps Script（バックエンド）
 *
 * ・スプレッドシートが「正」。職員はシートを直接編集してタスクを管理する。
 * ・学生はスマホ画面（GitHub Pages）から「状況」だけを変更できる。
 * ・このファイルはスプレッドシートの「拡張機能 → Apps Script」に貼り付けて使う。
 */

const CFG = {
  SHEET_TASKS: 'タスク',
  SHEET_SETTINGS: '設定',
  SHEET_LOG: 'ログ',
  STATUSES: ['未着手', '対応中', '完了', '保留'],
  PRIORITIES: ['高', '通常', '低'],
  HEADERS: ['No.', '状況', '優先度', 'やること', '締め切り', '職員メモ', '担当', '最終更新者', '最終更新日時'],
  // ヘッダー名 → 内部キー（列の順番を入れ替えても動くように、名前で列を探す）
  HEADER_KEYS: {
    'No.': 'no', 'No': 'no', '番号': 'no',
    '状況': 'status', '優先度': 'priority', 'やること': 'title', '締め切り': 'due',
    '職員メモ': 'memo', '職員メモ／変更事項': 'memo', '担当': 'owner',
    '最終更新者': 'by', '最終更新日時': 'at',
  },
  STAFF_EDITOR_NAME: '職員（シート）',
  MAX_ROWS_FORMAT: 300,
};

/* ============================================================
 * Web API
 * ============================================================ */

/** GET ?action=list&key=XXXX → ボード全体 */
function doGet(e) {
  return respond_(function () {
    const p = (e && e.parameter) || {};
    checkKey_(p.key);
    const action = p.action || 'list';
    if (action === 'list') return getBoard_();
    throw new Error('不明な操作です: ' + action);
  });
}

/** POST {action:'update', key, id, title, status, name} → 更新後のボード全体 */
function doPost(e) {
  return respond_(function () {
    const body = JSON.parse((e && e.postData && e.postData.contents) || '{}');
    checkKey_(body.key);
    if (body.action === 'update') {
      const result = updateStatus_(body);
      const board = getBoard_();
      board.updated = result;
      return board;
    }
    throw new Error('不明な操作です: ' + body.action);
  });
}

function respond_(fn) {
  let out;
  try {
    out = Object.assign({ ok: true }, fn());
  } catch (err) {
    out = { ok: false, error: String((err && err.message) || err), code: (err && err.code) || 'ERROR' };
  }
  return ContentService.createTextOutput(JSON.stringify(out)).setMimeType(ContentService.MimeType.JSON);
}

function checkKey_(key) {
  const code = String(getSettings_()['アクセスコード'] || '').trim();
  if (!code) return; // 空欄ならチェックしない
  if (String(key || '').trim() !== code) {
    const err = new Error('アクセスコードが違います');
    err.code = 'AUTH';
    throw err;
  }
}

/* ============================================================
 * 読み取り
 * ============================================================ */

function ss_() {
  return SpreadsheetApp.getActiveSpreadsheet();
}

function tz_() {
  return ss_().getSpreadsheetTimeZone() || 'Asia/Tokyo';
}

function getSettings_() {
  const sh = ss_().getSheetByName(CFG.SHEET_SETTINGS);
  const map = {};
  if (!sh || sh.getLastRow() < 1) return map;
  sh.getRange(1, 1, sh.getLastRow(), 2).getValues().forEach(function (r) {
    const k = String(r[0] || '').trim();
    if (k) map[k] = r[1];
  });
  return map;
}

/** タスクシートを読み、{sh, col, values} を返す。col は 1 始まりの列番号 */
function readTasks_() {
  const sh = ss_().getSheetByName(CFG.SHEET_TASKS);
  if (!sh) throw new Error('「' + CFG.SHEET_TASKS + '」シートがありません。メニューの「初期設定」を実行してください');
  const lastRow = sh.getLastRow();
  const lastCol = sh.getLastColumn();
  if (lastRow < 1 || lastCol < 1) throw new Error('タスクシートにヘッダー行がありません');
  const values = sh.getRange(1, 1, lastRow, lastCol).getValues();
  const col = {};
  values[0].forEach(function (h, i) {
    const name = String(h || '').split('\n')[0].trim();
    const key = CFG.HEADER_KEYS[name];
    if (key && !col[key]) col[key] = i + 1;
  });
  ['no', 'status', 'title'].forEach(function (k) {
    if (!col[k]) throw new Error('タスクシートに必要な列がありません（No. / 状況 / やること）');
  });
  return { sh: sh, col: col, values: values };
}

function cell_(row, col, key) {
  return col[key] ? row[col[key] - 1] : '';
}

function getBoard_() {
  ensureIds_();
  const s = getSettings_();
  const t = readTasks_();
  const tasks = [];
  for (let r = 1; r < t.values.length; r++) {
    const row = t.values[r];
    const title = String(cell_(row, t.col, 'title') || '').trim();
    if (!title) continue;
    tasks.push(toTask_(row, t.col));
  }
  return {
    title: String(s['タイトル'] || '学生スタッフ｜当日タスクボード'),
    date: fmtDate_(s['開催日']),
    notice: String(s['お知らせ'] || ''),
    statuses: CFG.STATUSES,
    tasks: tasks,
    serverTime: new Date().toISOString(),
  };
}

function toTask_(row, col) {
  const due = cell_(row, col, 'due');
  const at = cell_(row, col, 'at');
  let status = String(cell_(row, col, 'status') || '').trim();
  if (CFG.STATUSES.indexOf(status) < 0) status = '未着手';
  return {
    id: String(cell_(row, col, 'no')),
    status: status,
    priority: String(cell_(row, col, 'priority') || '通常').trim() || '通常',
    title: String(cell_(row, col, 'title') || '').trim(),
    due: fmtDate_(due),
    dueSort: due instanceof Date ? due.getTime() : null,
    memo: String(cell_(row, col, 'memo') || ''),
    owner: String(cell_(row, col, 'owner') || ''),
    by: String(cell_(row, col, 'by') || ''),
    at: at instanceof Date ? at.toISOString() : '',
  };
}

function fmtDate_(v) {
  if (!(v instanceof Date)) return v === null || v === undefined ? '' : String(v);
  const tz = tz_();
  const wd = ['日', '月', '火', '水', '木', '金', '土'][Number(Utilities.formatDate(v, tz, 'u')) % 7];
  const hm = Utilities.formatDate(v, tz, 'H:mm');
  return Utilities.formatDate(v, tz, 'M/d') + '(' + wd + ')' + (hm === '0:00' ? '' : ' ' + hm);
}

/** No. が空欄・重複のタスク行に新しい番号を振る（職員が行を追加・コピーしたとき用） */
function ensureIds_() {
  if (!findIdFixes_(readTasks_()).rows.length) return;
  const lock = LockService.getScriptLock();
  if (!lock.tryLock(5000)) return;
  try {
    const t = readTasks_(); // ロック取得後に読み直す
    const f = findIdFixes_(t);
    let max = f.max;
    f.rows.forEach(function (rowNum) {
      t.sh.getRange(rowNum, t.col.no).setValue(++max);
    });
    SpreadsheetApp.flush();
  } finally {
    lock.releaseLock();
  }
}

function findIdFixes_(t) {
  const rows = [];
  const seen = {};
  let max = 0;
  for (let r = 1; r < t.values.length; r++) {
    const no = t.values[r][t.col.no - 1];
    const n = Number(no);
    if (no !== '' && no !== null && !isNaN(n)) max = Math.max(max, n);
  }
  for (let r = 1; r < t.values.length; r++) {
    const no = t.values[r][t.col.no - 1];
    if (!String(t.values[r][t.col.title - 1] || '').trim()) continue;
    const k = String(no);
    if (no === '' || no === null || seen[k]) rows.push(r + 1);
    else seen[k] = true;
  }
  return { rows: rows, max: max };
}

/* ============================================================
 * 更新（学生のスマホから）
 * ============================================================ */

function updateStatus_(b) {
  const status = String(b.status || '').trim();
  if (CFG.STATUSES.indexOf(status) < 0) throw new Error('状況の値が不正です: ' + status);
  const name = String(b.name || '').trim().slice(0, 20) || '（名前なし）';

  const lock = LockService.getScriptLock();
  lock.waitLock(15000);
  try {
    const t = readTasks_();
    const rowNum = findRow_(t, String(b.id), String(b.title || ''));
    const row = t.values[rowNum - 1];
    const prev = String(cell_(row, t.col, 'status') || '');
    const now = new Date();
    t.sh.getRange(rowNum, t.col.status).setValue(status);
    if (t.col.by) t.sh.getRange(rowNum, t.col.by).setValue(name);
    if (t.col.at) t.sh.getRange(rowNum, t.col.at).setValue(now);
    appendLog_(now, String(b.id), String(cell_(row, t.col, 'title')), prev, status, name);
    SpreadsheetApp.flush();
    return { id: String(b.id), prev: prev, status: status };
  } finally {
    lock.releaseLock();
  }
}

function findRow_(t, id, title) {
  const hits = [];
  for (let r = 1; r < t.values.length; r++) {
    if (String(t.values[r][t.col.no - 1]) === id) hits.push(r + 1);
  }
  if (hits.length === 1) return hits[0];
  if (hits.length > 1) {
    // No. が重複している場合はタイトルで特定
    const m = hits.filter(function (n) {
      return String(t.values[n - 1][t.col.title - 1]).trim() === title.trim();
    });
    if (m.length === 1) return m[0];
    throw new Error('No.' + id + ' が重複しています。職員に連絡してください');
  }
  const err = new Error('このタスクは削除または変更されました。画面を更新してください');
  err.code = 'NOT_FOUND';
  throw err;
}

function appendLog_(when, id, title, prev, next, name) {
  let sh = ss_().getSheetByName(CFG.SHEET_LOG);
  if (!sh) {
    sh = ss_().insertSheet(CFG.SHEET_LOG);
    sh.appendRow(['日時', 'No.', 'やること', '変更前', '変更後', '変更者']);
  }
  sh.appendRow([when, id, title, prev, next, name]);
}

/* ============================================================
 * シート側のイベント（職員の手動編集）
 * ============================================================ */

/** 職員がシート上で「状況」を変えたら、更新者・日時を記録し、ログに残す */
function onEdit(e) {
  try {
    if (!e || !e.range) return;
    const sh = e.range.getSheet();
    if (sh.getName() !== CFG.SHEET_TASKS) return;
    const t = readTasks_();
    const r0 = e.range.getRow();
    const c0 = e.range.getColumn();
    const c1 = c0 + e.range.getNumColumns() - 1;
    if (t.col.status < c0 || t.col.status > c1) return;
    const now = new Date();
    for (let rowNum = Math.max(2, r0); rowNum < r0 + e.range.getNumRows(); rowNum++) {
      const row = t.values[rowNum - 1];
      if (!row || !String(row[t.col.title - 1] || '').trim()) continue;
      if (t.col.by) sh.getRange(rowNum, t.col.by).setValue(CFG.STAFF_EDITOR_NAME);
      if (t.col.at) sh.getRange(rowNum, t.col.at).setValue(now);
      const prev = e.range.getNumRows() === 1 && e.range.getNumColumns() === 1 ? (e.oldValue || '') : '';
      appendLog_(now, String(row[t.col.no - 1]), String(row[t.col.title - 1]), prev,
        String(row[t.col.status - 1]), CFG.STAFF_EDITOR_NAME);
    }
  } catch (err) {
    console.warn(err);
  }
}

function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu('タスクボード')
    .addItem('初期設定（シート作成・書式）', 'setup')
    .addItem('状況をすべて「未着手」に戻す', 'resetStatuses')
    .addSeparator()
    .addItem('ログを消去', 'clearLog')
    .addToUi();
}

/* ============================================================
 * 管理用メニュー
 * ============================================================ */

/** シートの作成・書式設定。何度実行しても既存データは消えない */
function setup() {
  const ss = ss_();

  // --- 設定シート
  let st = ss.getSheetByName(CFG.SHEET_SETTINGS);
  if (!st) {
    st = ss.insertSheet(CFG.SHEET_SETTINGS);
    const code = String(Math.floor(1000 + Math.random() * 9000));
    st.getRange(1, 1, 5, 3).setValues([
      ['項目', '値', '説明'],
      ['タイトル', '学生スタッフ｜当日タスクボード', 'スマホ画面の見出し'],
      ['開催日', new Date(new Date().setHours(0, 0, 0, 0)), '日付を入れると画面に表示'],
      ['アクセスコード', code, 'QR に埋め込む合言葉。変更すると古い QR は使えなくなる。空欄ならチェックなし'],
      ['お知らせ', '', '画面上部に黄色で表示（例：12:00〜13:00 は昼休憩）。空欄なら非表示'],
    ]);
    st.getRange('A1:C1').setFontWeight('bold').setBackground('#1f3a5f').setFontColor('#ffffff');
    st.getRange('B3').setNumberFormat('yyyy/mm/dd');
    st.setColumnWidth(1, 130); st.setColumnWidth(2, 280); st.setColumnWidth(3, 420);
    st.setFrozenRows(1);
  }

  // --- タスクシート
  let sh = ss.getSheetByName(CFG.SHEET_TASKS);
  const isNew = !sh;
  if (!sh) sh = ss.insertSheet(CFG.SHEET_TASKS, 0);
  if (sh.getLastRow() === 0) {
    sh.getRange(1, 1, 1, CFG.HEADERS.length).setValues([CFG.HEADERS]);
    const sample = sampleTasks_();
    if (isNew && sample.length) sh.getRange(2, 1, sample.length, sample[0].length).setValues(sample);
  }
  formatTaskSheet_(sh);

  // --- ログシート
  let lg = ss.getSheetByName(CFG.SHEET_LOG);
  if (!lg) {
    lg = ss.insertSheet(CFG.SHEET_LOG);
    lg.appendRow(['日時', 'No.', 'やること', '変更前', '変更後', '変更者']);
  }
  lg.getRange('A1:F1').setFontWeight('bold').setBackground('#1f3a5f').setFontColor('#ffffff');
  lg.getRange('A:A').setNumberFormat('m/d hh:mm:ss');
  lg.setFrozenRows(1);
  lg.setColumnWidth(3, 320);

  // 初期の空シート（シート1）があれば削除
  const blank = ss.getSheetByName('シート1') || ss.getSheetByName('Sheet1');
  if (blank && ss.getSheets().length > 1 && blank.getLastRow() === 0) ss.deleteSheet(blank);

  ss.setActiveSheet(sh);
  try {
    SpreadsheetApp.getUi().alert('初期設定が完了しました。\n「設定」シートのアクセスコードを確認してください。');
  } catch (e) { /* エディタから実行した場合は UI なし */ }
}

function formatTaskSheet_(sh) {
  const t = readTasks_();
  const n = CFG.MAX_ROWS_FORMAT;
  const lastCol = sh.getLastColumn();
  sh.setFrozenRows(1);
  sh.getRange(1, 1, 1, lastCol).setFontWeight('bold').setBackground('#1f3a5f').setFontColor('#ffffff')
    .setVerticalAlignment('middle').setWrap(true);
  sh.setRowHeight(1, 32);

  const widths = { no: 50, status: 80, priority: 70, title: 320, due: 110, memo: 380, owner: 100, by: 110, at: 130 };
  Object.keys(widths).forEach(function (k) { if (t.col[k]) sh.setColumnWidth(t.col[k], widths[k]); });
  ['title', 'memo'].forEach(function (k) { if (t.col[k]) sh.getRange(2, t.col[k], n, 1).setWrap(true).setVerticalAlignment('top'); });
  if (t.col.at) sh.getRange(2, t.col.at, n, 1).setNumberFormat('m/d hh:mm');
  if (t.col.due) sh.getRange(2, t.col.due, n, 1).setHorizontalAlignment('left');

  const statusRange = sh.getRange(2, t.col.status, n, 1);
  statusRange.setDataValidation(SpreadsheetApp.newDataValidation()
    .requireValueInList(CFG.STATUSES, true).setAllowInvalid(false).build())
    .setHorizontalAlignment('center');
  if (t.col.priority) {
    sh.getRange(2, t.col.priority, n, 1).setDataValidation(SpreadsheetApp.newDataValidation()
      .requireValueInList(CFG.PRIORITIES, true).setAllowInvalid(false).build())
      .setHorizontalAlignment('center');
  }
  // 最終更新者・日時は自動記録なので灰色に
  ['by', 'at'].forEach(function (k) { if (t.col[k]) sh.getRange(2, t.col[k], n, 1).setFontColor('#6b7280'); });

  // 条件付き書式（状況・優先度の色）
  const rules = [];
  const colors = { '未着手': ['#eef1f5', '#374151'], '対応中': ['#dbeafe', '#1d4ed8'], '完了': ['#dcfce7', '#15803d'], '保留': ['#fef3c7', '#b45309'] };
  Object.keys(colors).forEach(function (s) {
    rules.push(SpreadsheetApp.newConditionalFormatRule().whenTextEqualTo(s)
      .setBackground(colors[s][0]).setFontColor(colors[s][1]).setRanges([statusRange]).build());
  });
  if (t.col.priority) {
    rules.push(SpreadsheetApp.newConditionalFormatRule().whenTextEqualTo('高')
      .setBackground('#fee2e2').setFontColor('#b91c1c').setBold(true)
      .setRanges([sh.getRange(2, t.col.priority, n, 1)]).build());
  }
  sh.setConditionalFormatRules(rules);
}

/** 次回の開催前に：状況をすべて未着手に戻し、更新者・日時を消す */
function resetStatuses() {
  const ui = SpreadsheetApp.getUi();
  const ans = ui.alert('確認', 'すべてのタスクの状況を「未着手」に戻します。よろしいですか？', ui.ButtonSet.YES_NO);
  if (ans !== ui.Button.YES) return;
  const t = readTasks_();
  const rows = t.values.length - 1;
  if (rows < 1) return;
  const vals = t.values.slice(1).map(function (r) { return [String(r[t.col.title - 1] || '').trim() ? '未着手' : r[t.col.status - 1]]; });
  t.sh.getRange(2, t.col.status, rows, 1).setValues(vals);
  if (t.col.by) t.sh.getRange(2, t.col.by, rows, 1).clearContent();
  if (t.col.at) t.sh.getRange(2, t.col.at, rows, 1).clearContent();
  ui.alert('リセットしました。');
}

function clearLog() {
  const ui = SpreadsheetApp.getUi();
  const ans = ui.alert('確認', 'ログをすべて消去します。よろしいですか？', ui.ButtonSet.YES_NO);
  if (ans !== ui.Button.YES) return;
  const lg = ss_().getSheetByName(CFG.SHEET_LOG);
  if (lg && lg.getLastRow() > 1) lg.deleteRows(2, lg.getLastRow() - 1);
}

/** 初期設定時に入るサンプル（元の Excel ボードの内容） */
function sampleTasks_() {
  const d26 = new Date(2026, 8, 26), d27 = new Date(2026, 8, 27);
  return [
    [1, '未着手', '高', '【実技特待生講座】模擬試験のみ！会場準備', d26, '準備するセット数は3セットです！', '', '', ''],
    [2, '未着手', '高', '【実技特待生講座・試験】A31教室の控室準備', d26, '午前中の間に清掃・机を拭くなどの控室の準備をお願いいたします。26日は講座控室、27日は実技試験控室で使用します', '', '', ''],
    [3, '未着手', '高', '【実技特待生試験】試験準備3セット', d26, '27日の試験準備になります！\n受験者は3名、１ターンのみとなります。', '', '', ''],
    [4, '未着手', '低', '資料請求発送作業', d26, '26日は孫先生が対応してくださるので、\n孫先生に確認してください！', '', '', ''],
    [5, '未着手', '通常', '【依頼】試薬の重量測定', d27, '高村先生からの依頼です。\n対応できるときに運営の先生にお声かけください。試薬棚の鍵を開ける必要があります。', '', '', ''],
    [6, '未着手', '高', '【筆記特待生試験】会場準備', '9/27の朝イチまで', 'A33教室の会場準備です。清掃・机を拭くなど、会場準備をお願いいたします。', '', '', ''],
    [7, '未着手', '通常', 'クリーニングに出す白衣の選定', d27, '27日の試験後、クリーニングに出す白衣を選定し、運営担当の先生にお渡しください！', '', '', ''],
    [8, '未着手', '高', 'シラバス更新', d27, '進路相談会場にあるシラバスの更新をお願いします。バイオ医薬品コースは再チェック。それ以外は差し替えのみ', '', '', ''],
    [9, '未着手', '低', 'セット組作成', d27, '高3用2セット、高2用2セット、高1用1セットがあるか、確認', '', '', ''],
    [10, '未着手', '低', '資料台の資料チェック', d27, '不足している書類がないか確認してください', '', '', ''],
    [11, '未着手', '通常', '本館PCルームリセット', d27, '日曜日のイベントが終了しましたら、リセットをお願いいたします。', '', '', ''],
  ];
}
