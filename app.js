/* 当日タスクボード — フロントエンド（複数ボード対応：?b=<ボードID>） */
(function () {
  'use strict';

  const CONF = window.TASKBOARD_CONFIG || {};
  const QS = new URLSearchParams(location.search);
  // ボードの決定：?b=assist など。指定なしは DEFAULT_BOARD（旧形式の GAS_URL 単独設定にも対応）
  const BOARDS = CONF.BOARDS || { default: { label: '', GAS_URL: CONF.GAS_URL || '' } };
  const DEFAULT_BOARD = CONF.DEFAULT_BOARD && BOARDS[CONF.DEFAULT_BOARD] ? CONF.DEFAULT_BOARD : Object.keys(BOARDS)[0];
  const BOARD_ID = (QS.get('b') || DEFAULT_BOARD).trim();
  const BOARD = BOARDS[BOARD_ID];
  const GAS_URL = BOARD ? String(BOARD.GAS_URL || '') : '';
  const DEMO = !!BOARD && (!GAS_URL || QS.get('demo') === '1');
  const POLL_MS = Math.max(10, Number(CONF.POLL_SECONDS) || 20) * 1000;
  const STATUSES = ['未着手', '対応中', '完了', '保留'];
  const S_CLASS = { '未着手': 'todo', '対応中': 'doing', '完了': 'done', '保留': 'hold' };
  const PRI_ORDER = { '高': 0, '通常': 1, '低': 2 };
  // 名前は全ボード共通、アクセスコードと絞り込みはボードごとに保存
  const LS = { name: 'tb.name', key: 'tb.key.' + BOARD_ID, filter: 'tb.filter.' + BOARD_ID };

  const $ = (id) => document.getElementById(id);
  const store = {
    get(k) { try { return localStorage.getItem(k) || ''; } catch (e) { return ''; } },
    set(k, v) { try { localStorage.setItem(k, v); } catch (e) { /* noop */ } },
  };

  const state = {
    board: null,          // サーバーから受け取った最新のボード
    pending: {},          // { id: status } 送信中の楽観的更新
    filter: store.get(LS.filter) || 'すべて',
    name: store.get(LS.name),
    key: store.get(LS.key) || (BOARD_ID === DEFAULT_BOARD ? store.get('tb.key') : ''), // 旧バージョンの保存値を引き継ぐ
    listSeq: 0,           // 一覧取得の通し番号
    acceptFrom: 0,        // この番号より前に始まった一覧取得の結果は捨てる
    lastRender: '',
    loading: false,
  };

  // QR 経由で ?k=コード が付いていたら保存して URL から消す
  if (QS.get('k')) {
    state.key = QS.get('k').trim();
    store.set(LS.key, state.key);
    QS.delete('k');
    const rest = QS.toString();
    history.replaceState(null, '', location.pathname + (rest ? '?' + rest : '') + location.hash);
  }

  /* ============================================================
   * API
   * ============================================================ */
  async function fetchJSON(url, opts) {
    const ctrl = new AbortController();
    const timer = setTimeout(() => ctrl.abort(), 25000);
    try {
      const res = await fetch(url, Object.assign({ signal: ctrl.signal, cache: 'no-store', redirect: 'follow' }, opts || {}));
      if (!res.ok) throw new Error('通信エラー（' + res.status + '）');
      const data = await res.json();
      if (!data.ok) {
        const err = new Error(data.error || 'エラーが発生しました');
        err.code = data.code;
        throw err;
      }
      return data;
    } catch (e) {
      if (e.name === 'AbortError') throw new Error('通信がタイムアウトしました');
      if (e instanceof TypeError) throw new Error('通信できません。電波状況を確認してください');
      throw e;
    } finally {
      clearTimeout(timer);
    }
  }

  const api = DEMO ? demoApi() : {
    list() {
      const u = GAS_URL + '?action=list&key=' + encodeURIComponent(state.key) + '&_=' + Date.now();
      return fetchJSON(u);
    },
    update(task, status) {
      return fetchJSON(GAS_URL, {
        method: 'POST',
        headers: { 'Content-Type': 'text/plain;charset=utf-8' }, // CORS プリフライトを避ける
        body: JSON.stringify({ action: 'update', key: state.key, id: task.id, title: task.title, status: status, name: state.name }),
      });
    },
  };

  /* ============================================================
   * 読み込み・同期
   * ============================================================ */
  async function load(manual) {
    if (!BOARD) return;
    if (state.loading && !manual) return;
    const seq = ++state.listSeq;
    state.loading = true;
    setSync('busy', '同期中…');
    if (manual) $('refresh').classList.add('spin');
    try {
      const data = await api.list();
      if (seq < state.acceptFrom) return; // 更新より前に取得した古いデータは捨てる
      applyBoard(data);
      setSync('ok', '最新 ' + hm(new Date()));
    } catch (e) {
      if (e.code === 'AUTH') { askKey(!!state.key); setSync('err', 'アクセスコードが必要です'); }
      else { setSync('err', '接続できません'); if (manual) toast(e.message, { error: true }); }
    } finally {
      state.loading = false;
      $('refresh').classList.remove('spin');
    }
  }

  function applyBoard(data) {
    state.board = data;
    document.title = data.title || '当日タスクボード';
    render();
  }

  async function changeStatus(id, status, opts) {
    opts = opts || {};
    const task = findTask(id);
    if (!task) return;
    const prev = currentStatus(task);
    if (prev === status) return;
    if (!state.name) {
      const ok = await askName();
      if (!ok) return;
    }
    state.pending[id] = status;
    render();
    if (navigator.vibrate) navigator.vibrate(8);
    if (!opts.silent) {
      toast('「' + short(task.title) + '」を ' + status + ' にしました', {
        undo: () => changeStatus(id, prev, { silent: true }),
      });
    }
    try {
      const data = await api.update(task, status);
      state.acceptFrom = state.listSeq + 1;
      if (state.pending[id] === status) delete state.pending[id];
      applyBoard(data);
      setSync('ok', '保存しました ' + hm(new Date()));
    } catch (e) {
      if (state.pending[id] === status) delete state.pending[id];
      render();
      if (e.code === 'AUTH') askKey(true);
      toast('保存できませんでした：' + e.message, { error: true });
      if (e.code === 'NOT_FOUND') load();
    }
  }

  /* ============================================================
   * 描画
   * ============================================================ */
  function currentStatus(t) {
    return Object.prototype.hasOwnProperty.call(state.pending, t.id) ? state.pending[t.id] : t.status;
  }
  function findTask(id) {
    return state.board && state.board.tasks.find((t) => t.id === id);
  }

  function render() {
    const b = state.board;
    if (!b) return;
    $('title').textContent = b.title || '当日タスクボード';
    $('date').textContent = b.date ? '開催日 ' + b.date : '';
    const notice = $('notice');
    notice.hidden = !b.notice;
    notice.textContent = b.notice || '';
    $('whoName').textContent = state.name ? state.name + ' さん' : '名前を登録';

    const tasks = b.tasks.map((t) => Object.assign({}, t, { cur: currentStatus(t), sending: t.id in state.pending }));
    const counts = { 'すべて': tasks.length };
    STATUSES.forEach((s) => { counts[s] = tasks.filter((t) => t.cur === s).length; });
    const done = counts['完了'];
    $('bar').style.width = (tasks.length ? Math.round(done / tasks.length * 100) : 0) + '%';
    $('pct').textContent = done + ' / ' + tasks.length + ' 完了';

    // 絞り込みチップ
    const chips = ['すべて', '未着手', '対応中', '保留', '完了'];
    if (chips.indexOf(state.filter) < 0) state.filter = 'すべて';
    $('filters').innerHTML = chips.map((c) =>
      '<button type="button" class="chip" data-filter="' + c + '" aria-pressed="' + (state.filter === c) + '">' +
      c + '<b>' + counts[c] + '</b></button>').join('');

    const shown = tasks
      .filter((t) => state.filter === 'すべて' || t.cur === state.filter)
      .sort((a, b2) =>
        (a.cur === '完了') - (b2.cur === '完了') ||
        ((PRI_ORDER[a.priority] ?? 1) - (PRI_ORDER[b2.priority] ?? 1)) ||
        (Number(a.id) - Number(b2.id)) || String(a.id).localeCompare(String(b2.id)));

    const html = shown.length ? shown.map(cardHtml).join('') :
      '<p class="empty">' + (tasks.length ? '「' + esc(state.filter) + '」のタスクはありません' : 'タスクはまだありません') + '</p>';
    if (html !== state.lastRender) {
      $('list').innerHTML = html;
      state.lastRender = html;
    }
  }

  function cardHtml(t) {
    const sc = S_CLASS[t.cur] || 'todo';
    const pri = t.priority || '通常';
    const upd = t.sending
      ? '<span class="sending">送信中…</span>'
      : (t.by && t.at ? esc(t.by) + (t.by === state.name ? '（あなた）' : '') + ' が ' + esc(when(t.at)) + ' に更新' : '');
    return '<article class="card s-' + sc + '" data-id="' + esc(t.id) + '">' +
      '<div class="meta"><span class="pri' + (pri === '高' ? ' p-hi' : '') + '">' + esc(pri) + '</span>' +
      '<span class="no">No.' + esc(t.id) + '</span>' +
      (t.due ? '<span class="due">' + esc(t.due) + '</span>' : '') + '</div>' +
      '<h2 class="title">' + esc(t.title) + '</h2>' +
      (t.memo ? '<p class="memo">' + esc(t.memo) + '</p>' : '') +
      (t.owner ? '<div class="owner">担当：<b>' + esc(t.owner) + '</b></div>' : '') +
      '<div class="seg" role="group" aria-label="状況を変更">' +
      STATUSES.map((s) => '<button type="button" class="b-' + S_CLASS[s] + '" data-status="' + s + '" aria-pressed="' + (t.cur === s) + '">' + s + '</button>').join('') +
      '</div>' +
      '<div class="upd">' + upd + '</div>' +
      '</article>';
  }

  /* ============================================================
   * UI 部品
   * ============================================================ */
  let toastTimer = null;
  function toast(msg, opts) {
    opts = opts || {};
    const el = $('toast');
    const undo = $('toastUndo');
    $('toastMsg').textContent = msg;
    el.classList.toggle('err', !!opts.error);
    undo.hidden = !opts.undo;
    undo.onclick = opts.undo ? () => { hideToast(); opts.undo(); } : null;
    el.hidden = false;
    clearTimeout(toastTimer);
    toastTimer = setTimeout(hideToast, opts.error ? 6000 : 5000);
  }
  function hideToast() { $('toast').hidden = true; }

  function setSync(cls, text) {
    const el = $('sync');
    el.className = 'sync ' + cls;
    el.textContent = DEMO ? 'デモ' : text;
  }

  function askName() {
    return new Promise((resolve) => {
      const dlg = $('dlgName');
      const input = $('inName');
      input.value = state.name || '';
      const cleanup = (ok) => {
        $('formName').onsubmit = null; $('nameCancel').onclick = null;
        if (dlg.open) dlg.close();
        resolve(ok);
      };
      $('formName').onsubmit = (e) => {
        e.preventDefault();
        const v = input.value.trim();
        if (!v) return;
        state.name = v.slice(0, 20);
        store.set(LS.name, state.name);
        render();
        cleanup(true);
      };
      $('nameCancel').onclick = () => cleanup(false);
      dlg.oncancel = () => resolve(false);
      dlg.showModal();
      setTimeout(() => input.focus(), 50);
    });
  }

  function askKey(wrong) {
    const dlg = $('dlgKey');
    if (dlg.open) { $('keyErr').hidden = !wrong; return; }
    $('keyErr').hidden = !wrong;
    $('inKey').value = '';
    $('formKey').onsubmit = (e) => {
      e.preventDefault();
      const v = $('inKey').value.trim();
      if (!v) return;
      state.key = v;
      store.set(LS.key, v);
      dlg.close();
      load(true);
    };
    dlg.oncancel = (e) => e.preventDefault(); // コードなしでは閉じられない
    dlg.showModal();
  }

  /* ============================================================
   * イベント
   * ============================================================ */
  $('list').addEventListener('click', (e) => {
    const btn = e.target.closest('.seg button');
    if (!btn) return;
    const card = btn.closest('.card');
    changeStatus(card.dataset.id, btn.dataset.status);
  });
  $('filters').addEventListener('click', (e) => {
    const chip = e.target.closest('.chip');
    if (!chip) return;
    state.filter = chip.dataset.filter;
    store.set(LS.filter, state.filter);
    render();
    window.scrollTo({ top: 0 });
  });
  $('refresh').addEventListener('click', () => load(true));
  $('who').addEventListener('click', () => askName());

  document.addEventListener('visibilitychange', () => {
    if (document.visibilityState === 'visible') load();
  });
  window.addEventListener('online', () => load());
  setInterval(() => { if (document.visibilityState === 'visible') load(); }, POLL_MS);

  /* ============================================================
   * ユーティリティ
   * ============================================================ */
  function esc(s) {
    return String(s == null ? '' : s).replace(/[&<>"']/g, (c) => ({ '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;', "'": '&#39;' }[c]));
  }
  function short(s) { s = String(s || ''); return s.length > 18 ? s.slice(0, 18) + '…' : s; }
  function hm(d) { return d.getHours() + ':' + String(d.getMinutes()).padStart(2, '0'); }
  function when(iso) {
    const d = new Date(iso);
    if (isNaN(d)) return '';
    const now = new Date();
    const sameDay = d.toDateString() === now.toDateString();
    return sameDay ? hm(d) : (d.getMonth() + 1) + '/' + d.getDate() + ' ' + hm(d);
  }

  /* ============================================================
   * デモ（GAS 未設定時）
   * ============================================================ */
  function demoApi() {
    $('demo').hidden = false;
    const d26 = '9/26(土)', d27 = '9/27(日)';
    const board = {
      ok: true, title: '学生スタッフ｜当日タスクボード', date: '9/26(土)', notice: '12:00〜13:00 は昼休憩です。第2会議室へ。',
      tasks: [
        ['1', '高', '【実技特待生講座】模擬試験のみ！会場準備', d26, '準備するセット数は3セットです！'],
        ['2', '高', '【実技特待生講座・試験】A31教室の控室準備', d26, '午前中の間に清掃・机を拭くなどの控室の準備をお願いいたします。26日は講座控室、27日は実技試験控室で使用します'],
        ['3', '高', '【実技特待生試験】試験準備3セット', d26, '27日の試験準備になります！\n受験者は3名、１ターンのみとなります。'],
        ['4', '低', '資料請求発送作業', d26, '26日は孫先生が対応してくださるので、\n孫先生に確認してください！'],
        ['5', '通常', '【依頼】試薬の重量測定', d27, '高村先生からの依頼です。\n対応できるときに運営の先生にお声かけください。試薬棚の鍵を開ける必要があります。'],
        ['6', '高', '【筆記特待生試験】会場準備', '9/27の朝イチまで', 'A33教室の会場準備です。清掃・机を拭くなど、会場準備をお願いいたします。'],
        ['7', '通常', 'クリーニングに出す白衣の選定', d27, '27日の試験後、クリーニングに出す白衣を選定し、運営担当の先生にお渡しください！'],
        ['8', '高', 'シラバス更新', d27, '進路相談会場にあるシラバスの更新をお願いします。バイオ医薬品コースは再チェック。それ以外は差し替えのみ'],
        ['9', '低', 'セット組作成', d27, '高3用2セット、高2用2セット、高1用1セットがあるか、確認'],
        ['10', '低', '資料台の資料チェック', d27, '不足している書類がないか確認してください'],
        ['11', '通常', '本館PCルームリセット', d27, '日曜日のイベントが終了しましたら、リセットをお願いいたします。'],
      ].map((r) => ({ id: r[0], status: '未着手', priority: r[1], title: r[2], due: r[3], memo: r[4], owner: '', by: '', at: '' })),
    };
    const clone = () => JSON.parse(JSON.stringify(board));
    const wait = (ms) => new Promise((r) => setTimeout(r, ms));
    return {
      async list() { await wait(300); return clone(); },
      async update(task, status) {
        await wait(700);
        const t = board.tasks.find((x) => x.id === task.id);
        t.status = status; t.by = state.name; t.at = new Date().toISOString();
        return clone();
      },
    };
  }

  if (!BOARD) {
    document.getElementById('list').innerHTML = '<p class="empty">ボード「' + esc(BOARD_ID) + '」が見つかりません。<br>QRコードを読み直すか、職員に確認してください。</p>';
    setSync('err', 'ボードが見つかりません');
    return;
  }
  load(true);
})();
