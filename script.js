/* =========================================================
   My Budget — 購入・残金管理アプリ（GitHub Pages 用）
   ========================================================= */

// ▼ 教員が設定：GAS の Web アプリ URL（空のままなら送信しません）
const GAS_URL = 'https://script.google.com/macros/s/AKfycbycZvxZBHtkzFOIusxJ9FlCdICeY3sgvhqK0Uw0gMTPp-H_-f3haIAe5ml9ysV7lNUK/exec';

const DEFAULT_BUDGET = 10000;   // 初期画面に最初から入っている初期費用
const STORAGE_KEY = 'myBudgetApp_v3';

/* ---------------------------------------------------------
   【教員用】Google Apps Script のコード例
   ---------------------------------------------------------
   1. Google スプレッドシートを新規作成
   2. 「拡張機能 → Apps Script」を開き、以下を貼り付けて保存
   3. 「デプロイ → 新しいデプロイ → 種類：ウェブアプリ」
        次のユーザーとして実行：自分
        アクセスできるユーザー：全員
   4. 発行された URL（…/exec）を上の GAS_URL に貼り付ける
   ※ コードを修正したら「デプロイを管理 → 編集 → 新バージョン」で再デプロイ

   function doPost(e) {
     const lock = LockService.getScriptLock();
     lock.waitLock(10000);
     try {
       const data = JSON.parse(e.postData.contents);
       const ss = SpreadsheetApp.getActiveSpreadsheet();
       let sh = ss.getSheetByName('Log');
       if (!sh) {
         sh = ss.insertSheet('Log');
         sh.appendRow(['受信日時', '端末の記録日時', 'ユーザー名',
                       '種別', 'カテゴリ', '金額', '使用額合計', '残金', 'ID']);
         sh.setFrozenRows(1);
       }
       // ※「登録」の行の「残金」列が、その生徒の初期費用になります
       // 同じ ID の重複送信は記録しない
       const last = sh.getLastRow();
       if (last > 1) {
         const ids = sh.getRange(2, 9, last - 1, 1).getValues().flat();
         if (ids.indexOf(data.id) !== -1) return out_({ ok: true, dup: true });
       }
       const s = v => (/^[=+\-@]/.test(String(v)) ? "'" + v : v);
       sh.appendRow([new Date(), data.timestamp, s(data.name),
                     data.type, data.category, data.amount, data.spent, data.balance, data.id]);
       return out_({ ok: true });
     } finally {
       lock.releaseLock();
     }
   }
   function out_(obj) {
     return ContentService.createTextOutput(JSON.stringify(obj))
       .setMimeType(ContentService.MimeType.JSON);
   }

   ・生徒ごとの使用額合計と残金を一覧したい場合は、別シートの A1 に次の式を入れます
     =QUERY(Log!C:H, "select C, sum(F), min(H) where D = '購入' group by C", 1)
   --------------------------------------------------------- */


/* ---------- ユーティリティ ---------- */
const $ = id => document.getElementById(id);
const yen = n => '¥' + Number(n).toLocaleString('ja-JP');

function pad(n) { return String(n).padStart(2, '0'); }
function stamp() {
  const d = new Date();
  return d.getFullYear() + '/' + pad(d.getMonth() + 1) + '/' + pad(d.getDate()) + ' ' +
         pad(d.getHours()) + ':' + pad(d.getMinutes()) + ':' + pad(d.getSeconds());
}
function shortTime(ts) { return String(ts || '').slice(11, 16); } // HH:MM

function toHalfWidth(str) {
  return String(str).replace(/[０-９]/g, c => String.fromCharCode(c.charCodeAt(0) - 0xFEE0));
}
function newId() {
  return Date.now().toString(36) + Math.random().toString(36).slice(2, 8);
}
function setMsg(id, text, isError) {
  const el = $(id);
  if (!el) return;
  el.textContent = text || '';
  el.classList.toggle('error', !!isError);
}

/* ---------- 状態（端末に保存） ---------- */
function emptyState() {
  return { user: null, history: [], queue: [] };
}
function load() {
  try {
    const raw = localStorage.getItem(STORAGE_KEY);
    if (!raw) return emptyState();
    const s = JSON.parse(raw);
    // 壊れたデータや古い形式は使わない
    return {
      user: s && s.user && typeof s.user.name === 'string' && s.user.name
        ? { name: s.user.name, budget: Number(s.user.budget) > 0 ? Number(s.user.budget) : DEFAULT_BUDGET }
        : null,
      history: Array.isArray(s && s.history) ? s.history : [],
      queue: Array.isArray(s && s.queue) ? s.queue : []
    };
  } catch (e) {
    return emptyState();
  }
}
function save() {
  try { localStorage.setItem(STORAGE_KEY, JSON.stringify(state)); } catch (e) { /* 保存できない環境 */ }
}

let state = load();

function spentTotal() {
  return state.history.reduce((sum, h) => sum + (Number(h.amount) || 0), 0);
}
function initialBudget() {
  return state.user ? state.user.budget : DEFAULT_BUDGET;
}
function balance() {
  return initialBudget() - spentTotal();
}

/* ---------- 画面表示 ---------- */
function showScreen(name) {
  $('screen-register').hidden = name !== 'register';
  $('screen-main').hidden = name !== 'main';
}

function render() {
  $('user-label').textContent = state.user ? state.user.name : '';
  $('balance').textContent = yen(balance());
  $('spent').textContent = '使用額 ' + yen(spentTotal()) + ' / 初期費用 ' + yen(initialBudget());

  const tbody = $('history');
  tbody.textContent = '';
  state.history.slice().reverse().forEach(h => {
    const tr = document.createElement('tr');
    [shortTime(h.timestamp), h.category, yen(h.amount), yen(h.balance)].forEach((v, i) => {
      const td = document.createElement('td');
      td.textContent = v;
      if (i >= 2) td.classList.add('num');
      tr.appendChild(td);
    });
    tbody.appendChild(tr);
  });
  $('history-empty').hidden = state.history.length > 0;
  renderSync();
}

function renderSync() {
  const el = $('sync-status');
  if (!GAS_URL) { el.textContent = ''; return; }
  el.textContent = state.queue.length
    ? '未送信の記録が ' + state.queue.length + ' 件あります（通信が回復すると自動で送信します）'
    : '';
}

/* ---------- スプレッドシート送信 ---------- */
function enqueue(type, category, amount, id, timestamp) {
  if (!GAS_URL || !state.user) return;
  state.queue.push({
    id: id || newId(),
    timestamp: timestamp || stamp(),
    name: state.user.name,
    type: type,
    category: category,
    amount: amount,
    spent: spentTotal(),
    balance: balance()
  });
}

let sending = false;
async function flushQueue() {
  if (!GAS_URL || sending || !state.queue.length) { renderSync(); return; }
  sending = true;
  try {
    while (state.queue.length) {
      // GAS はリダイレクトを返すため no-cors で送信（text/plain でプリフライトを回避）
      await fetch(GAS_URL, {
        method: 'POST',
        mode: 'no-cors',
        headers: { 'Content-Type': 'text/plain;charset=utf-8' },
        body: JSON.stringify(state.queue[0])
      });
      state.queue.shift(); // 通信エラーがなければ送信済みとみなす（重複は GAS 側で ID 判定）
      save();
    }
  } catch (err) {
    // オフラインなど：キューに残して後で再送
  } finally {
    sending = false;
    renderSync();
  }
}

/* ---------- 登録 ---------- */
function handleRegister(e) {
  e.preventDefault();
  const name = $('reg-name').value.trim().replace(/\s+/g, ' ');
  if (!name) {
    setMsg('reg-msg', 'ユーザー名を入力してください。', true);
    return;
  }
  const budgetRaw = toHalfWidth($('reg-budget').value.trim()).replace(/[,，]/g, '');
  const budget = Number(budgetRaw);
  if (!/^[0-9]+$/.test(budgetRaw) || budget <= 0) {
    setMsg('reg-msg', '初期費用を半角数字（1円以上）で入力してください。', true);
    return;
  }
  state = { user: { name: name, budget: budget }, history: [], queue: state.queue };
  enqueue('登録', '', 0);
  save();
  setMsg('reg-msg', '');
  render();
  showScreen('main');
  flushQueue();
}

/* ---------- 購入 ---------- */
function handleAmountInput(e) {
  const v = toHalfWidth(e.target.value).replace(/[^0-9]/g, '');
  if (v !== e.target.value) e.target.value = v;
}

function handleBuy(e) {
  e.preventDefault();
  const checked = document.querySelector('input[name="category"]:checked');
  const category = checked ? checked.value : 'その他';
  const raw = toHalfWidth($('amount').value.trim());
  const amount = Number(raw);

  if (!/^[0-9]+$/.test(raw) || amount <= 0) {
    setMsg('buy-msg', '金額を半角数字（1円以上）で入力してください。', true);
    return;
  }
  if (amount > balance()) {
    setMsg('buy-msg', '残金が足りません（残金 ' + yen(balance()) + '）。', true);
    return;
  }
  if (!confirm(category + '  ' + yen(amount) + '\nこの内容で記録しますか？')) return;

  const entry = {
    id: newId(),
    timestamp: stamp(),
    category: category,
    amount: amount,
    balance: balance() - amount
  };
  state.history.push(entry);
  enqueue('購入', category, amount, entry.id, entry.timestamp);
  save();

  $('amount').value = '';
  setMsg('buy-msg', '記録しました：' + category + ' ' + yen(amount));
  render();
  flushQueue();
}

/* ---------- リセット ---------- */
function handleReset() {
  if (!confirm('登録と購入履歴をこの端末から削除して、最初からやり直しますか？')) return;
  if (!confirm('本当に削除しますか？（元に戻せません）')) return;
  state = { user: null, history: [], queue: state.queue }; // 未送信分は残して送信を続ける
  save();
  $('reg-name').value = '';
  $('reg-budget').value = String(DEFAULT_BUDGET);
  setMsg('buy-msg', '');
  setMsg('reg-msg', '');
  showScreen('register');
}

/* ---------- 起動 ---------- */
function init() {
  $('form-register').addEventListener('submit', handleRegister);
  $('form-buy').addEventListener('submit', handleBuy);
  $('amount').addEventListener('input', handleAmountInput);
  $('reg-budget').addEventListener('input', handleAmountInput);
  $('btn-reset').addEventListener('click', handleReset);
  window.addEventListener('online', flushQueue);
  setInterval(flushQueue, 30000);

  if (state.user) {
    render();
    showScreen('main');
  } else {
    showScreen('register');
  }
  flushQueue();
}

// 予期しないエラーが起きたら画面に表示（原因が分かるように）
window.addEventListener('error', ev => {
  setMsg($('screen-main').hidden ? 'reg-msg' : 'buy-msg',
    'エラーが発生しました。ページを再読み込みしてください。（' + ev.message + '）', true);
});

if (document.readyState === 'loading') {
  document.addEventListener('DOMContentLoaded', init);
} else {
  init();
}
