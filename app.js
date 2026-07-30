'use strict';

const LS_DATA = 'eiwq.wordData.v1';
const LS_SESSION = 'eiwq.session.v1';
const LS_CUSTOM = 'eiwq.custom.v1';
const LS_GEMINI_KEY = 'eiwq.geminiKey.v1';
const SHEET_NAME = 'Wrk';
const QUIZ_SIZE = 15;
const GEMINI_MODEL = 'gemini-2.5-flash';

// 和英表現練習：単語クイズとは完全に独立したデータ領域
const LS_WAEI_DATA = 'waei.data.v1';
const LS_WAEI_SESSION = 'waei.session.v1';
const WAEI_JSON_FILE = 'expressions.json';

const $ = (id) => document.getElementById(id);

const screens = {
  home: $('home'),
  quiz: $('quiz'),
  result: $('result'),
  waeiHome: $('waeiHome'),
  waeiForm: $('waeiForm'),
  waeiQuiz: $('waeiQuiz'),
  waeiResult: $('waeiResult'),
};

function showScreen(name) {
  for (const k of Object.keys(screens)) {
    screens[k].hidden = (k !== name);
  }
}

// ---------- localStorage helpers ----------

function loadWordData() {
  try {
    const s = localStorage.getItem(LS_DATA);
    return s ? JSON.parse(s) : null;
  } catch { return null; }
}
function saveWordData(d) { localStorage.setItem(LS_DATA, JSON.stringify(d)); }

function loadSession() {
  try {
    const s = localStorage.getItem(LS_SESSION);
    return s ? JSON.parse(s) : null;
  } catch { return null; }
}
function saveSession(s) { localStorage.setItem(LS_SESSION, JSON.stringify(s)); }
function clearSession() { localStorage.removeItem(LS_SESSION); }

function loadCustomSettings() {
  try { return JSON.parse(localStorage.getItem(LS_CUSTOM)) || { x: 15, y: 5000 }; }
  catch { return { x: 15, y: 5000 }; }
}
function saveCustomSettings(x, y) { localStorage.setItem(LS_CUSTOM, JSON.stringify({ x, y })); }

function loadGeminiKey() {
  try { return localStorage.getItem(LS_GEMINI_KEY) || ''; } catch { return ''; }
}
function saveGeminiKey(key) { localStorage.setItem(LS_GEMINI_KEY, key); }
function clearGeminiKey() { localStorage.removeItem(LS_GEMINI_KEY); }

// ---------- Excel import ----------

async function importExcelFile(file) {
  const buf = await file.arrayBuffer();
  const wb = XLSX.read(buf, { type: 'array' });
  if (!wb.SheetNames.includes(SHEET_NAME)) {
    throw new Error(`シート "${SHEET_NAME}" が見つかりません。実在するシート: ${wb.SheetNames.join(', ')}`);
  }
  const ws = wb.Sheets[SHEET_NAME];
  // header:1 → 各行を配列で取得 / defval:'' → 空セルを '' に
  const rows = XLSX.utils.sheet_to_json(ws, { header: 1, defval: '', raw: false });

  // E列(index 4)=英語, F列(index 5)=日本語。 1行目はヘッダー扱いでスキップ
  const words = [];
  for (let i = 1; i < rows.length; i++) {
    const r = rows[i];
    const en = (r[4] ?? '').toString().trim();
    const ja = (r[5] ?? '').toString().trim();
    if (!en || !ja) continue;
    words.push({ row: words.length + 1, en, ja });
  }
  if (words.length === 0) {
    throw new Error('E列・F列に有効なデータが見つかりませんでした。');
  }

  // 単語は消されない前提のため、前回との差分＝今回新たに追加された単語数
  const previous = loadWordData();
  const previousCount = (previous && Array.isArray(previous.words)) ? previous.words.length : 0;
  const latestAddedCount = previousCount > 0
    ? Math.max(words.length - previousCount, 0)
    : words.length;

  const data = {
    importedAt: new Date().toISOString(),
    fileName: file.name,
    words,
    latestAddedCount,
  };
  saveWordData(data);
  return data;
}

// ---------- Sampling ----------

function shuffle(arr) {
  const a = arr.slice();
  for (let i = a.length - 1; i > 0; i--) {
    const j = Math.floor(Math.random() * (i + 1));
    [a[i], a[j]] = [a[j], a[i]];
  }
  return a;
}

function pickWords(allWords, mode, latestAddedCount) {
  let pool;
  let label;
  let quizSize = QUIZ_SIZE;
  switch (mode) {
    case 'latest50': {
      const n = Math.max(1, latestAddedCount || 50);
      pool = allWords.slice(-n);
      label = `Latest単語(${n}個)`;
      quizSize = n;
      break;
    }
    case 'all':
      pool = allWords;
      label = '完全ランダム';
      break;
    case 'bottom300':
      pool = allWords.slice(-300);
      label = '下から300';
      break;
    case 'bottom100':
      pool = allWords.slice(-100);
      label = '下から100';
      break;
    default:
      throw new Error('unknown mode: ' + mode);
  }
  const n = Math.min(quizSize, pool.length);
  return { words: shuffle(pool).slice(0, n), label };
}

// ---------- words.json fetch & export ----------

async function tryLoadFromJson() {
  try {
    const res = await fetch('words.json');
    if (!res.ok) return;
    const jsonData = await res.json();
    if (!jsonData.words || !jsonData.importedAt) return;
    const stored = loadWordData();
    if (!stored || jsonData.importedAt > stored.importedAt) {
      saveWordData(jsonData);
    }
  } catch { /* ローカルファイル起動時やネットワークエラーは無視 */ }
}

const WORDS_JSON_FOLDER = 'R:\\PUBTEMP\\FY26\\AI_Experiment\\英単語';

// 保存できた場合は { saved: true, pickerUsed } を返す。
// pickerUsed=true ならブラウザのネイティブ保存先選択ダイアログで直接そのフォルダに書き込めた。
// pickerUsed=false ならダウンロードフォルダに保存されたので、手動で移動する必要がある。
async function exportWords() {
  const data = loadWordData();
  if (!data) return { saved: false };
  const exportData = {
    importedAt: data.importedAt,
    fileName: data.fileName,
    words: data.words.slice(-3000),
    latestAddedCount: data.latestAddedCount,
  };
  const json = JSON.stringify(exportData);

  if (window.showSaveFilePicker) {
    try {
      const handle = await window.showSaveFilePicker({
        suggestedName: 'words.json',
        types: [{ description: 'JSON', accept: { 'application/json': ['.json'] } }],
      });
      const writable = await handle.createWritable();
      await writable.write(json);
      await writable.close();
      return { saved: true, pickerUsed: true };
    } catch (err) {
      if (err && err.name === 'AbortError') return { saved: false };
      // それ以外のエラー時はダウンロード方式にフォールバック
    }
  }

  const blob = new Blob([json], { type: 'application/json' });
  const url = URL.createObjectURL(blob);
  const a = document.createElement('a');
  a.href = url;
  a.download = 'words.json';
  a.click();
  URL.revokeObjectURL(url);
  return { saved: true, pickerUsed: false };
}

function showGitCommitReminder(pickerUsed) {
  const commands =
    `cd "${WORDS_JSON_FOLDER}"\n` +
    `git add words.json\n` +
    `git commit -m "単語データ更新"\n` +
    `git push`;

  if (pickerUsed) {
    alert(
      `words.json を保存しました。\n\n` +
      `続けて、以下をPowerShellで実行してgitにコミット・pushしてください:\n\n${commands}`
    );
  } else {
    alert(
      `words.json をダウンロードフォルダに保存しました。\n\n` +
      `1. ダウンロードされた words.json を次のフォルダに上書きしてください:\n` +
      `   ${WORDS_JSON_FOLDER}\n\n` +
      `2. 上書き後、以下をPowerShellで実行してgitにコミット・pushしてください:\n\n${commands}`
    );
  }
}

// ================================================================
// 和英表現練習（単語クイズとはデータ・localStorageキーとも完全に独立）
// ================================================================

function genWaeiId() {
  return Date.now().toString(36) + Math.random().toString(36).slice(2, 8);
}

function loadWaeiData() {
  try {
    const s = localStorage.getItem(LS_WAEI_DATA);
    const d = s ? JSON.parse(s) : null;
    return (d && Array.isArray(d.items)) ? d : { items: [], updatedAt: null };
  } catch { return { items: [], updatedAt: null }; }
}
function saveWaeiData(d) {
  d.updatedAt = new Date().toISOString();
  localStorage.setItem(LS_WAEI_DATA, JSON.stringify(d));
}

function addWaeiItem(ja, en) {
  const data = loadWaeiData();
  data.items.push({ id: genWaeiId(), ja, en, createdAt: new Date().toISOString() });
  saveWaeiData(data);
}
function updateWaeiItem(id, ja, en) {
  const data = loadWaeiData();
  const item = data.items.find((i) => i.id === id);
  if (item) { item.ja = ja; item.en = en; }
  saveWaeiData(data);
}
function deleteWaeiItem(id) {
  const data = loadWaeiData();
  data.items = data.items.filter((i) => i.id !== id);
  saveWaeiData(data);
}

function loadWaeiSession() {
  try {
    const s = localStorage.getItem(LS_WAEI_SESSION);
    return s ? JSON.parse(s) : null;
  } catch { return null; }
}
function saveWaeiSession(s) { localStorage.setItem(LS_WAEI_SESSION, JSON.stringify(s)); }
function clearWaeiSession() { localStorage.removeItem(LS_WAEI_SESSION); }

// ページ同梱の expressions.json（gitで永続管理）をローカル保存より新しければ取り込む
async function tryLoadWaeiFromJson() {
  try {
    const res = await fetch(WAEI_JSON_FILE);
    if (!res.ok) return;
    const jsonData = await res.json();
    if (!Array.isArray(jsonData.items)) return;
    const stored = loadWaeiData();
    if (!stored.updatedAt || (jsonData.updatedAt && jsonData.updatedAt > stored.updatedAt)) {
      localStorage.setItem(LS_WAEI_DATA, JSON.stringify({
        items: jsonData.items,
        updatedAt: jsonData.updatedAt || new Date().toISOString(),
      }));
    }
  } catch { /* ローカルファイル起動時やネットワークエラーは無視 */ }
}

async function exportWaeiData() {
  const data = loadWaeiData();
  if (!data.items.length) return { saved: false };
  const json = JSON.stringify({ items: data.items, updatedAt: data.updatedAt });

  if (window.showSaveFilePicker) {
    try {
      const handle = await window.showSaveFilePicker({
        suggestedName: WAEI_JSON_FILE,
        types: [{ description: 'JSON', accept: { 'application/json': ['.json'] } }],
      });
      const writable = await handle.createWritable();
      await writable.write(json);
      await writable.close();
      return { saved: true, pickerUsed: true };
    } catch (err) {
      if (err && err.name === 'AbortError') return { saved: false };
    }
  }

  const blob = new Blob([json], { type: 'application/json' });
  const url = URL.createObjectURL(blob);
  const a = document.createElement('a');
  a.href = url;
  a.download = WAEI_JSON_FILE;
  a.click();
  URL.revokeObjectURL(url);
  return { saved: true, pickerUsed: false };
}

function showWaeiGitCommitReminder(pickerUsed) {
  const commands =
    `cd "${WORDS_JSON_FOLDER}"\n` +
    `git add ${WAEI_JSON_FILE}\n` +
    `git commit -m "和英表現データ更新"\n` +
    `git push`;

  if (pickerUsed) {
    alert(
      `${WAEI_JSON_FILE} を保存しました。\n\n` +
      `続けて、以下をPowerShellで実行してgitにコミット・pushしてください:\n\n${commands}`
    );
  } else {
    alert(
      `${WAEI_JSON_FILE} をダウンロードフォルダに保存しました。\n\n` +
      `1. ダウンロードされた ${WAEI_JSON_FILE} を次のフォルダに上書きしてください:\n` +
      `   ${WORDS_JSON_FOLDER}\n\n` +
      `2. 上書き後、以下をPowerShellで実行してgitにコミット・pushしてください:\n\n${commands}`
    );
  }
}

// ---------- 和英表現練習：ホーム ----------

function renderWaeiHome() {
  const data = loadWaeiData();
  const session = loadWaeiSession();
  const count = data.items.length;

  $('waeiCountText').textContent = count
    ? `登録問題数: ${count} 個`
    : '問題が登録されていません。「問題を登録・編集する」から追加してください。';

  $('waeiStartAllBtn').disabled = count === 0;
  $('waeiCustomStartBtn').disabled = count === 0;
  $('waeiExportBtn').disabled = count === 0;
  if (count) {
    $('waeiCustomX').value = Math.min(Math.max(1, parseInt($('waeiCustomX').value) || 10), count);
  }

  const resumeBtn = $('waeiResumeBtn');
  if (session && session.items && session.currentIndex < session.items.length) {
    resumeBtn.hidden = false;
    resumeBtn.textContent = `前回の続きから（${session.round}周目 ${session.currentIndex + 1}/${session.items.length}）`;
  } else {
    resumeBtn.hidden = true;
  }

  $('waeiErrorMsg').textContent = '';
  showScreen('waeiHome');
}

// ---------- 和英表現練習：登録・編集 ----------

let waeiEditingId = null;

function resetWaeiForm() {
  waeiEditingId = null;
  $('waeiInputJa').value = '';
  $('waeiInputEn').value = '';
  $('waeiSaveItemBtn').textContent = '登録する';
  $('waeiCancelEditBtn').hidden = true;
  $('waeiFormError').textContent = '';
}

function renderWaeiForm() {
  resetWaeiForm();
  renderWaeiItemList();
  showScreen('waeiForm');
}

function renderWaeiItemList() {
  const data = loadWaeiData();
  const ul = $('waeiItemList');
  ul.innerHTML = '';
  const items = data.items.slice().reverse(); // 新しい登録を上に
  for (const it of items) {
    const li = document.createElement('li');

    const jaDiv = document.createElement('div');
    jaDiv.className = 'waei-item-ja';
    jaDiv.textContent = it.ja;

    const enDiv = document.createElement('div');
    enDiv.className = 'waei-item-en';
    enDiv.textContent = it.en;

    const actions = document.createElement('div');
    actions.className = 'waei-item-actions';

    const editBtn = document.createElement('button');
    editBtn.type = 'button';
    editBtn.textContent = '編集';
    editBtn.addEventListener('click', () => startEditWaeiItem(it.id));

    const delBtn = document.createElement('button');
    delBtn.type = 'button';
    delBtn.textContent = '削除';
    delBtn.addEventListener('click', () => {
      if (!confirm('この問題を削除しますか？')) return;
      deleteWaeiItem(it.id);
      if (waeiEditingId === it.id) resetWaeiForm();
      renderWaeiItemList();
    });

    actions.appendChild(editBtn);
    actions.appendChild(delBtn);
    li.appendChild(jaDiv);
    li.appendChild(enDiv);
    li.appendChild(actions);
    ul.appendChild(li);
  }
}

function startEditWaeiItem(id) {
  const data = loadWaeiData();
  const it = data.items.find((i) => i.id === id);
  if (!it) return;
  waeiEditingId = id;
  $('waeiInputJa').value = it.ja;
  $('waeiInputEn').value = it.en;
  $('waeiSaveItemBtn').textContent = '更新する';
  $('waeiCancelEditBtn').hidden = false;
  $('waeiFormError').textContent = '';
  window.scrollTo({ top: 0, behavior: 'smooth' });
}

// ---------- 和英表現練習：テスト ----------

function startWaeiQuiz(count) {
  const data = loadWaeiData();
  if (!data.items.length) return;
  const n = Math.min(Math.max(1, count), data.items.length);
  const session = {
    round: 1,
    currentIndex: 0,
    items: shuffle(data.items).slice(0, n),
    wrongIds: [],
    revealed: false,
  };
  saveWaeiSession(session);
  renderWaeiQuiz();
}

function renderWaeiQuiz() {
  const session = loadWaeiSession();
  if (!session) { renderWaeiHome(); return; }
  if (session.currentIndex >= session.items.length) { renderWaeiResult(); return; }

  const it = session.items[session.currentIndex];
  $('waeiRoundLabel').textContent = `${session.round}周目`;
  $('waeiProgressLabel').textContent = `${session.currentIndex + 1} / ${session.items.length}`;
  $('waeiJaText').textContent = it.ja;
  $('waeiEnText').textContent = it.en;

  if (session.revealed) {
    $('waeiEnText').hidden = false;
    $('waeiRevealBtn').hidden = true;
    $('waeiJudgeButtons').hidden = false;
  } else {
    $('waeiEnText').hidden = true;
    $('waeiRevealBtn').hidden = false;
    $('waeiJudgeButtons').hidden = true;
  }

  showScreen('waeiQuiz');
}

function waeiReveal() {
  const session = loadWaeiSession();
  if (!session) return;
  session.revealed = true;
  saveWaeiSession(session);
  renderWaeiQuiz();
}

function waeiJudge(isCorrect) {
  const session = loadWaeiSession();
  if (!session) return;
  const it = session.items[session.currentIndex];
  if (!isCorrect) session.wrongIds.push(it.id);
  session.currentIndex += 1;
  session.revealed = false;
  saveWaeiSession(session);
  if (session.currentIndex >= session.items.length) {
    renderWaeiResult();
  } else {
    renderWaeiQuiz();
  }
}

// ---------- 和英表現練習：結果 ----------

function renderWaeiResult() {
  const session = loadWaeiSession();
  if (!session) { renderWaeiHome(); return; }
  const wrongs = session.items.filter((it) => session.wrongIds.includes(it.id));
  const allCorrect = wrongs.length === 0;

  if (allCorrect) {
    $('waeiResultTitle').textContent = `🎉 ${session.round}周目で全問正解！`;
    $('waeiResultDetail').textContent = `${session.items.length} 問すべてできました。お疲れさまでした。`;
    $('waeiWrongList').innerHTML = '';
    $('waeiNextRoundBtn').hidden = true;
  } else {
    $('waeiResultTitle').textContent = `${session.round}周目 結果`;
    $('waeiResultDetail').textContent = `できた ${session.items.length - wrongs.length} / ${session.items.length}　できなかった ${wrongs.length} 個`;
    const ul = $('waeiWrongList');
    ul.innerHTML = '';
    for (const it of wrongs) {
      const li = document.createElement('li');
      const ja = document.createElement('span');
      ja.className = 'en';
      ja.textContent = it.ja;
      const en = document.createElement('span');
      en.className = 'ja';
      en.textContent = it.en;
      li.appendChild(ja);
      li.appendChild(en);
      ul.appendChild(li);
    }
    $('waeiNextRoundBtn').hidden = false;
  }
  showScreen('waeiResult');
}

function waeiNextRound() {
  const session = loadWaeiSession();
  if (!session) { renderWaeiHome(); return; }
  const wrongs = session.items.filter((it) => session.wrongIds.includes(it.id));
  if (wrongs.length === 0) {
    clearWaeiSession();
    renderWaeiHome();
    return;
  }
  session.round += 1;
  session.currentIndex = 0;
  session.items = shuffle(wrongs);
  session.wrongIds = [];
  session.revealed = false;
  saveWaeiSession(session);
  renderWaeiQuiz();
}

// ---------- Home rendering ----------

function renderHome() {
  const data = loadWordData();
  const session = loadSession();
  const modeBtns = document.querySelectorAll('.mode-btn');

  if (data) {
    $('wordCountText').textContent = `登録単語数: ${data.words.length.toLocaleString()} 個`;
    const dt = new Date(data.importedAt);
    const fmt = dt.toLocaleString('ja-JP');
    $('lastImportText').textContent = `最終取り込み: ${fmt}（${data.fileName}）`;
    const latestN = Math.max(1, data.latestAddedCount || 50);
    $('latest50Btn').textContent = `Latest単語(${latestN}個)`;
    modeBtns.forEach(b => b.disabled = false);
  } else {
    $('wordCountText').textContent = '単語データが読み込まれていません。下の「データ再取り込み」から Excel を読み込んでください。';
    $('lastImportText').textContent = '';
    modeBtns.forEach(b => b.disabled = true);
  }

  $('exportBtn').disabled = !data;

  const custom = loadCustomSettings();
  $('customX').value = custom.x;
  $('customY').value = custom.y;
  $('customStartBtn').disabled = !data;

  const resumeBtn = $('resumeBtn');
  if (session && session.words && session.currentIndex < session.words.length) {
    resumeBtn.hidden = false;
    resumeBtn.textContent = `前回の続きから（${session.modeLabel} / ${session.round}周目 ${session.currentIndex + 1}/${session.words.length}）`;
  } else {
    resumeBtn.hidden = true;
  }

  $('errorMsg').textContent = '';
  renderAiKeyStatus();
  showScreen('home');
}

function renderAiKeyStatus() {
  const key = loadGeminiKey();
  $('aiKeyStatus').textContent = key
    ? '✅ APIキーを保存済みです。この端末ではAIリクエストが直接Googleに送られます。'
    : '未設定です。未設定の場合、ローカルサーバー(node server.js)経由での利用を試みます。';
  $('aiKeyInput').value = '';
  $('aiKeyInput').placeholder = key ? '(保存済み。変更する場合のみ入力)' : 'Gemini APIキーを入力';
}

// ---------- Quiz ----------

function startCustomSession() {
  const data = loadWordData();
  if (!data) return;
  const x = Math.max(1, parseInt($('customX').value) || 15);
  const y = Math.max(1, parseInt($('customY').value) || 5000);
  saveCustomSettings(x, y);
  const effectiveY = Math.min(y, data.words.length);
  const pool = data.words.slice(-effectiveY);
  const n = Math.min(x, pool.length);
  const session = {
    mode: 'custom',
    modeLabel: `最新${effectiveY}件から${n}問`,
    round: 1,
    currentIndex: 0,
    words: shuffle(pool).slice(0, n),
    wrongIndices: [],
    revealed: false,
  };
  saveSession(session);
  renderQuiz();
}

function startNewSession(mode) {
  const data = loadWordData();
  if (!data) return;
  const { words, label } = pickWords(data.words, mode, data.latestAddedCount);
  const session = {
    mode,
    modeLabel: label,
    round: 1,
    currentIndex: 0,
    words,
    wrongIndices: [],
    revealed: false,
  };
  saveSession(session);
  renderQuiz();
}

function renderQuiz() {
  const session = loadSession();
  if (!session) { renderHome(); return; }
  if (session.currentIndex >= session.words.length) {
    renderRoundResult();
    return;
  }
  const w = session.words[session.currentIndex];
  $('roundLabel').textContent = `${session.round}周目（${session.modeLabel}）`;
  $('progressLabel').textContent = `${session.currentIndex + 1} / ${session.words.length}`;
  $('rowNum').textContent = `# ${w.row}`;
  $('englishWord').textContent = w.en;
  $('japaneseWord').textContent = w.ja;

  if (session.revealed) {
    $('japaneseWord').hidden = false;
    $('revealBtn').hidden = true;
    $('judgeButtons').hidden = false;
    $('aiExampleBox').hidden = false;
  } else {
    $('japaneseWord').hidden = true;
    $('revealBtn').hidden = false;
    $('judgeButtons').hidden = true;
    $('aiExampleBox').hidden = true;
  }

  // AI例文リクエストの状態は単語ごとにリセット
  const aiBtn = $('aiExampleBtn');
  const aiResult = $('aiExampleResult');
  aiBtn.hidden = false;
  aiBtn.disabled = false;
  aiBtn.textContent = '✨ 例文をAIにリクエスト';
  aiResult.hidden = true;
  aiResult.textContent = '';

  showScreen('quiz');
}

// ---------- 簡易Markdownレンダリング ----------

function escapeHtml(s) {
  return s.replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;');
}

function renderInline(text) {
  let s = escapeHtml(text);
  s = s.replace(/`([^`]+)`/g, '<code>$1</code>');
  s = s.replace(/\*\*([^*]+)\*\*/g, '<strong>$1</strong>');
  s = s.replace(/(^|[^*])\*([^*]+)\*(?!\*)/g, '$1<em>$2</em>');
  s = s.replace(/(^|[^_])_([^_]+)_(?!_)/g, '$1<em>$2</em>');
  return s;
}

function markdownToHtml(md) {
  const lines = md.replace(/\r\n/g, '\n').split('\n');
  const htmlParts = [];
  let listType = null; // 'ul' | 'ol'
  let paragraphLines = [];

  const flushParagraph = () => {
    if (paragraphLines.length) {
      htmlParts.push(`<p>${paragraphLines.map(renderInline).join('<br>')}</p>`);
      paragraphLines = [];
    }
  };
  const closeList = () => {
    if (listType) {
      htmlParts.push(`</${listType}>`);
      listType = null;
    }
  };

  for (const rawLine of lines) {
    const line = rawLine.trim();
    if (line === '') { flushParagraph(); closeList(); continue; }

    const headerMatch = line.match(/^(#{1,6})\s+(.*)$/);
    if (headerMatch) {
      flushParagraph();
      closeList();
      const level = Math.min(headerMatch[1].length + 2, 6);
      htmlParts.push(`<h${level}>${renderInline(headerMatch[2])}</h${level}>`);
      continue;
    }

    const ulMatch = line.match(/^[-*]\s+(.*)$/);
    if (ulMatch) {
      flushParagraph();
      if (listType !== 'ul') { closeList(); htmlParts.push('<ul>'); listType = 'ul'; }
      htmlParts.push(`<li>${renderInline(ulMatch[1])}</li>`);
      continue;
    }

    const olMatch = line.match(/^\d+[.)]\s+(.*)$/);
    if (olMatch) {
      flushParagraph();
      if (listType !== 'ol') { closeList(); htmlParts.push('<ol>'); listType = 'ol'; }
      htmlParts.push(`<li>${renderInline(olMatch[1])}</li>`);
      continue;
    }

    closeList();
    paragraphLines.push(line);
  }
  flushParagraph();
  closeList();
  return htmlParts.join('');
}

// ---------- AI例文リクエスト ----------

function buildGeminiPrompt(en, ja) {
  return (
    `英単語「${en}」（日本語訳: 「${ja}」）について、日本語で簡潔に回答してください。\n` +
    `1. 発音記号（IPA）と、カタカナで表現するなら何に近いかを示してください。カタカナ表記のうち、アクセント（強く読む部分）に当たる箇所は **太字** にしてください。\n` +
    `2. 品詞分類を、あてはまるものをすべて列挙してください（例: 名詞 / 自動詞 / 他動詞 / 形容詞 / 副詞 など）。特に動詞の場合は、自動詞・他動詞のどちらか、あるいは両方の用法があるかを明確にしてください。複数の品詞・用法がある場合、実際によく使われるのはどれかという傾向があれば、その旨も記載してください（例:「主に他動詞として使われる」等）。特に補足すべき傾向がなければその旨は省略して構いません。\n` +
    `3. この単語を使った例文を3つ、英語とその日本語訳のペアで挙げてください。\n` +
    `4. 日本語訳「${ja}」だけでは伝わりにくいニュアンスや使い分けがあれば、2〜3行で補足してください。特になければ「特になし」としてください。\n` +
    `見出しや箇条書きを使い、読みやすく整形してください。`
  );
}

// ブラウザから直接Gemini APIを呼ぶ経路。
// この端末のlocalStorageに保存されたキーのみを使用し、キーはGoogleへのリクエスト以外には送信しない。
async function callGeminiDirect(en, ja, apiKey) {
  const res = await fetch(
    `https://generativelanguage.googleapis.com/v1beta/models/${GEMINI_MODEL}:generateContent?key=${encodeURIComponent(apiKey)}`,
    {
      method: 'POST',
      headers: { 'Content-Type': 'application/json' },
      body: JSON.stringify({ contents: [{ parts: [{ text: buildGeminiPrompt(en, ja) }] }] }),
    }
  );
  const data = await res.json();
  if (!res.ok) throw new Error(data?.error?.message || `Gemini API エラー (HTTP ${res.status})`);
  const text = data?.candidates?.[0]?.content?.parts?.map(p => p.text || '').join('') || '';
  if (!text) throw new Error('Gemini から有効な応答が得られませんでした。');
  return text;
}

// ローカル開発サーバー(node server.js)経由の経路。キーはサーバー側の環境変数から読まれ、ブラウザには渡らない。
async function callGeminiViaServer(en, ja) {
  const res = await fetch('/api/gemini-examples', {
    method: 'POST',
    headers: { 'Content-Type': 'application/json' },
    body: JSON.stringify({ en, ja }),
  });
  const data = await res.json();
  if (!res.ok) throw new Error(data.error || `エラー (HTTP ${res.status})`);
  return data.text;
}

async function requestAiExamples() {
  const session = loadSession();
  if (!session) return;
  const w = session.words[session.currentIndex];
  const btn = $('aiExampleBtn');
  const resultEl = $('aiExampleResult');

  btn.disabled = true;
  btn.textContent = '生成中…';
  resultEl.hidden = true;

  const key = loadGeminiKey();
  try {
    const text = key ? await callGeminiDirect(w.en, w.ja, key) : await callGeminiViaServer(w.en, w.ja);
    resultEl.innerHTML = markdownToHtml(text);
    resultEl.hidden = false;
    btn.hidden = true;
  } catch (err) {
    const hint = key
      ? '（保存されているAPIキーが正しいか、ホーム画面の「AI機能のAPIキー設定」から確認してください）'
      : '（この機能は node server.js でローカルサーバーを起動しているか、ホーム画面の「AI機能のAPIキー設定」でGemini APIキーを登録している場合のみ利用できます）';
    const msg = 'AI例文の取得に失敗しました: ' + err.message + '\n' + hint;
    resultEl.innerHTML = escapeHtml(msg).replace(/\n/g, '<br>');
    resultEl.hidden = false;
    btn.disabled = false;
    btn.textContent = '✨ 例文をAIにリクエスト';
  }
}

function reveal() {
  const session = loadSession();
  if (!session) return;
  session.revealed = true;
  saveSession(session);
  renderQuiz();
}

function judge(isCorrect) {
  const session = loadSession();
  if (!session) return;
  if (!isCorrect) session.wrongIndices.push(session.currentIndex);
  session.currentIndex += 1;
  session.revealed = false;
  saveSession(session);
  if (session.currentIndex >= session.words.length) {
    renderRoundResult();
  } else {
    renderQuiz();
  }
}

// ---------- Round result ----------

function renderRoundResult() {
  const session = loadSession();
  if (!session) { renderHome(); return; }
  const wrongs = session.wrongIndices.map(i => session.words[i]);
  const allCorrect = wrongs.length === 0;

  if (allCorrect) {
    $('resultTitle').textContent = `🎉 ${session.round}周目で全問正解！クリアです`;
    $('resultDetail').textContent = `${session.words.length} 問すべて正解しました。お疲れさまでした。`;
    $('wrongList').innerHTML = '';
    $('nextRoundBtn').hidden = true;
  } else {
    $('resultTitle').textContent = `${session.round}周目 結果`;
    $('resultDetail').textContent = `正解 ${session.words.length - wrongs.length} / ${session.words.length}　不正解 ${wrongs.length} 個`;
    const ul = $('wrongList');
    ul.innerHTML = '';
    for (const w of wrongs) {
      const li = document.createElement('li');
      const en = document.createElement('span');
      en.className = 'en';
      en.textContent = `# ${w.row}  ${w.en}`;
      const ja = document.createElement('span');
      ja.className = 'ja';
      ja.textContent = w.ja;
      li.appendChild(en);
      li.appendChild(ja);
      ul.appendChild(li);
    }
    $('nextRoundBtn').hidden = false;
  }
  showScreen('result');
}

function nextRound() {
  const session = loadSession();
  if (!session) { renderHome(); return; }
  const wrongs = session.wrongIndices.map(i => session.words[i]);
  if (wrongs.length === 0) {
    clearSession();
    renderHome();
    return;
  }
  session.round += 1;
  session.currentIndex = 0;
  session.words = shuffle(wrongs);
  session.wrongIndices = [];
  session.revealed = false;
  saveSession(session);
  renderQuiz();
}

// ---------- Event wiring ----------

function showError(msg) {
  $('errorMsg').textContent = msg;
}

function bindEvents() {
  $('fileInput').addEventListener('change', async (e) => {
    const f = e.target.files[0];
    if (!f) return;
    showError('読み込み中…');
    try {
      const data = await importExcelFile(f);
      e.target.value = '';
      renderHome();

      const doExport = confirm(
        `単語データの取り込みが完了しました。\n` +
        `（新規追加: ${data.latestAddedCount}個 / 登録単語数合計: ${data.words.length}個）\n\n` +
        `words.json をエクスポートしますか？\n\n` +
        `保存先は次のフォルダを指定してください:\n${WORDS_JSON_FOLDER}`
      );
      if (doExport) {
        const result = await exportWords();
        if (result.saved) showGitCommitReminder(result.pickerUsed);
      }
    } catch (err) {
      showError('読み込みに失敗しました: ' + err.message);
    }
  });

  document.querySelectorAll('.mode-btn').forEach(btn => {
    btn.addEventListener('click', () => {
      if (btn.disabled) return;
      const mode = btn.dataset.mode;
      const existing = loadSession();
      if (existing && existing.currentIndex < existing.words.length) {
        if (!confirm('進行中のクイズがあります。新しく始めると進捗は失われます。続けますか？')) return;
      }
      clearSession();
      startNewSession(mode);
    });
  });

  $('resumeBtn').addEventListener('click', () => {
    const s = loadSession();
    if (s) renderQuiz();
  });

  $('backBtn').addEventListener('click', () => {
    if (confirm('クイズを中断しますか？\n進捗は自動保存されているので、ホームの「前回の続きから」でいつでも再開できます。')) {
      renderHome();
    }
  });

  $('revealBtn').addEventListener('click', reveal);
  $('aiExampleBtn').addEventListener('click', requestAiExamples);
  $('correctBtn').addEventListener('click', () => judge(true));
  $('wrongBtn').addEventListener('click', () => judge(false));

  $('exportBtn').addEventListener('click', async () => {
    const result = await exportWords();
    if (result.saved) showGitCommitReminder(result.pickerUsed);
  });

  $('customStartBtn').addEventListener('click', () => {
    if ($('customStartBtn').disabled) return;
    const existing = loadSession();
    if (existing && existing.currentIndex < existing.words.length) {
      if (!confirm('進行中のクイズがあります。新しく始めると進捗は失われます。続けますか？')) return;
    }
    clearSession();
    startCustomSession();
  });

  $('aiKeyToggleBtn').addEventListener('click', () => {
    $('aiKeyPanel').hidden = !$('aiKeyPanel').hidden;
  });

  $('aiKeySaveBtn').addEventListener('click', () => {
    const v = $('aiKeyInput').value.trim();
    if (!v) { alert('APIキーを入力してください。'); return; }
    saveGeminiKey(v);
    renderAiKeyStatus();
    alert('この端末にAPIキーを保存しました。');
  });

  $('aiKeyClearBtn').addEventListener('click', () => {
    if (!loadGeminiKey()) return;
    if (!confirm('この端末に保存済みのAPIキーを削除しますか？')) return;
    clearGeminiKey();
    renderAiKeyStatus();
  });

  $('nextRoundBtn').addEventListener('click', nextRound);
  $('homeBtn').addEventListener('click', () => {
    const s = loadSession();
    if (s && s.wrongIndices && s.wrongIndices.length === 0 && s.currentIndex >= s.words.length) {
      clearSession();
    }
    renderHome();
  });

  // ---- 和英表現練習 ----

  $('goWaeiBtn').addEventListener('click', () => renderWaeiHome());
  $('waeiBackToTopBtn').addEventListener('click', () => renderHome());
  $('waeiManageBtn').addEventListener('click', () => renderWaeiForm());
  $('waeiFormBackBtn').addEventListener('click', () => renderWaeiHome());

  $('waeiStartAllBtn').addEventListener('click', () => {
    if ($('waeiStartAllBtn').disabled) return;
    const existing = loadWaeiSession();
    if (existing && existing.currentIndex < existing.items.length) {
      if (!confirm('進行中のテストがあります。新しく始めると進捗は失われます。続けますか？')) return;
    }
    clearWaeiSession();
    startWaeiQuiz(loadWaeiData().items.length);
  });

  $('waeiCustomStartBtn').addEventListener('click', () => {
    if ($('waeiCustomStartBtn').disabled) return;
    const existing = loadWaeiSession();
    if (existing && existing.currentIndex < existing.items.length) {
      if (!confirm('進行中のテストがあります。新しく始めると進捗は失われます。続けますか？')) return;
    }
    const n = Math.max(1, parseInt($('waeiCustomX').value) || 10);
    clearWaeiSession();
    startWaeiQuiz(n);
  });

  $('waeiResumeBtn').addEventListener('click', () => {
    const s = loadWaeiSession();
    if (s) renderWaeiQuiz();
  });

  $('waeiQuizBackBtn').addEventListener('click', () => {
    if (confirm('テストを中断しますか？\n進捗は自動保存されているので、ホームの「前回の続きから」でいつでも再開できます。')) {
      renderWaeiHome();
    }
  });

  $('waeiRevealBtn').addEventListener('click', waeiReveal);
  $('waeiCorrectBtn').addEventListener('click', () => waeiJudge(true));
  $('waeiWrongBtn').addEventListener('click', () => waeiJudge(false));

  $('waeiSaveItemBtn').addEventListener('click', () => {
    const ja = $('waeiInputJa').value.trim();
    const en = $('waeiInputEn').value.trim();
    if (!ja || !en) {
      $('waeiFormError').textContent = '日本語と英語の両方を入力してください。';
      return;
    }
    if (waeiEditingId) {
      updateWaeiItem(waeiEditingId, ja, en);
    } else {
      addWaeiItem(ja, en);
    }
    resetWaeiForm();
    renderWaeiItemList();
  });

  $('waeiCancelEditBtn').addEventListener('click', resetWaeiForm);

  $('waeiExportBtn').addEventListener('click', async () => {
    const result = await exportWaeiData();
    if (result.saved) showWaeiGitCommitReminder(result.pickerUsed);
  });

  $('waeiNextRoundBtn').addEventListener('click', waeiNextRound);
  $('waeiHomeBtn').addEventListener('click', () => {
    const s = loadWaeiSession();
    if (s && s.wrongIds && s.wrongIds.length === 0 && s.currentIndex >= s.items.length) {
      clearWaeiSession();
    }
    renderWaeiHome();
  });

  // キーボードショートカット
  document.addEventListener('keydown', (e) => {
    if (!screens.quiz.hidden) {
      if (e.key === ' ' || e.key === 'Enter') {
        if (!$('revealBtn').hidden) { e.preventDefault(); reveal(); }
      } else if (e.key === 'o' || e.key === 'O' || e.key === 'ArrowRight') {
        if (!$('judgeButtons').hidden) { e.preventDefault(); judge(true); }
      } else if (e.key === 'x' || e.key === 'X' || e.key === 'ArrowLeft') {
        if (!$('judgeButtons').hidden) { e.preventDefault(); judge(false); }
      }
    } else if (!screens.waeiQuiz.hidden) {
      if (e.key === ' ' || e.key === 'Enter') {
        if (!$('waeiRevealBtn').hidden) { e.preventDefault(); waeiReveal(); }
      } else if (e.key === 'o' || e.key === 'O' || e.key === 'ArrowRight') {
        if (!$('waeiJudgeButtons').hidden) { e.preventDefault(); waeiJudge(true); }
      } else if (e.key === 'x' || e.key === 'X' || e.key === 'ArrowLeft') {
        if (!$('waeiJudgeButtons').hidden) { e.preventDefault(); waeiJudge(false); }
      }
    }
  });
}

// ---------- init ----------

bindEvents();
Promise.all([tryLoadFromJson(), tryLoadWaeiFromJson()]).then(() => renderHome());
