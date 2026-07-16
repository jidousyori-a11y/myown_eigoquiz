'use strict';

const LS_DATA = 'eiwq.wordData.v1';
const LS_SESSION = 'eiwq.session.v1';
const LS_CUSTOM = 'eiwq.custom.v1';
const SHEET_NAME = 'Wrk';
const QUIZ_SIZE = 15;

const $ = (id) => document.getElementById(id);

const screens = {
  home: $('home'),
  quiz: $('quiz'),
  result: $('result'),
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

function exportWords() {
  const data = loadWordData();
  if (!data) return;
  const exportData = {
    importedAt: data.importedAt,
    fileName: data.fileName,
    words: data.words.slice(-3000),
    latestAddedCount: data.latestAddedCount,
  };
  const blob = new Blob([JSON.stringify(exportData)], { type: 'application/json' });
  const url = URL.createObjectURL(blob);
  const a = document.createElement('a');
  a.href = url;
  a.download = 'words.json';
  a.click();
  URL.revokeObjectURL(url);
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
  showScreen('home');
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

async function requestAiExamples() {
  const session = loadSession();
  if (!session) return;
  const w = session.words[session.currentIndex];
  const btn = $('aiExampleBtn');
  const resultEl = $('aiExampleResult');

  btn.disabled = true;
  btn.textContent = '生成中…';
  resultEl.hidden = true;

  try {
    const res = await fetch('/api/gemini-examples', {
      method: 'POST',
      headers: { 'Content-Type': 'application/json' },
      body: JSON.stringify({ en: w.en, ja: w.ja }),
    });
    const data = await res.json();
    if (!res.ok) throw new Error(data.error || `エラー (HTTP ${res.status})`);
    resultEl.innerHTML = markdownToHtml(data.text);
    resultEl.hidden = false;
    btn.hidden = true;
  } catch (err) {
    const msg =
      'AI例文の取得に失敗しました: ' + err.message +
      '\n（この機能は node server.js でローカルサーバーを起動している場合のみ利用できます）';
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
      await importExcelFile(f);
      e.target.value = '';
      renderHome();
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

  $('exportBtn').addEventListener('click', exportWords);

  $('customStartBtn').addEventListener('click', () => {
    if ($('customStartBtn').disabled) return;
    const existing = loadSession();
    if (existing && existing.currentIndex < existing.words.length) {
      if (!confirm('進行中のクイズがあります。新しく始めると進捗は失われます。続けますか？')) return;
    }
    clearSession();
    startCustomSession();
  });

  $('nextRoundBtn').addEventListener('click', nextRound);
  $('homeBtn').addEventListener('click', () => {
    const s = loadSession();
    if (s && s.wrongIndices && s.wrongIndices.length === 0 && s.currentIndex >= s.words.length) {
      clearSession();
    }
    renderHome();
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
    }
  });
}

// ---------- init ----------

bindEvents();
tryLoadFromJson().then(() => renderHome());
