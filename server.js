'use strict';
// ローカル専用サーバー。
// 静的ファイル（index.html / app.js / style.css など）を配信しつつ、
// POST /api/gemini-examples だけサーバー側で処理する。
// Gemini の API キーはこのプロセスの環境変数 GEMINI_API_KEY からのみ読み取り、
// ソースコードにも、ブラウザに返すレスポンスにも一切含めない。

const http = require('http');
const fs = require('fs');
const path = require('path');

const PORT = process.env.PORT || 10503;
const ROOT = __dirname;
const GEMINI_MODEL = 'gemini-2.5-flash';

const MIME = {
  '.html': 'text/html; charset=utf-8',
  '.js': 'text/javascript; charset=utf-8',
  '.css': 'text/css; charset=utf-8',
  '.json': 'application/json; charset=utf-8',
  '.png': 'image/png',
};

function serveStatic(req, res) {
  const urlPath = decodeURIComponent(req.url.split('?')[0]);
  const rel = urlPath === '/' ? 'index.html' : urlPath.replace(/^\/+/, '');
  const filePath = path.normalize(path.join(ROOT, rel));

  // ROOT の外に出るパス（ディレクトリトラバーサル）は拒否
  if (!filePath.startsWith(ROOT)) {
    res.writeHead(403);
    res.end('Forbidden');
    return;
  }

  fs.readFile(filePath, (err, data) => {
    if (err) {
      res.writeHead(404, { 'Content-Type': 'text/plain; charset=utf-8' });
      res.end('Not Found');
      return;
    }
    const ext = path.extname(filePath).toLowerCase();
    res.writeHead(200, { 'Content-Type': MIME[ext] || 'application/octet-stream' });
    res.end(data);
  });
}

function readJsonBody(req) {
  return new Promise((resolve, reject) => {
    let body = '';
    req.on('data', (chunk) => {
      body += chunk;
      if (body.length > 1e6) req.destroy(); // 念のための上限
    });
    req.on('end', () => {
      try { resolve(body ? JSON.parse(body) : {}); }
      catch (e) { reject(e); }
    });
    req.on('error', reject);
  });
}

async function handleGeminiExamples(req, res) {
  const apiKey = process.env.GEMINI_API_KEY;
  if (!apiKey) {
    res.writeHead(500, { 'Content-Type': 'application/json; charset=utf-8' });
    res.end(JSON.stringify({ error: '環境変数 GEMINI_API_KEY が設定されていません。' }));
    return;
  }

  let payload;
  try {
    payload = await readJsonBody(req);
  } catch {
    res.writeHead(400, { 'Content-Type': 'application/json; charset=utf-8' });
    res.end(JSON.stringify({ error: 'リクエストの形式が不正です。' }));
    return;
  }

  const en = (payload.en || '').toString().trim();
  const ja = (payload.ja || '').toString().trim();
  if (!en || !ja) {
    res.writeHead(400, { 'Content-Type': 'application/json; charset=utf-8' });
    res.end(JSON.stringify({ error: '単語情報が不足しています。' }));
    return;
  }

  const prompt =
    `英単語「${en}」（日本語訳: 「${ja}」）について、日本語で簡潔に回答してください。\n` +
    `1. この単語を使った例文を3つ、英語とその日本語訳のペアで挙げてください。\n` +
    `2. 日本語訳「${ja}」だけでは伝わりにくいニュアンスや使い分けがあれば、2〜3行で補足してください。特になければ「特になし」としてください。\n` +
    `見出しや箇条書きを使い、読みやすく整形してください。`;

  try {
    const apiRes = await fetch(
      `https://generativelanguage.googleapis.com/v1beta/models/${GEMINI_MODEL}:generateContent?key=${encodeURIComponent(apiKey)}`,
      {
        method: 'POST',
        headers: { 'Content-Type': 'application/json' },
        body: JSON.stringify({ contents: [{ parts: [{ text: prompt }] }] }),
      }
    );

    const data = await apiRes.json();

    if (!apiRes.ok) {
      const message = data?.error?.message || `Gemini API エラー (HTTP ${apiRes.status})`;
      res.writeHead(apiRes.status, { 'Content-Type': 'application/json; charset=utf-8' });
      res.end(JSON.stringify({ error: message }));
      return;
    }

    const text = data?.candidates?.[0]?.content?.parts?.map(p => p.text || '').join('') || '';
    if (!text) {
      res.writeHead(502, { 'Content-Type': 'application/json; charset=utf-8' });
      res.end(JSON.stringify({ error: 'Gemini から有効な応答が得られませんでした。' }));
      return;
    }

    res.writeHead(200, { 'Content-Type': 'application/json; charset=utf-8' });
    res.end(JSON.stringify({ text }));
  } catch (e) {
    res.writeHead(502, { 'Content-Type': 'application/json; charset=utf-8' });
    res.end(JSON.stringify({ error: 'Gemini API への接続に失敗しました: ' + e.message }));
  }
}

const server = http.createServer((req, res) => {
  if (req.method === 'POST' && req.url.split('?')[0] === '/api/gemini-examples') {
    handleGeminiExamples(req, res);
    return;
  }
  if (req.method === 'GET') {
    serveStatic(req, res);
    return;
  }
  res.writeHead(405);
  res.end('Method Not Allowed');
});

server.listen(PORT, () => {
  const keyStatus = process.env.GEMINI_API_KEY ? '検出済み' : '未設定（例文リクエスト機能は使えません）';
  console.log(`英単語クイズ: http://localhost:${PORT} で起動しました`);
  console.log(`GEMINI_API_KEY: ${keyStatus}`);
});
