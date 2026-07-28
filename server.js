// ============================================================
// 集成绩效管理平台 — 后端服务器
// Node.js 原生 http 模块，零依赖，JSON 文件存储
// ============================================================

const http = require('http');
const fs   = require('fs');
const path = require('path');

const PORT    = process.env.PORT || 3000;
const DB_FILE = path.join(__dirname, 'data.json');
const DB_BAK  = path.join(__dirname, 'data.json.bak');

// ---------- JSON 文件读写 ----------

function readDB() {
  try {
    const raw = fs.readFileSync(DB_FILE, 'utf-8');
    return JSON.parse(raw);
  } catch (e) {
    return { version: 0, lastModified: null, lastModifiedBy: '', stores: {} };
  }
}

function writeDB(state) {
  // 先备份旧文件
  if (fs.existsSync(DB_FILE)) {
    try { fs.copyFileSync(DB_FILE, DB_BAK); } catch (_) {}
  }
  fs.writeFileSync(DB_FILE, JSON.stringify(state, null, 2), 'utf-8');
}

// ---------- 静态文件 MIME ----------

const MIME = {
  '.html': 'text/html; charset=utf-8',
  '.js'  : 'application/javascript; charset=utf-8',
  '.css' : 'text/css; charset=utf-8',
  '.json': 'application/json; charset=utf-8',
  '.png' : 'image/png',
  '.svg' : 'image/svg+xml',
  '.ico' : 'image/x-icon',
};

function serveStatic(req, res) {
  let filePath = req.url === '/' ? '/dashboard_with_permissions.html' : req.url;
  filePath = path.join(__dirname, filePath);

  // 安全：禁止访问上级目录
  if (!filePath.startsWith(__dirname)) {
    res.writeHead(403);
    res.end('Forbidden');
    return;
  }

  const ext = path.extname(filePath).toLowerCase();
  const ct = MIME[ext] || 'application/octet-stream';

  fs.readFile(filePath, (err, data) => {
    if (err) {
      res.writeHead(404, { 'Content-Type': 'text/plain' });
      res.end('Not Found');
      return;
    }
    res.writeHead(200, { 'Content-Type': ct });
    res.end(data);
  });
}

// ---------- JSON body 解析 ----------

function parseBody(req) {
  return new Promise((resolve) => {
    let body = '';
    req.on('data', chunk => body += chunk);
    req.on('end', () => {
      try { resolve(JSON.parse(body)); } catch (_) { resolve(null); }
    });
  });
}

// ---------- 请求处理 ----------

const server = http.createServer(async (req, res) => {
  // CORS
  res.setHeader('Access-Control-Allow-Origin', '*');
  res.setHeader('Access-Control-Allow-Methods', 'GET, PUT, POST, OPTIONS');
  res.setHeader('Access-Control-Allow-Headers', 'Content-Type');

  if (req.method === 'OPTIONS') {
    res.writeHead(204);
    res.end();
    return;
  }

  const url = req.url.split('?')[0];

  // ---- API: 获取全量状态 ----
  if (url === '/api/state' && req.method === 'GET') {
    const state = readDB();
    res.writeHead(200, { 'Content-Type': 'application/json' });
    res.end(JSON.stringify(state));
    return;
  }

  // ---- API: 保存全量状态 ----
  if (url === '/api/state' && req.method === 'PUT') {
    const body = await parseBody(req);
    if (!body || !body.stores) {
      res.writeHead(400, { 'Content-Type': 'application/json' });
      res.end(JSON.stringify({ error: 'Invalid body, "stores" required' }));
      return;
    }
    const current = readDB();
    const newState = {
      version: current.version + 1,
      lastModified: new Date().toISOString(),
      lastModifiedBy: body.user || 'anonymous',
      stores: body.stores,
    };
    writeDB(newState);
    res.writeHead(200, { 'Content-Type': 'application/json' });
    res.end(JSON.stringify({ ok: true, version: newState.version, lastModified: newState.lastModified }));
    return;
  }

  // ---- API: 健康检查 ----
  if (url === '/api/health' && req.method === 'GET') {
    res.writeHead(200, { 'Content-Type': 'application/json' });
    res.end(JSON.stringify({ status: 'ok', time: new Date().toISOString() }));
    return;
  }

  // ---- 静态文件 ----
  serveStatic(req, res);
});

server.listen(PORT, () => {
  console.log(`✅ 绩效管理平台服务已启动: http://localhost:${PORT}`);
  console.log(`📁 数据文件: ${DB_FILE}`);
});
