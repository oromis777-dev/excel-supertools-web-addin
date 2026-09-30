const fs = require("fs");
const path = require("path");
const http = require("http");
const https = require("https");

const PORT = 3000;
const ROOT_DIR = __dirname;
const PFX_PATH = path.join(ROOT_DIR, "certs", "localhost.pfx");

const MIME_TYPES = {
  ".html": "text/html; charset=utf-8",
  ".js": "application/javascript; charset=utf-8",
  ".css": "text/css; charset=utf-8",
  ".png": "image/png",
  ".jpg": "image/jpeg",
  ".jpeg": "image/jpeg",
  ".svg": "image/svg+xml",
  ".xml": "application/xml; charset=utf-8",
  ".json": "application/json; charset=utf-8",
  ".ico": "image/x-icon"
};

function requestHandler(req, res) {
  // CORS Headers
  res.setHeader("Access-Control-Allow-Origin", "*");
  res.setHeader("Access-Control-Allow-Methods", "GET, POST, OPTIONS");
  res.setHeader("Access-Control-Allow-Headers", "*");

  if (req.method === "OPTIONS") {
    res.writeHead(200);
    res.end();
    return;
  }

  let reqPath = decodeURIComponent(req.url.split("?")[0]);
  if (reqPath === "/") reqPath = "/help.html";

  // URL Alias mappings for Office manifest compatibility
  const aliases = {
    "/taskpane.html": "/src/taskpane/taskpane.html",
    "/taskpane.js": "/src/taskpane/taskpane.js",
    "/taskpane.css": "/src/taskpane/taskpane.css",
    "/commands.html": "/src/commands/commands.html",
    "/commands.js": "/src/commands/commands.js"
  };

  if (aliases[reqPath]) {
    reqPath = aliases[reqPath];
  }

  const filePath = path.normalize(path.join(ROOT_DIR, reqPath));

  if (!filePath.startsWith(ROOT_DIR)) {
    res.writeHead(403, { "Content-Type": "text/plain; charset=utf-8" });
    res.end("403 Заборонено");
    return;
  }

  fs.stat(filePath, (err, stats) => {
    if (err || !stats.isFile()) {
      res.writeHead(404, { "Content-Type": "text/plain; charset=utf-8" });
      res.end(`404 Не знайдено: ${reqPath}`);
      return;
    }

    const ext = path.extname(filePath).toLowerCase();
    const contentType = MIME_TYPES[ext] || "application/octet-stream";

    res.writeHead(200, { "Content-Type": contentType });
    fs.createReadStream(filePath).pipe(res);
  });
}

function startServer() {
  function attachErrorHandler(srv) {
    srv.on("error", (err) => {
      if (err.code === "EADDRINUSE") {
        console.error(`\n❌ Помилка: Порт ${PORT} вже зайнятий іншим процесом!`);
        console.error(`Ймовірно, сервер уже запущено у сусідньому вікні або у фоні.`);
        console.error(`Щоб звільнити порт, виконайте команду: npm run stop\n`);
      } else {
        console.error("Помилка сервера:", err);
      }
      process.exit(1);
    });
  }

  if (fs.existsSync(PFX_PATH)) {
    try {
      const pfxData = fs.readFileSync(PFX_PATH);
      const server = https.createServer({ pfx: pfxData, passphrase: "exceldev" }, requestHandler);
      attachErrorHandler(server);
      server.listen(PORT, () => {
        console.log(`======================================================`);
        console.log(`🚀 ExcelSuperTools Web Add-in Server запущено (HTTPS)!`);
        console.log(`Адреса: https://localhost:${PORT}/`);
        console.log(`Taskpane: https://localhost:${PORT}/src/taskpane/taskpane.html`);
        console.log(`Commands: https://localhost:${PORT}/src/commands/commands.html`);
        console.log(`Help:     https://localhost:${PORT}/help.html`);
        console.log(`======================================================`);
      });
      return;
    } catch (e) {
      console.warn("Не вдалося завантажити PFX сертифікат:", e.message);
    }
  }

  // Fallback to HTTP
  const server = http.createServer(requestHandler);
  attachErrorHandler(server);
  server.listen(PORT, () => {
    console.log(`======================================================`);
    console.log(`⚠️  ExcelSuperTools Web Add-in Server запущено (HTTP)!`);
    console.log(`Адреса: http://localhost:${PORT}/`);
    console.log(`Підказка: для HTTPS запустіть certs\\setup_certs.ps1`);
    console.log(`======================================================`);
  });
}

startServer();
