/**
 * Servidor HTTPS do suplemento: serve o bundle de produção (dist/) e REPASSA as duas rotas do contrato público do turno
 * (`POST /api/chat`, `GET /api/chat/outbox`) ao NewClaw — o mesmo papel do harness da interface de voz.
 *
 * Por que repassar: o manifest.xml fixa https://localhost:3000, e o NewClaw escuta noutro endereço. Antes de 04/11/2026 o
 * núcleo abria uma exceção de origem só para o suplemento; com o suplemento fora do núcleo, ele fala com o NewClaw pela
 * mesma origem, e o núcleo não precisa conhecê-lo. Só essas duas rotas passam: aprovar ação perigosa, configuração e o
 * resto do painel continuam fora do alcance do suplemento.
 *
 * Destino: NEWCLAW_URL, ou o `serverUrl` que o install.ps1 grava em config.local.json, ou http://127.0.0.1:3090.
 */
const https = require("https");
const http = require("http");
const fs = require("fs");
const path = require("path");
const devCerts = require("office-addin-dev-certs");

const PORT = process.env.PORT || 3000;
const DIST_DIR = path.join(__dirname, "dist");
const ROTAS_REPASSADAS = { "/api/chat": "POST", "/api/chat/outbox": "GET" };

function destinoDoNewClaw() {
  if (process.env.NEWCLAW_URL) return process.env.NEWCLAW_URL;
  try {
    const cfg = JSON.parse(fs.readFileSync(path.join(DIST_DIR, "config.local.json"), "utf-8"));
    if (cfg.serverUrl) return cfg.serverUrl;
  } catch {
    // sem config.local.json (instalação manual): o padrão abaixo
  }
  return "http://127.0.0.1:3090";
}

const MIME_TYPES = {
  ".html": "text/html; charset=utf-8",
  ".js": "application/javascript",
  ".css": "text/css",
  ".json": "application/json",
  ".map": "application/json",
  ".png": "image/png",
  ".ico": "image/x-icon",
};

function repassar(req, res) {
  const alvo = new URL(req.url, destinoDoNewClaw());
  const cliente = alvo.protocol === "https:" ? https : http;
  const cabecalhos = { "content-type": req.headers["content-type"] || "application/json" };
  if (req.headers.authorization) cabecalhos.authorization = req.headers.authorization;
  const saida = cliente.request(alvo, { method: req.method, headers: cabecalhos }, (resp) => {
    const volta = { "content-type": resp.headers["content-type"] || "application/json" };
    if (resp.headers["retry-after"]) volta["retry-after"] = resp.headers["retry-after"];
    res.writeHead(resp.statusCode || 502, volta);
    resp.pipe(res);
  });
  saida.on("error", (err) => {
    res.writeHead(502, { "content-type": "application/json" });
    res.end(JSON.stringify({ success: false, error: `NewClaw não respondeu em ${alvo.origin}: ${err.message}` }));
  });
  req.pipe(saida);
}

function requestHandler(req, res) {
  const urlPath = decodeURIComponent(req.url.split("?")[0]);
  if (ROTAS_REPASSADAS[urlPath]) {
    if (req.method !== ROTAS_REPASSADAS[urlPath]) {
      res.writeHead(405);
      res.end("Method Not Allowed");
      return;
    }
    repassar(req, res);
    return;
  }

  const requested = path.join(DIST_DIR, urlPath === "/" ? "/taskpane.html" : urlPath);
  const relative = path.relative(DIST_DIR, requested);

  if (relative.startsWith("..") || path.isAbsolute(relative)) {
    res.writeHead(403);
    res.end("Forbidden");
    return;
  }

  fs.readFile(requested, (err, data) => {
    if (err) {
      res.writeHead(404);
      res.end("Not found");
      return;
    }
    res.writeHead(200, { "Content-Type": MIME_TYPES[path.extname(requested)] || "application/octet-stream" });
    res.end(data);
  });
}

module.exports = { requestHandler, destinoDoNewClaw };

if (require.main === module) {
  devCerts
    .getHttpsServerOptions()
    .then((options) => {
      https.createServer(options, requestHandler).listen(PORT, () => {
        console.log(`newclaw-powerpoint servindo ${DIST_DIR} em https://localhost:${PORT} → NewClaw em ${destinoDoNewClaw()}`);
      });
    })
    .catch((err) => {
      console.error("Não foi possível carregar os certificados de desenvolvimento:", err);
      process.exit(1);
    });
}
