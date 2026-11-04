// O server.js repassa ao NewClaw SÓ as duas rotas do contrato do turno, com o token, e devolve a resposta como veio.
// Roda com `npm test` (node:test). Sem certificado: o requestHandler é montado num servidor HTTP comum, contra um NewClaw falso.
const { test } = require("node:test");
const assert = require("node:assert");
const http = require("node:http");

function ouvir(handler) {
  return new Promise((ok) => {
    const s = http.createServer(handler).listen(0, "127.0.0.1", () => ok(s));
  });
}

test("repasse das rotas do contrato do turno", async () => {
  const recebidos = [];
  const newclaw = await ouvir((req, res) => {
    let corpo = "";
    req.on("data", (c) => (corpo += c));
    req.on("end", () => {
      recebidos.push({ metodo: req.method, url: req.url, auth: req.headers.authorization, corpo });
      res.writeHead(req.url.startsWith("/api/chat/outbox") ? 200 : 202, { "content-type": "application/json", "retry-after": "7" });
      res.end(JSON.stringify({ success: true, turnId: "t1", response: "oi" }));
    });
  });
  process.env.NEWCLAW_URL = `http://127.0.0.1:${newclaw.address().port}`;
  delete require.cache[require.resolve("../server.js")];
  const { requestHandler } = require("../server.js");
  const suplemento = await ouvir(requestHandler);
  const base = `http://127.0.0.1:${suplemento.address().port}`;

  try {
    const envio = await fetch(`${base}/api/chat`, {
      method: "POST", headers: { "content-type": "application/json", authorization: "Bearer abc" },
      body: JSON.stringify({ message: "olá", sessionId: "powerpoint-addin-x" }),
    });
    assert.strictEqual(envio.status, 202);
    assert.strictEqual(envio.headers.get("retry-after"), "7");
    assert.deepStrictEqual(await envio.json(), { success: true, turnId: "t1", response: "oi" });
    assert.strictEqual(recebidos[0].auth, "Bearer abc", "o token vai junto");
    assert.strictEqual(JSON.parse(recebidos[0].corpo).message, "olá", "o corpo vai inteiro");

    const outbox = await fetch(`${base}/api/chat/outbox?turnId=t1`, { headers: { authorization: "Bearer abc" } });
    assert.strictEqual(outbox.status, 200);
    assert.strictEqual(recebidos[1].url, "/api/chat/outbox?turnId=t1", "a consulta vai com o turnId");

    // Fora do contrato do turno nada é repassado: aprovação, configuração e o resto do painel.
    for (const rota of ["/api/chat/auth-decision", "/api/config", "/api/chat/active?sessionId=x"]) {
      const r = await fetch(`${base}${rota}`, { method: rota.includes("auth") ? "POST" : "GET" });
      assert.notStrictEqual(r.status, 202, `${rota} não é repassada`);
    }
    assert.strictEqual(recebidos.length, 2, "o NewClaw só recebeu as duas chamadas do contrato");

    // Método errado numa rota repassada não chega ao NewClaw.
    const errado = await fetch(`${base}/api/chat`, { method: "GET" });
    assert.strictEqual(errado.status, 405);
    assert.strictEqual(recebidos.length, 2);
  } finally {
    suplemento.close();
    newclaw.close();
  }
});

test("NewClaw fora do ar vira 502 com o endereço tentado", async () => {
  process.env.NEWCLAW_URL = "http://127.0.0.1:9";
  delete require.cache[require.resolve("../server.js")];
  const { requestHandler } = require("../server.js");
  const suplemento = await ouvir(requestHandler);
  try {
    const r = await fetch(`http://127.0.0.1:${suplemento.address().port}/api/chat`, { method: "POST", body: "{}" });
    assert.strictEqual(r.status, 502);
    assert.match((await r.json()).error, /127\.0\.0\.1:9/);
  } finally {
    suplemento.close();
  }
});
