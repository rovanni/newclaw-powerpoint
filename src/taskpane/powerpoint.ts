/* global document, localStorage, crypto, fetch, Office, PowerPoint */

import { waitForTurnResponse } from "./turnPolling";

interface ChatAttachment {
  type: string;
  fileName: string;
  mimeType: string;
  data: string; // base64
}

interface ChatApiResponse {
  success: boolean;
  response?: string;
  sessionId?: string;
  /** Resposta assíncrona (HTTP 202): a resposta só sai por GET /api/chat/outbox?turnId=… */
  turnId?: string;
  attachments?: ChatAttachment[];
  error?: string;
}

const STORAGE_KEY_SERVER = "newclaw_server_url";
const STORAGE_KEY_TOKEN = "newclaw_token";
const STORAGE_KEY_SESSION = "newclaw_session_id";
// Vazio = a mesma origem do suplemento: o server.js repassa /api/chat e /api/chat/outbox ao NewClaw (como o harness da voz).
// O núcleo não precisa abrir exceção de origem para o suplemento. Um endereço preenchido nas configurações ainda é aceito
// (outro NewClaw), mas aí vale a política de origem daquele servidor.
const DEFAULT_SERVER_URL = "";

function getServerUrl(): string {
  return (localStorage.getItem(STORAGE_KEY_SERVER) || DEFAULT_SERVER_URL).replace(/\/+$/, "");
}

function getToken(): string {
  return localStorage.getItem(STORAGE_KEY_TOKEN) || "";
}

function getSessionId(): string {
  let id = localStorage.getItem(STORAGE_KEY_SESSION);
  if (!id) {
    id = `powerpoint-addin-${crypto.randomUUID()}`;
    localStorage.setItem(STORAGE_KEY_SESSION, id);
  }
  return id;
}

/**
 * Captura o contexto do slide ativo usando a API Office.js.
 * Retorna informacoes sobre o slide atual (numero, total, textos) para
 * que o agente saiba sobre o que o usuario esta trabalhando.
 * Falhas sao silenciadas — o chat funciona normalmente sem contexto.
 */
async function getSlideContext(): Promise<Record<string, unknown> | null> {
  try {
    return await PowerPoint.run(async (context) => {
      const slides = context.presentation.slides;
      slides.load("items/id");
      await context.sync();

      const totalSlides = slides.items.length;

      // Identifica o slide ativo
      const activeSlide = context.presentation.getSelectedSlides();
      activeSlide.load("items/id");
      await context.sync();

      let currentSlideNumber = 1;
      if (activeSlide.items.length > 0) {
        const activeId = activeSlide.items[0].id;
        for (let i = 0; i < slides.items.length; i++) {
          if (slides.items[i].id === activeId) {
            currentSlideNumber = i + 1;
            break;
          }
        }
      }

      // Captura textos do slide ativo
      const slideTexts: string[] = [];
      if (activeSlide.items.length > 0) {
        const slide = activeSlide.items[0];
        const shapes = slide.shapes;
        shapes.load("items/name,items/textFrame/textRange/text,items/textFrame/hasText");
        await context.sync();

        for (const shape of shapes.items) {
          try {
            if (shape.textFrame && shape.textFrame.hasText) {
              const text = shape.textFrame.textRange.text.trim();
              if (text) {
                slideTexts.push(text);
              }
            }
          } catch {
            // Shapes sem textFrame (imagens, graficos) — ignora silenciosamente
          }
        }
      }

      // Captura o nome do arquivo da apresentacao via Office Common API (se disponivel/salvo)
      let presentationTitle: string | undefined;
      try {
        presentationTitle = await new Promise<string | undefined>((resolve) => {
          if (Office && Office.context && Office.context.document && Office.context.document.getFilePropertiesAsync) {
            Office.context.document.getFilePropertiesAsync((asyncResult) => {
              if (asyncResult.status === Office.AsyncResultStatus.Succeeded && asyncResult.value.url) {
                const url = asyncResult.value.url;
                // Extrai apenas o nome do arquivo no final da URL/Caminho
                const fileName = url.substring(url.lastIndexOf('/') + 1).substring(url.lastIndexOf('\\') + 1);
                resolve(fileName || undefined);
              } else {
                resolve(undefined);
              }
            });
          } else {
            resolve(undefined);
          }
        });
      } catch {
        // Ignora erros na captura do nome do arquivo
      }

      return {
        presentationTitle,
        currentSlide: currentSlideNumber,
        totalSlides,
        slideTexts: slideTexts.length > 0 ? slideTexts : undefined,
      };
    });
  } catch {
    // API indisponivel ou erro de permissao — nao bloqueia o chat
    return null;
  }
}

/** O pedido com uma linha de fatos do PowerPoint aberto, quando houver (só o que a API do Office informou). */
function comContexto(message: string, ctx: Record<string, unknown> | null): string {
  if (!ctx) return message;
  const partes: string[] = [];
  if (typeof ctx.presentationTitle === "string" && ctx.presentationTitle) partes.push(`arquivo ${ctx.presentationTitle}`);
  if (typeof ctx.currentSlide === "number" && typeof ctx.totalSlides === "number") partes.push(`slide ${ctx.currentSlide} de ${ctx.totalSlides}`);
  return partes.length ? `[Enviado do PowerPoint: ${partes.join(", ")}]\n${message}` : message;
}

/**
 * Carrega servidor/token gerados pelo install.ps1 (config.local.json, servido junto do
 * bundle) na primeira execução. localStorage sempre tem prioridade — isso só preenche o
 * que o usuário ainda não configurou manualmente pelo painel de configurações.
 */
async function bootstrapFromInstaller(): Promise<void> {
  if (localStorage.getItem(STORAGE_KEY_SERVER) && localStorage.getItem(STORAGE_KEY_TOKEN)) return;
  try {
    const res = await fetch("config.local.json", { cache: "no-store" });
    if (!res.ok) return;
    const cfg = (await res.json()) as { serverUrl?: string; token?: string };
    // cfg.serverUrl é o destino do repasse do server.js, não do suplemento.
    if (cfg.token && !localStorage.getItem(STORAGE_KEY_TOKEN)) {
      localStorage.setItem(STORAGE_KEY_TOKEN, cfg.token);
    }
  } catch {
    // config.local.json não existe (instalação manual, sem install.ps1) — segue com os padrões
  }
}

Office.onReady(async (info) => {
  if (info.host !== Office.HostType.PowerPoint) return;

  await bootstrapFromInstaller();

  document.getElementById("sideload-msg")!.style.display = "none";
  document.getElementById("app-body")!.style.display = "flex";

  (document.getElementById("server-url") as HTMLInputElement).value = getServerUrl();
  (document.getElementById("server-token") as HTMLInputElement).value = getToken();

  document.getElementById("settings-toggle")!.onclick = toggleSettings;
  document.getElementById("save-settings")!.onclick = saveSettings;
  document.getElementById("send-button")!.onclick = () => void sendMessage();
  document.getElementById("message-input")!.addEventListener("keydown", (evt) => {
    const event = evt as KeyboardEvent;
    if (event.key === "Enter" && !event.shiftKey) {
      event.preventDefault();
      if (event.isComposing || event.keyCode === 229) return;
      void sendMessage();
    }
  });
});

function toggleSettings(): void {
  const panel = document.getElementById("settings-panel")!;
  panel.style.display = panel.style.display === "none" ? "flex" : "none";
}

function saveSettings(): void {
  const url = (document.getElementById("server-url") as HTMLInputElement).value.trim();
  const token = (document.getElementById("server-token") as HTMLInputElement).value.trim();
  if (url) localStorage.setItem(STORAGE_KEY_SERVER, url);
  if (token) localStorage.setItem(STORAGE_KEY_TOKEN, token);
  else localStorage.removeItem(STORAGE_KEY_TOKEN);
  toggleSettings();
}

function addMessage(role: "user" | "assistant" | "status" | "error", text: string): HTMLElement {
  const log = document.getElementById("chat-log")!;
  const bubble = document.createElement("div");
  bubble.className = `msg msg--${role}`;
  bubble.textContent = text;
  log.appendChild(bubble);
  log.scrollTop = log.scrollHeight;
  return bubble;
}

function formatElapsed(ms: number): string {
  const total = Math.floor(ms / 1000);
  const min = Math.floor(total / 60);
  const sec = total % 60;
  return min > 0 ? `${min} min ${sec} s` : `${sec} s`;
}

/** Texto da bolha enquanto o pedido é processado: tempo decorrido + dicas progressivas, para o usuário nunca ter que adivinhar. */
function processingText(elapsedMs: number): string {
  const secs = Math.floor(elapsedMs / 1000);
  const time = formatElapsed(elapsedMs);
  if (secs < 20) return `newclaw está processando… (${time})`;
  if (secs < 180) return `newclaw está processando… (${time})\nTarefas com várias etapas podem levar alguns minutos.`;
  return `newclaw ainda está trabalhando… (${time})\nSe a tarefa precisar de aprovação, confirme no Dashboard do newclaw.`;
}

async function sendMessage(): Promise<void> {
  const input = document.getElementById("message-input") as HTMLTextAreaElement;
  const message = input.value.trim();
  if (!message) return;

  const sendButton = document.getElementById("send-button") as HTMLButtonElement;
  addMessage("user", message);
  input.value = "";
  sendButton.disabled = true;
  const statusBubble = addMessage("status", "newclaw está processando…");

  try {
    const serverUrl = getServerUrl();
    const token = getToken();
    const headers: Record<string, string> = { "Content-Type": "application/json" };
    if (token) headers["Authorization"] = `Bearer ${token}`;

    const slideContext = await getSlideContext();

    const res = await fetch(`${serverUrl}/api/chat`, {
      method: "POST",
      headers,
      body: JSON.stringify({
        message: comContexto(message, slideContext),
        sessionId: getSessionId(),
      }),
    });

    const data = (await res.json()) as ChatApiResponse;

    if (!res.ok || !data.success) {
      statusBubble.remove();
      addMessage("error", data.error || `Erro ${res.status} ao falar com o newclaw.`);
      return;
    }

    if (res.status === 202 && data.turnId) {
      await awaitTurnAndShow(data.turnId, serverUrl, token, statusBubble);
    } else {
      // Caminho legado (servidor antigo): a resposta vem no próprio POST.
      statusBubble.remove();
      await showReply(data);
    }
  } catch (err) {
    statusBubble.remove();
    const detail = err instanceof Error ? err.message : String(err);
    addMessage("error", `Falha de conexão com ${getServerUrl()}: ${detail}`);
  } finally {
    sendButton.disabled = false;
    input.focus();
  }
}

/**
 * Espera a resposta de um turno assíncrono e a exibe. Se a espera acabar sem resposta (limite de tempo ou conexão), o turno NÃO é
 * perdido: a resposta fica na outbox do servidor, e o aviso traz o botão "Verificar resposta", que chama esta função de novo com o
 * mesmo turnId (a outbox é consumida uma vez, por isso só um consulente de cada vez: o botão some ao ser clicado).
 */
async function awaitTurnAndShow(turnId: string, serverUrl: string, token: string | null, statusBubble: HTMLElement): Promise<void> {
  // Servidor assíncrono: o POST só aceitou o pedido. A bolha de status fica VISÍVEL, com o tempo decorrido, até a resposta chegar.
  statusBubble.style.whiteSpace = "pre-line";
  const getHeaders: Record<string, string> = {};
  if (token) getHeaders["Authorization"] = `Bearer ${token}`;
  const outcome = await waitForTurnResponse({
    fetchFn: (url, init) => fetch(url, init),
    serverUrl,
    headers: getHeaders,
    turnId,
    onTick: (elapsedMs) => { statusBubble.textContent = processingText(elapsedMs); },
  });
  statusBubble.remove();

  if (outcome.kind === "timeout" || outcome.kind === "unreachable") {
    const why = outcome.kind === "timeout"
      ? `O newclaw ainda não terminou depois de ${formatElapsed(outcome.elapsedMs)}.`
      : `Perdi a conexão com o newclaw enquanto esperava a resposta (${outcome.lastError}).`;
    const bubble = addMessage("error", `${why} O pedido continua em execução no servidor e a resposta fica guardada: use "Verificar resposta" para continuar esperando. (turno ${turnId})`);
    const retry = document.createElement("button");
    retry.type = "button";
    retry.className = "msg-retry";
    retry.textContent = "Verificar resposta";
    retry.addEventListener("click", () => {
      retry.remove();
      const sendButton = document.getElementById("send-button") as HTMLButtonElement;
      sendButton.disabled = true;
      const again = addMessage("status", "Verificando a resposta…");
      awaitTurnAndShow(turnId, serverUrl, token, again)
        .catch((err) => { again.remove(); addMessage("error", `Falha ao verificar a resposta: ${err instanceof Error ? err.message : String(err)}`); })
        .finally(() => { sendButton.disabled = false; });
    });
    bubble.appendChild(document.createElement("br"));
    bubble.appendChild(retry);
    return;
  }
  if (outcome.kind === "rejected") {
    addMessage("error", `O servidor recusou a consulta da resposta (HTTP ${outcome.status}). Verifique o token configurado no suplemento. (turno ${turnId})`);
    return;
  }
  await showReply(outcome.payload as { response?: string; attachments?: ChatAttachment[] });
}

async function showReply(reply: { response?: string; attachments?: ChatAttachment[] }): Promise<void> {
  if (reply.response) {
    addMessage("assistant", reply.response);
  } else {
    // Nunca em silêncio: o usuário não deve ter que adivinhar se terminou.
    addMessage("status", "O newclaw terminou, mas não devolveu nenhum texto.");
  }

  const attachments = reply.attachments || [];
  const pptxAttachment = attachments.find((a) => a.fileName?.toLowerCase().endsWith(".pptx"));

  if (pptxAttachment) {
    await insertSlidesFromAttachment(pptxAttachment);
  }

  for (const att of attachments) {
    if (att === pptxAttachment) continue;
    addMessage("status", `Anexo recebido (não inserido automaticamente): ${att.fileName}`);
  }
}

async function insertSlidesFromAttachment(attachment: ChatAttachment): Promise<void> {
  const statusBubble = addMessage("status", `Inserindo "${attachment.fileName}" na apresentação…`);
  try {
    await PowerPoint.run(async (context) => {
      const slides = context.presentation.slides;
      slides.load("items/id");
      await context.sync();

      const options: PowerPoint.InsertSlideOptions = {
        formatting: PowerPoint.InsertSlideFormatting.useDestinationTheme,
      };
      // Sem targetSlideId a inserção cai no início da apresentação — usamos o
      // último slide existente como alvo para anexar ao final.
      if (slides.items.length > 0) {
        options.targetSlideId = slides.items[slides.items.length - 1].id;
      }

      context.presentation.insertSlidesFromBase64(attachment.data, options);
      await context.sync();
    });
    statusBubble.textContent = `"${attachment.fileName}" inserido na apresentação.`;
    statusBubble.classList.add("msg--success");
  } catch (err) {
    statusBubble.remove();
    const detail = err instanceof Error ? err.message : String(err);
    addMessage(
      "error",
      `Não foi possível inserir os slides automaticamente (${detail}). O arquivo "${attachment.fileName}" foi gerado pelo newclaw — insira manualmente via Inserir > Reutilizar Slides, se necessário.`
    );
  }
}
