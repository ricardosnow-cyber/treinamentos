import express from "express";
import cors from "cors";
import multer from "multer";
import JSZip from "jszip";
import { XMLParser } from "fast-xml-parser";
import { promises as fs } from "node:fs";
import path from "node:path";
import { fileURLToPath } from "node:url";
import { randomUUID, timingSafeEqual } from "node:crypto";
const ROOT = path.dirname(fileURLToPath(import.meta.url));
const PUBLIC_DIR = path.join(ROOT, "public");
const STORAGE_DIR = path.join(ROOT, "storage");
const COSTS_FILE = path.join(STORAGE_DIR, "costs.json");
const ADMIN_USER = process.env.ADMIN_USER || "";
const ADMIN_PASSWORD = process.env.ADMIN_PASSWORD || "";
const HEYGEN_API_KEY = process.env.HEYGEN_API_KEY || "";
const DEFAULT_HEYGEN_AVATAR_ID = process.env.HEYGEN_AVATAR_ID || "";
const DEFAULT_HEYGEN_VOICE_ID = process.env.HEYGEN_VOICE_ID || "";
const DAILY_COST_LIMIT_USD = 5;
const COST_TIME_ZONE = process.env.COST_TIME_ZONE || "America/Sao_Paulo";
const COST_RATES = { heygen: 0.05, "openai-image": 0.04, "heygen-ghibli": 0.05 };
await fs.mkdir(STORAGE_DIR, { recursive: true });
try { await fs.access(COSTS_FILE); } catch { await fs.writeFile(COSTS_FILE, "[]", "utf8"); }
const app = express();
const port = process.env.PORT || 3000;
const parser = new XMLParser({
  ignoreAttributes: false,
  attributeNamePrefix: "",
  trimValues: false
});
// This app serves its own frontend; do not enable wildcard cross-origin API access.
// app.use(cors()) is intentionally disabled to reduce cross-site API exposure.
app.use(express.json({ limit: "25mb" }));
app.use(express.static(PUBLIC_DIR));
app.use("/storage", express.static(STORAGE_DIR));
const upload = multer({
  storage: multer.memoryStorage(),
  limits: { fileSize: 100 * 1024 * 1024 }
});
const array = value => (value == null ? [] : Array.isArray(value) ? value : [value]);
function extractTexts(value, result = []) {
  if (value == null) return result;
  if (Array.isArray(value)) {
    value.forEach(item => extractTexts(item, result));
    return result;
  }
  if (typeof value !== "object") return result;
  for (const [key, child] of Object.entries(value)) {
    if (key === "a:t") {
      array(child).forEach(item => result.push(String(item)));
    } else {
      extractTexts(child, result);
    }
  }
  return result;
}
async function zipText(zip, filename) {
  const file = zip.file(filename);
  return file ? file.async("string") : null;
}
function zipPath(base, target) {
  const clean = decodeURIComponent(String(target || "")).replaceAll("\\", "/").replace(/^\//, "");
  const result = clean.startsWith("ppt/") ? clean : path.posix.normalize(path.posix.join(base, clean));
  if (result.startsWith("../")) throw new Error("Caminho inválido no PPTX");
  return result;
}
function mime(filename) {
  const ext = path.posix.extname(filename).toLowerCase();
  return {
    ".png": "image/png",
    ".jpg": "image/jpeg",
    ".jpeg": "image/jpeg",
    ".gif": "image/gif",
    ".bmp": "image/bmp",
    ".svg": "image/svg+xml",
    ".webp": "image/webp"
  }[ext] || "application/octet-stream";
}
async function parsePptx(buffer) {
  const zip = await JSZip.loadAsync(buffer);
  const slideFiles = Object.keys(zip.files)
    .filter(name => /^ppt\/slides\/slide\d+\.xml$/.test(name))
    .sort((a, b) => {
      const na = Number(a.match(/slide(\d+)/)?.[1] || 0);
      const nb = Number(b.match(/slide(\d+)/)?.[1] || 0);
      return na - nb;
    });
  if (!slideFiles.length) {
    throw new Error("Nenhum slide foi encontrado no arquivo PPTX.");
  }
  const slides = [];
  for (const slideFile of slideFiles) {
    const number = Number(slideFile.match(/slide(\d+)/)?.[1] || 0);
    const slideXml = await zipText(zip, slideFile);
    const slideDocument = parser.parse(slideXml || "");
    const paragraphs = extractTexts(slideDocument)
      .map(text => text.replace(/\s+/g, " ").trim())
      .filter(Boolean);
    const baseName = path.posix.basename(slideFile);
    const relFile = `ppt/slides/_rels/${baseName}.rels`;
    const relXml = await zipText(zip, relFile);
    const relDocument = relXml ? parser.parse(relXml) : {};
    const relationships = array(relDocument.Relationships?.Relationship);
    const images = [];
    for (const relation of relationships) {
      if (!String(relation.Type || "").endsWith("/image")) continue;
      if (relation.TargetMode === "External") continue;
      const target = zipPath("ppt/slides", relation.Target);
      const imageFile = zip.file(target);
      if (!imageFile) continue;
      const base64 = await imageFile.async("base64");
      images.push({
        path: target,
        mimeType: mime(target),
        dataUrl: `data:${mime(target)};base64,${base64}`
      });
    }
    let notes = "";
    const notesRelation = relationships.find(relation => String(relation.Type || "").endsWith("/notesSlide"));
    if (notesRelation?.Target) {
      try {
        const notesFile = zipPath("ppt/slides", notesRelation.Target);
        const notesXml = await zipText(zip, notesFile);
        if (notesXml) {
          notes = extractTexts(parser.parse(notesXml))
            .map(text => text.replace(/\s+/g, " ").trim())
            .filter(Boolean)
            .join(" ");
        }
      } catch {
        notes = "";
      }
    }
    slides.push({
      number,
      title: paragraphs[0] || `Slide ${number}`,
      text: paragraphs.join(" "),
      notes,
      images,
      thumbnail: images[0]?.dataUrl || null,
      sourceFile: slideFile,
      validationStatus: paragraphs.length ? "AGUARDANDO_VALIDACAO" : "REVISAR"
    });
  }
  return {
    slideCount: slides.length,
    slides,
    extraction: {
      text: true,
      speakerNotes: true,
      images: true,
      renderedSlideThumbnails: false
    }
  };
}

function dateKeyInTimeZone(date = new Date()) {
  const parts = new Intl.DateTimeFormat("en-US", {
    timeZone: COST_TIME_ZONE, year: "numeric", month: "2-digit", day: "2-digit"
  }).formatToParts(date);
  const values = Object.fromEntries(parts.map(part => [part.type, part.value]));
  return [values.year, values.month, values.day].join("-");
}
async function readCostEntries() {
  try {
    const parsed = JSON.parse(await fs.readFile(COSTS_FILE, "utf8"));
    return Array.isArray(parsed) ? parsed : [];
  } catch { return []; }
}
let costWriteQueue = Promise.resolve();
function persistCostEntries(entries) {
  costWriteQueue = costWriteQueue.then(() =>
    fs.writeFile(COSTS_FILE, JSON.stringify(entries.slice(-5000), null, 2), "utf8")
  );
  return costWriteQueue;
}
function costFor(provider, durationSeconds, imageCount) {
  const duration = Math.max(0, Math.min(3600, Number(durationSeconds) || 0));
  const images = Math.max(0, Math.min(100, Math.floor(Number(imageCount) || 0)));
  const imageCostUSD = provider === "openai-image" || provider === "heygen-ghibli"
    ? images * COST_RATES["openai-image"] : 0;
  const videoCostUSD = provider === "heygen" || provider === "heygen-ghibli"
    ? duration * COST_RATES.heygen : 0;
  return {
    durationSeconds: duration, imageCount: images,
    imageCostUSD: Number(imageCostUSD.toFixed(4)),
    videoCostUSD: Number(videoCostUSD.toFixed(4)),
    costUSD: Number((imageCostUSD + videoCostUSD).toFixed(4))
  };
}
function summarizeCosts(entries) {
  const todayKey = dateKeyInTimeZone();
  const todayEntries = entries.filter(entry => dateKeyInTimeZone(new Date(entry.timestamp)) === todayKey);
  const todayCostUSD = Number(todayEntries.reduce((sum, entry) => sum + entry.costUSD, 0).toFixed(4));
  const totalCostUSD = Number(entries.reduce((sum, entry) => sum + entry.costUSD, 0).toFixed(4));
  const byDayMap = new Map();
  for (const entry of entries) {
    const day = dateKeyInTimeZone(new Date(entry.timestamp));
    byDayMap.set(day, (byDayMap.get(day) || 0) + entry.costUSD);
  }
  const byDay = [...byDayMap.entries()].sort(([a], [b]) => a.localeCompare(b)).slice(-7)
    .map(([date, costUSD]) => ({ date, costUSD: Number(costUSD.toFixed(4)) }));
  return {
    totalCostUSD, todayCostUSD, dailyLimitUSD: DAILY_COST_LIMIT_USD,
    dailyLimitExceeded: todayCostUSD > DAILY_COST_LIMIT_USD,
    timezone: COST_TIME_ZONE, byDay
  };
}

function safeEqual(left, right) {
  const a = Buffer.from(String(left));
  const b = Buffer.from(String(right));
  return a.length === b.length && timingSafeEqual(a, b);
}
function requireAdmin(req, res, next) {
  if (!ADMIN_USER || !ADMIN_PASSWORD) {
    return res.status(503).json({ error: "Acesso administrativo não configurado. Defina ADMIN_USER e ADMIN_PASSWORD no servidor." });
  }
  const header = req.get("authorization") || "";
  const match = header.match(/^Basic\\s+(.+)$/i);
  if (!match) {
    res.set("WWW-Authenticate", 'Basic realm="Treinamenos - Administração", charset="UTF-8"');
    return res.status(401).json({ error: "Autenticação necessária." });
  }
  let decoded = "";
  try { decoded = Buffer.from(match[1], "base64").toString("utf8"); } catch {}
  const splitAt = decoded.indexOf(":");
  const username = splitAt >= 0 ? decoded.slice(0, splitAt) : "";
  const password = splitAt >= 0 ? decoded.slice(splitAt + 1) : "";
  if (!safeEqual(username, ADMIN_USER) || !safeEqual(password, ADMIN_PASSWORD)) {
    res.set("WWW-Authenticate", 'Basic realm="Treinamenos - Administração", charset="UTF-8"');
    return res.status(401).json({ error: "Credenciais inválidas." });
  }
  next();
}

app.get("/custos", requireAdmin, (req, res) => res.sendFile(path.join(PUBLIC_DIR, "custos.html")));
app.get("/api/costs", requireAdmin, async (req, res, next) => {
  try {
    const entries = await readCostEntries();
    res.json({ entries: entries.slice(-500).reverse(), summary: summarizeCosts(entries) });
  } catch (error) { next(error); }
});
app.post("/api/costs", requireAdmin, async (req, res, next) => {
  try {
    const provider = String(req.body?.provider || "");
    if (!Object.hasOwn(COST_RATES, provider)) {
      return res.status(400).json({ error: "Provedor inválido. Use heygen, openai-image ou heygen-ghibli." });
    }
    const prompt = String(req.body?.prompt || "").slice(0, 5000);
    const requestId = String(req.body?.requestId || "").slice(0, 120);
    const entries = await readCostEntries();
    if (requestId) {
      const duplicate = entries.find(entry => entry.requestId === requestId);
      if (duplicate) return res.status(200).json({ entry: duplicate, summary: summarizeCosts(entries), duplicate: true });
    }
    const estimate = costFor(provider, req.body?.durationSeconds, provider === "heygen" ? 0 : (req.body?.imageCount ?? 1));
    if (estimate.costUSD <= 0) {
      return res.status(400).json({ error: "Informe duração maior que zero para vídeo ou quantidade de imagens maior que zero." });
    }
    const entry = {
      id: randomUUID(), timestamp: new Date().toISOString(), provider,
      label: String(req.body?.label || "").slice(0, 160),
      prompt, requestId: requestId || null,
      videoId: String(req.body?.videoId || "").slice(0, 200) || null,
      status: String(req.body?.status || "estimated").slice(0, 40),
      ...estimate
    };
    entries.push(entry);
    await persistCostEntries(entries);
    res.status(201).json({ entry, summary: summarizeCosts(entries), duplicate: false });
  } catch (error) { next(error); }
});

app.get("/api/health", (req, res) => {
  res.json({ ok: true, service: "studio-ia-treinamentos" });
});
app.post(
  "/api/projects",
  upload.fields([
    { name: "pptx", maxCount: 1 },
    { name: "logo", maxCount: 1 }
  ]),
  async (req, res, next) => {
    try {
      const pptx = req.files?.pptx?.[0];
      const logo = req.files?.logo?.[0];
      if (!pptx) {
        return res.status(400).json({ error: "Envie um arquivo PPTX." });
      }
      const id = randomUUID();
      const parsed = await parsePptx(pptx.buffer);
      const pptxName = `${id}.pptx`;
      await fs.writeFile(path.join(STORAGE_DIR, pptxName), pptx.buffer);
      let logoUrl = null;
      if (logo) {
        const ext = path.extname(logo.originalname).toLowerCase() || ".png";
        const logoName = `${id}-logo${ext}`;
        await fs.writeFile(path.join(STORAGE_DIR, logoName), logo.buffer);
        logoUrl = `/storage/${logoName}`;
      }
      const project = {
        id,
        name: req.body.name || "Novo treinamento",
        unit: req.body.unit || "",
        category: req.body.category || "",
        audience: req.body.audience || "",
        avatar: req.body.avatar || "",
        voice: req.body.voice || "",
        aspectRatio: req.body.aspectRatio || "16:9",
        source: {
          filename: pptx.originalname,
          pptxUrl: `/storage/${pptxName}`,
        },
        ...parsed,
        status: "EM_ANALISE",
        createdAt: new Date().toISOString()
      };
      await fs.writeFile(
        path.join(STORAGE_DIR, `${id}.json`),
        JSON.stringify(project, null, 2)
      );
      res.status(201).json(project);
    } catch (error) {
      next(error);
    }
  }
);
app.get("/api/projects/:id", async (req, res, next) => {
  try {
    const file = path.join(STORAGE_DIR, `${req.params.id}.json`);
    res.json(JSON.parse(await fs.readFile(file, "utf8")));
  } catch (error) {
    error.status = 404;
    next(error);
  }
});
app.put("/api/projects/:id", async (req, res, next) => {
  try {
    const file = path.join(STORAGE_DIR, `${req.params.id}.json`);
    const project = JSON.parse(await fs.readFile(file, "utf8"));
    if (Array.isArray(req.body.slides)) project.slides = req.body.slides;
    project.updatedAt = new Date().toISOString();
    await fs.writeFile(file, JSON.stringify(project, null, 2));
    res.json(project);
  } catch (error) {
    next(error);
  }
});

const inFlightVideoRequests = new Map();
function estimateSpokenDuration(text) {
  const words = String(text || "").trim().split(/\\s+/).filter(Boolean).length;
  return Math.max(1, Math.min(3600, Math.ceil(words / 2.5)));
}
async function heygenRequest(url, options = {}) {
  if (!HEYGEN_API_KEY) {
    const error = new Error("HEYGEN_API_KEY não está configurada no servidor.");
    error.status = 503;
    throw error;
  }
  const response = await fetch(url, {
    ...options,
    headers: {
      "X-Api-Key": HEYGEN_API_KEY,
      "Accept": "application/json",
      ...(options.body ? { "Content-Type": "application/json" } : {}),
      ...(options.headers || {})
    },
    signal: AbortSignal.timeout(30000)
  });
  const payload = await response.json().catch(() => ({}));
  if (!response.ok) {
    const error = new Error(payload?.message || payload?.error?.message || "O provedor HeyGen recusou a solicitação.");
    error.status = response.status >= 500 ? 502 : 400;
    throw error;
  }
  return payload;
}
async function generateHeyGenVideo(req, res, next) {
  const requestId = String(req.get("idempotency-key") || req.body?.requestId || "").trim().slice(0, 120) || randomUUID();
  if (inFlightVideoRequests.has(requestId)) {
    try { return res.status(202).json(await inFlightVideoRequests.get(requestId)); }
    catch (error) { return next(error); }
  }
  const task = (async () => {
    const script = String(req.body?.script || req.body?.prompt || "").trim().slice(0, 10000);
    const avatarId = String(req.body?.avatar_id || DEFAULT_HEYGEN_AVATAR_ID).trim().slice(0, 200);
    const voiceId = String(req.body?.voice_id || DEFAULT_HEYGEN_VOICE_ID).trim().slice(0, 200);
    const title = String(req.body?.title || req.body?.label || "Treinamento Treinamenos").trim().slice(0, 200);
    const caption = req.body?.caption === true;
    const ratio = String(req.body?.aspect_ratio || "16:9");
    if (script.length < 2) {
      const error = new Error("Informe o texto que o avatar deverá narrar no campo script.");
      error.status = 400; throw error;
    }
    if (!avatarId || !voiceId) {
      const error = new Error("Configure HEYGEN_AVATAR_ID e HEYGEN_VOICE_ID no servidor ou envie avatar_id e voice_id.");
      error.status = 400; throw error;
    }
    if (!["16:9", "9:16", "1:1"].includes(ratio)) {
      const error = new Error("aspect_ratio deve ser 16:9, 9:16 ou 1:1.");
      error.status = 400; throw error;
    }
    const existing = await readCostEntries();
    const prior = existing.find(entry => entry.requestId === requestId && entry.videoId);
    if (prior) return { requestId, videoId: prior.videoId, status: prior.status || "submitted", duplicate: true };
    const durationSeconds = Math.max(1, Math.min(3600, Math.ceil(Number(req.body?.estimated_duration) || estimateSpokenDuration(script))));
    const dimension = ratio === "9:16" ? { width: 720, height: 1280 } : ratio === "1:1" ? { width: 1080, height: 1080 } : { width: 1280, height: 720 };
    const providerResult = await heygenRequest("https://api.heygen.com/v2/video/generate", {
      method: "POST",
      body: JSON.stringify({
        title,
        caption,
        dimension,
        video_inputs: [{
          character: { type: "avatar", avatar_id: avatarId, avatar_style: "normal" },
          voice: { type: "text", input_text: script, voice_id: voiceId }
        }]
      })
    });
    const videoId = providerResult?.data?.video_id || providerResult?.video_id;
    if (!videoId) {
      const error = new Error("HeyGen não retornou um identificador de vídeo.");
      error.status = 502; throw error;
    }
    const entry = {
      id: randomUUID(), timestamp: new Date().toISOString(), provider: "heygen",
      label: title, prompt: script, requestId, videoId, status: "submitted",
      estimated: true, ...costFor("heygen", durationSeconds, 0)
    };
    existing.push(entry);
    await persistCostEntries(existing);
    return { requestId, videoId, status: "submitted", estimatedCostUSD: entry.costUSD, estimatedDurationSeconds: durationSeconds, summary: summarizeCosts(existing) };
  })();
  inFlightVideoRequests.set(requestId, task);
  try {
    const result = await task;
    return res.status(result.duplicate ? 200 : 202).json(result);
  } catch (error) {
    return next(error);
  } finally {
    inFlightVideoRequests.delete(requestId);
  }
}
app.post("/api/heygen-ghibli", requireAdmin, generateHeyGenVideo);
app.post("/api/videos/generate", requireAdmin, generateHeyGenVideo);
app.get("/api/heygen-ghibli/:videoId", requireAdmin, async (req, res, next) => {
  try {
    const videoId = String(req.params.videoId || "").slice(0, 200);
    const entries = await readCostEntries();
    const entry = entries.find(item => item.videoId === videoId);
    if (!entry) return res.status(404).json({ error: "Vídeo não encontrado no histórico desta aplicação." });
    const result = await heygenRequest("https://api.heygen.com/v1/video_status.get?video_id=" + encodeURIComponent(videoId));
    const data = result?.data || result;
    const status = String(data?.status || entry.status || "unknown").toLowerCase();
    entry.status = status;
    if (status === "completed" || status === "success") {
      const actualDuration = Number(data?.duration);
      if (Number.isFinite(actualDuration) && actualDuration > 0) {
        Object.assign(entry, costFor("heygen", actualDuration, 0));
        entry.estimated = false;
      }
      entry.videoUrl = typeof data?.video_url === "string" ? data.video_url : null;
    }
    if (["failed", "error"].includes(status)) {
      entry.costUSD = 0;
      entry.videoCostUSD = 0;
      entry.failedAt = new Date().toISOString();
    }
    await persistCostEntries(entries);
    return res.json({
      requestId: entry.requestId, videoId, status: entry.status,
      videoUrl: entry.videoUrl || data?.video_url || null,
      error: data?.error || null,
      costUSD: entry.costUSD, estimated: Boolean(entry.estimated),
      summary: summarizeCosts(entries)
    });
  } catch (error) { next(error); }
});

app.use((error, req, res, next) => {
  console.error(error);
  res.status(error.status || 500).json({
    error: error.message || "Erro interno."
  });
});
app.listen(port, () => {
  console.log(`Studio IA de Treinamentos: http://localhost:${port}`);
});
