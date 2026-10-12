mport express from "express";
import cors from "cors";
import multer from "multer";
import JSZip from "jszip";
import { XMLParser } from "fast-xml-parser";
import { promises as fs } from "node:fs";
import path from "node:path";
import { fileURLToPath } from "node:url";
import { randomUUID } from "node:crypto";
const ROOT = path.dirname(fileURLToPath(import.meta.url));
const PUBLIC_DIR = path.join(ROOT, "public");
const STORAGE_DIR = path.join(ROOT, "storage");
await fs.mkdir(STORAGE_DIR, { recursive: true });
const app = express();
const port = process.env.PORT || 3000;
const parser = new XMLParser({
  ignoreAttributes: false,
  attributeNamePrefix: "",
  trimValues: false
});
app.use(cors());
app.use(express.json({ limit: "25mb" }));
app.use(express.static(PUBLIC_DIR));
app.use("/storage", express.static(STORAGE_DIR));
const upload = multer({
  storage: multer.memoryStorage(),
  limits: { fileSize: 100 * 1024 * 1024 }
});
const array = value =&gt; value == null ? [] : Array.isArray(value) ? value : [value];
function extractTexts(value, result = []) {
  if (value == null) return result;
  if (Array.isArray(value)) {
    value.forEach(item =&gt; extractTexts(item, result));
    return result;
  }
  if (typeof value !== "object") return result;
  for (const [key, child] of Object.entries(value)) {
    if (key === "a:t") {
      array(child).forEach(item =&gt; result.push(String(item)));
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
  const clean = decodeURIComponent(String(target || ""))
    .replaceAll("\\", "/")
    .replace(/^\//, "");
  const result = clean.startsWith("ppt/")
    ? clean
    : path.posix.normalize(path.posix.join(base, clean));
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
    .filter(name =&gt; /^ppt\/slides\/slide\d+\.xml$/.test(name))
    .sort((a, b) =&gt; {
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
      .map(text =&gt; text.replace(/\s+/g, " ").trim())
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
    const notesRelation = relationships.find(relation =&gt;
      String(relation.Type || "").endsWith("/notesSlide")
    );
    if (notesRelation?.Target) {
      try {
        const notesFile = zipPath("ppt/slides", notesRelation.Target);
        const notesXml = await zipText(zip, notesFile);
        if (notesXml) {
          notes = extractTexts(parser.parse(notesXml))
            .map(text =&gt; text.replace(/\s+/g, " ").trim())
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
app.get("/api/health", (req, res) =&gt; {
  res.json({ ok: true, service: "studio-ia-treinamentos" });
});
app.post(
  "/api/projects",
  upload.fields([
    { name: "pptx", maxCount: 1 },
    { name: "logo", maxCount: 1 }
  ]),
  async (req, res, next) =&gt; {
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
app.get("/api/projects/:id", async (req, res, next) =&gt; {
  try {
    const file = path.join(STORAGE_DIR, `${req.params.id}.json`);
    res.json(JSON.parse(await fs.readFile(file, "utf8")));
  } catch (error) {
    error.status = 404;
    next(error);
  }
});
app.put("/api/projects/:id", async (req, res, next) =&gt; {
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
app.use((error, req, res, next) =&gt; {
  console.error(error);
  res.status(error.status || 500).json({
    error: error.message || "Erro interno."
  });
});
app.listen(port, () =&gt; {
  console.log(`Studio IA de Treinamentos: http://localhost:${port}`);
});
3. Arquivo public/index.html
Na criação de novo arquivo do GitHub, digite o caminho completo:


public/index.html
Cole:


&lt;!doctype html&gt;
&lt;html lang="pt-BR"&gt;
&lt;head&gt;
  &lt;meta charset="utf-8"&gt;
  &lt;meta name="viewport" content="width=device-width, initial-scale=1"&gt;
  &lt;title&gt;Studio IA de Treinamentos&lt;/title&gt;
  &lt;style&gt;
  &lt;/style&gt;
&lt;/head&gt;
&lt;body&gt;
  &lt;header class="top"&gt;
    &lt;h1&gt;Studio IA de Treinamentos&lt;/h1&gt;
    &lt;p&gt;Processamento real de arquivos PPTX&lt;/p&gt;
  &lt;/header&gt;
  &lt;main class="wrap"&gt;
    &lt;section class="card"&gt;
      &lt;h2&gt;Novo treinamento&lt;/h2&gt;
      &lt;form id="form"&gt;
        &lt;div class="grid"&gt;
          &lt;div class="field"&gt;&lt;label&gt;Nome&lt;/label&gt;&lt;input name="name" value="Reciclagem Operação Segura" required&gt;&lt;/div&gt;
          &lt;div class="field"&gt;&lt;label&gt;Unidade&lt;/label&gt;&lt;input name="unit" value="CD-SP"&gt;&lt;/div&gt;
          &lt;div class="field"&gt;&lt;label&gt;Categoria&lt;/label&gt;&lt;input name="category" value="Segurança e Saúde no Trabalho"&gt;&lt;/div&gt;
          &lt;div class="field"&gt;&lt;label&gt;Público-alvo&lt;/label&gt;&lt;input name="audience" value="Operadores de equipamentos"&gt;&lt;/div&gt;
          &lt;div class="field"&gt;&lt;label&gt;Avatar&lt;/label&gt;&lt;input name="avatar" placeholder="Instrutor corporativo"&gt;&lt;/div&gt;
          &lt;div class="field"&gt;&lt;label&gt;Voz&lt;/label&gt;&lt;input name="voice" placeholder="Português BR didática"&gt;&lt;/div&gt;
          &lt;div class="field"&gt;&lt;label&gt;Arquivo PPTX&lt;/label&gt;&lt;input name="pptx" type="file" accept=".pptx" required&gt;&lt;/div&gt;
          &lt;div class="field"&gt;&lt;label&gt;Logotipo&lt;/label&gt;&lt;input name="logo" type="file" accept="image/*"&gt;&lt;/div&gt;
          &lt;div class="field"&gt;&lt;label&gt;Proporção&lt;/label&gt;&lt;input name="aspectRatio" value="16:9"&gt;&lt;/div&gt;
        &lt;/div&gt;
        &lt;button class="btn"&gt;Processar PPTX&lt;/button&gt;
      &lt;/form&gt;
      &lt;div id="status" class="status"&gt;&lt;/div&gt;
    &lt;/section&gt;
    &lt;section id="result" class="card" hidden&gt;
      &lt;h2 id="title"&gt;&lt;/h2&gt;
      &lt;p id="summary"&gt;&lt;/p&gt;
      &lt;div id="slides" class="slides"&gt;&lt;/div&gt;
      &lt;button id="save" class="btn"&gt;Salvar alterações&lt;/button&gt;
    &lt;/section&gt;
  &lt;/main&gt;
  &lt;script&gt;
    const esc = value =&gt; String(value ?? '').replace(/[&&lt;&gt;"']/g, c =&gt; ({
      '&':'&amp;','&lt;':'&lt;','&gt;':'&gt;','"':'&quot;',"'":'&#039;'
      slidesBox.innerHTML = project.slides.map((slide, index) =&gt; `
        &lt;article class="slide" draggable="true" data-index="${index}"&gt;
          &lt;div&gt;
            &lt;div class="thumb"&gt;${slide.thumbnail ? `&lt;img src="${slide.thumbnail}" alt="Slide ${slide.number}"&gt;` : 'Sem imagem'}&lt;/div&gt;
            &lt;span class="tag"&gt;Slide ${slide.number}&lt;/span&gt;
          &lt;/div&gt;
          &lt;div&gt;
            &lt;strong&gt;${esc(slide.title)}&lt;/strong&gt;
            &lt;textarea class="text"&gt;${esc(slide.text)}&lt;/textarea&gt;
            &lt;textarea class="notes" placeholder="Notas do apresentador"&gt;${esc(slide.notes)}&lt;/textarea&gt;
          &lt;/div&gt;
        &lt;/article&gt;
      document.querySelectorAll('.slide').forEach(card =&gt; {
    form.addEventListener('submit', async event =&gt; {
    document.querySelector('#save').addEventListener('click', async () =&gt; {
  &lt;/script&gt;
&lt;/body&gt;
&lt;/html&gt;
