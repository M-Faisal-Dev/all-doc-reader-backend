import express, { type ErrorRequestHandler } from "express";
import multer from "multer";
import cors from "cors";
import helmet from "helmet";
import morgan from "morgan";
import { rateLimit } from "express-rate-limit";

const MAX_BYTES = 30 * 1024 * 1024;
export interface Options { gotenbergUrl: string; origins?: string[]; timeoutMs?: number; fetcher?: typeof fetch }
export function createApp(options: Options) {
  const upstream = new URL(options.gotenbergUrl);
  if (!["http:", "https:"].includes(upstream.protocol)) throw new Error("GOTENBERG_URL must use HTTP or HTTPS");
  const base = upstream.toString().replace(/\/$/, "");
  const fetcher = options.fetcher ?? fetch;
  const timeoutMs = options.timeoutMs ?? 90000;
  const app = express();
  app.disable("x-powered-by");
  app.use(morgan(process.env.NODE_ENV === "production" ? "combined" : "dev"));
  app.use(helmet());
  app.use(cors({ origin(origin, callback) { callback(null, !origin || (options.origins ?? []).includes(origin)); }, exposedHeaders: ["Content-Disposition"] }));
  app.get("/health", async (_request, response) => {
    try {
      const result = await fetcher(`${base}/health`, { signal: AbortSignal.timeout(5000) });
      response.status(result.ok ? 200 : 503).json({ status: result.ok ? "up" : "degraded", gotenberg: result.ok ? "up" : "down" });
    } catch { response.status(503).json({ status: "degraded", gotenberg: "down" }); }
  });
  const upload = multer({ storage: multer.memoryStorage(), limits: { fileSize: MAX_BYTES, files: 1, fields: 0, parts: 1 } });
  const limit = rateLimit({ windowMs: 60_000, limit: 20, standardHeaders: "draft-8", legacyHeaders: false, message: { error: "Too many conversions. Try again shortly." } });
  let active = 0;
  app.post(["/api/v1/convert/word", "/api/v1/convert/excel", "/api/v1/convert/powerpoint"], limit, (request, response, next) => {
    // Reserve capacity before buffering uploads.
    if (active >= 2) { response.status(503).json({ error: "Conversion server is busy. Try again shortly." }); return; }
    active++;
    let released = false;
    const release = () => { if (!released) { released = true; active--; } };
    response.once("finish", release); response.once("close", release);
    next();
  }, upload.single("files"), async (request, response) => {
    const excel = request.path.endsWith("/excel");
    const powerpoint = request.path.endsWith("/powerpoint");
    const label = powerpoint ? "PowerPoint" : excel ? "Excel" : "Word";
    const file = request.file;
    if (!file) { response.status(400).json({ error: `Upload one ${label} file in the files field.` }); return; }
    const extension = /\.(docx|doc|xlsx|xls|pptx|ppt)$/i.exec(file.originalname)?.[1].toLowerCase();
    const signature = file.buffer.subarray(0, 8);
    const allowed = powerpoint ? ["ppt", "pptx"] : excel ? ["xls", "xlsx"] : ["doc", "docx"];
    const valid = !!extension && allowed.includes(extension) && (extension.endsWith("x") ? signature.length === 8 && signature[0] === 0x50 && signature[1] === 0x4b && signature[2] === 3 && signature[3] === 4 : signature.equals(Buffer.from([0xd0,0xcf,0x11,0xe0,0xa1,0xb1,0x1a,0xe1])));
    if (!valid) { response.status(415).json({ error: powerpoint ? "Only valid PPT and PPTX files are supported." : excel ? "Only valid XLS and XLSX files are supported." : "Only valid DOC and DOCX files are supported." }); return; }
    const controller = new AbortController();
    const timer = setTimeout(() => controller.abort(), timeoutMs);
    const cancel = () => { if (!response.writableEnded) controller.abort(); };
    response.once("close", cancel);
    try {
      const form = new FormData();
      form.append("files", new Blob([new Uint8Array(file.buffer)], { type: extension === "pptx" ? "application/vnd.openxmlformats-officedocument.presentationml.presentation" : extension === "ppt" ? "application/vnd.ms-powerpoint" : extension === "xlsx" ? "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet" : extension === "xls" ? "application/vnd.ms-excel" : extension === "docx" ? "application/vnd.openxmlformats-officedocument.wordprocessingml.document" : "application/msword" }), `document.${extension}`);
      const result = await fetcher(`${base}/forms/libreoffice/convert`, { method: "POST", body: form, signal: controller.signal, headers: { Accept: "application/pdf" } });
      if (!result.ok) { response.status(result.status === 400 ? 422 : 502).json({ error: result.status === 400 ? `This ${label} file could not be converted. It may be damaged or password protected.` : "Gotenberg could not complete the conversion. Please retry." }); return; }
      if (Number(result.headers.get("content-length")) > MAX_BYTES) throw new Error("PDF too large");
      const chunks: Uint8Array[] = []; let size = 0;
      if (result.body) for await (const part of result.body) { size += part.length; if (size > MAX_BYTES) { controller.abort(); throw new Error("PDF too large"); } chunks.push(part); }
      const pdf = Buffer.concat(chunks);
      if (pdf.length < 8 || pdf.subarray(0, 5).toString() !== "%PDF-") throw new Error("Invalid PDF response");
      response.set({ "Content-Type": "application/pdf", "Content-Disposition": 'attachment; filename="preview.pdf"', "Cache-Control": "no-store" }).send(pdf);
    } catch {
      if (!response.destroyed) response.status(controller.signal.aborted ? 504 : 502).json({ error: controller.signal.aborted ? "Conversion timed out. Please retry." : "Cannot convert this file right now. Check Gotenberg and try again." });
    } finally { clearTimeout(timer); response.off("close", cancel); }
  });
  const errors: ErrorRequestHandler = (error, _request, response, _next) => {
    if (error instanceof multer.MulterError) { response.status(error.code === "LIMIT_FILE_SIZE" ? 413 : 400).json({ error: error.code === "LIMIT_FILE_SIZE" ? "Maximum document file size is 30 MB." : "Upload exactly one file in the files field, with no extra form fields." }); return; }
    console.error("Request failed:", error instanceof Error ? error.name : "Unknown error");
    response.status(500).json({ error: "Unexpected server error." });
  };
  app.use(errors);
  return app;
}
