"use strict";

/**
 * SR Micro-Agent — local file operations for Speaker Ready desks
 *
 * Runs on each SR laptop. Only two jobs:
 *   POST /open   — download PPTX from PHV2, save to Working/, open in PowerPoint
 *   POST /publish — upload edited file from Working/ back to PHV2
 *   GET  /health  — status check
 *   GET  /config  — return current config for the browser UI
 *
 * All UI lives on PHV2 (speakerready.html). This agent is invisible.
 */

const path = require("path");
const fs = require("fs");
const http = require("http");
const https = require("https");
const { spawn } = require("child_process");
const os = require("os");

const PORT = parseInt(process.env.PORT || "8899", 10);

// ── Config ───────────────────────────────────────────────────────────────────

const ROOT = process.env.SR_ROOT || path.resolve(__dirname, "..");
const CFG_PATH = path.join(ROOT, "sr-config.json");

let _cfg = {};
try { if (fs.existsSync(CFG_PATH)) _cfg = JSON.parse(fs.readFileSync(CFG_PATH, "utf8")); } catch (_) {}

const PH_SERVER = process.env.PH_SERVER || _cfg.phServer || "http://10.0.0.166:8088";
const AGENT_TOKEN = process.env.AGENT_TOKEN || _cfg.agentToken || "";
const WORKING = path.join(ROOT, "Working");
const CACHE = path.join(ROOT, "cache");

function ensureDir(p) { if (!fs.existsSync(p)) fs.mkdirSync(p, { recursive: true }); }
ensureDir(WORKING);
ensureDir(CACHE);

function log(...args) { console.log(`[sr-agent]`, ...args); }

// ── Session folder resolution (same logic as old SR sidecar) ─────────────────

function isShortCodeFolder(sf) {
  return /^S\d+$/i.test(String(sf || "").trim()) || String(sf || "").trim().length < 12;
}

async function fetchJson(url) {
  return new Promise((resolve, reject) => {
    const mod = url.startsWith("https") ? https : http;
    mod.get(url, { headers: { "Cache-Control": "no-cache" } }, (res) => {
      let data = "";
      res.on("data", (c) => data += c);
      res.on("end", () => {
        try { resolve(JSON.parse(data)); } catch (e) { reject(e); }
      });
    }).on("error", reject);
  });
}

async function resolveSessionFolder({ id, sessionFolder }) {
  if (sessionFolder && !isShortCodeFolder(sessionFolder)) return String(sessionFolder).trim();

  // Try cache
  const cacheFile = path.join(CACHE, "agenda.json");
  try {
    if (fs.existsSync(cacheFile)) {
      const arr = JSON.parse(fs.readFileSync(cacheFile, "utf8"));
      const hit = Array.isArray(arr) ? arr.find(x => String(x.id) === String(id)) : null;
      const sf = hit && (hit.sessionFolder || hit.session_folder);
      if (sf && !isShortCodeFolder(sf)) return String(sf);
    }
  } catch (_) {}

  // Try live agenda
  try {
    const arr = await fetchJson(`${PH_SERVER}/api/agenda`);
    // Cache it
    try { fs.writeFileSync(cacheFile, JSON.stringify(arr)); } catch (_) {}
    const hit = Array.isArray(arr) ? arr.find(x => String(x.id) === String(id)) : null;
    const sf = hit && (hit.sessionFolder || hit.session_folder);
    if (sf && !isShortCodeFolder(sf)) return String(sf);
  } catch (_) {}

  return String(sessionFolder || "").trim();
}

// ── Download file from PHV2 ──────────────────────────────────────────────────

function downloadFile(url, dest) {
  return new Promise((resolve, reject) => {
    ensureDir(path.dirname(dest));
    const file = fs.createWriteStream(dest);
    const mod = url.startsWith("https") ? https : http;
    mod.get(url, (res) => {
      if (res.statusCode === 302 || res.statusCode === 301) {
        // Follow redirect
        file.close();
        fs.unlinkSync(dest);
        return downloadFile(res.headers.location, dest).then(resolve).catch(reject);
      }
      if (res.statusCode !== 200) {
        file.close();
        try { fs.unlinkSync(dest); } catch (_) {}
        return reject(new Error(`Download failed: HTTP ${res.statusCode}`));
      }
      res.pipe(file);
      file.on("finish", () => { file.close(); resolve(); });
    }).on("error", (e) => {
      file.close();
      try { fs.unlinkSync(dest); } catch (_) {}
      reject(e);
    });
  });
}

// ── Upload file to PHV2 ─────────────────────────────────────────────────────

function uploadFile(filePath, sessionId, room) {
  return new Promise((resolve, reject) => {
    const fileName = path.basename(filePath);
    const stat = fs.statSync(filePath);
    const boundary = "----SRAgent" + Date.now().toString(36);

    // Build multipart body
    const fields = [
      ["id", sessionId],
      ["room", room],
    ];

    let header = "";
    for (const [key, val] of fields) {
      header += `--${boundary}\r\nContent-Disposition: form-data; name="${key}"\r\n\r\n${val}\r\n`;
    }
    header += `--${boundary}\r\nContent-Disposition: form-data; name="file"; filename="${fileName}"\r\nContent-Type: application/octet-stream\r\n\r\n`;
    const footer = `\r\n--${boundary}--\r\n`;

    const headerBuf = Buffer.from(header, "utf-8");
    const footerBuf = Buffer.from(footer, "utf-8");
    const contentLength = headerBuf.length + stat.size + footerBuf.length;

    const url = new URL(`${PH_SERVER}/api/upload`);
    const mod = url.protocol === "https:" ? https : http;

    const req = mod.request({
      hostname: url.hostname,
      port: url.port,
      path: url.pathname,
      method: "POST",
      headers: {
        "Content-Type": `multipart/form-data; boundary=${boundary}`,
        "Content-Length": contentLength,
      },
    }, (res) => {
      let data = "";
      res.on("data", (c) => data += c);
      res.on("end", () => {
        try {
          const json = JSON.parse(data);
          if (json.ok) resolve(json);
          else reject(new Error(json.error || "Upload failed"));
        } catch (_) {
          if (res.statusCode >= 400) reject(new Error(`HTTP ${res.statusCode}: ${data}`));
          else resolve({ ok: true, raw: data });
        }
      });
    });

    req.on("error", reject);
    req.write(headerBuf);
    const stream = fs.createReadStream(filePath);
    stream.on("end", () => { req.end(footerBuf); });
    stream.pipe(req, { end: false });
  });
}

// ── HTTP Server ──────────────────────────────────────────────────────────────

function parseBody(req) {
  return new Promise((resolve) => {
    let data = "";
    req.on("data", (c) => data += c);
    req.on("end", () => {
      try { resolve(JSON.parse(data)); } catch (_) { resolve({}); }
    });
  });
}

function cors(res) {
  res.setHeader("Access-Control-Allow-Origin", "*");
  res.setHeader("Access-Control-Allow-Methods", "GET, POST, OPTIONS");
  res.setHeader("Access-Control-Allow-Headers", "Content-Type");
}

function json(res, status, obj) {
  cors(res);
  res.writeHead(status, { "Content-Type": "application/json" });
  res.end(JSON.stringify(obj));
}

const server = http.createServer(async (req, res) => {
  cors(res);

  // CORS preflight
  if (req.method === "OPTIONS") { res.writeHead(204); return res.end(); }

  // GET /health
  if (req.method === "GET" && req.url === "/health") {
    return json(res, 200, { ok: true, agent: "sr-micro", port: PORT, server: PH_SERVER, working: WORKING });
  }

  // GET /config
  if (req.method === "GET" && req.url === "/config") {
    return json(res, 200, {
      ok: true,
      phServer: PH_SERVER,
      working: WORKING,
      activeRoom: _cfg.activeRoom || "",
      pcname: os.hostname(),
    });
  }

  // POST /open — download from PHV2 and open in PowerPoint
  if (req.method === "POST" && req.url === "/open") {
    try {
      const body = await parseBody(req);
      const { id, room, sessionFolder, fileName } = body;

      if (!id || !room) return json(res, 400, { ok: false, error: "Missing id or room" });

      const sfResolved = await resolveSessionFolder({ id, sessionFolder });
      let sfFinal = sfResolved || "";
      if (!sfFinal || isShortCodeFolder(sfFinal)) {
        // Derive from fileName
        sfFinal = fileName ? path.parse(fileName).name : "";
      }
      if (!sfFinal) return json(res, 400, { ok: false, error: "Cannot determine session folder" });

      const destDir = path.join(WORKING, room, sfFinal);
      ensureDir(destDir);

      const canonicalName = `${sfFinal}.pptx`;
      const dest = path.join(destDir, canonicalName);

      // Download from PHV2
      const dlUrl = `${PH_SERVER}/api/upload/download/${encodeURIComponent(id)}`;
      log(`Downloading ${dlUrl} -> ${dest}`);
      await downloadFile(dlUrl, dest);

      // Open in PowerPoint
      log(`Opening: ${dest}`);
      spawn("cmd.exe", ["/c", "start", "", dest], { detached: true, stdio: "ignore" }).unref();

      return json(res, 200, { ok: true, path: dest, sessionFolder: sfFinal, fileName: canonicalName });
    } catch (e) {
      log("OPEN ERROR:", e.message);
      return json(res, 500, { ok: false, error: e.message });
    }
  }

  // POST /publish — upload edited file from Working/ back to PHV2
  if (req.method === "POST" && req.url === "/publish") {
    try {
      const body = await parseBody(req);
      const { id, room, sessionFolder, fileName } = body;

      if (!id || !room) return json(res, 400, { ok: false, error: "Missing id or room" });

      const sfResolved = await resolveSessionFolder({ id, sessionFolder });
      let sfFinal = sfResolved || "";
      if (!sfFinal || isShortCodeFolder(sfFinal)) {
        sfFinal = fileName ? path.parse(fileName).name : "";
      }
      if (!sfFinal) return json(res, 400, { ok: false, error: "Cannot determine session folder" });

      const canonicalName = `${sfFinal}.pptx`;
      const localPath = path.join(WORKING, room, sfFinal, canonicalName);

      if (!fs.existsSync(localPath)) {
        return json(res, 404, { ok: false, error: "Working file not found", path: localPath });
      }

      log(`Publishing ${localPath} -> ${PH_SERVER}/api/upload`);
      const result = await uploadFile(localPath, id, room);

      return json(res, 200, { ok: true, result });
    } catch (e) {
      log("PUBLISH ERROR:", e.message);
      return json(res, 500, { ok: false, error: e.message });
    }
  }

  // POST /open-folder — open Working folder in Explorer
  if (req.method === "POST" && req.url === "/open-folder") {
    try {
      const body = await parseBody(req);
      const { room, sessionFolder } = body;
      let target = WORKING;
      if (room) target = path.join(target, room);
      if (sessionFolder) target = path.join(target, sessionFolder);
      ensureDir(target);
      spawn("explorer.exe", [target], { detached: true, stdio: "ignore" }).unref();
      return json(res, 200, { ok: true, path: target });
    } catch (e) {
      return json(res, 500, { ok: false, error: e.message });
    }
  }

  json(res, 404, { ok: false, error: "Not found" });
});

// ── Heartbeat to PHV2 ───────────────────────────────────────────────────────

let activeRoom = _cfg.activeRoom || "";

async function sendHeartbeat() {
  if (!activeRoom) return;
  try {
    const body = JSON.stringify({ room: activeRoom, type: "speakerready", pcname: os.hostname() });
    const url = new URL(`${PH_SERVER}/api/agents/heartbeat`);
    const mod = url.protocol === "https:" ? https : http;
    const req = mod.request({
      hostname: url.hostname, port: url.port, path: url.pathname,
      method: "POST",
      headers: {
        "Content-Type": "application/json",
        "Content-Length": Buffer.byteLength(body),
        ...(AGENT_TOKEN ? { "X-Agent-Token": AGENT_TOKEN } : {}),
      },
    }, () => {});
    req.on("error", () => {});
    req.end(body);
  } catch (_) {}
}

setInterval(sendHeartbeat, 10000);

// ── Startup ──────────────────────────────────────────────────────────────────

server.listen(PORT, () => {
  log(`Micro-agent running on http://localhost:${PORT}`);
  log(`PHV2 server: ${PH_SERVER}`);
  log(`Working folder: ${WORKING}`);
  log(`Active room: ${activeRoom || "(none — will track from requests)"}`);
  if (activeRoom) setTimeout(sendHeartbeat, 2000);
});
