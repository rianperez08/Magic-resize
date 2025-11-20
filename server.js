// server.js
// Local dev app:
// 1) OAuth to Canva (PKCE)
// 2) Paste a Canva design URL (16:9)
// 3) Resize to:
//      - 9:16 (1080x1920)
//      - 1:1  (1080x1080)
//    via /rest/v1/resizes
// 4) Export original + both resizes as PNG via /rest/v1/exports
// 5) NEW: Download PNGs, store locally in canva_received/, then upload to S3
//    bucket canva-bridge-storage in the form:
//    images/cv_<groupCode>/cv_<groupCode>_<imageCode>.png

require("dotenv").config();

const express = require("express");
const session = require("express-session");
const fetch = require("node-fetch");
const crypto = require("crypto");
const path = require("path");
const fs = require("fs");
const fsp = fs.promises;

// ---------- AWS S3 (NEW) ----------
const {
  S3Client,
  PutObjectCommand,
  ListObjectsV2Command,
  GetObjectCommand,
} = require("@aws-sdk/client-s3");

const AWS_REGION = process.env.AWS_REGION || "ap-southeast-1";
const S3_BUCKET =
  process.env.CANVA_S3_BUCKET || process.env.AWS_S3_BUCKET || "canva-bridge-storage";

// Local folder where PNGs will be stored first
const LOCAL_RECEIVE_DIR =
  process.env.LOCAL_RECEIVE_DIR || path.join(__dirname, "canva_received");

// Ensure local receive dir exists
(async () => {
  try {
    await fsp.mkdir(LOCAL_RECEIVE_DIR, { recursive: true });
    console.log("[INIT] Local receive dir:", LOCAL_RECEIVE_DIR);
  } catch (err) {
    console.error("[INIT] Failed to create local receive dir:", err);
  }
})();

// S3 client (uses IAM access key from env)
const s3 = new S3Client({
  region: AWS_REGION,
  credentials: {
    accessKeyId: process.env.AWS_ACCESS_KEY_ID || "AKIA6MZCBU7RK5EGDECW",
    secretAccessKey: process.env.AWS_SECRET_ACCESS_KEY || "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",
  },
});

const app = express();

// ---------- Middleware ----------
app.use(express.urlencoded({ extended: true }));
app.use(express.json());

// Use express-session to store OAuth tokens + last result
app.use(
  session({
    name: "canva-session",
    // session secret only; NOT your Canva client secret
    secret: process.env.CANVA_CLIENT_SECRET_CNVCA || "cnvca-cnvcazDN127UZFWjhHoXAMs6j6tctw9ysiUIFpgMxRBygnn45ba19c9c",
    resave: false,
    saveUninitialized: false,
    cookie: {
      maxAge: 24 * 60 * 60 * 1000, // 1 day
    },
  })
);

// ---------- Canva constants ----------
const CANVA_AUTH_BASE = "https://www.canva.com/api/oauth/authorize";
const CANVA_TOKEN_URL = "https://api.canva.com/rest/v1/oauth/token";
const CANVA_EXPORTS_URL = "https://api.canva.com/rest/v1/exports";
const CANVA_RESIZES_URL = "https://api.canva.com/rest/v1/resizes";

// From .env
const CLIENT_ID = "OC-AZqP9sNKUNOp";
const CLIENT_SECRET = "cnvcaU5jePVWcyy7-1xescatSB4EwLsDfskwhmQwDBODD0b8a77baa8e";
const REDIRECT_URI = "http://127.0.0.1:3001/oauth/redirect";
const SCOPES =
  process.env.CANVA_SCOPES || "design:content:read design:content:write";

// ---------- Utility helpers ----------

// Random base64url string
function generateRandomBase64Url() {
  return crypto.randomBytes(32).toString("base64url");
}

// PKCE pair
function generatePkcePair() {
  const codeVerifier = generateRandomBase64Url();
  const codeChallenge = crypto
    .createHash("sha256")
    .update(codeVerifier)
    .digest("base64url");
  return { codeVerifier, codeChallenge };
}

// Basic auth for token endpoint
function basicAuthHeader(clientId, clientSecret) {
  const creds = Buffer.from(`${clientId}:${clientSecret}`).toString("base64");
  return `Basic ${creds}`;
}

// Sleep helper (for polling)
function sleep(ms) {
  return new Promise((resolve) => setTimeout(resolve, ms));
}

// Extract Canva design ID from URL
function extractDesignIdFromUrl(urlString) {
  try {
    const url = new URL(urlString);
    const parts = url.pathname.split("/").filter(Boolean); // remove empty

    // Typical: /design/DAGirp_1ZUA/some-slug/edit
    const designIndex = parts.indexOf("design");
    if (designIndex !== -1 && parts.length > designIndex + 1) {
      return parts[designIndex + 1];
    }
    return null;
  } catch (e) {
    return null;
  }
}

// ---------- NEW HELPERS: grouping, download, local save + S3 upload ----------

// Generate a unique group code for this export batch.
// Result looks like: cv_<shortDesign>_<TIMESTAMP_BASE36>
function generateGroupCode(designId) {
  const safeDesign = (designId || "design")
    .replace(/[^A-Za-z0-9]/g, "")
    .slice(0, 6)
    .toUpperCase();
  const ts = Date.now().toString(36).toUpperCase();
  // This is the "cv_uniqueGroupCode(generated from canva app receiver)"
  return `cv_${safeDesign}_${ts}`;
}

// Ensure a directory exists
async function ensureDir(dirPath) {
  await fsp.mkdir(dirPath, { recursive: true });
}

// Download a Canva PNG URL into a Buffer
async function downloadPngToBuffer(url) {
  console.log("[DOWNLOAD] Fetching PNG from Canva:", url);
  const res = await fetch(url);
  if (!res.ok) {
    const text = await res.text().catch(() => "");
    throw new Error(
      `Failed to download PNG (${res.status}): ${text.slice(0, 300)}`
    );
  }
  const ab = await res.arrayBuffer();
  return Buffer.from(ab);
}

// Core helper: download image, save locally, then upload to S3
// groupCode: e.g. "cv_DESIGN1_ABC123"
// imageCode: e.g. "orig_p1", "v916_p1", "sq_p1"
async function storeLocallyAndUploadToS3(url, groupCode, imageCode) {
  const buffer = await downloadPngToBuffer(url);

  // Local path: <LOCAL_RECEIVE_DIR>/<groupCode>/<groupCode>_<imageCode>.png
  const localDir = path.join(LOCAL_RECEIVE_DIR, groupCode);
  await ensureDir(localDir);

  const filename = `${groupCode}_${imageCode}.png`;
  const localPath = path.join(localDir, filename);

  await fsp.writeFile(localPath, buffer);
  console.log("[LOCAL] Saved PNG:", localPath);

  // S3 key: images/<groupCode>/<groupCode>_<imageCode>.png
  const key = `images/${groupCode}/${filename}`;

  const putCmd = new PutObjectCommand({
    Bucket: S3_BUCKET,
    Key: key,
    Body: buffer,
    ContentType: "image/png",
  });

  await s3.send(putCmd);
  console.log("[S3] Uploaded PNG to:", `${S3_BUCKET}/${key}`);

  return { localPath, s3Key: key };
}

// ---------- Canva API helpers (Resize + Export) ----------

async function createResizeJob(accessToken, designId, { width, height }) {
  const body = {
    design_id: designId,
    design_type: {
      type: "custom",
      width,
      height,
    },
  };

  const res = await fetch(CANVA_RESIZES_URL, {
    method: "POST",
    headers: {
      Authorization: `Bearer ${accessToken}`,
      "Content-Type": "application/json",
      Accept: "application/json",
    },
    body: JSON.stringify(body),
  });

  const text = await res.text();
  console.log("[RESIZE:create] status:", res.status, "body:", text);

  if (!res.ok) {
    throw new Error(
      `Failed to create resize job (${res.status}): ${text.slice(0, 500)}`
    );
  }

  return JSON.parse(text); // { job: { id, ... } }
}

async function pollResizeJob(accessToken, jobId) {
  const url = `${CANVA_RESIZES_URL}/${encodeURIComponent(jobId)}`;

  const maxAttempts = 20;
  for (let attempt = 0; attempt < maxAttempts; attempt++) {
    const res = await fetch(url, {
      headers: {
        Authorization: `Bearer ${accessToken}`,
        Accept: "application/json",
      },
    });

    const text = await res.text();
    console.log(
      `[RESIZE:poll] attempt=${attempt + 1} status=${res.status} body=${text}`
    );

    if (!res.ok) {
      throw new Error(
        `Failed to get resize job (${res.status}): ${text.slice(0, 500)}`
      );
    }

    const json = JSON.parse(text);
    const status = json.job?.status;

    if (status === "success" || status === "failed") {
      return json;
    }

    await sleep(1500);
  }

  throw new Error("Resize job polling timed out.");
}

async function createExportJobPng(accessToken, designId) {
  if (!designId) {
    throw new Error("createExportJobPng called with empty designId");
  }

  const body = {
    design_id: designId,
    format: {
      type: "png",
      export_quality: "regular",
    },
  };

  const res = await fetch(CANVA_EXPORTS_URL, {
    method: "POST",
    headers: {
      Authorization: `Bearer ${accessToken}`,
      "Content-Type": "application/json",
      Accept: "application/json",
    },
    body: JSON.stringify(body),
  });

  const text = await res.text();
  console.log("[EXPORT:create] status:", res.status, "body:", text);

  if (!res.ok) {
    throw new Error(
      `Failed to create export job (${res.status}): ${text.slice(0, 500)}`
    );
  }

  return JSON.parse(text); // { job: { id, ... } }
}

async function pollExportJob(accessToken, exportId) {
  const url = `${CANVA_EXPORTS_URL}/${encodeURIComponent(exportId)}`;

  const maxAttempts = 20;
  for (let attempt = 0; attempt < maxAttempts; attempt++) {
    const res = await fetch(url, {
      headers: {
        Authorization: `Bearer ${accessToken}`,
        Accept: "application/json",
      },
    });

    const text = await res.text();
    console.log(
      `[EXPORT:poll] attempt=${attempt + 1} status=${res.status} body=${text}`
    );

    if (!res.ok) {
      throw new Error(
        `Failed to get export job (${res.status}): ${text.slice(0, 500)}`
      );
    }

    const json = JSON.parse(text);
    const status = json.job?.status;

    if (status === "success" || status === "failed") {
      return json;
    }

    await sleep(1500);
  }

  throw new Error("Export job polling timed out.");
}

// ---------- HTML view ----------
function renderHomeHtml({ isAuthed, error, result }) {
  // Helper: render list of PNG URLs as Page 1, Page 2, ...
  function renderPngList(label, urls) {
    if (!urls || !urls.length) {
      return `<li>${label}: <span style="color:#ffb74d;">(no URLs – export failed)</span></li>`;
    }

    const links = urls
      .map(
        (url, idx) =>
          `<li>Page ${idx + 1}: <a class="download-link" href="${url}" target="_blank">Download PNG</a></li>`
      )
      .join("");

    return `
      <li>
        ${label}:
        <ul style="margin:4px 0 8px 18px;padding-left:0;list-style:square;">
          ${links}
        </ul>
      </li>
    `;
  }

  return `
<!DOCTYPE html>
<html lang="en">
<head>
  <meta charset="utf-8" />
  <title>Canva 16:9 → 9:16 + 1:1 PNG Exporter (Local)</title>
  <style>
    body {
      font-family: system-ui, -apple-system, BlinkMacSystemFont, "Segoe UI", sans-serif;
      background: #0b0b0b;
      color: #f7f7f7;
      margin: 0;
      padding: 32px;
    }
    .card {
      max-width: 720px;
      margin: 0 auto;
      background: #151515;
      border-radius: 16px;
      padding: 24px 28px;
      box-shadow: 0 18px 40px rgba(0,0,0,0.6);
    }
    h1 {
      margin-top: 0;
      font-size: 24px;
    }
    label {
      display: block;
      margin-bottom: 4px;
      font-size: 14px;
      color: #cccccc;
    }
    input[type="text"] {
      width: 100%;
      padding: 10px 12px;
      border-radius: 10px;
      border: 1px solid #333;
      background: #111;
      color: #f7f7f7;
      font-size: 14px;
      box-sizing: border-box;
    }
    .btn {
      display: inline-flex;
      align-items: center;
      justify-content: center;
      padding: 10px 16px;
      border-radius: 999px;
      border: none;
      cursor: pointer;
      font-size: 14px;
      font-weight: 600;
      margin-top: 12px;
    }
    .btn-primary {
      background: #00c4cc;
      color: #051015;
    }
    .btn-secondary {
      background: #262626;
      color: #f7f7f7;
      margin-left: 8px;
    }
    .badge {
      display: inline-flex;
      align-items: center;
      gap: 6px;
      font-size: 12px;
      padding: 4px 10px;
      border-radius: 999px;
      background: #111;
      border: 1px solid #333;
      margin-bottom: 12px;
    }
    .status {
      margin-top: 12px;
      font-size: 13px;
      color: #ffb74d;
    }
    .error {
      margin-top: 12px;
      font-size: 13px;
      color: #ff5252;
      white-space: pre-wrap;
    }
    .result {
      margin-top: 16px;
      padding-top: 12px;
      border-top: 1px solid #333;
    }
    a.download-link {
      color: #80deea;
      text-decoration: none;
      font-size: 13px;
    }
    a.download-link:hover {
      text-decoration: underline;
    }
    small {
      color: #9e9e9e;
      font-size: 12px;
    }
    .row {
      display: flex;
      flex-wrap: wrap;
      gap: 12px;
      margin-top: 8px;
    }
    .row > div {
      flex: 1 1 260px;
    }
    code {
      font-size: 12px;
      background:#111;
      padding:2px 4px;
      border-radius:4px;
    }
  </style>
</head>
<body>
  <div class="card">
    <div class="badge">
      <span>⚙️ Local dev</span>
      <span>·</span>
      <span>Canva Connect – PNG Exporter</span>
    </div>
    <h1>16:9 → 9:16 + 1:1 PNG Exporter</h1>
    <p style="font-size:13px;color:#bdbdbd;margin-top:4px;">
      Paste a Canva <strong>presentation or design link</strong> you own. This app will:
      <br/>1) Export the original design as <strong>PNG (16:9)</strong>
      <br/>2) <strong>Magic Resize</strong> to <strong>9:16</strong> (1080×1920) and export PNG
      <br/>3) <strong>Magic Resize</strong> to <strong>square</strong> (1080×1080) and export PNG
      <br/>4) <strong>NEW:</strong> Save all PNGs locally under <code>canva_received/</code> and upload them to S3.
    </p>

    ${
      !isAuthed
        ? `
      <form method="GET" action="/auth/canva">
        <button class="btn btn-primary" type="submit">
          Connect Canva
        </button>
      </form>
      <p class="status">You need to connect your Canva account first.</p>
    `
        : `
      <p class="status">✅ Connected to Canva. Paste a design URL below.</p>
      <form method="POST" action="/export">
        <label for="designUrl">Canva design URL</label>
        <input
          type="text"
          id="designUrl"
          name="designUrl"
          placeholder="https://www.canva.com/design/XXXXXXXXX/..."
          required
        />
        <div class="row">
          <div>
            <small>Original: exported at design's native 16:9 size.</small>
          </div>
          <div>
            <small>Resizes:<br/>
            • 9:16 → 1080×1920<br/>
            • Square → 1080×1080</small>
          </div>
        </div>
        <button class="btn btn-primary" type="submit">
          Export + Save Locally + Upload to S3
        </button>
        <a href="/disconnect" class="btn btn-secondary" style="text-decoration:none;">
          Disconnect
        </a>
      </form>
    `
    }

    ${
      error
        ? `<div class="error"><strong>Error:</strong> ${error}</div>`
        : ""
    }

    ${
      result
        ? `
      <div class="result">
        <h3 style="font-size:16px;margin:0 0 6px;">Outputs</h3>
        <div style="font-size:13px;margin-bottom:8px;">
          Base design ID: <code>${result.designId}</code><br/>
          9:16 design ID: <code>${result.designId916}</code><br/>
          1:1 design ID:  <code>${result.designIdSquare}</code><br/>
          Group code (folder name): <code>${result.groupCode}</code><br/>
          Local base folder: <code>${result.localBaseDir}</code><br/>
          S3 bucket: <code>${result.s3Bucket}</code>
        </div>
        <ul style="padding-left:18px;font-size:13px;margin-top:6px;">
          ${renderPngList("Original PNG(s) – 16:9", result.originalPngUrls)}
          ${renderPngList("Resized PNG(s) – 9:16 (1080×1920)", result.png916Urls)}
          ${renderPngList("Resized PNG(s) – 1:1 (1080×1080)", result.pngSquareUrls)}
        </ul>
        <small>
          Files were stored under <code>canva_received/${result.groupCode}/</code> locally,
          and uploaded to <code>${result.s3Bucket}/images/${result.groupCode}/</code> in S3.
        </small>
      </div>
    `
        : ""
    }
  </div>
</body>
</html>
`;
}

// ---------- Routes ----------

// Home
app.get("/", (req, res) => {
  console.log("[HOME] session =", req.session);
  const isAuthed = Boolean(req.session && req.session.accessToken);
  res.send(
    renderHomeHtml({
      isAuthed,
      error: req.session.lastError || null,
      result: req.session.lastResult || null,
    })
  );
  // clear flash
  req.session.lastError = null;
  req.session.lastResult = null;
});

// Start OAuth with Canva
app.get("/auth/canva", (req, res) => {
  if (!CLIENT_ID || !CLIENT_SECRET || !REDIRECT_URI) {
    return res.send(
      "Missing Canva env vars. Please set CANVA_CLIENT_ID and CANVA_CLIENT_SECRET."
    );
  }

  const { codeVerifier, codeChallenge } = generatePkcePair();
  const state = generateRandomBase64Url();

  req.session.codeVerifier = codeVerifier;
  req.session.state = state;

  const params = new URLSearchParams({
    code_challenge: codeChallenge,
    code_challenge_method: "S256",
    scope: SCOPES,
    response_type: "code",
    client_id: CLIENT_ID,
    state,
    redirect_uri: REDIRECT_URI,
  });

  const authUrl = `${CANVA_AUTH_BASE}?${params.toString()}`;
  console.log("[AUTH] redirecting to:", authUrl);
  res.redirect(authUrl);
});

// OAuth redirect handler
app.get("/oauth/redirect", async (req, res) => {
  const { code, state, error } = req.query;
  console.log("[OAUTH] redirect query:", req.query);

  if (error) {
    console.log("[OAUTH] error from Canva:", error);
    req.session.lastError = `Authorization error: ${error}`;
    return res.redirect("/");
  }

  if (!code || !state) {
    req.session.lastError = "Missing code or state in OAuth redirect.";
    return res.redirect("/");
  }

  if (!req.session.state || state !== req.session.state) {
    console.log(
      "[OAUTH] state mismatch. expected:",
      req.session.state,
      "got:",
      state
    );
    req.session.lastError = "State mismatch. Please try connecting again.";
    return res.redirect("/");
  }

  const codeVerifier = req.session.codeVerifier;
  if (!codeVerifier) {
    console.log("[OAUTH] no codeVerifier in session");
    req.session.lastError =
      "Missing code_verifier in session. Start auth again.";
    return res.redirect("/");
  }

  try {
    console.log("[OAUTH] exchanging code for token…");

    const body = new URLSearchParams({
      grant_type: "authorization_code",
      code: code.toString(),
      code_verifier: codeVerifier,
      redirect_uri: REDIRECT_URI,
    });

    const tokenRes = await fetch(CANVA_TOKEN_URL, {
      method: "POST",
      headers: {
        Authorization: basicAuthHeader(CLIENT_ID, CLIENT_SECRET),
        "Content-Type": "application/x-www-form-urlencoded",
        Accept: "application/json",
      },
      body,
    });

    const text = await tokenRes.text();
    console.log("[OAUTH] token response status:", tokenRes.status);
    console.log("[OAUTH] token response body:", text);

    if (!tokenRes.ok) {
      throw new Error(`Token endpoint error (${tokenRes.status}): ${text}`);
    }

    const tokenJson = JSON.parse(text);
    req.session.accessToken = tokenJson.access_token;
    req.session.refreshToken = tokenJson.refresh_token;
    req.session.tokenScope = tokenJson.scope;
    req.session.tokenExpiresIn = tokenJson.expires_in;

    console.log("[OAUTH] accessToken stored in session");
    req.session.lastError = null;
  } catch (e) {
    console.error("[OAuth] Token exchange failed", e);
    req.session.lastError = `Token exchange failed: ${e.message}`;
  }

  // cleanup
  req.session.codeVerifier = null;
  req.session.state = null;

  res.redirect("/");
});

// Disconnect
app.get("/disconnect", (req, res) => {
  if (req.session) {
    req.session.destroy((err) => {
      if (err) {
        console.error("[DISCONNECT] error destroying session", err);
      }
      res.redirect("/");
    });
  } else {
    res.redirect("/");
  }
});

// Export route: original 16:9 + 9:16 + 1:1 PNGs
// NEW: For each PNG URL, we download, store locally, then upload to S3.
app.post("/export", async (req, res) => {
  const accessToken = req.session.accessToken;

  if (!accessToken) {
    req.session.lastError = "You must connect Canva first.";
    return res.redirect("/");
  }

  const designUrl = (req.body.designUrl || "").trim();
  const designId = extractDesignIdFromUrl(designUrl);

  if (!designId) {
    req.session.lastError =
      "Could not extract design ID from the URL. Make sure it's a full Canva design link.";
    return res.redirect("/");
  }

  console.log("[EXPORT] designUrl:", designUrl, "designId:", designId);

  try {
    // 1) Export original design as PNG
    console.log("[EXPORT] creating export job for original PNG…");
    const exportOriginalJob = await createExportJobPng(accessToken, designId);
    const exportOriginalResult = await pollExportJob(
      accessToken,
      exportOriginalJob.job.id
    );
    console.log(
      "[EXPORT] original PNG export status:",
      exportOriginalResult.job.status
    );

    const originalPngUrls =
      exportOriginalResult.job.status === "success"
        ? exportOriginalResult.job.urls || []
        : [];

    // 2) Resize to 9:16 (1080x1920) and export PNG
    console.log("[EXPORT] creating 9:16 resize job…");
    const resize916Job = await createResizeJob(accessToken, designId, {
      width: 1080,
      height: 1920,
    });

    console.log("[EXPORT] 9:16 resize job created:", resize916Job.job?.id);
    const resize916Result = await pollResizeJob(accessToken, resize916Job.job.id);
    console.log(
      "[EXPORT] 9:16 resize job final status:",
      resize916Result.job.status
    );

    if (resize916Result.job.status !== "success") {
      throw new Error(
        `9:16 resize failed: ${JSON.stringify(
          resize916Result.job.error || {},
          null,
          2
        )}`
      );
    }

    const designId916 = resize916Result.job.result.design.id;
    console.log("[EXPORT] 9:16 resized designId:", designId916);

    console.log("[EXPORT] creating export job for 9:16 PNG…");
    const export916Job = await createExportJobPng(accessToken, designId916);
    const export916Result = await pollExportJob(
      accessToken,
      export916Job.job.id
    );
    console.log(
      "[EXPORT] 9:16 PNG export status:",
      export916Result.job.status
    );

    const png916Urls =
      export916Result.job.status === "success"
        ? export916Result.job.urls || []
        : [];

    // 3) Resize to 1:1 (1080x1080) and export PNG
    console.log("[EXPORT] creating 1:1 resize job…");
    const resizeSquareJob = await createResizeJob(accessToken, designId, {
      width: 1080,
      height: 1080,
    });

    console.log("[EXPORT] 1:1 resize job created:", resizeSquareJob.job?.id);
    const resizeSquareResult = await pollResizeJob(
      accessToken,
      resizeSquareJob.job.id
    );
    console.log(
      "[EXPORT] 1:1 resize job final status:",
      resizeSquareResult.job.status
    );

    if (resizeSquareResult.job.status !== "success") {
      throw new Error(
        `1:1 resize failed: ${JSON.stringify(
          resizeSquareResult.job.error || {},
          null,
          2
        )}`
      );
    }

    const designIdSquare = resizeSquareResult.job.result.design.id;
    console.log("[EXPORT] 1:1 resized designId:", designIdSquare);

    console.log("[EXPORT] creating export job for 1:1 PNG…");
    const exportSquareJob = await createExportJobPng(
      accessToken,
      designIdSquare
    );
    const exportSquareResult = await pollExportJob(
      accessToken,
      exportSquareJob.job.id
    );
    console.log(
      "[EXPORT] 1:1 PNG export status:",
      exportSquareResult.job.status
    );

    const pngSquareUrls =
      exportSquareResult.job.status === "success"
        ? exportSquareResult.job.urls || []
        : [];

    // ---------- NEW: Local save + S3 upload for all PNG URLs ----------
    const groupCode = generateGroupCode(designId);
    console.log("[EXPORT] Using groupCode:", groupCode);

    const uploadTasks = [];

    // Original 16:9 PNGs
    originalPngUrls.forEach((url, idx) => {
      const imageCode = `orig_p${idx + 1}`;
      uploadTasks.push(
        storeLocallyAndUploadToS3(url, groupCode, imageCode).catch((err) => {
          console.error(
            `[EXPORT] Failed handling original PNG ${idx + 1}:`,
            err
          );
        })
      );
    });

    // 9:16 PNGs
    png916Urls.forEach((url, idx) => {
      const imageCode = `v916_p${idx + 1}`;
      uploadTasks.push(
        storeLocallyAndUploadToS3(url, groupCode, imageCode).catch((err) => {
          console.error(
            `[EXPORT] Failed handling 9:16 PNG ${idx + 1}:`,
            err
          );
        })
      );
    });

    // 1:1 PNGs
    pngSquareUrls.forEach((url, idx) => {
      const imageCode = `sq_p${idx + 1}`;
      uploadTasks.push(
        storeLocallyAndUploadToS3(url, groupCode, imageCode).catch((err) => {
          console.error(
            `[EXPORT] Failed handling 1:1 PNG ${idx + 1}:`,
            err
          );
        })
      );
    });

    // Wait for all local saves + S3 uploads to finish
    await Promise.all(uploadTasks);
    console.log("[EXPORT] All PNGs handled locally + S3 for group:", groupCode);

    // Save for UI
    req.session.lastResult = {
      designId,
      designId916,
      designIdSquare,
      originalPngUrls,
      png916Urls,
      pngSquareUrls,
      groupCode,
      localBaseDir: LOCAL_RECEIVE_DIR,
      s3Bucket: S3_BUCKET,
    };
    req.session.lastError = null;
  } catch (e) {
    console.error("[EXPORT] Error", e);
    req.session.lastError = e.message;
    req.session.lastResult = null;
  }

  res.redirect("/");
});

// ---------- Unity helper endpoints ----------

// List all PNG images under images/
app.get("/unity/list-images", async (req, res) => {
  try {
    const prefix = "images/";

    const cmd = new ListObjectsV2Command({
      Bucket: S3_BUCKET,
      Prefix: prefix,
    });

    const out = await s3.send(cmd);

    const items =
      (out.Contents || [])
        .filter((o) => o.Key && o.Key.endsWith(".png"))
        .map((o) => ({
          key: o.Key,
          size: o.Size || 0,
          lastModified: o.LastModified
            ? o.LastModified.toISOString?.() || null
            : null,
        })) || [];

    res.json({ ok: true, items });
  } catch (e) {
    console.error("[UNITY] /unity/list-images error", e);
    res.status(500).json({ ok: false, error: String(e) });
  }
});

// Download a single PNG by S3 key (proxy)
app.get("/unity/download-image", async (req, res) => {
  const key = req.query.key;
  if (!key) {
    return res.status(400).send("Missing 'key' query param");
  }

  try {
    console.log("[UNITY] download-image for key:", key);

    const cmd = new GetObjectCommand({
      Bucket: S3_BUCKET,
      Key: key,
    });

    const out = await s3.send(cmd);

    res.setHeader("Content-Type", "image/png");

    if (out.Body && typeof out.Body.pipe === "function") {
      // Node stream – pipe directly
      out.Body.pipe(res);
    } else if (out.Body && typeof out.Body.transformToByteArray === "function") {
      // Some runtimes expose transformToByteArray
      const bytes = await out.Body.transformToByteArray();
      res.end(Buffer.from(bytes));
    } else if (out.Body && out.Body.arrayBuffer) {
      const buf = Buffer.from(await out.Body.arrayBuffer());
      res.end(buf);
    } else {
      res.status(500).send("Unexpected S3 body type.");
    }
  } catch (e) {
    console.error("[UNITY] /unity/download-image error", e);
    res.status(500).send("Error fetching image: " + String(e));
  }
});


// ---------- Start server ----------
const port = process.env.PORT || 3001;
app.listen(port, () => {
  console.log(`✅ Server running at http://127.0.0.1:${port}`);
  console.log(`   Local receive dir: ${LOCAL_RECEIVE_DIR}`);
  console.log(`   S3 bucket: ${S3_BUCKET} (region: ${AWS_REGION})`);
});
