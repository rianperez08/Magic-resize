// server.js
require("dotenv").config();

const express = require("express");
const session = require("cookie-session");
const fetch = require("node-fetch");
const crypto = require("crypto");

const app = express();

app.use(express.urlencoded({ extended: true }));
app.use(express.json());

// --- Session setup ---
app.use(
  session({
    name: "canva-session",
    keys: [process.env.SESSION_SECRET || "super-secret-dev"],
    maxAge: 24 * 60 * 60 * 1000, // 1 day
  })
);

// --- Canva constants ---
const CANVA_AUTH_BASE = "https://www.canva.com/api/oauth/authorize";
const CANVA_TOKEN_URL = "https://api.canva.com/rest/v1/oauth/token";
const CANVA_EXPORTS_URL = "https://api.canva.com/rest/v1/exports";
const CANVA_RESIZES_URL = "https://api.canva.com/rest/v1/resizes";

const CLIENT_ID = process.env.CANVA_CLIENT_ID;
const CLIENT_SECRET = process.env.CANVA_CLIENT_SECRET;
const REDIRECT_URI = process.env.CANVA_REDIRECT_URI;
const SCOPES = process.env.CANVA_SCOPES;

// --- Helpers: PKCE + state ---
function generateRandomBase64Url() {
  return crypto.randomBytes(96).toString("base64url");
}

function generatePkcePair() {
  const codeVerifier = generateRandomBase64Url();
  const codeChallenge = crypto
    .createHash("sha256")
    .update(codeVerifier)
    .digest("base64url");
  return { codeVerifier, codeChallenge };
}

// --- Helper: parse Canva design ID from URL ---
function extractDesignIdFromUrl(urlString) {
  try {
    const url = new URL(urlString);
    const parts = url.pathname.split("/").filter(Boolean); // remove empty

    // Typical: /design/DAGirp_1ZUA/some-slug/view
    const designIndex = parts.indexOf("design");
    if (designIndex !== -1 && parts.length > designIndex + 1) {
      return parts[designIndex + 1];
    }
    return null;
  } catch (e) {
    return null;
  }
}

// --- Helper: basic auth header for token endpoint ---
function basicAuthHeader(clientId, clientSecret) {
  const creds = Buffer.from(`${clientId}:${clientSecret}`).toString("base64");
  return `Basic ${creds}`;
}

// --- Helper: simple sleep ---
function sleep(ms) {
  return new Promise((resolve) => setTimeout(resolve, ms));
}

// --- Views: simple HTML ---
// You can later replace this with React/Vite/etc. For localhost, keep it simple.
function renderHomeHtml({ isAuthed, error, result }) {
  return `
<!DOCTYPE html>
<html lang="en">
<head>
  <meta charset="utf-8" />
  <title>Canva Magic PPTX + 4:3 (Local)</title>
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
  </style>
</head>
<body>
  <div class="card">
    <div class="badge">
      <span>⚙️ Local dev</span>
      <span>·</span>
      <span>Canva Connect APIs</span>
    </div>
    <h1>Magic PPTX + 4:3 Resize (localhost)</h1>
    <p style="font-size:13px;color:#bdbdbd;margin-top:4px;">
      Paste any Canva <strong>design link</strong> you have access to. The app will:
      <br/>1) Export the original design as PPTX
      <br/>2) Create a 4:3 resized copy and export that as PPTX
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
      <form method="POST" action="/convert">
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
            <small>Tip: use a Presentation design for PPTX exports.</small>
          </div>
          <div>
            <small>Resize target: <strong>4:3</strong> (1440x1080) presentation</small>
          </div>
        </div>
        <button class="btn btn-primary" type="submit">
          Convert to PPTX + 4:3
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
        <h3 style="font-size:16px;margin:0 0 6px;">Download links</h3>
        <div style="font-size:13px;margin-bottom:8px;">
          Design ID: <code>${result.designId}</code><br/>
          Resized design ID: <code>${result.resizedDesignId}</code>
        </div>
        <ul style="padding-left:18px;font-size:13px;margin-top:6px;">
          <li>
            Original PPTX:
            ${
              result.originalUrl
                ? `<a class="download-link" href="${result.originalUrl}" target="_blank">Download</a>`
                : `<span style="color:#ffb74d;">(no URL – export failed)</span>`
            }
          </li>
          <li>
            4:3 PPTX:
            ${
              result.resizedUrl
                ? `<a class="download-link" href="${result.resizedUrl}" target="_blank">Download</a>`
                : `<span style="color:#ffb74d;">(no URL – export failed)</span>`
            }
          </li>
        </ul>
        <small>Note: Canva export URLs expire after 24 hours.</small>
      </div>
    `
        : ""
    }
  </div>
</body>
</html>
`;
}

// --- Routes ---

// Home page
app.get("/", (req, res) => {
  const isAuthed = Boolean(req.session.accessToken);
  res.send(
    renderHomeHtml({
      isAuthed,
      error: req.session.lastError || null,
      result: req.session.lastResult || null,
    })
  );
  // Clear flash-ish data
  req.session.lastError = null;
  req.session.lastResult = null;
});

// Start OAuth with Canva
app.get("/auth/canva", (req, res) => {
  if (!CLIENT_ID || !CLIENT_SECRET || !REDIRECT_URI) {
    return res.send(
      "Missing Canva env vars. Please set CANVA_CLIENT_ID, CANVA_CLIENT_SECRET, CANVA_REDIRECT_URI."
    );
  }

  const { codeVerifier, codeChallenge } = generatePkcePair();
  const state = generateRandomBase64Url();

  // Save in session for later verification
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
  res.redirect(authUrl);
});

// OAuth redirect handler
app.get("/oauth/redirect", async (req, res) => {
  const { code, state, error } = req.query;

  if (error) {
    req.session.lastError = `Authorization error: ${error}`;
    return res.redirect("/");
  }

  if (!code || !state) {
    req.session.lastError = "Missing code or state in OAuth redirect.";
    return res.redirect("/");
  }

  if (!req.session.state || state !== req.session.state) {
    req.session.lastError = "State mismatch. Please try connecting again.";
    return res.redirect("/");
  }

  const codeVerifier = req.session.codeVerifier;
  if (!codeVerifier) {
    req.session.lastError = "Missing code_verifier in session. Start auth again.";
    return res.redirect("/");
  }

  try {
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

    if (!tokenRes.ok) {
      const text = await tokenRes.text();
      throw new Error(
        `Token endpoint error (${tokenRes.status}): ${text.slice(0, 500)}`
      );
    }

    const tokenJson = await tokenRes.json();
    // Store basic token info in session
    req.session.accessToken = tokenJson.access_token;
    req.session.refreshToken = tokenJson.refresh_token;
    req.session.tokenScope = tokenJson.scope;
    req.session.tokenExpiresIn = tokenJson.expires_in;

    req.session.lastError = null;
  } catch (e) {
    console.error("[OAuth] Token exchange failed", e);
    req.session.lastError = `Token exchange failed: ${e.message}`;
  }

  // Cleanup
  req.session.codeVerifier = null;
  req.session.state = null;

  res.redirect("/");
});

// Disconnect (clear session tokens)
app.get("/disconnect", (req, res) => {
  req.session = null;
  res.redirect("/");
});

// Core convert route
app.post("/convert", async (req, res) => {
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

  try {
    // 1) Create a 4:3 resize job (e.g. 1440x1080)
    const resizeJob = await createResizeJob(accessToken, designId, {
      width: 1440,
      height: 1080,
    });

    const resizeResult = await pollResizeJob(accessToken, resizeJob.job.id);
    if (resizeResult.job.status !== "success") {
      throw new Error(
        `Resize failed: ${JSON.stringify(resizeResult.job.error || {}, null, 2)}`
      );
    }

    const resizedDesignId = resizeResult.job.result.design.id;

    // 2) Export original design to PPTX
    const exportOriginalJob = await createExportJobPptx(
      accessToken,
      designId
    );
    const exportOriginalResult = await pollExportJob(
      accessToken,
      exportOriginalJob.job.id
    );

    // 3) Export resized 4:3 design to PPTX
    const exportResizedJob = await createExportJobPptx(
      accessToken,
      resizedDesignId
    );
    const exportResizedResult = await pollExportJob(
      accessToken,
      exportResizedJob.job.id
    );

    const originalUrl =
      exportOriginalResult.job.status === "success"
        ? exportOriginalResult.job.urls?.[0] || null
        : null;

    const resizedUrl =
      exportResizedResult.job.status === "success"
        ? exportResizedResult.job.urls?.[0] || null
        : null;

    req.session.lastResult = {
      designId,
      resizedDesignId,
      originalUrl,
      resizedUrl,
    };
    req.session.lastError = null;
  } catch (e) {
    console.error("[Convert] Error", e);
    req.session.lastError = e.message;
    req.session.lastResult = null;
  }

  res.redirect("/");
});

// --- Canva API helpers (Resize + Export) ---

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

  if (!res.ok) {
    const text = await res.text();
    throw new Error(
      `Failed to create resize job (${res.status}): ${text.slice(0, 500)}`
    );
  }

  return res.json();
}

async function pollResizeJob(accessToken, jobId) {
  const url = `${CANVA_RESIZES_URL}/${encodeURIComponent(jobId)}`;

  // Simple polling loop; tune timing & maxAttempts as needed
  const maxAttempts = 20;
  for (let attempt = 0; attempt < maxAttempts; attempt++) {
    const res = await fetch(url, {
      headers: {
        Authorization: `Bearer ${accessToken}`,
        Accept: "application/json",
      },
    });

    if (!res.ok) {
      const text = await res.text();
      throw new Error(
        `Failed to get resize job (${res.status}): ${text.slice(0, 500)}`
      );
    }

    const json = await res.json();
    const status = json.job?.status;

    if (status === "success" || status === "failed") {
      return json;
    }

    await sleep(1500);
  }

  throw new Error("Resize job polling timed out.");
}

async function createExportJobPptx(accessToken, designId) {
  const body = {
    design_id: designId,
    format: {
      type: "pptx",
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

  if (!res.ok) {
    const text = await res.text();
    throw new Error(
      `Failed to create export job (${res.status}): ${text.slice(0, 500)}`
    );
  }

  return res.json();
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

    if (!res.ok) {
      const text = await res.text();
      throw new Error(
        `Failed to get export job (${res.status}): ${text.slice(0, 500)}`
      );
    }

    const json = await res.json();
    const status = json.job?.status;

    if (status === "success" || status === "failed") {
      return json;
    }

    await sleep(1500);
  }

  throw new Error("Export job polling timed out.");
}

// --- Start server ---
const port = process.env.PORT || 3001;
app.listen(port, () => {
  console.log(`✅ Server running at http://127.0.0.1:${port}`);
});
