import crypto from 'crypto';
const AUTH = "https://www.canva.com/api/oauth/authorize";
const TOKEN = "https://www.canva.com/api/oauth/token";
const API   = "https://api.canva.com/rest/v1";

export function challengeFromVerifier(verifier) {
  return crypto.createHash("sha256").update(verifier).digest("base64url");
}

export async function tokenExchange(code, verifier) {
  const basic = Buffer.from(`${process.env.CANVA_CLIENT_ID}:${process.env.CANVA_CLIENT_SECRET}`).toString("base64");
  const r = await fetch(TOKEN, {
    method: "POST",
    headers: { "Authorization": `Basic ${basic}`, "Content-Type":"application/x-www-form-urlencoded" },
    body: new URLSearchParams({
      grant_type: "authorization_code",
      code, code_verifier: verifier,
      redirect_uri: process.env.CANVA_REDIRECT
    })
  });
  if (!r.ok) throw new Error("OAuth token exchange failed");
  return r.json();
}

export async function pollJob(url, accessToken, timeoutMs = 120000) {
  const start = Date.now();
  let delay = 1200;
  for (;;) {
    const r = await fetch(url, { headers: { Authorization: `Bearer ${accessToken}` } });
    const data = await r.json();
    if (data.status === "completed") return data;
    if (data.status === "failed") throw new Error(data.error?.message || "job_failed");
    if (Date.now() - start > timeoutMs) throw new Error("job_timeout");
    await new Promise(s => setTimeout(s, delay));
    delay = Math.min(delay * 1.3, 4000);
  }
}

export async function createExport(accessToken, designId) {
  const r = await fetch(`${API}/exports`, {
    method: "POST",
    headers: { Authorization: `Bearer ${accessToken}`, "Content-Type": "application/json" },
    body: JSON.stringify({ design_id: designId, format: { type: "pptx" } })
  });
  const j = await r.json();
  if (!r.ok) throw new Error(j.error?.message || "create_export_failed");
  return j.id;
}

export async function getExportUrl(accessToken, jobId) {
  const done = await pollJob(`${API}/exports/${jobId}`, accessToken);
  return done.result.url;
}

export async function createResize(accessToken, designId, width, height) {
  const r = await fetch(`${API}/resizes`, {
    method: "POST",
    headers: { Authorization: `Bearer ${accessToken}`, "Content-Type": "application/json" },
    body: JSON.stringify({ design_id: designId, design_type: { type: "custom", width, height } })
  });
  const j = await r.json();
  if (!r.ok) throw new Error(j.error?.message || "create_resize_failed");
  return j.id;
}

export async function getResizedDesignId(accessToken, jobId) {
  const done = await pollJob(`${API}/resizes/${jobId}`, accessToken);
  return done.result.design.id;
}