import { unseal } from "./_common/crypto.js";
import { createExport, getExportUrl } from "./_common/canva.js";
import { uploadFromUrlToS3 } from "./_common/s3.js";

export const handler = async (event, context) => {
  if (event.httpMethod !== "POST") return { statusCode: 405, body: "Method Not Allowed" };
  const cookie = event.headers.cookie || "";
  const sess = parseCookie(cookie, "sess");
  const s = sess ? unseal(sess) : null;
  if (!s?.token?.access_token) return { statusCode: 401, body: "unauthorized" };

  const body = JSON.parse(event.body || "{}");
  const { designId } = body;
  if (!designId) return { statusCode: 400, body: "designId required" };

  try {
    const jobId = await createExport(s.token.access_token, designId);
    const canvaUrl = await getExportUrl(s.token.access_token, jobId);
    const key = `canva/original/${designId}.pptx`;
    const finalUrl = await uploadFromUrlToS3(canvaUrl, key, "application/vnd.openxmlformats-officedocument.presentationml.presentation");
    return json200({ s3Url: finalUrl, s3Key: key });
  } catch (e) {
    return json500({ error: e.message || "export_failed" });
  }
};

function parseCookie(cookie, name) {
  const map = Object.fromEntries(cookie.split(/;\s*/).map(s=>s.split("=")));
  return map[name];
}
function json200(obj){ return { statusCode: 200, headers:{ "Content-Type":"application/json"}, body: JSON.stringify(obj)}}
function json500(obj){ return { statusCode: 500, headers:{ "Content-Type":"application/json"}, body: JSON.stringify(obj)}}