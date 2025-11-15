import { unseal, seal } from "./_common/crypto.js";
import { tokenExchange } from "./_common/canva.js";

export const handler = async (event, context) => {
  const params = new URLSearchParams(event.rawQuery || "");
  const code = params.get("code");
  const state = params.get("state");
  const cookie = event.headers.cookie || "";
  const cookieMap = Object.fromEntries(cookie.split(/;\s*/).map(s=>s.split("=").map(decodeURIComponent)).filter(a=>a.length===2));
  const pkce = cookieMap["pkce"] ? unseal(cookieMap["pkce"]) : null;
  if (!pkce || pkce.state !== state) return { statusCode: 400, body: "Bad state" };

  try {
    const tok = await tokenExchange(code, pkce.verifier);
    const sess = seal({ token: tok });
    return {
      statusCode: 302,
      headers: {
        Location: "/",
        "Set-Cookie": `sess=${sess}; Path=/; HttpOnly; SameSite=Lax; Secure; Max-Age=2592000`
      }
    };
  } catch (e) {
    return { statusCode: 500, body: "OAuth failed" };
  }
};