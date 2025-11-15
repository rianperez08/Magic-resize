import crypto from 'crypto';
import { seal } from "./_common/crypto.js";
import { challengeFromVerifier } from "./_common/canva.js";

export const handler = async (event, context) => {
  const state = crypto.randomBytes(16).toString("base64url");
  const verifier = crypto.randomBytes(64).toString("base64url");
  const scopes = "design:content:read design:content:write profile:read";
  const challenge = challengeFromVerifier(verifier);
  const url = `https://www.canva.com/api/oauth/authorize?code_challenge=${challenge}&code_challenge_method=s256&scope=${encodeURIComponent(scopes)}&response_type=code&client_id=${process.env.CANVA_CLIENT_ID}&state=${state}&redirect_uri=${encodeURIComponent(process.env.CANVA_REDIRECT)}`;
  const cookie = `pkce=${seal({verifier, state})}; Path=/; HttpOnly; SameSite=Lax; Secure; Max-Age=600`;
  return { statusCode: 302, headers: { Location: url, "Set-Cookie": cookie } };
};