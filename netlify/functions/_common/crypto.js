import crypto from 'crypto';

const ALGO = 'aes-256-gcm';
const KEY = crypto.createHash('sha256').update(process.env.SESSION_SECRET||'dev').digest();

export function seal(obj) {
  const iv = crypto.randomBytes(12);
  const cipher = crypto.createCipheriv(ALGO, KEY, iv);
  const data = Buffer.from(JSON.stringify(obj));
  const enc = Buffer.concat([cipher.update(data), cipher.final()]);
  const tag = cipher.getAuthTag();
  return Buffer.concat([iv, tag, enc]).toString('base64url');
}

export function unseal(token) {
  if (!token) return null;
  const buf = Buffer.from(token, 'base64url');
  const iv = buf.subarray(0,12);
  const tag = buf.subarray(12,28);
  const enc = buf.subarray(28);
  const decipher = crypto.createDecipheriv(ALGO, KEY, iv);
  decipher.setAuthTag(tag);
  const dec = Buffer.concat([decipher.update(enc), decipher.final()]);
  return JSON.parse(dec.toString());
}