export const handler = async (event, context) => {
  return {
    statusCode: 200,
    headers: { "Set-Cookie": "sess=; Path=/; HttpOnly; SameSite=Lax; Secure; Max-Age=0" },
    body: JSON.stringify({ ok: true })
  };
};