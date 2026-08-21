export const config = { runtime: 'edge' };

// Decrypts a submission's sensitive blob (bank account/routing, SSNs).
// Requires two server secrets set in Vercel:
//   DECRYPT_PRIVATE_KEY  — base64 PKCS8 RSA private key (pairs with the public
//                          key embedded in index.html)
//   DECRYPT_PASSWORD     — shared password gating access to this endpoint
// The private key never leaves the server; the browser only ever has the
// public key. Only holders of the password can recover plaintext.

const b64ToBytes = s => Uint8Array.from(atob(s), c => c.charCodeAt(0));

// Constant-time string compare to avoid leaking the password via timing.
function safeEqual(a, b) {
  const ab = new TextEncoder().encode(a);
  const bb = new TextEncoder().encode(b);
  if (ab.length !== bb.length) return false;
  let diff = 0;
  for (let i = 0; i < ab.length; i++) diff |= ab[i] ^ bb[i];
  return diff === 0;
}

async function decryptBlob(blobB64, privB64) {
  const bundle = JSON.parse(atob(blobB64));
  const subtle = crypto.subtle;
  const privKey = await subtle.importKey('pkcs8', b64ToBytes(privB64), { name: 'RSA-OAEP', hash: 'SHA-256' }, false, ['decrypt']);
  const rawKey = new Uint8Array(await subtle.decrypt({ name: 'RSA-OAEP' }, privKey, b64ToBytes(bundle.k)));
  const aesKey = await subtle.importKey('raw', rawKey, { name: 'AES-GCM' }, false, ['decrypt']);
  const pt = await subtle.decrypt({ name: 'AES-GCM', iv: b64ToBytes(bundle.iv) }, aesKey, b64ToBytes(bundle.ct));
  return JSON.parse(new TextDecoder().decode(pt));
}

export default async function handler(req) {
  const headers = { 'Content-Type': 'application/json', 'Cache-Control': 'no-store' };
  if (req.method !== 'POST') {
    return new Response(JSON.stringify({ error: 'Method not allowed' }), { status: 405, headers });
  }

  const privB64 = process.env.DECRYPT_PRIVATE_KEY;
  const password = process.env.DECRYPT_PASSWORD;
  if (!privB64 || !password) {
    return new Response(JSON.stringify({ error: 'Decryption is not configured on the server.' }), { status: 500, headers });
  }

  let body;
  try { body = await req.json(); } catch { body = null; }
  if (!body || !body.blob || typeof body.password !== 'string') {
    return new Response(JSON.stringify({ error: 'Missing blob or password.' }), { status: 400, headers });
  }

  if (!safeEqual(body.password, password)) {
    // Small delay to slow brute-force attempts
    await new Promise(r => setTimeout(r, 600));
    return new Response(JSON.stringify({ error: 'Incorrect password.' }), { status: 401, headers });
  }

  try {
    const fields = await decryptBlob(body.blob, privB64);
    return new Response(JSON.stringify({ ok: true, fields }), { status: 200, headers });
  } catch (err) {
    return new Response(JSON.stringify({ error: 'Could not decrypt — the data may be corrupted or from a different key.' }), { status: 422, headers });
  }
}
