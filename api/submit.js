export const config = { runtime: 'edge' };

export default async function handler(req) {
  const corsHeaders = {
    'Access-Control-Allow-Origin': '*',
    'Access-Control-Allow-Methods': 'POST, OPTIONS',
    'Access-Control-Allow-Headers': 'Content-Type',
  };

  if (req.method === 'OPTIONS') {
    return new Response(null, { status: 200, headers: corsHeaders });
  }

  if (req.method !== 'POST') {
    return new Response(JSON.stringify({ error: 'Method not allowed' }), {
      status: 405, headers: { ...corsHeaders, 'Content-Type': 'application/json' },
    });
  }

  const scriptUrl   = process.env.APPS_SCRIPT_URL;
  const icoreUrl    = process.env.ICORE_URL;      // e.g. https://icore.icans.ai
  const icoreSecret = process.env.WEBHOOK_SECRET;

  if (!scriptUrl) {
    return new Response(JSON.stringify({ error: 'APPS_SCRIPT_URL is not configured.' }), {
      status: 500, headers: { ...corsHeaders, 'Content-Type': 'application/json' },
    });
  }

  try {
    const body = await req.text();

    // Fan out to Apps Script + iCore in parallel (iCore failure is non-fatal)
    const [upstream, icoreResult] = await Promise.allSettled([

      // ── Apps Script: Sheets + emails ──────────────────────────────────────
      fetch(scriptUrl, {
        method: 'POST',
        headers: { 'Content-Type': 'text/plain' },
        body,
        redirect: 'follow',
      }),

      // ── iCore: checklist checkbox + account note ──────────────────────────
      icoreUrl
        ? fetch(`${icoreUrl}/api/webhooks/tiger-merchant`, {
            method: 'POST',
            headers: {
              'Content-Type': 'application/json',
              ...(icoreSecret ? { 'x-webhook-secret': icoreSecret } : {}),
            },
            body,
          })
        : Promise.resolve(null),

    ]);

    // Log iCore result — non-fatal so we don't block the form on a CRM error
    if (icoreResult.status === 'rejected') {
      console.warn('[submit] iCore webhook error:', icoreResult.reason?.message);
    } else if (icoreResult.value) {
      const icoreText = await icoreResult.value.text().catch(() => '');
      console.log('[submit] iCore response:', icoreResult.value.status, icoreText.slice(0, 200));
    }

    // Apps Script response drives success/failure back to the form
    if (upstream.status === 'rejected') {
      return new Response(JSON.stringify({ success: false, error: upstream.reason?.message }), {
        status: 500, headers: { ...corsHeaders, 'Content-Type': 'application/json' },
      });
    }

    const text = await upstream.value.text();
    let parsed;
    try { parsed = JSON.parse(text); } catch { parsed = null; }

    if (parsed) {
      return new Response(JSON.stringify(parsed), {
        status: 200,
        headers: { ...corsHeaders, 'Content-Type': 'application/json' },
      });
    }

    return new Response(JSON.stringify({ success: upstream.value.ok, raw: text }), {
      status: upstream.value.ok ? 200 : 502,
      headers: { ...corsHeaders, 'Content-Type': 'application/json' },
    });

  } catch (err) {
    return new Response(JSON.stringify({ error: err.message }), {
      status: 500, headers: { ...corsHeaders, 'Content-Type': 'application/json' },
    });
  }
}
