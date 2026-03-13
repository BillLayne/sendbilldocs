// Cloudflare Worker — Twilio SMS Proxy for SendBillDocs Agent
// Deploy to Cloudflare Workers and add these secrets via dashboard:
//   TWILIO_SID, TWILIO_TOKEN, TWILIO_FROM, AGENT_PASS

const CORS_HEADERS = {
  'Access-Control-Allow-Origin': '*',
  'Access-Control-Allow-Methods': 'POST, OPTIONS',
  'Access-Control-Allow-Headers': 'Content-Type',
};

export default {
  async fetch(request, env) {
    // Handle CORS preflight
    if (request.method === 'OPTIONS') {
      return new Response(null, { status: 204, headers: CORS_HEADERS });
    }

    if (request.method !== 'POST') {
      return jsonResponse({ error: 'Method not allowed' }, 405);
    }

    try {
      const body = await request.json();
      const { to, message, passcode } = body;

      // Verify agent passcode
      if (passcode !== env.AGENT_PASS) {
        return jsonResponse({ error: 'Unauthorized' }, 401);
      }

      // Validate inputs
      if (!to || !message) {
        return jsonResponse({ error: 'Missing "to" or "message"' }, 400);
      }

      // Clean phone number — keep only digits, prepend +1 if needed
      let phone = to.replace(/\D/g, '');
      if (phone.length === 10) phone = '1' + phone;
      if (!phone.startsWith('+')) phone = '+' + phone;

      // Call Twilio REST API
      const twilioUrl = `https://api.twilio.com/2010-04-01/Accounts/${env.TWILIO_SID}/Messages.json`;
      const auth = btoa(env.TWILIO_SID + ':' + env.TWILIO_TOKEN);

      const twilioResponse = await fetch(twilioUrl, {
        method: 'POST',
        headers: {
          'Authorization': 'Basic ' + auth,
          'Content-Type': 'application/x-www-form-urlencoded',
        },
        body: new URLSearchParams({
          From: env.TWILIO_FROM,
          To: phone,
          Body: message,
        }),
      });

      const result = await twilioResponse.json();

      if (!twilioResponse.ok) {
        return jsonResponse({
          error: 'Twilio error',
          detail: result.message || result.code || 'Unknown error',
          code: result.code
        }, twilioResponse.status);
      }

      return jsonResponse({
        success: true,
        sid: result.sid,
        to: result.to,
        status: result.status
      }, 200);

    } catch (err) {
      return jsonResponse({ error: 'Server error', detail: err.message }, 500);
    }
  }
};

function jsonResponse(data, status) {
  return new Response(JSON.stringify(data), {
    status,
    headers: { 'Content-Type': 'application/json', ...CORS_HEADERS },
  });
}
