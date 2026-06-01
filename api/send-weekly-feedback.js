const {
  MS_TENANT_ID,
  MS_CLIENT_ID,
  MS_CLIENT_SECRET,
  MS_SENDER_EMAIL,
  SEND_EMAIL_TOKEN,
  ADMIN_ACCESS_TOKEN
} = process.env;

const MAX_BODY_LENGTH = 50000;

function json(res, status, body) {
  res.status(status).setHeader('Content-Type', 'application/json');
  res.end(JSON.stringify(body));
}

function isEmail(value) {
  return /^[^\s@]+@[^\s@]+\.[^\s@]+$/.test(String(value || '').trim());
}

async function readJson(req) {
  if (req.body && typeof req.body === 'object') return req.body;
  if (req.body && typeof req.body === 'string') return JSON.parse(req.body);

  const chunks = [];
  for await (const chunk of req) chunks.push(chunk);
  const raw = Buffer.concat(chunks).toString('utf8');
  return raw ? JSON.parse(raw) : {};
}

async function getGraphAccessToken() {
  const tokenUrl = `https://login.microsoftonline.com/${encodeURIComponent(MS_TENANT_ID)}/oauth2/v2.0/token`;
  const body = new URLSearchParams({
    client_id: MS_CLIENT_ID,
    client_secret: MS_CLIENT_SECRET,
    scope: 'https://graph.microsoft.com/.default',
    grant_type: 'client_credentials'
  });

  const response = await fetch(tokenUrl, {
    method: 'POST',
    headers: { 'Content-Type': 'application/x-www-form-urlencoded' },
    body
  });

  const data = await response.json().catch(() => ({}));
  if (!response.ok || !data.access_token) {
    throw new Error(data.error_description || data.error || 'Unable to authorize Microsoft Graph');
  }

  return data.access_token;
}

async function sendGraphMail({ to, subject, html, text, replyTo }) {
  const accessToken = await getGraphAccessToken();
  const sendUrl = `https://graph.microsoft.com/v1.0/users/${encodeURIComponent(MS_SENDER_EMAIL)}/sendMail`;
  const message = {
    subject,
    body: {
      contentType: html ? 'HTML' : 'Text',
      content: html || text
    },
    toRecipients: [
      { emailAddress: { address: to } }
    ]
  };

  if (isEmail(replyTo)) {
    message.replyTo = [
      { emailAddress: { address: replyTo } }
    ];
  }

  const response = await fetch(sendUrl, {
    method: 'POST',
    headers: {
      Authorization: `Bearer ${accessToken}`,
      'Content-Type': 'application/json'
    },
    body: JSON.stringify({
      message,
      saveToSentItems: true
    })
  });

  if (!response.ok) {
    const details = await response.text();
    throw new Error(details || 'Microsoft Graph sendMail failed');
  }
}

export default async function handler(req, res) {
  if (req.method !== 'POST') {
    return json(res, 405, { error: 'Method not allowed' });
  }

  const expectedToken = SEND_EMAIL_TOKEN || ADMIN_ACCESS_TOKEN;
  if (!MS_TENANT_ID || !MS_CLIENT_ID || !MS_CLIENT_SECRET || !MS_SENDER_EMAIL || !expectedToken) {
    return json(res, 500, {
      error: 'Email sending is not configured',
      requiredEnv: ['MS_TENANT_ID', 'MS_CLIENT_ID', 'MS_CLIENT_SECRET', 'MS_SENDER_EMAIL', 'SEND_EMAIL_TOKEN']
    });
  }

  const token = req.headers['x-send-token'];
  if (!token || token !== expectedToken) {
    return json(res, 401, { error: 'Unauthorized' });
  }

  try {
    const payload = await readJson(req);
    const to = String(payload.to || '').trim();
    const subject = String(payload.subject || '').trim();
    const html = String(payload.html || '').trim();
    const text = String(payload.text || '').trim();
    const replyTo = String(payload.replyTo || '').trim();

    if (!isEmail(to)) return json(res, 400, { error: 'Valid recipient email is required' });
    if (!subject) return json(res, 400, { error: 'Subject is required' });
    if (!html && !text) return json(res, 400, { error: 'Email body is required' });
    if (html.length > MAX_BODY_LENGTH || text.length > MAX_BODY_LENGTH) {
      return json(res, 400, { error: 'Email body is too large' });
    }

    await sendGraphMail({ to, subject, html, text, replyTo });
    return json(res, 200, { ok: true });
  } catch (error) {
    return json(res, 500, {
      error: 'Failed to send weekly feedback email',
      details: error.message
    });
  }
}
