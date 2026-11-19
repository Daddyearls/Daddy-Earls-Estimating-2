const express = require('express');
const session = require('express-session');
const pgSession = require('connect-pg-simple')(session);
const { Pool } = require('pg');
const { ConfidentialClientApplication } = require('@azure/msal-node');
const PDFDocument = require('pdfkit');
const crypto = require('crypto');
const path = require('path');

const app = express();
const port = process.env.PORT || 3000;
if (process.env.NODE_ENV === 'production' && !process.env.SESSION_SECRET) throw new Error('SESSION_SECRET is required in production');
const pool = new Pool({
  connectionString: process.env.DATABASE_URL,
  ssl: process.env.PGSSLMODE === 'require' ? { rejectUnauthorized: false } : false
});

app.set('trust proxy', 1);
app.use(express.json({ limit: '100kb' }));
app.use(express.urlencoded({ extended: false, limit: '10kb' }));
app.use(session({
  store: new pgSession({ pool, tableName: 'user_sessions', createTableIfMissing: true }),
  secret: process.env.SESSION_SECRET || 'development-only-change-me',
  resave: false,
  saveUninitialized: false,
  cookie: { httpOnly: true, sameSite: 'lax', secure: process.env.NODE_ENV === 'production', maxAge: 8 * 60 * 60 * 1000 }
}));

const LABOR_RATE = 135;
const PARTS_MARKUP = 0.40;
const money = n => new Intl.NumberFormat('en-US', { style: 'currency', currency: 'USD' }).format(Number(n || 0));
const esc = value => String(value ?? '').replace(/[&<>"']/g, c => ({'&':'&amp;','<':'&lt;','>':'&gt;','"':'&quot;',"'":'&#39;'}[c]));
const baseUrl = () => (process.env.APP_BASE_URL || 'https://daddy-earls-estimating.onrender.com').replace(/\/+$/, '');

async function initDb() {
  await pool.query(\`
    CREATE TABLE IF NOT EXISTS estimate_requests (
      id BIGSERIAL PRIMARY KEY,
      created_at TIMESTAMPTZ NOT NULL DEFAULT NOW(),
      customer_name TEXT NOT NULL,
      phone TEXT NOT NULL,
      email TEXT,
      year TEXT NOT NULL,
      make TEXT NOT NULL,
      model TEXT NOT NULL,
      engine TEXT,
      vin TEXT,
      mileage TEXT,
      vehicle_type TEXT NOT NULL DEFAULT 'Automotive',
      concern TEXT NOT NULL,
      status TEXT NOT NULL DEFAULT 'New',
      estimate JSONB,
      owner_notes TEXT
    );
    CREATE INDEX IF NOT EXISTS estimate_requests_created_at_idx ON estimate_requests(created_at DESC);
    ALTER TABLE estimate_requests ADD COLUMN IF NOT EXISTS estimate_email_sent_at TIMESTAMPTZ;
    ALTER TABLE estimate_requests ADD COLUMN IF NOT EXISTS estimate_email_error TEXT;
    ALTER TABLE estimate_requests ADD COLUMN IF NOT EXISTS customer_response_token_hash TEXT;
    ALTER TABLE estimate_requests ADD COLUMN IF NOT EXISTS customer_response_expires_at TIMESTAMPTZ;
    ALTER TABLE estimate_requests ADD COLUMN IF NOT EXISTS customer_decision TEXT;
    ALTER TABLE estimate_requests ADD COLUMN IF NOT EXISTS customer_decision_at TIMESTAMPTZ;\n    CREATE TABLE IF NOT EXISTS microsoft_email_tokens (id INTEGER PRIMARY KEY CHECK (id=1), encrypted_cache TEXT NOT NULL, updated_at TIMESTAMPTZ NOT NULL DEFAULT NOW());
  \`);
}

function requireOwner(req, res, next) {
  if (!req.session || !req.session.owner) return res.status(401).json({ error: 'Please sign in to continue.' });
  next();
}


const EMAIL_SCOPES = ['User.Read', 'Mail.Send', 'offline_access'];
const EMAIL_REDIRECT_URI = () => process.env.MICROSOFT_REDIRECT_URI || baseUrl() + '/auth/microsoft/callback';
function encryptionKey() {
  const raw = process.env.EMAIL_TOKEN_ENCRYPTION_KEY || '';
  const key = /^[a-f0-9]{64}$/i.test(raw) ? Buffer.from(raw, 'hex') : Buffer.from(raw, 'base64');
  if (key.length !== 32) throw new Error('EMAIL_TOKEN_ENCRYPTION_KEY must be a 32-byte key encoded as 64 hex characters or base64.');
  return key;
}
function encryptCache(value) {
  const iv = crypto.randomBytes(12);
  const cipher = crypto.createCipheriv('aes-256-gcm', encryptionKey(), iv);
  const encrypted = Buffer.concat([cipher.update(value, 'utf8'), cipher.final()]);
  return [iv.toString('base64'), cipher.getAuthTag().toString('base64'), encrypted.toString('base64')].join('.');
}
function decryptCache(value) {
  const [iv, tag, data] = String(value).split('.');
  const decipher = crypto.createDecipheriv('aes-256-gcm', encryptionKey(), Buffer.from(iv, 'base64'));
  decipher.setAuthTag(Buffer.from(tag, 'base64'));
  return Buffer.concat([decipher.update(Buffer.from(data, 'base64')), decipher.final()]).toString('utf8');
}
function msalClient() {
  const missing = ['MICROSOFT_CLIENT_ID','MICROSOFT_CLIENT_SECRET','EMAIL_TOKEN_ENCRYPTION_KEY'].filter(k => !process.env[k]);
  if (missing.length) throw new Error('Microsoft email setup is incomplete. Configure ' + missing.join(', ') + ' in Render.');
  return new ConfidentialClientApplication({
    auth: { clientId: process.env.MICROSOFT_CLIENT_ID, authority: 'https://login.microsoftonline.com/consumers', clientSecret: process.env.MICROSOFT_CLIENT_SECRET },
    cache: { cachePlugin: {
      beforeCacheAccess: async context => {
        const result = await pool.query('SELECT encrypted_cache FROM microsoft_email_tokens WHERE id=1');
        if (result.rowCount) context.tokenCache.deserialize(decryptCache(result.rows[0].encrypted_cache));
      },
      afterCacheAccess: async context => {
        if (context.cacheHasChanged) {
          const encrypted = encryptCache(context.tokenCache.serialize());
          await pool.query('INSERT INTO microsoft_email_tokens(id,encrypted_cache,updated_at) VALUES(1,$1,NOW()) ON CONFLICT(id) DO UPDATE SET encrypted_cache=EXCLUDED.encrypted_cache,updated_at=NOW()', [encrypted]);
        }
      }
    }}
  });
}
async function sendMicrosoftMail({ to, subject, html, text: plainText, attachments }) {
  const client = msalClient();
  const accounts = await client.getTokenCache().then(cache => cache.getAllAccounts());
  if (!accounts.length) throw new Error('Connect the shop Outlook account in the dashboard before sending estimates.');
  const token = await client.acquireTokenSilent({ account: accounts[0], scopes: EMAIL_SCOPES });
  const message = { subject, body: { contentType: html ? 'HTML' : 'Text', content: html || plainText || '' }, toRecipients: [{ emailAddress: { address: to } }] };
  if (attachments && attachments.length) message.attachments = attachments.map(a => ({ '@odata.type': '#microsoft.graph.fileAttachment', name: a.filename, contentType: a.contentType, contentBytes: a.content.toString('base64') }));
  const response = await fetch('https://graph.microsoft.com/v1.0/me/sendMail', {
    method: 'POST', headers: { Authorization: 'Bearer ' + token.accessToken, 'Content-Type': 'application/json' },
    body: JSON.stringify({ message, saveToSentItems: true })
  });
  if (!response.ok) {
    const detail = await response.text();
    throw new Error('Outlook could not send the email (' + response.status + '): ' + detail.slice(0, 500));
  }
}

function estimatePdf(row, estimate) {
  return new Promise((resolve, reject) => {
    try {
      const doc = new PDFDocument({ size: 'LETTER', margin: 52 });
      const chunks = [];
      doc.on('data', chunk => chunks.push(chunk));
      doc.on('end', () => resolve(Buffer.concat(chunks)));
      doc.on('error', reject);
      const items = Array.isArray(estimate.items) ? estimate.items : [];
      const totals = estimate.totals || {};
      doc.fontSize(20).font('Helvetica-Bold').text("DADDY EARL'S AUTO REPAIR");
      doc.moveDown(0.25).fontSize(10).font('Helvetica').text('ESTIMATE');
      doc.moveDown().fontSize(11).text('Prepared for: ' + row.customer_name);
      doc.text('Vehicle: ' + [row.year, row.make, row.model].filter(Boolean).join(' '));
      doc.text('Vehicle type: ' + (row.vehicle_type || 'Automotive'));
      doc.text('Date: ' + new Date().toLocaleDateString('en-US'));
      doc.moveDown().font('Helvetica-Bold').text('Requested service');
      doc.font('Helvetica').text(row.concern || '—', { width: 500 });
      doc.moveDown().font('Helvetica-Bold').text('Estimate details');
      items.forEach((item, index) => {
        const lineTotal = item.kind === 'Labor' ? Number(item.quantity || 0) * LABOR_RATE : Number(item.quantity || 0) * Number(item.unitCost || 0) * (estimate.partsType !== 'Customer Supplied' ? 1 + PARTS_MARKUP : 1);
        const label = item.kind === 'Labor' ? 'Labor' : 'Parts';
        doc.font('Helvetica').text((index + 1) + '. ' + label + ' — ' + (item.description || 'Service item') + ' | Qty: ' + Number(item.quantity || 0) + ' | ' + money(lineTotal), { width: 500 });
      });
      doc.moveDown().font('Helvetica-Bold').text('Estimated totals');
      doc.font('Helvetica').text('Labor: ' + money(totals.labor));
      doc.text('Parts: ' + money(totals.parts));
      doc.moveDown(0.25).font('Helvetica-Bold').fontSize(14).text('Estimated total: ' + money(totals.total));
      if (estimate.customerMessage) doc.moveDown().fontSize(10).font('Helvetica').text('Note from our team: ' + estimate.customerMessage, { width: 500 });
      doc.moveDown(2).fontSize(9).fillColor('#555').text("This estimate is based on the information currently available. Additional work or parts may require an updated estimate and customer authorization. Please contact Daddy Earl's Auto Repair with any questions.");
      doc.end();
    } catch (error) { reject(error); }
  });
}

async function sendEstimate(row) {
  if (!row.email || !String(row.email).trim()) throw new Error('This customer does not have an email address on file.');
  if (row.estimate_email_sent_at) return { sent: true, alreadySent: true };

  const estimate = row.estimate || {};
  const totals = estimate.totals || {};
  const token = crypto.randomBytes(32).toString('hex');
  const tokenHash = crypto.createHash('sha256').update(token).digest('hex');
  await pool.query(
    'UPDATE estimate_requests SET customer_response_token_hash=$2, customer_response_expires_at=NOW()+INTERVAL \'30 days\', estimate_email_error=NULL WHERE id=$1',
    [row.id, tokenHash]
  );
  const responseUrl = baseUrl() + '/estimate-response/' + token;
  const pdf = await estimatePdf(row, estimate);
  const subject = "Your repair estimate from Daddy Earl's Auto Repair";
  const html = \`<div style="font-family:Arial,sans-serif;color:#202124;max-width:620px;margin:auto;line-height:1.55"><div style="background:#171717;color:#fff;padding:24px"><div style="font-size:21px;font-weight:bold;letter-spacing:1px">DADDY EARL'S AUTO REPAIR</div><div style="color:#e6b84b;margin-top:5px">YOUR REPAIR ESTIMATE</div></div><div style="padding:24px"><p>Hello \${esc(row.customer_name)},</p><p>Thank you for choosing Daddy Earl's Auto Repair. Your estimate is ready. A PDF copy is attached for your records.</p><div style="background:#f4f4f4;padding:16px;border-radius:8px"><strong>Vehicle:</strong> \${esc([row.year,row.make,row.model].filter(Boolean).join(' '))}<br><strong>Service:</strong> \${esc(row.concern)}<br><strong>Estimated total:</strong> <span style="font-size:22px;font-weight:bold">\${money(totals.total)}</span></div>\${estimate.customerMessage ? '<p>' + esc(estimate.customerMessage) + '</p>' : ''}<p>Please review the attached estimate and choose an option below. Your response will be recorded and sent to our team.</p><p style="margin:28px 0"><a href="\${responseUrl}" style="background:#171717;color:#fff;padding:13px 20px;text-decoration:none;border-radius:5px">Review estimate & respond</a></p><p style="font-size:13px;color:#666">This link is valid for 30 days. Selecting approve records your response; our team will confirm scheduling before work begins.</p><p>Thank you,<br>Daddy Earl's Auto Repair</p></div></div>\`;
  await transport.sendMail({
    from: process.env.SMTP_FROM,
    to: row.email,
    subject,
    html,
    attachments: [{ filename: 'Daddy-Earls-Estimate-' + row.id + '.pdf', content: pdf, contentType: 'application/pdf' }]
  });
  await pool.query('UPDATE estimate_requests SET estimate_email_sent_at=NOW(), estimate_email_error=NULL, status=CASE WHEN status=\'Approved\' THEN \'Sent\' ELSE status END WHERE id=$1', [row.id]);
  const notifyTo = process.env.OWNER_NOTIFICATION_EMAIL;
  if (notifyTo) {
    await transport.sendMail({
      from: process.env.SMTP_FROM, to: notifyTo,
      subject: 'Estimate emailed to ' + row.customer_name + ' (#' + row.id + ')',
      text: 'Estimate #' + row.id + ' was emailed to ' + row.customer_name + ' (' + row.email + '). Estimated total: ' + money(totals.total) + '.'
    }).catch(error => console.error('Owner notification email failed:', error.message));
  }
  return { sent: true, alreadySent: false };
}

app.get('/api/health', async (req, res) => {
  try { await pool.query('SELECT 1'); res.json({ ok: true, database: 'connected', emailConfigured: !!(process.env.MICROSOFT_CLIENT_ID && process.env.MICROSOFT_CLIENT_SECRET && process.env.EMAIL_TOKEN_ENCRYPTION_KEY) }); }
  catch (e) { res.status(503).json({ ok: false, database: 'unavailable' }); }
});

app.post('/api/login', (req, res) => {
  const supplied = String(req.body.password || '');
  const expected = process.env.ADMIN_PASSWORD || '';
  if (!expected) return res.status(503).json({ error: 'Owner sign-in is not configured yet.' });
  const a = Buffer.from(supplied);
  const b = Buffer.from(expected);
  const valid = a.length === b.length && require('crypto').timingSafeEqual(a, b);
  if (!valid) return res.status(401).json({ error: 'Incorrect password.' });
  req.session.owner = true;
  res.json({ ok: true });
});

app.post('/api/logout', requireOwner, (req, res) => req.session.destroy(() => res.json({ ok: true })));
app.get('/api/email/status', requireOwner, async (req, res) => {
  try {
    const result = await pool.query('SELECT updated_at FROM microsoft_email_tokens WHERE id=1');
    res.json({ connected: !!result.rowCount, updatedAt: result.rowCount ? result.rows[0].updated_at : null });
  } catch (e) { res.status(500).json({ error: 'Could not check Outlook connection.' }); }
});
app.get('/api/email/connect', requireOwner, async (req, res) => {
  try {
    const state = crypto.randomBytes(24).toString('hex');
    req.session.microsoftOAuthState = state;
    const url = await msalClient().getAuthCodeUrl({ scopes: EMAIL_SCOPES, redirectUri: EMAIL_REDIRECT_URI(), state, prompt: 'select_account' });
    res.redirect(url);
  } catch (e) { res.status(500).send(esc(e.message)); }
});
app.get('/auth/microsoft/callback', async (req, res) => {
  try {
    if (!req.session || !req.session.owner) return res.status(401).send('Please sign in to the owner dashboard and try connecting Outlook again.');
    if (!req.query.code || !req.query.state || req.query.state !== req.session.microsoftOAuthState) return res.status(400).send('Outlook sign-in could not be verified. Please try again.');
    delete req.session.microsoftOAuthState;
    await msalClient().acquireTokenByCode({ code: String(req.query.code), scopes: EMAIL_SCOPES, redirectUri: EMAIL_REDIRECT_URI() });
    res.redirect('/dashboard?outlook=connected');
  } catch (e) { console.error('Microsoft OAuth callback failed:', e); res.status(500).send('Could not connect Outlook. Return to the dashboard and try again.');
  }
});
app.get('/api/me', (req, res) => res.json({ signedIn: !!(req.session && req.session.owner) }));

app.post('/api/requests', async (req, res) => {
  const b = req.body || {};
  const required = ['customerName','phone','year','make','model','concern'];
  if (required.some(k => !String(b[k] || '').trim())) return res.status(400).json({ error: 'Please complete all required fields.' });
  try {
    const result = await pool.query(
      \`INSERT INTO estimate_requests (customer_name, phone, email, year, make, model, engine, vin, mileage, vehicle_type, concern)
       VALUES ($1,$2,$3,$4,$5,$6,$7,$8,$9,$10,$11) RETURNING id\`,
      [b.customerName.trim(), b.phone.trim(), (b.email || '').trim(), String(b.year).trim(), b.make.trim(), b.model.trim(), b.engine || '', b.vin || '', b.mileage || '', b.vehicleType || 'Automotive', b.concern.trim()]
    );
    res.status(201).json({ ok: true, id: result.rows[0].id });
  } catch (e) { console.error(e); res.status(500).json({ error: 'We could not submit your request. Please call the shop.' }); }
});

app.get('/api/requests', requireOwner, async (req, res) => {
  try {
    const result = await pool.query('SELECT * FROM estimate_requests ORDER BY created_at DESC LIMIT 500');
    res.json(result.rows);
  } catch (e) { console.error(e); res.status(500).json({ error: 'Could not load requests.' }); }
});

app.patch('/api/requests/:id', requireOwner, async (req, res) => {
  const id = Number(req.params.id);
  const { status, estimate, ownerNotes } = req.body || {};
  const allowed = ['New','In Review','Estimate Draft','Approved','Sent','Declined','Completed'];
  if (!Number.isSafeInteger(id) || id < 1) return res.status(400).json({ error: 'Invalid request.' });
  if (status && !allowed.includes(status)) return res.status(400).json({ error: 'Invalid status.' });
  try {
    const result = await pool.query(
      \`UPDATE estimate_requests SET
        status = COALESCE($2,status),
        estimate = COALESCE($3::jsonb,estimate),
        owner_notes = COALESCE($4,owner_notes)
       WHERE id = $1 RETURNING *\`,
      [id, status || null, estimate ? JSON.stringify(estimate) : null, ownerNotes === undefined ? null : String(ownerNotes)]
    );
    if (!result.rowCount) return res.status(404).json({ error: 'Request not found.' });
    let row = result.rows[0];
    if (status === 'Approved' && !row.estimate_email_sent_at) {
      try {
        const delivery = await sendEstimate(row);
        const refreshed = await pool.query('SELECT * FROM estimate_requests WHERE id=$1', [id]);
        row = refreshed.rows[0];
        return res.json({ ...row, emailSent: delivery.sent, emailAlreadySent: delivery.alreadySent });
      } catch (error) {
        await pool.query('UPDATE estimate_requests SET estimate_email_error=$2 WHERE id=$1', [id, error.message]);
        return res.status(200).json({ ...row, emailSent: false, emailError: error.message });
      }
    }
    res.json({ ...row, emailSent: !!row.estimate_email_sent_at, emailAlreadySent: !!row.estimate_email_sent_at });
  } catch (e) { console.error(e); res.status(500).json({ error: 'Could not update request.' }); }
});

app.get('/estimate-response/:token', async (req, res) => {
  const hash = crypto.createHash('sha256').update(String(req.params.token || '')).digest('hex');
  try {
    const result = await pool.query('SELECT id,customer_name,year,make,model,estimate,customer_decision,customer_response_expires_at FROM estimate_requests WHERE customer_response_token_hash=$1', [hash]);
    if (!result.rowCount || !result.rows[0].customer_response_expires_at || new Date(result.rows[0].customer_response_expires_at) < new Date()) {
      return res.status(404).send('<!doctype html><html><body style="font-family:Arial;padding:40px"><h1>Link unavailable</h1><p>This estimate response link is invalid or has expired. Please contact Daddy Earl\'s Auto Repair.</p></body></html>');
    }
    const r = result.rows[0];
    if (r.customer_decision) return res.send('<!doctype html><html><body style="font-family:Arial;padding:40px"><h1>Response already received</h1><p>We have recorded your response: ' + esc(r.customer_decision) + '. Thank you.</p></body></html>');
    const total = r.estimate && r.estimate.totals ? money(r.estimate.totals.total) : 'To be confirmed';
    res.send('<!doctype html><html><head><meta name="viewport" content="width=device-width,initial-scale=1"><title>Review estimate</title></head><body style="font-family:Arial,sans-serif;max-width:620px;margin:40px auto;padding:20px;color:#222"><h1 style="background:#171717;color:white;padding:22px">Daddy Earl\'s Auto Repair</h1><h2>Review your estimate</h2><p>Hello ' + esc(r.customer_name) + ',</p><p>Vehicle: ' + esc([r.year,r.make,r.model].filter(Boolean).join(' ')) + '</p><h2>Estimated total: ' + total + '</h2><p>Choose whether you would like us to proceed. Our team will contact you to confirm scheduling. Your selection records your response but does not by itself authorize additional work beyond this estimate.</p><form method="POST" action="/estimate-response/' + encodeURIComponent(req.params.token) + '"><button name="decision" value="Approved" style="padding:14px 20px;background:#176b3a;color:white;border:0;border-radius:5px;margin-right:8px">Approve estimate</button><button name="decision" value="Declined" style="padding:14px 20px;background:#555;color:white;border:0;border-radius:5px">Decline</button></form></body></html>');
  } catch (e) { console.error(e); res.status(500).send('Unable to load estimate response.'); }
});

app.post('/estimate-response/:token', async (req, res) => {
  const decision = req.body.decision;
  if (!['Approved','Declined'].includes(decision)) return res.status(400).send('Invalid response.');
  const hash = crypto.createHash('sha256').update(String(req.params.token || '')).digest('hex');
  try {
    const result = await pool.query(
      \`UPDATE estimate_requests SET customer_decision=$2, customer_decision_at=NOW()
       WHERE customer_response_token_hash=$1 AND customer_response_expires_at>NOW() AND customer_decision IS NULL
       RETURNING id,customer_name,email,estimate\`,
      [hash, decision]
    );
    if (!result.rowCount) return res.status(404).send('This link has expired or a response was already recorded.');
    const r = result.rows[0];
    if (process.env.OWNER_NOTIFICATION_EMAIL) {
      try {
        await mailer().then(m => m.sendMail({
          from: process.env.SMTP_FROM, to: process.env.OWNER_NOTIFICATION_EMAIL,
          subject: 'Customer ' + decision.toLowerCase() + ' estimate #' + r.id,
          text: r.customer_name + ' (' + (r.email || 'no email') + ') has ' + decision.toLowerCase() + ' estimate #' + r.id + '.'
        }));
      } catch (e) { console.error('Customer response notification failed:', e.message); }
    }
    res.send('<!doctype html><html><body style="font-family:Arial;padding:40px;max-width:620px;margin:auto"><h1>Thank you</h1><p>Your response has been recorded: <strong>' + esc(decision) + '</strong>.</p><p>Daddy Earl\'s Auto Repair has been notified and will follow up with you.</p></body></html>');
  } catch (e) { console.error(e); res.status(500).send('We could not record your response. Please call the shop.'); }
});

app.use(express.static(path.join(__dirname, 'public')));
app.get('/dashboard', (req, res) => res.sendFile(path.join(__dirname, 'public', 'dashboard.html')));
app.get('*', (req, res) => res.sendFile(path.join(__dirname, 'public', 'index.html')));

initDb().then(() => {
  app.listen(port, '0.0.0.0', () => console.log(\`Daddy Earl's Estimating listening on \${port}\`));
}).catch(err => { console.error('Database initialization failed:', err); process.exit(1); });
