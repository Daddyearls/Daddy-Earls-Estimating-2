const express = require('express');
const session = require('express-session');
const pgSession = require('connect-pg-simple')(session);
const { Pool } = require('pg');
const path = require('path');

const app = express();
const port = process.env.PORT || 3000;
const pool = new Pool({
  connectionString: process.env.DATABASE_URL,
  ssl: process.env.PGSSLMODE === 'require' ? { rejectUnauthorized: false } : false
});

app.set('trust proxy', 1);
app.use(express.json({ limit: '100kb' }));
app.use(session({
  store: new pgSession({ pool, tableName: 'user_sessions', createTableIfMissing: true }),
  secret: process.env.SESSION_SECRET || 'development-only-change-me',
  resave: false,
  saveUninitialized: false,
  cookie: { httpOnly: true, sameSite: 'lax', secure: process.env.NODE_ENV === 'production', maxAge: 8 * 60 * 60 * 1000 }
}));

const LABOR_RATE = 135;
const PARTS_MARKUP = 0.40;

async function initDb() {
  await pool.query(`
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
  `);
}

function requireOwner(req, res, next) {
  if (!req.session || !req.session.owner) return res.status(401).json({ error: 'Please sign in to continue.' });
  next();
}

app.get('/api/health', async (req, res) => {
  try { await pool.query('SELECT 1'); res.json({ ok: true, database: 'connected' }); }
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
app.get('/api/me', (req, res) => res.json({ signedIn: !!(req.session && req.session.owner) }));

app.post('/api/requests', async (req, res) => {
  const b = req.body || {};
  const required = ['customerName','phone','year','make','model','concern'];
  if (required.some(k => !String(b[k] || '').trim())) return res.status(400).json({ error: 'Please complete all required fields.' });
  try {
    const result = await pool.query(
      `INSERT INTO estimate_requests (customer_name, phone, email, year, make, model, engine, vin, mileage, vehicle_type, concern)
       VALUES ($1,$2,$3,$4,$5,$6,$7,$8,$9,$10,$11) RETURNING id`,
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
      `UPDATE estimate_requests SET
        status = COALESCE($2,status),
        estimate = COALESCE($3::jsonb,estimate),
        owner_notes = COALESCE($4,owner_notes)
       WHERE id = $1 RETURNING *`,
      [id, status || null, estimate ? JSON.stringify(estimate) : null, ownerNotes === undefined ? null : String(ownerNotes)]
    );
    if (!result.rowCount) return res.status(404).json({ error: 'Request not found.' });
    res.json(result.rows[0]);
  } catch (e) { console.error(e); res.status(500).json({ error: 'Could not update request.' }); }
});

app.use(express.static(path.join(__dirname, 'public')));
app.get('*', (req, res) => res.sendFile(path.join(__dirname, 'public', 'index.html')));

initDb().then(() => {
  app.listen(port, '0.0.0.0', () => console.log(`Daddy Earl's Estimating listening on ${port}`));
}).catch(err => { console.error('Database initialization failed:', err); process.exit(1); });
