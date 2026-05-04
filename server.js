import express from 'express';
import cors from 'cors';
import dotenv from 'dotenv';
import { google } from 'googleapis';
import fs from 'fs';
import path from 'path';
import { fileURLToPath } from 'url';

dotenv.config();

const __dirname = path.dirname(fileURLToPath(import.meta.url));
const app = express();

app.use(cors());
app.use(express.json());

const SPREADSHEET_ID = process.env.SPREADSHEET_ID;
const SHEET_NAME = 'Sheet1';

let authClient;

async function initGoogleSheets() {
  try {
    let credentials;
    if (process.env.GOOGLE_CREDENTIALS) {
      credentials = JSON.parse(process.env.GOOGLE_CREDENTIALS);
    } else {
      credentials = JSON.parse(fs.readFileSync(path.join(__dirname, 'credentials.json'), 'utf8'));
    }
    authClient = new google.auth.GoogleAuth({
      credentials,
      scopes: ['https://www.googleapis.com/auth/spreadsheets'],
    });
    console.log('✓ Google Sheets connected');
  } catch (err) {
    console.error('Google Sheets init failed:', err.message);
  }
}

async function getSheets() {
  if (!authClient) await initGoogleSheets();
  return google.sheets({ version: 'v4', auth: authClient });
}

function daysUntilExpiry(dateStr) {
  if (!dateStr) return null;
  const diff = new Date(dateStr) - new Date();
  return Math.ceil(diff / (1000 * 60 * 60 * 24));
}

function computeStatus(row, manualStatus) {
  if (['Archived', 'Terminated', 'On Hold'].includes(manualStatus)) return manualStatus;
  if (['Proposal Sent', 'Due for Discussion', 'Under Negotiation'].includes(manualStatus)) return manualStatus;
  const days = daysUntilExpiry(row['Period To']);
  if (days === null) return manualStatus || 'Unknown';
  if (days < 0) return 'Expired';
  if (days <= 90) return 'Renewal Due';
  return 'Active';
}

function parseRows(rows) {
  if (!rows || rows.length < 2) return [];
  const headers = rows[0];
  return rows.slice(1)
    .map((row, i) => {
      const obj = { id: i + 1 };
      headers.forEach((h, j) => { obj[h] = row[j] || null; });
      obj.daysUntilExpiry = daysUntilExpiry(obj['Period To']);
      obj.status = computeStatus(obj, obj['Status']);
      return obj;
    })
    .filter(r => r['Client Name']);
}

// GET all contracts
app.get('/api/contracts', async (req, res) => {
  try {
    const sheets = await getSheets();
    const response = await sheets.spreadsheets.values.get({
      spreadsheetId: SPREADSHEET_ID,
      range: `${SHEET_NAME}!A1:T`,
    });
    let contracts = parseRows(response.data.values);
    const showArchived = req.query.archived === 'true';
    if (!showArchived) {
      contracts = contracts.filter(c => !['Archived', 'Terminated'].includes(c.status));
    }
    res.json({ contracts, count: contracts.length });
  } catch (err) {
    res.status(500).json({ error: err.message });
  }
});

// GET dashboard metrics
app.get('/api/metrics', async (req, res) => {
  try {
    const sheets = await getSheets();
    const response = await sheets.spreadsheets.values.get({
      spreadsheetId: SPREADSHEET_ID,
      range: `${SHEET_NAME}!A1:T`,
    });
    const all = parseRows(response.data.values);
    res.json({
      total: all.filter(c => !['Archived', 'Terminated'].includes(c.status)).length,
      active: all.filter(c => c.status === 'Active').length,
      renewalDue: all.filter(c => c.status === 'Renewal Due').length,
      expired: all.filter(c => c.status === 'Expired').length,
      inNegotiation: all.filter(c => ['Proposal Sent', 'Due for Discussion', 'Under Negotiation'].includes(c.status)).length,
      archived: all.filter(c => ['Archived', 'Terminated'].includes(c.status)).length,
    });
  } catch (err) {
    res.status(500).json({ error: err.message });
  }
});

// GET renewals due within N days
app.get('/api/renewals', async (req, res) => {
  try {
    const days = parseInt(req.query.days) || 90;
    const sheets = await getSheets();
    const response = await sheets.spreadsheets.values.get({
      spreadsheetId: SPREADSHEET_ID,
      range: `${SHEET_NAME}!A1:T`,
    });
    const contracts = parseRows(response.data.values)
      .filter(c => c.daysUntilExpiry !== null && c.daysUntilExpiry >= 0 && c.daysUntilExpiry <= days)
      .sort((a, b) => a.daysUntilExpiry - b.daysUntilExpiry);
    res.json({ contracts, count: contracts.length });
  } catch (err) {
    res.status(500).json({ error: err.message });
  }
});

// GET search
app.get('/api/search', async (req, res) => {
  try {
    const q = (req.query.q || '').toLowerCase();
    const sheets = await getSheets();
    const response = await sheets.spreadsheets.values.get({
      spreadsheetId: SPREADSHEET_ID,
      range: `${SHEET_NAME}!A1:T`,
    });
    const results = parseRows(response.data.values).filter(c =>
      [c['Client Name'], c['Services in Agreement'], c['Document Type'], c['Status']]
        .some(f => f && f.toLowerCase().includes(q))
    );
    res.json({ results, count: results.length });
  } catch (err) {
    res.status(500).json({ error: err.message });
  }
});

// GET health check
app.get('/api/health', (req, res) => {
  res.json({ status: 'ok', timestamp: new Date().toISOString() });
});

const PORT = process.env.PORT || 3001;
(async () => {
  await initGoogleSheets();
  app.listen(PORT, () => console.log(`✓ Server running on port ${PORT}`));
})();
