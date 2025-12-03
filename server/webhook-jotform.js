#!/usr/bin/env node
/**
 * JotForm webhook receiver
 * Place this file at ./server/webhook-jotform.js
 * Run: node server/webhook-jotform.js
 *
 * Env vars (example .env provided below):
 *   JOTFORM_API_KEY
 *   JOTFORM_WEBHOOK_SECRET
 *   GOOGLE_SHEET_ID (optional)
 *   GOOGLE_CREDENTIALS_PATH (optional, default google-sheets-creds.json)
 */
require('dotenv').config();
const express = require('express');
const bodyParser = require('body-parser');
const axios = require('axios');
const fs = require('fs-extra');
const path = require('path');

const PORT = process.env.PORT || 3000;
const JOTFORM_API_KEY = process.env.JOTFORM_API_KEY || '';
const WEBHOOK_SECRET = process.env.JOTFORM_WEBHOOK_SECRET || '';
const GOOGLE_SHEET_ID = process.env.GOOGLE_SHEET_ID || '';
const GOOGLE_CREDENTIALS_PATH = process.env.GOOGLE_CREDENTIALS_PATH || 'google-sheets-creds.json';

if (!JOTFORM_API_KEY) console.warn('WARNING: JOTFORM_API_KEY not set');
if (!WEBHOOK_SECRET) console.warn('WARNING: JOTFORM_WEBHOOK_SECRET not set');

const app = express();
app.use(bodyParser.urlencoded({ extended: true }));
app.use(bodyParser.json());

function normalizeDateToISO(raw, preferDM = false) {
  if (!raw) return null;
  raw = String(raw).trim();
  const ts = Date.parse(raw);
  if (!Number.isNaN(ts)) return new Date(ts).toISOString().slice(0,10);
  const parts = raw.match(/^\s*(\d{1,2})[\/\-\.\s](\d{1,2})[\/\-\.\s](\d{2,4})\s*$/);
  if (parts) {
    let a = parseInt(parts[1],10), b = parseInt(parts[2],10), y = parseInt(parts[3],10);
    if (y < 100) y += (y > 50 ? 1900 : 2000);
    let day, month;
    if (a > 12) { day = a; month = b; } else {
      if (preferDM) { day = a; month = b; } else { month = a; day = b; }
    }
    if (month >=1 && month <=12 && day >=1 && day <=31) {
      return `${String(y).padStart(4,'0')}-${String(month).padStart(2,'0')}-${String(day).padStart(2,'0')}`;
    }
  }
  return null;
}

async function fetchSubmission(submissionId) {
  if (!JOTFORM_API_KEY) throw new Error('Missing JOTFORM_API_KEY');
  const url = `https://api.jotform.com/submission/${submissionId}?apiKey=${JOTFORM_API_KEY}`;
  const resp = await axios.get(url, { timeout: 10000 });
  if (resp.data && resp.data.content) return resp.data.content;
  throw new Error('Bad response from JotForm API');
}

async function downloadFile(url, destDir) {
  await fs.ensureDir(destDir);
  const filename = path.basename(url.split('?')[0]);
  const dest = path.join(destDir, filename);
  const writer = fs.createWriteStream(dest);
  const resp = await axios.get(url, { responseType: 'stream', timeout: 20000 });
  resp.data.pipe(writer);
  return new Promise((resolve, reject) => {
    writer.on('finish', () => resolve(dest));
    writer.on('error', reject);
  });
}

// Replace this with your actual generation function
async function generate(mapped) {
  console.log('Generator would run with mapped data:', mapped);
  // e.g. spawn a process or call internal function
  return true;
}

// Optional helper: append to Google Sheets (non-blocking)
async function appendRowToSheet(rowValues) {
  try {
    const sheetsHelper = require('./append-to-sheets');
    const spreadsheetId = process.env.GOOGLE_SHEET_ID;
    if (!spreadsheetId) { console.warn('GOOGLE_SHEET_ID not set, skipping append'); return; }
    await sheetsHelper.appendRow(spreadsheetId, rowValues, GOOGLE_CREDENTIALS_PATH);
  } catch (err) {
    console.error('appendRowToSheet failed', err.message);
  }
}

app.post('/jotform-webhook', async (req, res) => {
  try {
    const secret = (req.query && req.query.secret) || '';
    if (WEBHOOK_SECRET && secret !== WEBHOOK_SECRET) {
      console.warn('Webhook secret mismatch; rejecting');
      return res.status(401).send('Invalid webhook secret');
    }

    const submissionId = req.body.submission_id || req.body.id || req.body.sid;
    let submission = null;
    if (submissionId) {
      try { submission = await fetchSubmission(submissionId); } catch (err) {
        console.warn('Could not fetch submission via API:', err.message);
      }
    }

    const mapped = {};
    const answers = submission && submission.answers ? submission.answers : null;
    if (answers) {
      for (const key of Object.keys(answers)) {
        const ans = answers[key];
        const value = ans.answer ?? ans.prettyFormat ?? ans.text ?? '';
        const fieldName = (ans.name || ans.text || `q${key}`).toString();
        mapped[fieldName] = value;
      }
    } else {
      for (const k of Object.keys(req.body)) mapped[k] = req.body[k];
    }

    // Normalize date fields (adjust keys to your actual field names)
    const rawDate = mapped.meeting_date || mapped['Meeting Date'] || mapped.date || mapped['date'];
    const isoDate = normalizeDateToISO(rawDate);
    if (isoDate) mapped.meeting_date = isoDate;

    // Download uploads if present
    const tempDir = path.join(__dirname, 'tmp', String(submissionId || Date.now()));
    if (answers) {
      for (const key of Object.keys(answers)) {
        const ans = answers[key];
        if (ans.type === 'control_fileupload' && ans.answer) {
          const urls = Array.isArray(ans.answer) ? ans.answer : [ans.answer];
          mapped.uploads = mapped.uploads || [];
          for (const u of urls) {
            try {
              const downloadUrl = u.includes('?') ? u : `${u}?apiKey=${JOTFORM_API_KEY}`;
              const saved = await downloadFile(downloadUrl, tempDir);
              mapped.uploads.push(saved);
            } catch (err) {
              console.warn('Failed to download upload:', err.message);
            }
          }
        }
      }
    }

    // Respond fast
    res.status(200).send('OK');

    // Async: append to sheet (best-effort) and run generator
    const row = [
      mapped.name || mapped['Name'] || '',
      mapped.email || mapped['Email'] || '',
      mapped.meeting_date || '',
      new Date().toISOString()
    ];
    appendRowToSheet(row); // fire-and-forget

    try { await generate(mapped); } catch (err) { console.error('Generator error', err); }

  } catch (err) {
    console.error('Webhook handler error', err);
    try { res.status(500).send('Server error'); } catch (_) {}
  }
});

app.listen(PORT, () => console.log(`JotForm webhook listening on port ${PORT}`));