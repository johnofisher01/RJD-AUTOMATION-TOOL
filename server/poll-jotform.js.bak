#!/usr/bin/env node
/**
 * server/poll-jotform.js
 *
 * Simple JotForm poller for desktop/local use.
 * - Polls the JotForm submissions API every POLL_INTERVAL_MS
 * - Skips already processed submissions (saved to last_submission.txt)
 * - Maps answers into a simple object and calls generate(mapped)
 * - Tries to require ./fillTemplate.js and call a function export; if that fails,
 *   falls back to spawning `node ./fillTemplate.js` and piping JSON to stdin.
 *
 * Config (put these in your project .env):
 *   JOTFORM_API_KEY        - required
 *   FORM_ID               - required (default 253362621119048)
 *   JOTFORM_API_HOST      - optional (e.g. eu-api.jotform.com for EU Safe mode). Defaults to api.jotform.com
 *   POLL_INTERVAL_MS      - optional (default 15000)
 *   GOOGLE_SHEET_ID       - optional (if you want to append to Google Sheets)
 *   GOOGLE_CREDENTIALS_PATH - optional (default google-sheets-creds.json)
 *
 * Usage:
 *   npm install axios dotenv
 *   node server/poll-jotform.js
 */
require('dotenv').config();
const axios = require('axios');
const fs = require('fs');
const path = require('path');
const child_process = require('child_process');

const FORM_ID = process.env.FORM_ID || '253362621119048';
const API_KEY = process.env.JOTFORM_API_KEY || '';
const API_HOST = process.env.JOTFORM_API_HOST || 'api.jotform.com';
const API_BASE = `https://${API_HOST}`;
const POLL_INTERVAL_MS = Number(process.env.POLL_INTERVAL_MS || 15000);
const LAST_FILE = path.join(__dirname, 'last_submission.txt');
const GOOGLE_SHEET_ID = process.env.GOOGLE_SHEET_ID || '';
const GOOGLE_CREDENTIALS_PATH = process.env.GOOGLE_CREDENTIALS_PATH || 'google-sheets-creds.json';

if (!API_KEY) {
  console.error('Missing JOTFORM_API_KEY in .env — add it and restart this script.');
  process.exit(1);
}

async function listSubmissions() {
  const url = `${API_BASE}/form/${FORM_ID}/submissions?apiKey=${API_KEY}&limit=50&orderby=created_at`;
  const resp = await axios.get(url, { timeout: 10000 });
  if (resp.data && resp.data.content) return resp.data.content;
  return [];
}

function readLastProcessed() {
  try { return fs.readFileSync(LAST_FILE, 'utf8').trim(); } catch (_) { return ''; }
}
function writeLastProcessed(id) {
  try { fs.writeFileSync(LAST_FILE, String(id), 'utf8'); } catch (err) { console.warn('Could not write last file', err.message); }
}

function mapSubmission(submission) {
  const mapped = {};
  if (submission && submission.answers) {
    for (const k of Object.keys(submission.answers)) {
      const ans = submission.answers[k];
      const key = (ans.name || ans.text || `q${k}`).toString();
      mapped[key] = ans.answer ?? ans.prettyFormat ?? ans.text ?? '';
    }
  } else if (submission) {
    // fallback: copy top-level fields
    Object.assign(mapped, submission);
  }
  mapped._submission_id = submission.submission_id || submission.id;
  mapped._created_at = submission.created_at || submission.created_at;
  return mapped;
}

// Optional helper: append to Google Sheets (best-effort)
async function appendRowToSheet(rowValues) {
  if (!GOOGLE_SHEET_ID) return;
  try {
    const sheetsHelper = require('./append-to-sheets');
    await sheetsHelper.appendRow(GOOGLE_SHEET_ID, rowValues, GOOGLE_CREDENTIALS_PATH);
  } catch (err) {
    console.warn('appendRowToSheet failed:', err && err.message ? err.message : err);
  }
}

/**
 * Robust generate(mapped):
 * 1) If server/fillTemplate.js exists:
 *    - try to require() it and call a function export (common names).
 *    - if no callable export found or require() throws, fallback to spawning
 *      `node server/fillTemplate.js` and piping JSON to stdin.
 * 2) If no file, return false.
 */
async function generate(mapped) {
  const modulePath = path.join(__dirname, 'fillTemplate.js');
  if (!fs.existsSync(modulePath)) {
    console.warn('No fillTemplate.js found at server/fillTemplate.js — please create it or edit generate() to call your generator.');
    return false;
  }

  // Try require() first
  try {
    // clear cache to allow quick changes during development
    try { delete require.cache[require.resolve(modulePath)]; } catch (_) {}
    const mod = require(modulePath);

    // Build candidate functions to call
    const candidates = [];
    if (typeof mod === 'function') candidates.push(mod);
    if (mod && typeof mod.default === 'function') candidates.push(mod.default);

    ['generate','fillTemplate','fill','create','run','main','process','createDocument'].forEach(name => {
      if (mod && typeof mod[name] === 'function') candidates.push(mod[name]);
    });

    if (candidates.length) {
      const fn = candidates[0];
      try {
        const maybePromise = fn.length > 0 ? fn(mapped) : fn();
        if (maybePromise && typeof maybePromise.then === 'function') await maybePromise;
        return true;
      } catch (err) {
        console.warn('Calling exported generator function failed:', err && err.message ? err.message : err);
        // fall through to spawn fallback
      }
    } else {
      console.warn('fillTemplate.js found but no callable export detected. Will try spawning it as a CLI.');
    }
  } catch (err) {
    console.warn('Error requiring fillTemplate.js (will try spawn fallback):', err && err.message ? err.message : err);
  }

  // Spawn fallback: run node fillTemplate.js and pipe mapped JSON to stdin
  try {
    return await new Promise((resolve) => {
      const cp = child_process.spawn(process.execPath, [modulePath], { stdio: ['pipe','inherit','inherit'] });
      cp.on('error', (e) => {
        console.warn('Failed to spawn fillTemplate.js:', e && e.message ? e.message : e);
        resolve(false);
      });
      cp.on('exit', (code) => {
        resolve(code === 0);
      });
      try { cp.stdin.write(JSON.stringify(mapped)); } catch (_) {}
      try { cp.stdin.end(); } catch (_) {}
    });
  } catch (err) {
    console.warn('Spawn fallback failed:', err && err.message ? err.message : err);
    return false;
  }
}

async function processNew() {
  try {
    const submissions = await listSubmissions();
    if (!submissions || !submissions.length) return;

    // submissions are newest-first; process oldest-first so ordering is sensible
    submissions.reverse();

    const last = readLastProcessed();
    let newestSeen = last;

    for (const s of submissions) {
      const sid = s.submission_id || s.id;
      if (!sid) continue;
      // if we've already seen this (or it's older), skip
      if (last && sid <= last) continue;

      const mapped = mapSubmission(s);

      // DEBUG: print the mapped submission so you can confirm exact JotForm field keys
      console.log('Mapped submission:', JSON.stringify(mapped, null, 2));

      console.log('Processing submission', sid);
      try {
        const ok = await generate(mapped);
        if (ok) {
          newestSeen = sid;
          // best-effort: append a row to Google Sheets (if configured)
          const row = [
            // Name column
            mapped.name || mapped.Name || mapped['Full Name'] || mapped['Full name'] || mapped['your_name'] || '',

            // Email column
            mapped.email || mapped.Email || mapped['E-mail'] || mapped['your_email'] || '',

            // Meeting / Date column
            mapped.meeting_date || mapped.date || mapped['Date'] || mapped._created_at || '',

            // Work Still to do / Need to go back (explicit header and fallbacks)
            mapped['Work Still to do/Need to go back'] ||
              mapped['work still to do/need to go back'] ||
              mapped['WORK STILL TO DO/NEED TO GO BACK'] ||
              mapped['Work Still to do'] ||
              mapped['works still to do'] ||
              mapped.WORK_STILL_TO_DO ||
              mapped.work_still_to_do ||
              mapped['WORK STILL TO DO'] ||
              '',

            // Timestamp column added by poller
            new Date().toISOString()
          ];
          appendRowToSheet(row);
        } else {
          console.warn('Generator returned false/failed for', sid);
        }
      } catch (err) {
        console.error('Generator failed for', sid, err && err.message ? err.message : err);
        // do not update last processed so it can be retried later
      }
    }

    if (newestSeen && newestSeen !== last) writeLastProcessed(newestSeen);

  } catch (err) {
    console.error('Poll error', err && err.message ? err.message : err);
  }
}

async function run() {
  console.log('Starting JotForm poller for form', FORM_ID, 'using host', API_HOST, 'poll interval', POLL_INTERVAL_MS, 'ms');
  await processNew(); // initial run
  setInterval(processNew, POLL_INTERVAL_MS);
}

if (require.main === module) run();

module.exports = { run };