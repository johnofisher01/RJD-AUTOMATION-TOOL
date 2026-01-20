// Final authoritative generator — deterministic British D-M-YYYY filenames only.
// Includes lock-file protection to avoid concurrent runs.
const fs = require('fs-extra');
const path = require('path');
const { google } = require('googleapis');
const PizZip = require('pizzip');
const Docxtemplater = require('docxtemplater');
const minimist = require('minimist');

require('dotenv').config({ path: path.join(__dirname, '..', '.env') });

const argv = minimist(process.argv.slice(2));
const SHEET_ID = process.env.GOOGLE_SHEET_ID;
const CREDS_PATH = process.env.GOOGLE_CREDENTIALS_PATH || path.join(__dirname, 'google-sheets-creds.json');
const TEMPLATE_PATH = path.join(__dirname, 'templates', 'worksheet_template_10.docx');
const OUTPUT_DIR = argv.output ? path.resolve(argv.output) : path.join(__dirname, 'output');
const SHEET_RANGE = process.env.GOOGLE_SHEET_RANGE || 'Form Responses 1';

const LAST_N = Number(argv.last || argv.lastN || argv.max || 0);
const PRUNE_N = Number(argv.prune || 0);
const FORCE = !!argv.force;
const DRY = !!argv.dry || !!argv['dry-run'];
const DEBUG = !!argv.debug;
const DAYS = Number(argv.days || argv['since-days'] || process.env.DEFAULT_DAYS || 7);

// Lock-file to prevent concurrent runs
const LOCK_PATH = path.join(__dirname, '.generator.lock');
if (fs.existsSync(LOCK_PATH)) {
  console.error('Another generator run appears active (lock present). Exiting.');
  process.exit(1);
}
try { fs.writeFileSync(LOCK_PATH, String(process.pid)); } catch (e) {}
function removeLock() { try { if (fs.existsSync(LOCK_PATH)) fs.unlinkSync(LOCK_PATH); } catch (e) {} }
process.on('exit', removeLock);
process.on('SIGINT', () => process.exit(1));
process.on('SIGTERM', () => process.exit(1));

const FINAL_DATE_REGEX = /^\d{1,2}-\d{1,2}-\d{4}$/;

function safeFilename(str) {
  return String(str || '').replace(/[\/\\?%*:|"<>]/g, '-').trim();
}

async function getRows() {
  if (!SHEET_ID) throw new Error('Missing GOOGLE_SHEET_ID in .env!');
  if (!CREDS_PATH) throw new Error('Missing GOOGLE_CREDENTIALS_PATH in .env!');
  if (!fs.existsSync(CREDS_PATH)) throw new Error(`Credentials file not found at: ${CREDS_PATH}`);

  const creds = require(path.resolve(CREDS_PATH));
  const auth = new google.auth.GoogleAuth({
    credentials: creds,
    scopes: ['https://www.googleapis.com/auth/spreadsheets.readonly'],
  });
  const sheets = google.sheets({ version: 'v4', auth });

  console.log(`Using spreadsheetId=${SHEET_ID} range="${SHEET_RANGE}"`);

  const response = await sheets.spreadsheets.values.get({
    spreadsheetId: SHEET_ID,
    range: SHEET_RANGE,
  });

  const values = response.data.values || [];
  const [header, ...rows] = values;
  if (!header) return [];

  const mappedRows = rows.map(row =>
    header.reduce((obj, key, i) => {
      obj[key] = row[i] || '';
      return obj;
    }, { __row: row })
  );

  if (mappedRows.length) {
    console.log('Sample mapped row (first):', mappedRows[0]);
    console.log('Mapped to template fields (sample):', mapSheetRowToTemplateFields(mappedRows[0]));
  }

  return mappedRows;
}

/**
 * Deterministic parser — numeric strings ARE ALWAYS day-first (DMY).
 * Recognises numeric DMY, ISO, and common textual month formats.
 */
function parseDateDeterministic(s) {
  if (!s) return { date: null, kind: null };
  const str = String(s).trim();

  // 1) numeric DMY (allow leading zeros)
  let m = str.match(/^0*(\d{1,2})[\/\-\.\s]0*(\d{1,2})[\/\-\.\s](\d{2,4})$/);
  if (m) {
    const day = Number(m[1]), month = Number(m[2]), rawYear = Number(m[3]);
    const year = rawYear < 100 ? 2000 + rawYear : rawYear;
    if (month >= 1 && month <= 12 && day >= 1 && day <= 31) {
      const d = new Date(year, month - 1, day);
      if (d.getFullYear() === year && d.getMonth() === month - 1 && d.getDate() === day) return { date: d, kind: 'dmy' };
    }
    return { date: null, kind: null };
  }

  // 2) ISO yyyy-mm-dd
  m = str.match(/^(\d{4})-(\d{1,2})-(\d{1,2})$/);
  if (m) {
    const year = Number(m[1]), month = Number(m[2]), day = Number(m[3]);
    const d = new Date(year, month - 1, day);
    if (d.getFullYear() === year && d.getMonth() === month - 1 && d.getDate() === day) return { date: d, kind: 'iso' };
    return { date: null, kind: null };
  }

  // 3) textual months e.g. "5 Jan 2026", "January 5 2026", "5th January 2026"
  const monthNames = {
    january:1, february:2, march:3, april:4, may:5, june:6,
    july:7, august:8, september:9, october:10, november:11, december:12,
    jan:1, feb:2, mar:3, apr:4, may:5, jun:6, jul:7, aug:8, sep:9, sept:9, oct:10, nov:11, dec:12
  };
  const cleaned = str.replace(/,/g,' ').replace(/\s+/g,' ').trim();
  m = cleaned.match(/^(\d{1,2})(?:st|nd|rd|th)?\s+([A-Za-z]{3,9})\s+(\d{4})$/i);
  if (m) {
    const day = Number(m[1]), month = monthNames[m[2].toLowerCase()], year = Number(m[3]);
    if (month && day >=1 && day <=31) {
      const d = new Date(year, month -1, day);
      if (d.getFullYear() === year && d.getMonth() === month -1 && d.getDate() === day) return { date: d, kind: 'text' };
    }
  }
  m = cleaned.match(/^([A-Za-z]{3,9})\s+(\d{1,2})(?:st|nd|rd|th)?\s+(\d{4})$/i);
  if (m) {
    const month = monthNames[m[1].toLowerCase()], day = Number(m[2]), year = Number(m[3]);
    if (month && day >=1 && day <=31) {
      const d = new Date(year, month -1, day);
      if (d.getFullYear() === year && d.getMonth() === month -1 && d.getDate() === day) return { date: d, kind: 'text' };
    }
  }

  return { date: null, kind: null };
}

function formatDateForFilename(date) {
  if (!date || !(date instanceof Date) || Number.isNaN(date.getTime())) return '';
  return `${String(date.getDate())}-${String(date.getMonth() + 1)}-${String(date.getFullYear())}`;
}

function normalizeRawDateForFilename(raw) {
  if (!raw) return null;
  const s = String(raw).trim();
  const parts = s.match(/^0*(\d{1,2})[\/\-\.\s]0*(\d{1,2})[\/\-\.\s](\d{2,4})$/);
  if (!parts) return null;
  let day = Number(parts[1]), month = Number(parts[2]), rawYear = Number(parts[3]);
  const year = rawYear < 100 ? 2000 + rawYear : rawYear;
  const d = new Date(year, month - 1, day);
  if (d.getFullYear() === year && d.getMonth() === month - 1 && d.getDate() === day) return `${day}-${month}-${year}`;
  return null;
}

/**
 * Map sheet row to template fields (keeps supplier O and S handling)
 */
function mapSheetRowToTemplateFields(row) {
  const normalize = (s) => String(s || '').toLowerCase().replace(/[^a-z0-9]/g, '');
  const normRow = {};
  for (const k of Object.keys(row || {})) { if (k === '__row') continue; normRow[normalize(k)] = row[k]; }

  const pick = (...names) => { for (const name of names) { const v = normRow[normalize(name)]; if (v !== undefined && v !== null && String(v).trim() !== '') return String(v).trim(); } return ''; };
  const pickByIndex = (idx) => { const arr = row && row.__row ? row.__row : []; if (!Array.isArray(arr)) return ''; const v = arr[idx]; return (v !== undefined && v !== null && String(v).trim() !== '') ? String(v).trim() : ''; };

  const addrParts = [
    pick('Address - Street Address', 'Address - Street Address', 'street address', 'address'),
    pick('Address - Street Address Line 2', 'Address - Street Address Line 2', 'street address line 2', 'address line 2'),
    pick('Address - City', 'Address - City', 'city'),
    pick('Address - State / Province', 'Address - State / Province', 'state', 'province'),
    pick('Address - Postal / Zip Code', 'Address - Postal / Zip Code', 'postal', 'zip', 'postcode')
  ].filter(Boolean);
  const ADDRESS = addrParts.join(', ');
  const supplierFromO = pickByIndex(14);
  const supplierFromS = pickByIndex(18);

  return {
    NAME: pick('Name','Full Name','full name','name'),
    DATE: pick('Date','DATE','_created_at','date'),
    JOB_NO: pick('Job Number','JOB NUMBER','Job No','job number','job_no'),
    CUSTOMER: pick('Customer','Customer Name','CLIENT','client'),
    ADDRESS,
    WORKS_CARRIED_OUT: pick('Works carried out','WORKS CARRIED OUT','works carried out','works'),
    HOURS: pick('Hours','HOURS','hour'),
    WORK_STILL_TO_DO: pick('Work Still to do/Need to go back','WORK STILL TO DO/NEED TO GO BACK','Work Still to do','works still to do','to go back'),
    WORKED_WITH: pick('Worked with','WORKED WITH','worked with'),
    CERTIFICATE_SHARED: pick('Certificate Shared','CERTIFICATE_SHARED','certificate shared'),
    MATERIALS: pick('Materials','MATERIALS'),
    SUPPLIER_O: supplierFromO || pick('First Supplier','Supplier O','Supplier (1)','supplier1','supplier'),
    EXTRAS: pick('VARIATIONS - Extras (works outside scope of works / specification of job)','VARIATIONS - Extras','Extras','EXTRAS','variations','works outside scope'),
    HOURS_EXTRA: pick('Hours Extra','HOURS EXTRA','Hours Extra','hours extra'),
    EXTRA_MATERIALS: pick('Extra Matierials','Extra Materials','EXTRA MATERIALS','Extra Materials','Extra Matierials'),
    SUPPLIER: supplierFromS || pick('Supplier','SUPPLIER','Supplier for Extras','Supplier for extras','supplier for extras')
  };
}

function createDocx(data) {
  if (!fs.existsSync(TEMPLATE_PATH)) throw new Error(`Template not found at: ${TEMPLATE_PATH}`);
  fs.ensureDirSync(OUTPUT_DIR);
  const content = fs.readFileSync(TEMPLATE_PATH, 'binary');
  const zip = new PizZip(content);
  const doc = new Docxtemplater(zip, { paragraphLoop: true, linebreaks: true });
  doc.render(data);

  const rawDateStr = data.DATE || '';
  const parsed = parseDateDeterministic(rawDateStr);
  const parsedDate = parsed.date;
  const parsedKind = parsed.kind;

  let formattedDate = null;
  if (parsedDate) formattedDate = formatDateForFilename(parsedDate);
  if (!formattedDate) {
    const normalized = normalizeRawDateForFilename(rawDateStr);
    if (normalized) formattedDate = normalized;
  }
  const safeDate = formattedDate || 'nodate';

  if (safeDate !== 'nodate' && !FINAL_DATE_REGEX.test(safeDate)) {
    console.error('FATAL: computed filename date does not match D-M-YYYY:', safeDate);
    console.error(`[filename-debug] raw="${rawDateStr}" kind="${parsedKind}" parsed="${parsedDate ? parsedDate.toISOString() : ''}" formatted="${formattedDate || ''}"`);
    throw new Error('Invalid filename date format - aborting');
  }

  const safeName = safeFilename(data.NAME || 'NONAME');
  const safeJobNumber = safeFilename(data.JOB_NO || 'NOJOBNO');
  const outputFile = path.join(OUTPUT_DIR, `${safeDate}-${safeName}-${safeJobNumber}.docx`);
  const buf = doc.getZip().generate({ type: 'nodebuffer' });

  console.log(`[filename-debug] raw="${rawDateStr}" kind="${parsedKind || ''}" parsed="${parsedDate ? parsedDate.toISOString() : ''}" formatted="${formattedDate || ''}" final="${safeDate}" output="${outputFile}"`);

  if (fs.existsSync(outputFile) && !FORCE) {
    console.log('Skipping existing file:', outputFile);
    return { skipped: true, path: outputFile };
  }

  if (DRY) {
    console.log('[dry-run] would write:', outputFile);
    return { skipped: false, path: outputFile, dry: true };
  }

  fs.writeFileSync(outputFile, buf);
  return { skipped: false, path: outputFile };
}

(async () => {
  try {
    let rows = await getRows();
    if (!rows.length) { console.log('No rows in sheet!'); return; }
    console.log(`Total rows read from sheet: ${rows.length}`);

    if (DAYS > 0) {
      const now = Date.now();
      const cutoff = now - (DAYS * 24 * 60 * 60 * 1000);
      const kept = [];
      for (const r of rows) {
        const mapped = mapSheetRowToTemplateFields(r);
        const dateStr = mapped.DATE || r['Date'] || r['_created_at'] || '';
        const parsed = parseDateDeterministic(dateStr);
        const chosenDate = parsed.date;
        const keep = chosenDate && chosenDate.getTime() >= cutoff;
        if (DEBUG) {
          console.log('--- row debug ---');
          console.log('raw DATE field:', dateStr);
          console.log('parsed kind/date:', parsed.kind, parsed.date ? parsed.date.toISOString() : '');
          console.log('kept (within last', DAYS, 'days)?', keep);
        }
        if (keep) kept.push(r);
      }
      rows = kept;
      console.log(`After --days ${DAYS} filter: ${rows.length} rows`);
    }

    if (LAST_N > 0) rows = rows.slice(-LAST_N);

    let generatedCount = 0;
    for (const [i, row] of rows.entries()) {
      const tplData = mapSheetRowToTemplateFields(row);
      const out = createDocx(tplData);
      if (out && !out.skipped) generatedCount++;
    }

    console.log(`\nSuccess! Total worksheets generated: ${generatedCount}`);
  } catch (err) {
    console.error('\nFailed:', err && err.message ? err.message : err);
    process.exit(1);
  } finally {
    removeLock();
  }
})();