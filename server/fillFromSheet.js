const fs = require('fs-extra');
const path = require('path');
const { google } = require('googleapis');
const PizZip = require('pizzip');
const Docxtemplater = require('docxtemplater');
const minimist = require('minimist');

// Load .env from project root (one level above server/)
require('dotenv').config({ path: path.join(__dirname, '..', '.env') });

const argv = minimist(process.argv.slice(2));
const SHEET_ID = process.env.GOOGLE_SHEET_ID;
const CREDS_PATH = process.env.GOOGLE_CREDENTIALS_PATH || path.join(__dirname, 'google-sheets-creds.json');
const TEMPLATE_PATH = path.join(__dirname, 'templates', 'worksheet_template_10.docx');
const OUTPUT_DIR = argv.output ? path.resolve(argv.output) : path.join(__dirname, 'output');
const SHEET_RANGE = process.env.GOOGLE_SHEET_RANGE || 'Form Responses 1';

// CLI flags and defaults
const LAST_N = Number(argv.last || argv.lastN || argv.max || 0);
const PRUNE_N = Number(argv.prune || 0);
const FORCE = !!argv.force;
const DRY = !!argv.dry || !!argv['dry-run'];
const DEBUG = !!argv.debug;
// Default to 7 days if not provided
const DAYS = Number(argv.days || argv['since-days'] || process.env.DEFAULT_DAYS || 7);

function safeFilename(str) {
  return String(str || '').replace(/[\/\\?%*:|"<>]/g, '-');
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

  console.log('Header row detected:', header);

  // produce array of objects and attach raw __row array for index-based picks
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
 * Safe parseDateCandidates
 * - returns { date: Date|null, candidates: [{kind,date}], chosenKind }
 * - validates numeric parts so JS Date overflow cannot produce misleading dates
 */
function parseDateCandidates(s) {
  const out = { date: null, candidates: [], chosenKind: null };
  if (!s) return out;
  const str = String(s).trim();

  // Native parse (may be ambiguous)
  const native = new Date(str);
  if (!Number.isNaN(native.getTime())) out.candidates.push({ kind: 'native', date: native });

  // Numeric pattern e.g. 12/15/2025 or 08-12-2025
  const parts = str.match(/^(\d{1,2})[\/\-\.\s](\d{1,2})[\/\-\.\s](\d{2,4})$/);
  if (parts) {
    const p1 = Number(parts[1]), p2 = Number(parts[2]), p3 = Number(parts[3]);
    const year = p3 < 100 ? (2000 + p3) : p3;

    // DMY candidate (day = p1, month = p2) — only if month/day valid and construction matches inputs
    if (p2 >= 1 && p2 <= 12 && p1 >= 1 && p1 <= 31) {
      const dmy = new Date(year, p2 - 1, p1);
      if (dmy.getFullYear() === year && dmy.getMonth() === p2 - 1 && dmy.getDate() === p1) {
        out.candidates.push({ kind: 'dmy', date: dmy });
      }
    }

    // MDY candidate (month = p1, day = p2) — only if month/day valid and construction matches inputs
    if (p1 >= 1 && p1 <= 12 && p2 >= 1 && p2 <= 31) {
      const mdy = new Date(year, p1 - 1, p2);
      if (mdy.getFullYear() === year && mdy.getMonth() === p1 - 1 && mdy.getDate() === p2) {
        out.candidates.push({ kind: 'mdy', date: mdy });
      }
    }
  }

  // Choose preferred candidate:
  // prefer DMY if available (common UK), else prefer native if present, else first candidate
  if (out.candidates.length) {
    const dmyC = out.candidates.find(c => c.kind === 'dmy');
    if (dmyC) {
      out.date = dmyC.date;
      out.chosenKind = 'dmy';
    } else {
      const nativeC = out.candidates.find(c => c.kind === 'native');
      const chosen = nativeC ? nativeC : out.candidates[0];
      out.date = chosen.date;
      out.chosenKind = chosen.kind;
    }
  }

  return out;
}

/**
 * Map sheet row to template fields (keeps supplier O and S handling)
 */
function mapSheetRowToTemplateFields(row) {
  const normalize = (s) => String(s || '').toLowerCase().replace(/[^a-z0-9]/g, '');
  const normRow = {};
  for (const k of Object.keys(row || {})) {
    if (k === '__row') continue;
    normRow[normalize(k)] = row[k];
  }

  const pick = (...names) => {
    for (const name of names) {
      const v = normRow[normalize(name)];
      if (v !== undefined && v !== null && String(v).trim() !== '') return String(v).trim();
    }
    return '';
  };

  const pickByIndex = (idx) => {
    const arr = row && row.__row ? row.__row : [];
    if (!Array.isArray(arr)) return '';
    const v = arr[idx];
    return (v !== undefined && v !== null && String(v).trim() !== '') ? String(v).trim() : '';
  };

  const addrParts = [
    pick('Address - Street Address', 'Address - Street Address', 'street address', 'address'),
    pick('Address - Street Address Line 2', 'Address - Street Address Line 2', 'street address line 2', 'address line 2'),
    pick('Address - City', 'Address - City', 'city'),
    pick('Address - State / Province', 'Address - State / Province', 'state', 'province'),
    pick('Address - Postal / Zip Code', 'Address - Postal / Zip Code', 'postal', 'zip', 'postcode')
  ].filter(Boolean);

  const ADDRESS = addrParts.join(', ');

  // Column indexes (0-based): O = 14, S = 18
  const supplierFromO = pickByIndex(14);
  const supplierFromS = pickByIndex(18);

  return {
    NAME: pick('Name', 'Full Name', 'full name', 'name'),
    DATE: pick('Date', 'DATE', '_created_at', 'date'),
    JOB_NO: pick('Job Number', 'JOB NUMBER', 'Job No', 'job number', 'job_no'),
    CUSTOMER: pick('Customer', 'Customer Name', 'CLIENT', 'client'),
    ADDRESS,
    WORKS_CARRIED_OUT: pick('Works carried out', 'WORKS CARRIED OUT', 'works carried out', 'works'),
    HOURS: pick('Hours', 'HOURS', 'hour'),
    WORK_STILL_TO_DO: pick('Work Still to do/Need to go back', 'WORK STILL TO DO/NEED TO GO BACK', 'Work Still to do', 'works still to do', 'to go back'),
    WORKED_WITH: pick('Worked with', 'WORKED WITH', 'worked with'),
    CERTIFICATE_SHARED: pick('Certificate Shared', 'CERTIFICATE_SHARED', 'certificate shared'),
    MATERIALS: pick('Materials', 'MATERIALS'),
    SUPPLIER_O: supplierFromO || pick('First Supplier', 'Supplier O', 'Supplier (1)', 'supplier1', 'supplier'),
    EXTRAS: pick('VARIATIONS - Extras (works outside scope of works / specification of job)', 'VARIATIONS - Extras', 'Extras', 'EXTRAS', 'variations', 'works outside scope'),
    HOURS_EXTRA: pick('Hours Extra', 'HOURS EXTRA', 'Hours Extra', 'hours extra'),
    EXTRA_MATERIALS: pick('Extra Matierials', 'Extra Materials', 'EXTRA MATERIALS', 'Extra Materials', 'Extra Matierials'),
    SUPPLIER: supplierFromS || pick('Supplier', 'SUPPLIER', 'Supplier for Extras', 'Supplier for extras', 'supplier for extras')
  };
}

function createDocx(data) {
  if (!fs.existsSync(TEMPLATE_PATH)) throw new Error(`Template not found at: ${TEMPLATE_PATH}`);
  fs.ensureDirSync(OUTPUT_DIR);
  const content = fs.readFileSync(TEMPLATE_PATH, 'binary');
  const zip = new PizZip(content);
  const doc = new Docxtemplater(zip, { paragraphLoop: true, linebreaks: true });
  doc.render(data);
  const safeDate = safeFilename(data.DATE || 'nodate');
  const safeName = safeFilename(data.NAME || 'NONAME');
  const safeJobNumber = safeFilename(data.JOB_NO || 'NOJOBNO');
  const outputFile = path.join(OUTPUT_DIR, `${safeDate}-${safeName}-${safeJobNumber}.docx`);
  const buf = doc.getZip().generate({ type: 'nodebuffer' });

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

function pruneOutputDir(keepN) {
  if (!keepN || keepN <= 0) return;
  if (!fs.existsSync(OUTPUT_DIR)) return;
  const files = fs.readdirSync(OUTPUT_DIR)
    .map(name => ({ name, full: path.join(OUTPUT_DIR, name) }))
    .filter(f => f.name.toLowerCase().endsWith('.docx'))
    .map(f => ({ ...f, mtime: fs.statSync(f.full).mtimeMs }))
    .sort((a, b) => b.mtime - a.mtime);
  if (files.length <= keepN) return;
  const toDelete = files.slice(keepN);
  for (const f of toDelete) {
    try {
      if (DRY) {
        console.log('[dry-run] would prune old file:', f.full);
      } else {
        fs.unlinkSync(f.full);
        console.log('Pruned old file:', f.full);
      }
    } catch (err) {
      console.warn('Failed to delete', f.full, err.message || err);
    }
  }
}

(async () => {
  try {
    let rows = await getRows();
    if (!rows.length) {
      console.log('No rows in sheet!');
      return;
    }

    console.log(`Total rows read from sheet: ${rows.length}`);

    // Days filter (default 7)
    if (DAYS > 0) {
      const now = Date.now();
      const cutoff = now - (DAYS * 24 * 60 * 60 * 1000);
      const preFilterCount = rows.length;
      const kept = [];

      for (const r of rows) {
        const mapped = mapSheetRowToTemplateFields(r);
        const dateStr = mapped.DATE || r['Date'] || r['_created_at'] || '';
        const parsed = parseDateCandidates(dateStr);

        let chosenDate = parsed.date;
        let chosenKind = parsed.chosenKind;

        const candDMY = parsed.candidates.find(c => c.kind === 'dmy');
        const candMDY = parsed.candidates.find(c => c.kind === 'mdy');
        if (candDMY && candMDY) {
          const dmyOk = candDMY.date.getTime() >= cutoff;
          const mdyOk = candMDY.date.getTime() >= cutoff;
          if (dmyOk && !mdyOk) {
            chosenDate = candDMY.date;
            chosenKind = 'dmy';
          } else if (mdyOk && !dmyOk) {
            chosenDate = candMDY.date;
            chosenKind = 'mdy';
          }
        }

        const keep = chosenDate && chosenDate.getTime() >= cutoff;

        if (DEBUG) {
          console.log('--- row debug ---');
          console.log('raw DATE field:', dateStr);
          console.log('candidates:', parsed.candidates.map(t => ({ kind: t.kind, iso: t.date.toISOString() })));
          console.log('chosen kind:', chosenKind, chosenDate ? chosenDate.toISOString() : null);
          console.log('kept (within last', DAYS, 'days)?', keep);
        }

        if (keep) kept.push(r);
      }

      rows = kept;
      console.log(`After --days ${DAYS} filter: ${rows.length} rows (from ${preFilterCount})`);
    }

    // LAST_N slicing
    if (LAST_N > 0) {
      rows = rows.slice(-LAST_N);
      console.log(`Processing last ${LAST_N} rows (of selected ${rows.length + Math.max(0, (rows.length - LAST_N))}).`);
    } else {
      console.log(`Processing selected rows (${rows.length}).`);
    }

    let generatedCount = 0;
    for (const [i, row] of rows.entries()) {
      const tplData = mapSheetRowToTemplateFields(row);
      const out = createDocx(tplData);
      if (out && !out.skipped) {
        console.log('Generated for row', i + 1, ':', out.path);
        generatedCount++;
      }
    }

    console.log(`\nSuccess! Total worksheets generated: ${generatedCount}`);

    if (PRUNE_N > 0) {
      console.log(`Pruning output directory to keep only ${PRUNE_N} newest .docx files.`);
      pruneOutputDir(PRUNE_N);
    }
  } catch (err) {
    console.error('\nFailed:', err && err.message ? err.message : err);
    process.exit(1);
  }
})();