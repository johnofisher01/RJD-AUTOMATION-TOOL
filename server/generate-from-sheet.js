const fs = require('fs-extra');
const path = require('path');
const { google } = require('googleapis');
const PizZip = require('pizzip');
const Docxtemplater = require('docxtemplater');
const minimist = require('minimist');

// Always load .env from the project root (../.env from server/)
require('dotenv').config({ path: path.join(__dirname, '..', '.env') });

const argv = minimist(process.argv.slice(2));
const SHEET_ID = process.env.GOOGLE_SHEET_ID;
const CREDS_PATH = process.env.GOOGLE_CREDENTIALS_PATH || path.join(__dirname, 'google-sheets-creds.json');
const TEMPLATE_PATH = path.join(__dirname, 'templates', 'worksheet_template_10.docx');
const OUTPUT_DIR = argv.output ? path.resolve(argv.output) : path.join(__dirname, 'output');
// Use env var if set; fallback keeps backwards compatibility
const SHEET_RANGE = process.env.GOOGLE_SHEET_RANGE || 'Form Responses 1';

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

    // Log which sheet/range is being used (helps debug)
    console.log(`Using spreadsheetId=${SHEET_ID} range="${SHEET_RANGE}"`);

    const response = await sheets.spreadsheets.values.get({
        spreadsheetId: SHEET_ID,
        range: SHEET_RANGE,
    });

    const [header, ...rows] = response.data.values || [];
    if (!header) return [];
    return rows.map(row =>
        header.reduce((obj, key, i) => {
            obj[key] = row[i] || '';
            return obj;
        }, {})
    );
}

/**
 * Robust mapping: normalizes header names and picks by several fallback variants.
 * Ensures your column "Work Still to do/Need to go back" (and variants) are matched.
 */
function mapSheetRowToTemplateFields(row) {
  const normalize = (s) => String(s || '').toLowerCase().replace(/[^a-z0-9]/g, '');
  const normRow = {};
  for (const k of Object.keys(row || {})) normRow[normalize(k)] = row[k];

  const pick = (...names) => {
    for (const n of names) {
      const v = normRow[normalize(n)];
      if (v !== undefined && v !== null && String(v).trim() !== '') return v;
    }
    return '';
  };

  return {
    NAME: pick('NAME', 'Full Name', 'Full name'),
    DATE: pick('DATE', 'Date', '_created_at'),
    JOB_NO: pick('Job Number', 'JOB NUMBER', 'Job Number', 'JOB_NO', 'job_no'),
    CUSTOMER: pick('Customer', 'CLIENT', 'client'),
    ADDRESS: pick('Address', 'ADDRESS', 'address'),
    WORKS_CARRIED_OUT: pick('WORKS CARRIED OUT', 'Works carried out', 'works carried out', 'WORKS_CARRIED_OUT'),
    HOURS: pick('HOURS', 'Hours', 'hours'),
    WORK_STILL_TO_DO: pick(
      'Work Still to do/Need to go back',
      'WORK STILL TO DO/NEED TO GO BACK',
      'WORKS STILL TO DO',
      'Work Still to do',
      'Works still to do',
      'works still to do',
      'WORK_STILL_TO_DO',
      'work_still_to_do'
    ),
    WORKED_WITH: pick('WORKED WITH', 'Worked With', 'worked_with'),
    CERTIFICATE_SHARED: pick('Certificate Shared', 'CERTIFICATE_SHARED', 'certificate_shared'),
    MATERIALS: pick('MATERIALS', 'Materials', 'materials'),
    EXTRAS: pick('VARIATIONS - Extras (works outside scope of works / specification of job)', 'EXTRAS', 'Extras'),
    HOURS_EXTRA: pick('HOURS EXTRA', 'Hours Extra', 'hours_extra'),
    EXTRA_MATERIALS: pick('EXTRA MATERIALS:', 'Extra Materials', 'extra_materials'),
    SUPPLIER: pick('SUPPLIER', 'SUPPLIER EXTRAS', 'supplier')
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
    fs.writeFileSync(outputFile, buf);
    return outputFile;
}

(async () => {
    try {
        const rows = await getRows();
        if (!rows.length) {
            console.log('No rows in sheet!');
            return;
        }
        let generatedCount = 0;
        for (const [i, row] of rows.entries()) {
            const tplData = mapSheetRowToTemplateFields(row);
            const out = createDocx(tplData);
            console.log('Generated for row', i + 1, ':', out);
            generatedCount++;
        }
        console.log(`\nSuccess! Total worksheets generated: ${generatedCount}`);
    } catch (err) {
        console.error('\nFailed:', err && err.message ? err.message : err);
        process.exit(1);
    }
})();