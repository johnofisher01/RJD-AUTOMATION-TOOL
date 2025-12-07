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
// configurable sheet/tab name (exact match required)
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

    console.log(`Using spreadsheetId=${SHEET_ID} range="${SHEET_RANGE}"`);

    const response = await sheets.spreadsheets.values.get({
        spreadsheetId: SHEET_ID,
        range: SHEET_RANGE,
    });

    const values = response.data.values || [];
    const [header, ...rows] = values;
    if (!header) return [];

    console.log('Header row detected:', header);

    const mappedRows = rows.map(row =>
        header.reduce((obj, key, i) => {
            obj[key] = row[i] || '';
            return obj;
        }, {})
    );

    if (mappedRows.length) {
        console.log('Sample mapped row (first):', mappedRows[0]);
        console.log('Mapped to template fields (sample):', mapSheetRowToTemplateFields(mappedRows[0]));
    }

    return mappedRows;
}

/**
 * Robust mapping (same as fillfromsheet.js) — normalized keys + address join
 */
function mapSheetRowToTemplateFields(row) {
  const normalize = (s) => String(s || '').toLowerCase().replace(/[^a-z0-9]/g, '');
  const normRow = {};
  for (const k of Object.keys(row || {})) normRow[normalize(k)] = row[k];

  const pick = (...names) => {
    for (const n of names) {
      const v = normRow[normalize(n)];
      if (v !== undefined && v !== null && String(v).trim() !== '') return String(v).trim();
    }
    return '';
  };

  const addrParts = [
    pick('Address - Street Address', 'Address - Street Address', 'street address', 'address'),
    pick('Address - Street Address Line 2', 'Address - Street Address Line 2', 'street address line 2', 'address line 2'),
    pick('Address - City', 'Address - City', 'city'),
    pick('Address - State / Province', 'Address - State / Province', 'state', 'province'),
    pick('Address - Postal / Zip Code', 'Address - Postal / Zip Code', 'postal', 'zip', 'postcode')
  ].filter(Boolean);
  const ADDRESS = addrParts.join(', ');

  return {
    NAME: pick('Name', 'Full Name', 'full name', 'name'),
    DATE: pick('Date', 'DATE', '_created_at', 'date'),
    JOB_NO: pick('Job Number', 'JOB NUMBER', 'Job No', 'job number', 'job_no'),
    CUSTOMER: pick('Customer', 'CLIENT', 'client'),
    ADDRESS,
    WORKS_CARRIED_OUT: pick('Works carried out', 'WORKS CARRIED OUT', 'works carried out', 'works'),
    HOURS: pick('Hours', 'HOURS', 'hour'),
    WORK_STILL_TO_DO: pick('Work Still to do/Need to go back', 'WORK STILL TO DO/NEED TO GO BACK', 'Work Still to do', 'works still to do', 'to go back'),
    WORKED_WITH: pick('Worked with', 'WORKED WITH', 'worked with'),
    CERTIFICATE_SHARED: pick('Certificate Shared', 'CERTIFICATE_SHARED', 'certificate shared'),
    MATERIALS: pick('Materials', 'MATERIALS'),
    EXTRAS: pick('VARIATIONS - Extras (works outside scope of works / specification of job)', 'VARIATIONS - Extras', 'Extras', 'EXTRAS', 'variations', 'works outside scope'),
    HOURS_EXTRA: pick('Hours Extra', 'HOURS EXTRA', 'Hours Extra', 'hours extra'),
    EXTRA_MATERIALS: pick('Extra Matierials', 'Extra Materials', 'EXTRA MATERIALS', 'Extra Materials', 'Extra Matierials'),
    SUPPLIER: pick('Supplier', 'SUPPLIER', 'Supplier for Extras', 'Supplier for extras', 'supplier for extras')
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