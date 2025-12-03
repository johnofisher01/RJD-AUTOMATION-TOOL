
const fs = require('fs-extra');
const path = require('path');
const PizZip = require('pizzip');
const Docxtemplater = require('docxtemplater');

/**
 * fillTemplate.js
 *
 * Exports an async function(main) that accepts a `mapped` object (from poller).
 * If run directly (CLI), it reads JSON from stdin and invokes main(mapped).
 *
 * Behavior:
 * - Uses server/templates/worksheet_template_10.docx as the template.
 * - Produces server/output/worksheet_<JOB_NO>_<DATE>.docx
 * - Does NOT run on require(), only when called.
 */

const TEMPLATE_NAME = 'worksheet_template_10.docx';
const TEMPLATE_PATH = path.join(__dirname, 'templates', TEMPLATE_NAME);
const OUTPUT_DIR = path.join(__dirname, 'output');

async function main(mapped = {}) {
  try {
    // Ensure template exists
    if (!fs.existsSync(TEMPLATE_PATH)) {
      throw new Error(`Template not found: ${TEMPLATE_PATH}`);
    }

    // Provide sensible defaults and map incoming keys (case-insensitive common keys)
    const get = (keys, fallback = '') => {
      for (const k of keys) {
        if (mapped[k] !== undefined && mapped[k] !== null && String(mapped[k]).length) return mapped[k];
        const low = k.toLowerCase();
        const up = k.toUpperCase();
        if (mapped[low] !== undefined && mapped[low] !== null && String(mapped[low]).length) return mapped[low];
        if (mapped[up] !== undefined && mapped[up] !== null && String(mapped[up]).length) return mapped[up];
      }
      return fallback;
    };

    // Compose the data object used by the docxtemplater template
    const data = {
      NAME: String(get(['NAME', 'name', 'Full Name', 'FullName'], 'No name provided')),
      DATE: String(get(['DATE', 'date', '_created_at'], new Date().toISOString().slice(0, 10))),
      JOB_NO: String(get(['JOB_NO', 'job_no', 'Job No', 'jobNo', 'JOBNO'], 'unknown')),
      CUSTOMER: String(get(['CUSTOMER', 'customer', 'Client'], '')),
      ADDRESS: String(get(['ADDRESS', 'address'], '')),
      WORKS_CARRIED_OUT: String(get(['WORKS_CARRIED_OUT', 'works', 'works_carried_out'], '')),
      HOURS: String(get(['HOURS', 'hours'], '')),
      WORK_STILL_TO_DO: String(get(['WORK_STILL_TO_DO', 'to_do', 'work_still_to_do'], '')),
      WORKED_WITH: String(get(['WORKED_WITH', 'worked_with'], '')),
      CERTIFICATE_SHARED: String(get(['CERTIFICATE_SHARED', 'certificate_shared'], '')),
      MATERIALS: String(get(['MATERIALS', 'materials'], '')),
      EXTRAS: String(get(['EXTRAS', 'extras'], '')),
      HOURS_EXTRA: String(get(['HOURS_EXTRA', 'hours_extra'], '')),
      EXTRA_MATERIALS: String(get(['EXTRA_MATERIALS', 'extra_materials'], '')),
      SUPPLIER: String(get(['SUPPLIER', 'supplier'], '')),
      _raw_mapped: mapped
    };

    // Read template and render
    const content = fs.readFileSync(TEMPLATE_PATH, 'binary');
    const zip = new PizZip(content);
    const doc = new Docxtemplater(zip, {
      paragraphLoop: true,
      linebreaks: true,
    });

    doc.render(data);

    // Ensure output folder exists
    fs.ensureDirSync(OUTPUT_DIR);

    // Create output file — sanitize filename parts
    const safe = (s = '') => String(s).replace(/[\/\\:?<>|"]/g, '-').replace(/\s+/g, '_');
    const outputFile = path.join(OUTPUT_DIR, `worksheet_${safe(data.JOB_NO)}_${safe(data.DATE)}.docx`);
    const buf = doc.getZip().generate({ type: 'nodebuffer' });

    fs.writeFileSync(outputFile, buf);
    console.log('Generated:', outputFile);

    return outputFile;
  } catch (err) {
    console.error('fillTemplate error:', err && err.message ? err.message : err);
    throw err;
  }
}

module.exports = main;

// If invoked directly, read JSON from stdin and run main()
if (require.main === module) {
  (async () => {
    try {
      let input = '';
      process.stdin.setEncoding('utf8');
      for await (const chunk of process.stdin) input += chunk;
      const mapped = input ? JSON.parse(input) : {};
      await main(mapped);
      process.exit(0);
    } catch (err) {
      console.error(err && err.stack ? err.stack : err);
      process.exit(1);
    }
  })();
}
