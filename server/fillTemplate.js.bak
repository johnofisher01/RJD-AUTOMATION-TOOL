const fs = require('fs-extra');
const path = require('path');
const PizZip = require('pizzip');
const Docxtemplater = require('docxtemplater');

// Path setup
const TEMPLATE_NAME = 'worksheet_template_10.docx';
const TEMPLATE_PATH = path.join(__dirname, 'templates', TEMPLATE_NAME);
const OUTPUT_DIR = path.join(__dirname, 'output');

// Sample data (replace with Google Sheets row later)
const data = {
    NAME: "Jane Electrician",
    DATE: "2025-11-24",
    JOB_NO: "12345",
    CUSTOMER: "ACME Ltd",
    ADDRESS: "123 Fake Street, Somewhere",
    WORKS_CARRIED_OUT: "Installed lights & checked circuits",
    HOURS: "6",
    WORK_STILL_TO_DO: "Fit 2nd fix sockets next visit",
    WORKED_WITH: "Joe Bloggs",
    CERTIFICATE_SHARED: "Yes",
    MATERIALS: "LED lights, cable",
    EXTRAS: "Added extra socket in kitchen",
    HOURS_EXTRA: "2",
    EXTRA_MATERIALS: "1 double socket",
    SUPPLIER: "City Electrical Wholesale"
};

// Load template
const content = fs.readFileSync(TEMPLATE_PATH, 'binary');
const zip = new PizZip(content);
const doc = new Docxtemplater(zip, {
  paragraphLoop: true,
  linebreaks: true,
});

// Render (fill) template
doc.render(data);

// Ensure output folder exists
fs.ensureDirSync(OUTPUT_DIR);

// Create output file
const outputFile = path.join(OUTPUT_DIR, `worksheet_${data.JOB_NO}_${data.DATE}.docx`);
const buf = doc.getZip().generate({ type: 'nodebuffer' });

fs.writeFileSync(outputFile, buf);
console.log('Generated:', outputFile);