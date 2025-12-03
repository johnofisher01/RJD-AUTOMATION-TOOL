const fs = require('fs');
const path = require('path');

(async () => {
  try {
    const helper = require('./append-to-sheets');

    // Read .env from project root
    const env = fs.readFileSync(path.join(__dirname, '..', '.env'), 'utf8');
    const id = (env.match(/^GOOGLE_SHEET_ID=(.*)$/m) || [])[1];
    const cred = (env.match(/^GOOGLE_CREDENTIALS_PATH=(.*)$/m) || [])[1] || 'google-sheets-creds.json';

    const row = [
      'Manual Test Name',
      'manual@example.com',
      '2025-12-12 00:00:00', // date as string
      'Work still test value',
      new Date().toISOString()
    ];

    console.log('Using spreadsheetId=', id, 'credentials=', cred);
    console.log('Row to append:', JSON.stringify(row));

    await helper.appendRow(id, row, cred);
    console.log('Manual append succeeded');
  } catch (e) {
    console.error('Manual append failed:', e && e.message ? e.message : e);
    process.exit(1);
  }
})();