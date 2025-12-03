(async () => {
  try {
    const helper = require('./server/append-to-sheets');
    const env = require('fs').readFileSync('.env','utf8');
    const id = (env.match(/^GOOGLE_SHEET_ID=(.*)$/m)||[])[1];
    const cred = (env.match(/^GOOGLE_CREDENTIALS_PATH=(.*)$/m)||[])[1] || 'google-sheets-creds.json';
    console.log('Using spreadsheetId=', id, 'credentials=', cred);
    await helper.appendRow(id, ['TEST_ROW', new Date().toISOString()], cred);
    console.log('Append succeeded');
  } catch (e) {
    console.error('Append failed:', e && e.message ? e.message : e);
    process.exit(1);
  }
})();