'use strict';
const fs = require('fs');
const { google } = require('googleapis');

async function getAuth(credentialsPath) {
  if (!credentialsPath || !fs.existsSync(credentialsPath)) {
    throw new Error('Google credentials file not found: ' + credentialsPath);
  }
  return new google.auth.GoogleAuth({
    keyFile: credentialsPath,
    scopes: ['https://www.googleapis.com/auth/spreadsheets'],
  });
}

async function appendRow(spreadsheetId, rowValues = [], credentialsPath = 'google-sheets-creds.json') {
  if (!spreadsheetId) throw new Error('spreadsheetId is required');
  const auth = await getAuth(credentialsPath);
  const client = await auth.getClient();
  const sheets = google.sheets({ version: 'v4', auth: client });

  const meta = await sheets.spreadsheets.get({ spreadsheetId });
  const sheetsList = (meta.data && meta.data.sheets) || [];
  const firstSheetTitle = (sheetsList[0] && sheetsList[0].properties && sheetsList[0].properties.title) || 'Sheet1';
  const range = `${firstSheetTitle}!A1`;

  await sheets.spreadsheets.values.append({
    spreadsheetId,
    range,
    valueInputOption: 'USER_ENTERED',
    insertDataOption: 'INSERT_ROWS',
    requestBody: {
      values: [rowValues],
    },
  });
}

module.exports = { appendRow };
