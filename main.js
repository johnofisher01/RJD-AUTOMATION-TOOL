const { app, BrowserWindow, ipcMain, dialog } = require('electron');
const path = require('path');
const fs = require('fs');
const { spawn } = require('child_process');
const dotenv = require('dotenv');

// --- Load .env safely: prefer a resources .env when packaged, otherwise use project .env ---
function loadDotenv() {
  try {
    if (app && app.isPackaged) {
      // When packaged, extraResources or an asar-unpacked file will be placed under process.resourcesPath
      const packagedEnv = path.join(process.resourcesPath, '.env');
      if (fs.existsSync(packagedEnv)) {
        dotenv.config({ path: packagedEnv });
        console.log(`[dotenv] loaded from packaged .env: ${packagedEnv}`);
        return;
      }
      // also try app.asar (less ideal) — if .env ended up inside app.asar
      const asarEnv = path.join(process.resourcesPath, 'app.asar', '.env');
      if (fs.existsSync(asarEnv)) {
        dotenv.config({ path: asarEnv });
        console.log(`[dotenv] loaded from app.asar .env: ${asarEnv}`);
        return;
      }
    }
  } catch (e) {
    console.warn('[dotenv] load error', e && e.message);
  }

  // Development fallback — .env next to main.js (project root)
  const devEnv = path.join(__dirname, '.env');
  dotenv.config({ path: devEnv });
  console.log(`[dotenv] loaded from dev .env: ${devEnv}`);
}
loadDotenv();
// ------------------------------------------------------------------------------

// Normalize credentials path when packaged so it's an absolute path the worker can read
if (process.env.GOOGLE_CREDENTIALS_PATH) {
  if (app && app.isPackaged && !path.isAbsolute(process.env.GOOGLE_CREDENTIALS_PATH)) {
    process.env.GOOGLE_CREDENTIALS_PATH = path.join(process.resourcesPath, process.env.GOOGLE_CREDENTIALS_PATH);
    console.log('[env] normalized GOOGLE_CREDENTIALS_PATH ->', process.env.GOOGLE_CREDENTIALS_PATH);
  } else {
    console.log('[env] GOOGLE_CREDENTIALS_PATH (as provided) ->', process.env.GOOGLE_CREDENTIALS_PATH);
  }
}

// --- Ensure cache and userData use a per-user writable folder and avoid GPU cache issues ---
app.disableHardwareAcceleration();
const userDataPath = path.join(app.getPath('appData'), 'Worksheet Generator');
try { fs.mkdirSync(userDataPath, { recursive: true }); } catch (e) {}
app.setPath('userData', userDataPath);
app.commandLine.appendSwitch('disk-cache-dir', path.join(userDataPath, 'Cache'));
// ------------------------------------------------------------------------------

function createWindow() {
  const win = new BrowserWindow({
    width: 800,
    height: 600,
    webPreferences: {
      preload: path.join(__dirname, 'server', 'electron-preload.js'),
      contextIsolation: true,
      nodeIntegration: false
    }
  });

  const startUrl = process.env.ELECTRON_START_URL || `file://${path.join(__dirname, 'client', 'build', 'index.html')}`;
  win.loadURL(startUrl);
}

app.whenReady().then(createWindow);

ipcMain.handle('chooseOutputFolder', async () => {
  const result = await dialog.showOpenDialog({ properties: ['openDirectory'] });
  if (result.canceled || !result.filePaths[0]) return null;
  return result.filePaths[0];
});

ipcMain.handle('generateWorksheets', async (event, outputFolder) => {
  return new Promise((resolve) => {
    let scriptPath = path.join(__dirname, 'server', 'fillFromSheet.js');

    if (process.resourcesPath && !fs.existsSync(scriptPath)) {
      const unpacked = path.join(process.resourcesPath, 'app.asar.unpacked', 'server', 'fillFromSheet.js');
      if (fs.existsSync(unpacked)) scriptPath = unpacked;
    }

    if (!fs.existsSync(scriptPath)) {
      const msg = `Error: worker script not found: ${scriptPath}`;
      console.error(msg);
      return resolve(msg);
    }

    // Use explicit child environment so the worker definitely sees the needed values
    const childEnv = Object.assign({}, process.env);
    if (app.isPackaged) childEnv.ELECTRON_RUN_AS_NODE = '1';

    // Make sure required env vars exist for the child process (use normalized creds path)
    childEnv.GOOGLE_SHEET_ID = process.env.GOOGLE_SHEET_ID || '';
    childEnv.GOOGLE_CREDENTIALS_PATH = process.env.GOOGLE_CREDENTIALS_PATH ||
      path.join(process.resourcesPath || __dirname, 'google-sheets-creds.json');

    const runner = app.isPackaged ? process.execPath : 'node';

    // Debug snapshot so we can see exactly what the main will pass to the child
    console.log('Spawning child:', {
      runner,
      scriptPath,
      outputFolder,
      appIsPackaged: app.isPackaged,
      childEnvSnapshot: {
        GOOGLE_SHEET_ID: childEnv.GOOGLE_SHEET_ID ? '[REDACTED]' : '(empty)',
        GOOGLE_CREDENTIALS_PATH: childEnv.GOOGLE_CREDENTIALS_PATH
      }
    });

    const child = spawn(runner, [scriptPath, '--output', outputFolder], { env: childEnv, stdio: ['ignore', 'pipe', 'pipe'] });

    console.log(`Spawned child pid=${child.pid}`);

    let output = '';
    child.stdout.on('data', (data) => {
      const s = data.toString();
      output += s;
      console.log('[child stdout]', s.trim());
    });
    child.stderr.on('data', (data) => {
      const s = data.toString();
      output += '\n' + s;
      console.error('[child stderr]', s.trim());
    });
    child.on('close', (code, signal) => {
      console.log(`Child exited code=${code} signal=${signal}`);
      resolve(output || `Done! (exit ${code})`);
    });
    child.on('error', (err) => {
      console.error('Child process error', err);
      resolve(`Error: ${err.message}`);
    });
  });
});

app.on('window-all-closed', () => {
  if (process.platform !== 'darwin') app.quit();
});