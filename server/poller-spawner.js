const path = require('path');
const fs = require('fs');
const { spawn } = require('child_process');
const os = require('os');

/**
 * spawnPoller(options)
 * - options: { sheetId, sheetTab, keepAttached=false }
 * - When packaged: uses process.execPath with ELECTRON_RUN_AS_NODE=1
 * - When dev: uses 'node'
 * - Uses app.getPath('userData') for logs if available; otherwise os.tmpdir()
 */
function spawnPoller(options = {}) {
  const { sheetId, sheetTab, keepAttached = false } = options;

  // locate script: prefer project server/fillFromSheet.js, fallback to resources unpacked path
  let scriptPath = path.join(__dirname, 'fillFromSheet.js');

  // If running from a packaged app, __dirname will be inside resources/app.asar.
  // Many packagers create app.asar.unpacked for large or native files under resourcesPath.
  try {
    const resourcesPath = (process.resourcesPath) ? process.resourcesPath : null;
    if (resourcesPath) {
      // Check the likely asar-unpacked location
      const unpacked = path.join(resourcesPath, 'app.asar.unpacked', 'server', 'fillFromSheet.js');
      if (fs.existsSync(unpacked)) {
        scriptPath = unpacked;
      } else {
        // Also check resources/server (if you copied files via extraResources)
        const resourcesServer = path.join(resourcesPath, 'server', 'fillFromSheet.js');
        if (fs.existsSync(resourcesServer)) scriptPath = resourcesServer;
      }
    }
  } catch (e) {
    // ignore and use default scriptPath
  }

  if (!fs.existsSync(scriptPath)) {
    console.warn('[poller-spawner] worker script not found:', scriptPath);
    return null;
  }

  // Choose runner
  const isPackaged = !!(process.execPath && process.resourcesPath && process.pkg === undefined);
  const runner = (process.env.NODE_ENV !== 'development' && isPackaged) ? process.execPath : 'node';

  // Child env
  const childEnv = Object.assign({}, process.env);
  if (isPackaged) childEnv.ELECTRON_RUN_AS_NODE = '1';
  if (sheetId) childEnv.GOOGLE_SHEET_ID = sheetId;
  if (sheetTab) childEnv.SHEET_TAB = sheetTab;

  // Create log files in userData or tmp for debugging
  let logDir = os.tmpdir();
  try {
    const { app } = require('electron');
    if (app && typeof app.getPath === 'function') {
      logDir = path.join(app.getPath('userData'), 'poller-logs');
      fs.mkdirSync(logDir, { recursive: true });
    }
  } catch (e) {
    // ignore if electron not available at require time
  }
  const outLog = path.join(logDir, `poller-out-${Date.now()}.log`);
  const errLog = path.join(logDir, `poller-err-${Date.now()}.log`);

  // Build args: pass script path and flags (worker expects node script)
  const args = [scriptPath, '--poll']; // child script may accept --poll or adapt as needed

  // spawn detached so it continues while app runs; if keepAttached true, don't detach (useful for debugging)
  const spawnOptions = {
    env: childEnv,
    detached: !keepAttached,
    stdio: keepAttached ? 'inherit' : ['ignore', 'ignore', 'ignore']
  };

  try {
    // If detached and we want logs, open file streams and redirect stdio
    let child;
    if (!keepAttached) {
      // spawn with stdio ignored, but also write logs by spawning a small wrapper is complex;
      // simpler: spawn detached and let child create its own logs (fillFromSheet should log to userData).
      child = spawn(runner, args, spawnOptions);
      child.unref();
    } else {
      // attached - useful for debugging
      child = spawn(runner, args, spawnOptions);
    }

    // Best-effort: write a small starter log
    try {
      fs.appendFileSync(path.join(logDir, 'poller-starter.log'),
        `[poller] started pid=${child.pid} runner=${runner} script=${scriptPath}\n`);
    } catch (_) {}

    return child.pid;
  } catch (err) {
    console.error('[poller-spawner] failed to spawn poller:', err && err.message);
    return null;
  }
}

module.exports = { spawnPoller };