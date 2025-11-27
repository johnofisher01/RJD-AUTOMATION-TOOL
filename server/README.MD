Worksheet Generator — Windows build & release cheat‑sheet

Purpose
This Word document contains a clear, step‑by‑step guide you can follow to produce a fresh Windows build (unpacked app + installer) for Worksheet Generator. It includes the exact Git Bash commands to run, when to run the client build, troubleshooting steps for common Windows file-lock issues, and a short checklist for handling secrets (.env and google-sheets-creds.json).

How to use this document
- Copy the whole content below into a new Microsoft Word document (File → New → Blank document), or save this page as text and import it to Word.
- Follow the numbered command blocks in Git Bash (one command per line). Run Git Bash "Run as administrator" if you see access errors.
- Do not commit your google-sheets-creds.json or .env.

1) Quick summary (what to run, in order)
- If you changed frontend code (React in client/): build the frontend
  cd client
  npm install
  npm run build
  cd ..
- Clean old output:
  rm -rf dist-out
- Build Windows artifacts:
  npx electron-builder --win --x64 --config.directories.output=dist-out 2>&1 | tee build-output.log
- Verify packaging:
  npx asar list "dist-out/win-unpacked/resources/app.asar" | grep "client/build/index.html" || true
  ls -la dist-out/win-unpacked/resources/.env dist-out/win-unpacked/resources/google-sheets-creds.json
- Quick local smoke test:
  cd dist-out/win-unpacked
  "./Worksheet Generator.exe" 2>&1 | tee ../runtime.log
  (Use the app UI → Choose Output Folder → Generate → Ctrl+C to stop)
  tail -n 200 ../runtime.log

2) Detailed step‑by‑step (copy/paste safe commands, Git Bash)
A — Optional: bump version
  npm version patch

B — Ensure secrets exist in project root (do NOT commit them)
  ls -la .env google-sheets-creds.json

C — Build client (only when UI/frontend changed)
  cd client
  npm install
  npm run build
  cd ..

D — Clean packaging output (start clean)
  rm -rf dist-out

E — Create Windows build (packager)
  npx electron-builder --win --x64 --config.directories.output=dist-out 2>&1 | tee build-output.log

F — Verify renderer files are inside app.asar
  npx asar list "dist-out/win-unpacked/resources/app.asar" | grep "client/build/index.html" || true

G — Verify .env and creds are copied outside app.asar
  ls -la dist-out/win-unpacked/resources/.env dist-out/win-unpacked/resources/google-sheets-creds.json

H — Smoke test the unpacked exe locally
  cd dist-out/win-unpacked
  "./Worksheet Generator.exe" 2>&1 | tee ../runtime.log
  (Use the UI → Choose Output Folder → Generate → when finished press Ctrl+C)
  tail -n 200 ../runtime.log

I — Locate installer & portable exe for distribution
  ls -la dist-out/*.exe dist-out/*.exe.blockmap || true

3) When to run npm run build (React client)
- Run npm run build inside client/ only when you changed code in:
  - client/src
  - client/public
  - static assets referenced by the UI
- If you only changed main.js, server/*, or templates, you can skip the client build and go straight to steps D→E.

4) Managing secrets (correct & safe packaging)
- package.json build MUST:
  - Exclude .env and google-sheets-creds.json from app.asar (so they are not buried inside app.asar).
  - Use extraResources to copy them into dist-out/win-unpacked/resources at pack time.
- Example build settings (package.json.build):
  "files": [
    "main.js",
    "server/**/*",
    "client/build/**",
    "server/electron-preload.js",
    "package.json",
    "!**/.env",
    "!google-sheets-creds.json"
  ],
  "extraResources": [
    { "from": ".env", "to": ".env" },
    { "from": "google-sheets-creds.json", "to": "google-sheets-creds.json" }
  ]
- Do NOT commit google-sheets-creds.json or .env. Add to .gitignore:
  echo ".env" >> .gitignore
  echo "google-sheets-creds.json" >> .gitignore

5) Common Windows problems & fixes (file locked / Access is denied)
Symptoms:
- electron-builder fails with "The process cannot access the file because it is being used by another process" or "Access is denied".

Simple fix (recommended):
- Reboot Windows, open Git Bash "Run as administrator", then repeat steps D→E.

Find & kill the lock (no reboot):
- Download Sysinternals handle.exe and search who holds the file:
  curl -L -o /tmp/handle.zip "https://download.sysinternals.com/files/Handle.zip"
  unzip -o /tmp/handle.zip -d /tmp/handle
  /tmp/handle/handle.exe -accepteula "C:\Users\youruser\Desktop\RJD-AUTOMATION-TOOL\dist-out\win-unpacked\resources\app.asar"
- Kill reported PIDs:
  taskkill /PID <pid> /F
- Then:
  rm -rf "dist-out/win-unpacked"
  npx electron-builder --win --x64 --config.directories.output=dist-out 2>&1 | tee build-output.log

If rm -rf fails with "Device or resource busy" and handle.exe shows explorer.exe as a holder, restart explorer:
  taskkill /F /IM explorer.exe
  start explorer.exe

6) Troubleshooting runtime errors
- ERR_FILE_NOT_FOUND when launching the exe:
  - Cause: client/build not included in app.asar.
  - Fix: build client (npm run build) and ensure package.json.build includes client/build/** in "files", then rm -rf dist-out and rebuild.

- dotenv loaded from app.asar/.env but values missing:
  - Cause: .env was packaged inside app.asar with empty values.
  - Fix: Exclude .env from asar and add it to extraResources (see section 4). Rebuild.

- "Missing GOOGLE_SHEET_ID in .env!" at runtime:
  - Ensure the .env that lives in dist-out/win-unpacked/resources/.env contains the correct GOOGLE_SHEET_ID line:
    GOOGLE_SHEET_ID=<your-sheet-id>
  - Or copy your local .env into the resources folder for testing:
    cp .env dist-out/win-unpacked/resources/.env
    cp google-sheets-creds.json dist-out/win-unpacked/resources/google-sheets-creds.json

7) Quick recovery (if you accidentally lost .env or creds)
- If files were accidentally removed from project root but exist in dist-out:
  cp dist-out/win-unpacked/resources/.env ./.env
  cp dist-out/win-unpacked/resources/google-sheets-creds.json ./

- If google-sheets-creds.json is lost permanently, re-download from Google Cloud Console (Service Accounts → Keys → Create key → JSON).

8) Automation idea (optional build script: build-windows.sh)
Create a small shell script in project root to run the common sequence:
  #!/usr/bin/env bash
  set -e
  cd client
  npm ci
  npm run build
  cd ..
  rm -rf dist-out
  npx electron-builder --win --x64 --config.directories.output=dist-out

Make executable:
  chmod +x build-windows.sh
Run:
  ./build-windows.sh 2>&1 | tee build-output.log

9) Final checklist before distribution
- [ ] client build is up-to-date (if UI changed)
- [ ] .env contains correct GOOGLE_SHEET_ID
- [ ] google-sheets-creds.json is present locally (not committed)
- [ ] package.json.build contains client/build/**, excludes .env and creds from asar, and extraResources copies them
- [ ] rm -rf dist-out run before packaging
- [ ] Build succeeded and dist-out contains installer (.exe)
- [ ] Smoke test passed (Worksheet Generator ran and produced outputs)
- [ ] Add .env and creds to .gitignore (do not commit)

10) Where outputs land
- Unpacked app & runtime: dist-out/win-unpacked/
- Installers: dist-out/Worksheet Generator Setup <version>.exe (NSIS) and dist-out/Worksheet Generator <version>.exe (portable)

Support notes / security
- Never check service account JSON into source control.
- Consider moving secret-dependent actions to a server (where you can restrict keys safely) if you plan to distribute widely.
- For signing installers in production, obtain a code signing certificate and configure the signing credentials (signtool) in your CI or local build environment.

If you want this delivered as a real .docx file:
- Copy this content into a Word document, adjust the title fonts and spacing as you like, then save as .docx.
- If you'd like, I can produce a ready-made build-windows.sh script or a README.md file you can add to the repo.

End of document.