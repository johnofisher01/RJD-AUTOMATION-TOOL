import React, { useState } from 'react';
import './App.css';

const USERS = {
  "john.dorey@rjdorey.co.uk": "John1234!",
  "claire.broadhurst@rjdorey.co.uk": "Claire1234!"
};
const APP_TITLE = "RJDorey Worksheet Automater";

function App() {
  // Auth state
  const [signedIn, setSignedIn] = useState(false);
  const [userEmail, setUserEmail] = useState("");
  const [emailInput, setEmailInput] = useState("");
  const [pwInput, setPwInput] = useState("");
  const [pwShow, setPwShow] = useState(false);
  const [authError, setAuthError] = useState("");

  // Main app state
  const [error, setError] = useState('');
  const [output, setOutput] = useState('');
  const [loading, setLoading] = useState(false);
  const [outputFolder, setOutputFolder] = useState(() => localStorage.getItem('outputFolder') || '');

  // Login logic
  const handleLogin = (e) => {
    e.preventDefault();
    const emailLower = emailInput.trim().toLowerCase();
    if (!USERS[emailLower]) {
      setAuthError("Unrecognized email address.");
      return;
    }
    if (USERS[emailLower] !== pwInput) {
      setAuthError("Incorrect password.");
      return;
    }
    setUserEmail(emailLower);
    setSignedIn(true);
    setAuthError('');
    setPwInput('');
  };

  const handleChooseFolder = async () => {
    setError("");
    setOutput("");
    if (window.electronAPI && window.electronAPI.chooseOutputFolder) {
      const folder = await window.electronAPI.chooseOutputFolder();
      if (folder) {
        setOutputFolder(folder);
        localStorage.setItem('outputFolder', folder);
        setError('');
      }
    } else {
      setError('Could not choose folder: Electron API not available!');
    }
  };

  const handleGenerate = async () => {
    setError('');
    setOutput('');
    setLoading(true);
    try {
      if (!window.electronAPI || !window.electronAPI.generateWorksheets) {
        throw new Error('Electron API not found. Are you running inside Electron?');
      }
      if (!outputFolder) {
        throw new Error('Please select an output folder before generating.');
      }
      const result = await window.electronAPI.generateWorksheets(outputFolder);
      setOutput(result);
    } catch (err) {
      setError(err.message || 'Unknown error');
    }
    setLoading(false);
  };

  const handleFolderClick = () => {
    if (!outputFolder) {
      setError('No output folder selected!');
      return;
    }
    if (window.electronAPI && window.electronAPI.openOneDriveFolder) {
      window.electronAPI.openOneDriveFolder();
    } else if (window.electronAPI && window.electronAPI.openPath) {
      window.electronAPI.openPath(outputFolder);
    } else {
      alert('Could not open folder: Electron API not available!');
    }
  };

  // -------- UI --------
  if (!signedIn) {
    return (
      <div className="login-bg">
        <form className="login-panel" onSubmit={handleLogin}>
          <img
            src="rjdorey-logo.png"
            alt="RJDorey Logo"
            className="login-logo"
          />
          <div className="login-title">{APP_TITLE}</div>
          <div className="login-subtitle">Sign in to continue</div>
          <input
            className="login-input"
            type="email"
            required
            autoFocus
            placeholder="Email address"
            value={emailInput}
            autoComplete="username"
            onChange={e => setEmailInput(e.target.value)}
          />
          <div className="pw-row">
            <input
              className="login-input"
              type={pwShow ? "text" : "password"}
              required
              placeholder="Password"
              value={pwInput}
              autoComplete="current-password"
              onChange={e => setPwInput(e.target.value)}
              style={{ marginRight: 6 }}
            />
           
          </div>
          {authError && <div className="login-error">{authError}</div>}
          <button className="login-submit" type="submit">
            Sign In
          </button>
          <div className="login-footer">
            <span style={{ opacity: 0.65 }}>Access restricted to:<br /><b>john.dorey@rjdorey.co.uk</b> &amp; <b>claire.broadhurst@rjdorey.co.uk</b></span>
          </div>
        </form>
      </div>
    );
  }

  return (
    <div className="app-bg">
      <div className="generator-container">
        <div className="main-title">{APP_TITLE}</div>
        <div className="main-byline">
          Signed in as <span className="signed-email">{userEmail}</span>
        </div>

        <div className="main-section">
          <button
            className="choose-folder-btn"
            onClick={handleChooseFolder}
            disabled={loading}
            title="Select the location to save worksheets"
          >
            <span className="btn-icon" role="img" aria-label="folder">📁</span>
            Choose Output Folder
          </button>
        </div>
        {outputFolder && (
          <div className="output-folder-label">
            <span className="btn-icon" role="img" aria-label="location">📂</span>
            <b style={{wordBreak: "break-all"}}>{outputFolder}</b>
          </div>
        )}

        <button
          className="generate-btn"
          onClick={handleGenerate}
          disabled={loading || !outputFolder}
        >
          <span className="btn-icon" role="img" aria-label="doc-folder">📝</span>
          {loading ? "Generating..." : "Generate Worksheets"}
        </button>

        <div className="folder-btn-under">
          <button
            className="folder-btn"
            title="Go to Worksheets Folder"
            onClick={handleFolderClick}
            disabled={!outputFolder}
          >
            <span className="btn-icon" role="img" aria-label="worksheets">📁</span>
            <span className="folder-label">Worksheets</span>
          </button>
        </div>

        {(error || output) && (
          <div className={`result-box${error ? " error" : ""}`}>
            {error ? `Error: ${error}` : output}
          </div>
        )}
      </div>
      <footer className="footer">
        &copy; {new Date().getFullYear()} RJ Dorey &mdash; All rights reserved.
      </footer>
    </div>
  );
}

export default App;