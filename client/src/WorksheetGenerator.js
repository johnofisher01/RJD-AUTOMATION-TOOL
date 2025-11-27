import React, { useState } from "react";
import "./WorksheetGenerator.css";

export default function WorksheetGenerator() {
  const [outputFolder, setOutputFolder] = useState("");
  const [status, setStatus] = useState("");
  const [inProgress, setInProgress] = useState(false);

  const handleChooseFolder = async () => {
    setStatus("");
    if (window.electronAPI && window.electronAPI.chooseOutputFolder) {
      const folder = await window.electronAPI.chooseOutputFolder();
      if (folder) setOutputFolder(folder);
    } else {
      setStatus("Error: Electron API not available!");
    }
  };

  const handleGenerate = async () => {
    setStatus("");
    setInProgress(true);
    if (window.electronAPI && window.electronAPI.generateWorksheets) {
      try {
        const result = await window.electronAPI.generateWorksheets(outputFolder);
        setStatus(result || "Worksheets generated!");
      } catch (e) {
        setStatus("Error: " + (e.message || e));
      }
    } else {
      setStatus("Error: Electron API not available!");
    }
    setInProgress(false);
  };

  return (
    <div className="wg-container">
      <div className="wg-header">
        <span className="wg-icon">🎓</span>
        <div>
          <h1 className="wg-title">Worksheet Generator</h1>
          <h2 className="wg-subtitle">RJ Dorey Ltd</h2>
        </div>
      </div>

      <button
        className="wg-choose-folder-btn"
        onClick={handleChooseFolder}
        disabled={inProgress}
      >
        <span role="img" aria-label="folder" className="wg-folder-emoji">📁</span>
        Choose Output Folder
      </button>

      {outputFolder && (
        <div className="wg-output-path-label">
          <span role="img" aria-label="location" style={{marginRight: 6}}>📂</span>
          <span style={{fontSize: "0.97rem"}}>{outputFolder}</span>
        </div>
      )}

      <button
        className="wg-generate-btn"
        onClick={handleGenerate}
        disabled={!outputFolder || inProgress}
      >
        <span role="img" aria-label="doc-folder" className="wg-folder-emoji">📝</span>
        Generate Worksheets
      </button>

      {status && (
        <div className={status.startsWith("Error") ? "wg-error" : "wg-status"}>
          {status}
        </div>
      )}
    </div>
  );
}