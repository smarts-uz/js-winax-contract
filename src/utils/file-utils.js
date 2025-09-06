import fs from 'fs';
import path from 'path';
import { execSync } from "child_process";


function exists(filePath) {
  return fs.existsSync(filePath);
}

function mkdirIfNotExists(dirPath) {
  if (!fs.existsSync(dirPath)) {
    fs.mkdirSync(dirPath, { recursive: true });
  }
}

function getBaseName(filePath, ext) {
  return path.basename(filePath, ext);
}

function getDirName(filePath) {
  return path.dirname(path.resolve(filePath));
}


function openFileDialog(initialDir = "D:\\Projects") {
  const psScript = `
Add-Type -AssemblyName System.Windows.Forms
$dlg = New-Object System.Windows.Forms.OpenFileDialog
$dlg.InitialDirectory = '${initialDir}'
$dlg.Filter = 'Word Documents (*.doc;*.docx)|*.doc;*.docx|All Files (*.*)|*.*'
if ($dlg.ShowDialog() -eq [System.Windows.Forms.DialogResult]::OK) {
    Write-Output $dlg.FileName
}
`;

  try {
    // Inline PowerShell script with -NoProfile to avoid user profile issues
    const filePath = execSync(
      `powershell -NoProfile -Command "${psScript.replace(/\n/g, ';')}"`,
      { encoding: "utf8" }
    ).trim();

    if (filePath) {
      console.log("Selected file:", filePath);
      return filePath;
    } else {
      console.log("No file selected.");
      return null;
    }
  } catch (err) {
    console.error("Error opening dialog:", err.message);
    return null;
  }
}

export { exists, mkdirIfNotExists, getBaseName, getDirName, openFileDialog }; 
