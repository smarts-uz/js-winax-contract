#!/usr/bin/env node

import fs from "fs";
import path from "path";
import { execSync } from "child_process";

// === Get arguments ===
// Usage: node run-all.js <actrecoFile> <sourceExcelFile>
const [,, actrecoFile, sourceExcelPath] = process.argv;

if (!actrecoFile || !sourceExcelPath) {
  console.error("❌ Usage: node all.js <actrecoFilePath> <sourceExcelFilePath>");
  process.exit(1);
}

// Get parent folder of the .actreco file
const rootDir = path.dirname(actrecoFile);

// Folders to ignore
const ignoredFolders = ["@ Weak", "@ Bads", "ALL", "App"];

// === Recursively find all ALL.contract files ===
function findAllContractFiles(dir) {
  let results = [];
  const entries = fs.readdirSync(dir, { withFileTypes: true });

  for (const entry of entries) {
    const fullPath = path.join(dir, entry.name);

    if (entry.isDirectory()) {
      if (ignoredFolders.includes(entry.name)) {
        console.log(`⚠️ Ignoring folder: ${entry.name}`);
        continue;
      }
      results = results.concat(findAllContractFiles(fullPath));
    } else if (entry.isFile() && entry.name === "ALL.contract") {
      results.push(fullPath);
    }
  }

  return results;
}

const contractFiles = findAllContractFiles(rootDir);

if (contractFiles.length === 0) {
  console.log("No ALL.contract files found.");
  process.exit(0);
}

console.log(`Found ${contractFiles.length} ALL.contract files.`);

// === Run processing script for each file ===
for (const contractFile of contractFiles) {
  try {
    const cmd = `node index.js "${contractFile}" "${sourceExcelPath}"`;
    console.log(`\n=== Running: ${cmd}`);
    execSync(cmd, { stdio: "inherit" });
  } catch (err) {
    console.error(`❌ Error processing ${contractFile}:`, err.message);
  }
}


setTimeout(() => {
  console.log('Exiting...');
  process.exit(0);
}, 2000);
