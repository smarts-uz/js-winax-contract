import fs from "fs";
import path from "path";
import yaml from "js-yaml";
import winax from "winax";
import dotenv from "dotenv";
import { getNumberWordOnly, getRussianMonthName } from "../utils/number-to-text.js";
import { exists, mkdirIfNotExists, getBaseName, getDirName } from "../utils/file-utils.js";
import { PDF_FORMAT_CODE } from "../config/constants.js";

// Load .env
dotenv.config();

/* ============================
   Helper: Get Company Initials
   ============================ */
function getComNameInitials(name) {
  if (!name || typeof name !== "string") return "";
  let cleaned = name.replace(/[«»"']/g, "").trim();
  return cleaned
    .split(/\s+/)
    .map(word => word[0] ? word[0].toUpperCase() : "")
    .join("");
}

/* ============================
   Function: Generate Contract Number
   ============================ */
function generateContractNumFromFormat(data) {
  const prefix = data["ContractPrefix"] && data["ContractPrefix"].trim()
    ? data["ContractPrefix"].trim()
    : (process.env.ContractPrefix || "RC");

  const format = data["ContractFormat"] && data["ContractFormat"].trim()
    ? data["ContractFormat"].trim()
    : (process.env.ContractFormat || "RC-{Year}-{Month}-{Day}");

  const values = {
    ContractPrefix: prefix,
    Prefix: prefix,
    ComName: getComNameInitials(data["ComName"]),
    CName: getComNameInitials(data["ComName"]),
    Day: String(data["Day"] || "").padStart(2, "0"),
    Month: String(data["Month"] || "").padStart(2, "0"),
    Year: String(data["Year"] || "")
  };

  return format.replace(
    /\{(ContractPrefix|Prefix|ComName|CName|Day|Month|Year)\}/g,
    (_, key) => values[key] || ""
  );
}

/* ============================
   Function: Generate Contract Files (DOCX and PDF)
   ============================ */
function generateContractFiles(data, ymlFilePath, templatePath) {
  // Contract number from YAML → fallback to generated → fallback to ENV
  const contractNum = (data["ContractNumber"] && String(data["ContractNumber"]).trim() !== "")
    ? String(data["ContractNumber"]).trim()
    : generateContractNumFromFormat(data);

  const area = data["Area"];
  const company = data["MyName"] && data["MyName"].includes("SMART TEAMS") ? "LLC" : "Person";

  // Start Word application
  const word = new winax.Object("Word.Application");
  word.Visible = false;

  // Prepare paths
  const docPath = path.resolve(templatePath);
  const docBaseName = getBaseName(docPath, ".docx");
  const ymlFolder = getDirName(ymlFilePath);
  const contractFolder = path.join(ymlFolder, "Contract");
  mkdirIfNotExists(contractFolder);
  const contractNumFolder = path.join(contractFolder, contractNum);
  mkdirIfNotExists(contractNumFolder);

  // Output files
  const outputDocxPath = path.join(contractNumFolder, `${contractNum}, ${area}-kv, ${company}, ${docBaseName}.docx`);
  const outputPdfPath = path.join(contractNumFolder, `${contractNum}, ${area}-kv, ${company}, ${docBaseName}.pdf`);

  // Open template
  const doc = word.Documents.Open(docPath);

  // Prepare replacement
  const find = doc.Content.Find;
  find.ClearFormatting();

  // Collect placeholders
  const docContent = doc.Content.Text;
  const regex = /\[([A-Za-z0-9_]+)\]/g;
  let match;
  const placeholders = new Set();
  while ((match = regex.exec(docContent)) !== null) {
    placeholders.add(match[1]);
  }

  // Replace placeholders
  for (const placeholder of placeholders) {
    let replacementText = "";

    switch (true) {
      case (placeholder === "ContractNum"):
        replacementText = contractNum;
        break;
      case (placeholder === "MonthText"):
        replacementText = getRussianMonthName(Number(data["Month"]));
        break;
      case (placeholder.endsWith("Text")): {
        const key = placeholder.replace(/Text$/, "");
        const value = data[key];
        replacementText = value !== undefined && value !== null
          ? getNumberWordOnly(Number(value))
          : "";
        break;
      }
      case (placeholder.endsWith("Phone")): {
        const keyPhone = placeholder.replace(/Phone$/, "");
        const valuePhone = data[keyPhone + "Phone"];
        replacementText = valuePhone
          ? String(valuePhone).replace(/^998/, "+998")
          : "";
        break;
      }
      default:
        replacementText = data[placeholder] !== undefined && data[placeholder] !== null
          ? data[placeholder].toString()
          : "";
    }

    // Execute Word find/replace
    find.Text = `[${placeholder}]`;
    find.Replacement.ClearFormatting();
    find.Replacement.Text = replacementText;

    find.Execute(
      find.Text,
      false, false, false, false, false,
      true, 1, false,
      find.Replacement.Text,
      2 // wdReplaceAll
    );
  }

  // Save as DOCX and PDF
  doc.SaveAs(outputDocxPath);
  doc.SaveAs(outputPdfPath, PDF_FORMAT_CODE);

  // Close Word
  doc.Close(false);
  word.Quit();

  return { outputDocxPath, outputPdfPath };
}

/* ============================
   Main Execution (YAML contract number preview)
   ============================ */
// Only run if called directly (not imported)
if (import.meta.url === process.argv[1] || import.meta.url === `file://${process.argv[1]}`) {
  const yamlFilePath = path.resolve(process.argv[2] || "./ALL.contract");

  if (!fs.existsSync(yamlFilePath)) {
    console.error(`❌ YAML file not found: ${yamlFilePath}`);
    process.exit(1);
  }

  try {
    // Load YAML directly
    const content = fs.readFileSync(yamlFilePath, "utf8");
    const data = yaml.load(content);

    // Determine contract number with fallback
    const contractNum = data["ContractNumber"]
      ? String(data["ContractNumber"]).replace(/\s+/g, "")
      : generateContractNumFromFormat(data);

    console.log("✅ Contract Number:", contractNum);
  } catch (err) {
    console.error("❌ Error processing YAML:", err.message);
    process.exit(1);
  }
}

export { generateContractFiles };
