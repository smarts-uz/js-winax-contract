import winax from 'winax';
import path from 'path';
import { getNumberWordOnly, getRussianMonthName } from '../utils/number-to-text.js';
import { exists, mkdirIfNotExists, getBaseName, getDirName } from '../utils/file-utils.js';
import { PDF_FORMAT_CODE } from '../config/constants.js';

/**
 * Extracts and returns the uppercase initials from a company name.
 * Removes special characters and splits by whitespace.
 * @param {string} comName - The company name.
 * @returns {string} - The initials in uppercase.
 */
function getComNameInitials(comName) {
    if (!comName || typeof comName !== 'string') return '';
    // Remove special characters and trim
    let cleaned = comName.replace(/[«»"']/g, '').trim();
    let words = cleaned.split(/\s+/);
    // Take the first letter of each word, uppercase
    return words.map(w => w[0] ? w[0].toUpperCase() : '').join('');
}

/**
 * Generates a contract number in the format: RC-<Initials>-<DD>-<MM>-<YYYY>
 * @param {Object} data - The contract data object.
 * @returns {string} - The generated contract number.
 */
function generateContractNum(data) {
    const day = String(data['Day']).padStart(2, '0');
    const month = String(data['Month']).padStart(2, '0');
    const year = String(data['Year']);
    const comName = data['ComName'];
    const initials = getComNameInitials(comName);
    return `RC-${initials}-${day}-${month}-${year}`;
}

/**
 * Generates contract files (Word and PDF) by replacing placeholders in a template.
 * @param {Object} data - The contract data object.
 * @param {string} ymlFilePath - Path to the YAML data file.
 * @param {string} templatePath - Path to the Word template file.
 * @returns {Object} - Paths to the generated DOCX and PDF files.
 */
function generateContractFiles(data, ymlFilePath, templatePath) {
    // Generate contract number
    const contractNum = generateContractNum(data);

    // Start Word application (invisible)
    const word = new winax.Object('Word.Application');
    word.Visible = false;

    // Prepare paths and folder structure
    const docPath = path.resolve(templatePath);
    const docBaseName = getBaseName(docPath, '.docx');
    const ymlFolder = getDirName(ymlFilePath);
    const contractFolder = path.join(ymlFolder, 'Contract');
    mkdirIfNotExists(contractFolder);
    const contractNumFolder = path.join(contractFolder, contractNum);
    mkdirIfNotExists(contractNumFolder);

    // Output file paths
    const outputDocxPath = path.join(contractNumFolder, `${docBaseName}.docx`);
    const outputPdfPath = path.join(contractNumFolder, `${docBaseName}.pdf`);

    // Open the Word template document
    const doc = word.Documents.Open(docPath);

    // Prepare to find and replace placeholders
    const find = doc.Content.Find;
    find.ClearFormatting();

    // Extract all placeholders in the format [Placeholder]
    const docContent = doc.Content.Text;
    const regex = /\[([A-Za-z0-9_]+)\]/g;
    let match;
    const placeholders = new Set();
    while ((match = regex.exec(docContent)) !== null) {
        placeholders.add(match[1]);
    }

    // Replace each placeholder with the appropriate value
    for (const placeholder of placeholders) {
        let replacementText = '';

        switch (true) {
            case (placeholder === 'ContractNum'):
                // Maxsus placeholder - shartnoma raqami
                replacementText = contractNum;
                break;
            case (placeholder === 'MonthText'):
                // Maxsus placeholder - rus tilida oy nomi
                {
                    const monthNumber = data['Month'];
                    replacementText = getRussianMonthName(Number(monthNumber));
                }
                break;
            case (placeholder.endsWith('Text')):
                // 'Text' bilan tugaydigan placeholderlar sonni so'zga aylantiradi
                {
                    const key = placeholder.replace(/Text$/, '');
                    const value = data[key];
                    if (value !== undefined && value !== null) {
                        replacementText = getNumberWordOnly(Number(value));
                    } else {
                        replacementText = '';
                    }
                }
                break;
            case (placeholder.endsWith('Phone')):
                // 'Phone' bilan tugaydigan placeholderlar sonni so'zga aylantiradi
                {
                    const keyPhone = placeholder.replace(/Phone$/, '');
                    const valuePhone = data[keyPhone + 'Phone'];
                    if (valuePhone !== undefined && valuePhone !== null && valuePhone !== "") {
                        // valuePhone ni stringga aylantirib, replace ishlatamiz
                        replacementText = String(valuePhone).replace(/^998/, '+998');
                    } else {
                        replacementText = '';
                    }
                }
                break;
            default:
                // Default: data dan qiymatni oladi yoki bo'sh string
                if (data[placeholder] !== undefined && data[placeholder] !== null) {
                    // Fix: Only call toString if not null/undefined
                    replacementText = data[placeholder].toString();
                } else {
                    replacementText = '';
                }
        }

        // Set up the find/replace operation
        find.Text = `[${placeholder}]`;
        find.Replacement.ClearFormatting();
        find.Replacement.Text = replacementText;

        // Execute the replacement throughout the document
        find.Execute(
            find.Text,
            false, false, false, false, false,
            true, 1, false,
            find.Replacement.Text,
            2 // wdReplaceAll
        );
    }

    // Save the filled document as DOCX and PDF
    doc.SaveAs(outputDocxPath);
    doc.SaveAs(outputPdfPath, PDF_FORMAT_CODE);

    // Close the document and quit Word
    doc.Close(false);
    word.Quit();

    // Return the output file paths
    return { outputDocxPath, outputPdfPath };
}

export { generateContractFiles };