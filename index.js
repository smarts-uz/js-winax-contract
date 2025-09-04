import fs from 'fs';
import yaml from 'js-yaml';
import path from 'path';
import { fileURLToPath } from 'url';
import { generateContractFiles } from './src/services/contract-generator.js';
import { exec } from 'child_process';

const __filename = fileURLToPath(import.meta.url);
const __dirname = path.dirname(__filename);




if (process.argv.length < 4) {
    console.error('Usage: node index.js <data.yml> <template.docx> [isOpen=false]');
    process.exit(1);
}
const ymlFilePath = process.argv[2];
let templatePath = process.argv[3];
let isOpen = process.argv[4] || 'false';


if (!fs.existsSync(templatePath)) {
    console.warn(`Template file not found at ${templatePath}. Using default template.`);
    templatePath = "d:\\Humans\\Building\\Rentalls\\Contract\\Projects\\Rentals 282.docx";
}


if (!fs.existsSync(ymlFilePath)) {
    console.error(`YAML file not found: ${ymlFilePath}`);
    process.exit(1);
}
if (!fs.existsSync(templatePath)) {
    console.error(`Template file not found: ${templatePath}`);
    process.exit(1);
}



const yamlOptions = {
    schema: yaml.JSON_SCHEMA,
    onWarning: (e) => { console.warn('YAML ogohlantirishi:', e); }
};

const ymlRaw = fs.readFileSync(ymlFilePath, 'utf8');


const ymlPatched = ymlRaw.split('\n').map(line => {
    if (!line.includes(':') || line.trim().startsWith('#')) return line;

    const idx = line.indexOf(':');
    const key = line.slice(0, idx);
    let value = line.slice(idx + 1).trim();

    if (
        (value.startsWith('"') && value.endsWith('"')) ||
        (value.startsWith("'") && value.endsWith("'")) ||
        value === 'null' || value === 'true' || value === 'false'
    ) {
        return line;
    }

    if (value === '' || value.startsWith('#')) {
        return line;
    }

    if (/^\d{1,}$/.test(value)) {
        return `${key}: "${value}"`;
    }

    if (/[",]/.test(value)) {
        // Ichki ikki tirnoqlarni ekranga chiqarish uchun almashtiramiz
        const safeValue = value.replace(/"/g, '\\"');
        return `${key}: "${safeValue}"`;
    }


    return line;
}).join('\n');

const data = yaml.load(ymlPatched, yamlOptions);
// console.log(data);

const { outputDocxPath, outputPdfPath } = generateContractFiles(data, ymlFilePath, templatePath);

console.log('✅ Word yaratildi:', outputDocxPath);
console.log('✅ PDF yaratildi:', outputPdfPath);


if (isOpen === 'true' || isOpen === '1') {
    if (process.platform === 'win32') {
        exec(`start "" "${outputDocxPath}"`);
    } else if (process.platform === 'darwin') {
        exec(`open "${outputDocxPath}"`);
    } else {
        console.warn('Platformani qo\'llab-quvvatlanmaydi:', process.platform);
    }
}



setTimeout(() => {
    console.log('Exiting...');
    process.exit(0);
}, 2000);
