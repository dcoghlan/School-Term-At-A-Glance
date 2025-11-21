const shell = require('shelljs');
const fs = require('fs');
const { google } = require('googleapis');
const path = require('path');

// Alternative Step 1: Read from JS file
const codeContent = fs.readFileSync(path.join(__dirname, '../src/Code.js'), 'utf8');
// Looks for: const VERSION = "1.0.0" or '1.0.0'
const versionMatch = codeContent.match(/const\s+VERSION\s*=\s*['"]([^'"]+)['"]/);
const version = versionMatch ? versionMatch[1] : '0.0.0';

// 1. Setup Auth from Secrets
const credentials = JSON.parse(process.env.CLASP_CREDENTIALS);

// Create the .clasprc.json file that clasp expects in the home directory
const homeDir = require('os').homedir();
fs.writeFileSync(path.join(homeDir, '.clasprc.json'), JSON.stringify(credentials));

console.log('✅ Credentials configured.');

process.exit(0);
