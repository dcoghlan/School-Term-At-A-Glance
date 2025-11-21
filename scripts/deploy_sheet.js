const shell = require('shelljs');
const fs = require('fs');
const { google } = require('googleapis');
const path = require('path');

// 0. Read version from from JS file
const codeContent = fs.readFileSync(path.join(__dirname, '../src/Code.js'), 'utf8');
// Looks for: const VERSION = "1.0.0" or '1.0.0'
const versionMatch = codeContent.match(/const\s+VERSION\s*=\s*['"]([^'"]+)['"]/);
const version = versionMatch ? versionMatch[1] : '0.0.0';

console.log('Version from Code.js: ' + version);

// 1. Setup Auth from Secrets
const credentials = JSON.parse(process.env.CLASP_CREDENTIALS);

// Create the .clasprc.json file that clasp expects in the home directory
const homeDir = require('os').homedir();
fs.writeFileSync(path.join(homeDir, '.clasprc.json'), JSON.stringify(credentials));

console.log('✅ Credentials configured.');


// 2. Create the Sheet and Script using Clasp
// --type sheets creates a standalone sheet with a bound script
console.log('🚀 Creating new Google Sheet and Script...');
const createCmd = 'npx @google/clasp create --type sheets --title "School-Term-At-A-Glance v' + version + '" --rootDir ./src';

if (shell.exec(createCmd).code !== 0) {
  console.error('❌ Error: Failed to create project.');
  process.exit(1);
}

// Locate and Configure .clasp.json ---
// Sometimes clasp creates the config file inside the source folder when rootDir is specified.
// We need to move it to the project root and ensure it points to src.

let configPath = '.clasp.json';
const srcConfigPath = path.join('src', '.clasp.json');

// 1. Move file if it is in the wrong place
if (!fs.existsSync(configPath) && fs.existsSync(srcConfigPath)) {
    console.log('⚠️  .clasp.json created inside src folder. Moving to root...');
    shell.mv(srcConfigPath, configPath);
} else if (!fs.existsSync(configPath)) {
    console.error('❌ Error: .clasp.json not found in root or src folder.');
    process.exit(1);
}

// 2. Update the config to ensure rootDir is correct
const config = JSON.parse(fs.readFileSync(configPath, 'utf8'));
if (config.rootDir !== './src') {
    console.log('🔧 Updating rootDir in .clasp.json to "./src"...');
    config.rootDir = './src';
    fs.writeFileSync(configPath, JSON.stringify(config, null, 2));
}

// 3. Push the code into the new sheet project
console.log('📤 Pushing code...');
if (shell.exec('npx @google/clasp push --force').code !== 0) {
  console.error('❌ Error: Failed to push code.');
  process.exit(1);
}

// 4. Retrieve the Sheet ID
// Clasp generates a .clasp.json file containing the scriptId and parentId (Sheet ID)
if (!fs.existsSync('.clasp.json')) {
  console.error('❌ Error: .clasp.json not found after creation.');
  process.exit(1);
}

const claspConfig = JSON.parse(fs.readFileSync('.clasp.json', 'utf8'));
// In a bound script, parentId is an array, usually [ "SPREADSHEET_ID" ]
const sheetId = claspConfig.parentId ? claspConfig.parentId[0] : null;

if (!sheetId) {
  console.error('❌ Error: Could not determine Sheet ID from .clasp.json');
  process.exit(1);
}

console.log(`📍 Sheet created with ID: ${sheetId}`);

// 6. Configure Drive (Move Folder + Set Permissions)
async function configureDriveFile() {
    const auth = new google.auth.OAuth2(
      credentials.oauth2ClientSettings.clientId,
      credentials.oauth2ClientSettings.clientSecret
    );
  
    auth.setCredentials(credentials.token);
  
    const drive = google.drive({ version: 'v3', auth });
    const folderId = process.env.DRIVE_FOLDER_ID;
  
    try {
      // --- A. Move to specific folder (if ID provided) ---
      if (folderId) {
          console.log(`📂 Moving Sheet to folder ID: ${folderId}...`);
          
          // 1. Get current parents (usually 'root')
          const file = await drive.files.get({
              fileId: sheetId,
              fields: 'parents'
          });
          
          const previousParents = file.data.parents ? file.data.parents.join(',') : '';
  
          // 2. Move: Add new parent, remove old parent
          await drive.files.update({
              fileId: sheetId,
              addParents: folderId,
              removeParents: previousParents,
              fields: 'id, parents'
          });
          console.log('✅ File moved successfully.');
      }
  
      // --- B. Set Permissions ---
      console.log('🔓 Setting permissions to "Anyone with link can view"...');
      
      await drive.permissions.create({
        fileId: sheetId,
        requestBody: {
          role: 'reader',
          type: 'anyone',
        },
      });
  
      console.log('✅ Permissions updated successfully!');
      console.log(`🔗 Link to Sheet: https://docs.google.com/spreadsheets/d/${sheetId}/copy`);
      
    } catch (error) {
      console.error('❌ Error configuring Drive file:', error.message);
      process.exit(1);
    }
  }
  
  configureDriveFile();