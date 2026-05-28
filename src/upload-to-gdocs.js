const { google } = require('googleapis');
const { drive_v3 } = google;
const fs = require('fs-extra');
const path = require('path');
const config = require('./config');

async function uploadToGoogleDocs() {
  console.log('[GOOGLE DRIVE] Uploading document to Google Docs...');

  try {
    const auth = await google.auth.getClient({
      credentials: JSON.parse(fs.readFileSync(config.GOOGLE_DRIVE.SERVICE_ACCOUNT_KEY)),
      scopes: ['https://www.googleapis.com/auth/drive.file']
    });

    const gdrive = google.drive({ version: 'v3', auth });

    const fileContent = fs.readFileSync(config.DOCX_PATH);
    const fileName = path.basename(config.DOCX_PATH, '.docx');

    const fileMetadata = {
      name: `${fileName}.gdoc`,
      mimeType: 'application/vnd.openxmlformats-officedocument.wordprocessingml.document',
      parents: config.GOOGLE_DRIVE.FOLDER_ID ? [config.GOOGLE_DRIVE.FOLDER_ID] : undefined
    };

    const response = await gdrive.files.create({
      resource: fileMetadata,
      media: { mimeType: 'application/vnd.openxmlformats-officedocument.wordprocessingml.document', content: fileContent }
    });

    // Convert to Google Docs format
    await gdrive.files.convert(fileMetadata.id, {
      requestBody: { mimeType: 'application/vnd.google-apps.document' }
    });

    console.log(`[SUCCESS] Google Doc uploaded. URL: https://docs.google.com/document/d/${response.data.id}`);
    console.log(`[INFO] Access ID: ${response.data.id}`);

  } catch (error) {
    console.error('[ERROR] Google Drive upload failed:', error.message);
    if (error.response) {
      console.error('[DEBUG]', JSON.stringify(error.response.data, null, 2));
    }
    throw error;
  }
}

module.exports = { uploadToGoogleDocs };
if (require.main === module) {
  uploadToGoogleDocs().catch(err => {
    console.error('Upload failed:', err);
    process.exit(1);
  });
}
