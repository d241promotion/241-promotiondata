const express = require('express');
const bodyParser = require('body-parser');
const ExcelJS = require('exceljs');
const path = require('path');
const fs = require('fs').promises;
const { google } = require('googleapis');
const disk = require('diskusage');

const app = express();
const PORT = process.env.PORT || 3000;
const LOCAL_EXCEL_FILE = path.join(__dirname, 'customers.xlsx');
const GOOGLE_DRIVE_FOLDER_ID = '1l4e6cq0LaFS2IFkJlWKLFJ_CVIEqPqTK';

let isFileLocked = false;

const auth = new google.auth.GoogleAuth({
  credentials: JSON.parse(process.env.GOOGLE_SERVICE_ACCOUNT),
  scopes: ['https://www.googleapis.com/auth/drive'],
});
const drive = google.drive({ version: 'v3', auth });

app.use(bodyParser.json());
app.use(express.static(__dirname));

// Initialize the Excel workbook
async function initializeExcel() {
  const workbook = new ExcelJS.Workbook();
  const sheet = workbook.addWorksheet('Customers');
  sheet.columns = [
    { header: 'Name', key: 'name', width: 20 },
    { header: 'Email', key: 'email', width: 30 },
    { header: 'Phone', key: 'phone', width: 15 },
  ];
  return workbook;
}

// Check disk space and file permissions
async function checkDiskSpaceAndPermissions(filePath) {
  try {
    const { available, total } = await disk.check(path.dirname(filePath));
    const availableMB = available / (1024 * 1024);
    console.log(`Available disk space: ${availableMB.toFixed(2)} MB`);
    if (availableMB < 10) {
      throw new Error(`Insufficient disk space: ${availableMB.toFixed(2)} MB available`);
    }

    try {
      await fs.access(filePath, fs.constants.R_OK | fs.constants.W_OK);
      console.log(`File ${filePath} is readable and writable`);
    } catch (error) {
      console.log(`File ${filePath} not accessible, attempting to fix permissions`);
      await fs.chmod(filePath, 0o666); // Fixed typo: removed 'durdurur'
      console.log(`Permissions fixed for ${filePath}`);
    }
  } catch (error) {
    console.error('Disk space or permission check failed:', error.message);
    throw error;
  }
}

// Load the local Excel file or initialize a new one
async function loadLocalExcel() {
  let workbook;
  try {
    await checkDiskSpaceAndPermissions(LOCAL_EXCEL_FILE);
    workbook = new ExcelJS.Workbook();
    await workbook.xlsx.readFile(LOCAL_EXCEL_FILE);
    console.log('Loaded local Excel file:', LOCAL_EXCEL_FILE);
  } catch (error) {
    console.log('Local Excel file not found, inaccessible, or corrupted, initializing new one:', error.message);
    workbook = await initializeExcel();
    try {
      await checkDiskSpaceAndPermissions(LOCAL_EXCEL_FILE);
      await workbook.xlsx.writeFile(LOCAL_EXCEL_FILE);
      console.log('Initialized new Excel file:', LOCAL_EXCEL_FILE);
    } catch (writeError) {
      console.error('Failed to initialize new Excel file:', writeError.message, writeError.stack);
      throw writeError;
    }
  }
  return workbook;
}

// Upload the local Excel file to Google Drive
async function uploadToGoogleDrive() {
  isFileLocked = true;
  try {
    // Verify the local file exists and is not empty
    const stats = await fs.stat(LOCAL_EXCEL_FILE);
    if (stats.size === 0) {
      throw new Error('Local Excel file is empty');
    }
    console.log('Local Excel file size:', stats.size, 'bytes');

    // Verify Google Drive authentication
    const authClient = await auth.getClient();
    console.log('Google Drive authentication successful:', !!authClient);

    await checkDiskSpaceAndPermissions(LOCAL_EXCEL_FILE);
    const existingFiles = await drive.files.list({
      q: `'${GOOGLE_DRIVE_FOLDER_ID}' in parents and name = 'customers.xlsx' and trashed = false`,
      fields: 'files(id, name)',
    });

    const fileMetadata = {
      name: 'customers.xlsx',
      parents: [GOOGLE_DRIVE_FOLDER_ID],
    };

    const media = {
      mimeType: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
      body: require('fs').createReadStream(LOCAL_EXCEL_FILE),
    };

    let file;
    if (existingFiles.data.files.length > 0) {
      const fileId = existingFiles.data.files[0].id;
      file = await drive.files.update({
        fileId: fileId,
        media: media,
        fields: 'id',
      });
      console.log('Updated file in Google Drive, ID:', file.data.id);
    } else {
      file = await drive.files.create({
        resource: fileMetadata,
        media: media,
        fields: 'id',
      });
      console.log('Created new file in Google Drive, ID:', file.data.id);
    }
  } catch (error) {
    console.error('Failed to upload to Google Drive:', error.message, error.stack);
    throw error;
  } finally {
    isFileLocked = false;
  }
}

// Periodic sync with Google Drive (every 5 minutes)
function startGoogleDriveSync() {
  setInterval(async () => {
    try {
      console.log('Starting periodic sync with Google Drive...');
      await uploadToGoogleDrive();
      console.log('Periodic sync completed.');
    } catch (error) {
      console.error('Periodic sync failed:', error.message, error.stack);
    }
  }, 5 * 60 * 1000);
}

// Download the Excel file from Google Drive on server start
async function initializeFromGoogleDrive() {
  try {
    const response = await drive.files.list({
      q: `'${GOOGLE_DRIVE_FOLDER_ID}' in parents and name = 'customers.xlsx' and trashed = false`,
      fields: 'files(id)',
    });

    if (response.data.files.length > 0) {
      const fileId = response.data.files[0].id;
      const file = await drive.files.get(
        { fileId, alt: 'media' },
        { responseType: 'stream' }
      );

      await new Promise((resolve, reject) => {
        const dest = require('fs').createWriteStream(LOCAL_EXCEL_FILE);
        file.data
          .on('error', reject)
          .pipe(dest)
          .on('error', reject)
          .on('finish', resolve);
      });
      console.log('Downloaded Excel file from Google Drive to local:', LOCAL_EXCEL_FILE);
    } else {
      console.log('No Excel file found in Google Drive, initializing new one locally.');
      const workbook = await initializeExcel();
      await checkDiskSpaceAndPermissions(LOCAL_EXCEL_FILE);
      await workbook.xlsx.writeFile(LOCAL_EXCEL_FILE);
    }
  } catch (error) {
    console.error('Error initializing from Google Drive:', error.message, error.stack);
    const workbook = await initializeExcel();
    await checkDiskSpaceAndPermissions(LOCAL_EXCEL_FILE);
    await workbook.xlsx.writeFile(LOCAL_EXCEL_FILE);
  }
}

// Handle form submission
app.post('/submit', async (req, res) => {
  const { name, email, phone } = req.body;

  console.log('Received initial submission:', { name, email, phone });

  if (!name || !email || !phone) {
    console.log('Validation failed: Missing required fields');
    return res.status(400).json({ success: false, error: 'Missing required fields' });
  }

  const emailRegex = /^[a-zA-Z0-9._%+-]+@[a-zA-Z0-9.-]+\.(com|org|net|edu|gov|co|io|me|biz)$/i;
  if (!emailRegex.test(email)) {
    console.log('Validation failed: Invalid email address');
    return res.status(400).json({ success: false, error: 'Please check the email address' });
  }

  const domain = email.split('@')[1].toLowerCase();
  const commonMisspellings = ['gmil.com', 'gail.com', 'gmai.com', 'gnail.com'];
  if (commonMisspellings.includes(domain)) {
    console.log(`Detected common email domain misspelling: ${domain}`);
    return res.status(400).json({ success: false, error: 'Please check the email address' });
  }

  if (!/^\d{10}$/.test(phone)) {
    console.log('Validation failed: Invalid phone number');
    return res.status(400).json({ success: false, error: 'Invalid phone number (10 digits required)' });
  }

  if (isFileLocked) {
    console.log('File is locked, please try again later');
    return res.status(503).json({ success: false, error: 'Server is busy, please try again later.' });
  }

  isFileLocked = true;
  try {
    const workbook = await loadLocalExcel();
    const sheet = workbook.getWorksheet('Customers');

    let emailExists = false;
    let phoneExists = false;
    sheet.eachRow((row, rowNumber) => {
      if (rowNumber === 1) return;
      const existingEmail = row.getCell('email').value;
      const existingPhone = row.getCell('phone').value;
      if (existingEmail && existingEmail.toLowerCase() === email.toLowerCase()) {
        emailExists = true;
      }
      if (existingPhone && existingPhone.toString() === phone.toString()) {
        phoneExists = true;
      }
    });

    if (emailExists || phoneExists) {
      console.log(`Duplicate found - Email exists: ${emailExists}, Phone exists: ${phoneExists}`);
      return res.status(400).json({ success: false, error: 'Details already exist' });
    }

    const newRow = sheet.addRow([name, email, phone]);
    newRow.commit();

    await checkDiskSpaceAndPermissions(LOCAL_EXCEL_FILE);
    await workbook.xlsx.writeFile(LOCAL_EXCEL_FILE);
    console.log('Data saved to local Excel file:', LOCAL_EXCEL_FILE);

    try {
      await uploadToGoogleDrive();
    } catch (syncError) {
      console.error('Google Drive sync failed after local write:', syncError.message, syncError.stack);
      res.status(200).json({ 
        success: true, 
        name, 
        warning: 'Data saved locally, but failed to sync to Google Drive. Please contact support.'
      });
      return;
    }

    res.status(200).json({ success: true, name });
  } catch (error) {
    console.error('Failed to save to local Excel:', error.message, error.stack);
    if (error.message.includes('Insufficient disk space')) {
      res.status(500).json({ success: false, error: 'Server disk space is full. Please contact support.' });
    } else if (error.message.includes('Permission denied')) {
      res.status(500).json({ success: false, error: 'File permission error. Please contact support.' });
    } else if (error.message.includes('Corrupt')) {
      console.log('Excel file appears to be corrupted, attempting to recreate...');
      try {
        const workbook = await initializeExcel();
        await checkDiskSpaceAndPermissions(LOCAL_EXCEL_FILE);
        await workbook.xlsx.writeFile(LOCAL_EXCEL_FILE);
        console.log('Recreated Excel file:', LOCAL_EXCEL_FILE);
        res.status(503).json({ success: false, error: 'File was corrupted, please try again.' });
      } catch (recreateError) {
        console.error('Failed to recreate Excel file:', recreateError.message, recreateError.stack);
        res.status(500).json({ success: false, error: 'Unable to save your submission. Please try again later.' });
      }
    } else {
      res.status(500).json({ success: false, error: 'Unable to save your submission. Please try again later.' });
    }
  } finally {
    isFileLocked = false;
  }
});

// Handle file download
app.get('/download', async (req, res) => {
  try {
    const fileExists = await fs.access(LOCAL_EXCEL_FILE).then(() => true).catch(() => false);
    if (!fileExists) {
      return res.status(404).send('No customer data available yet');
    }

    res.setHeader('Content-Disposition', 'attachment; filename=customers.xlsx');
    res.setHeader('Content-Type', 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet');
    const fileStream = require('fs').createReadStream(LOCAL_EXCEL_FILE);
    fileStream.pipe(res);
  } catch (error) {
    console.error('Error downloading local file:', error.message, error.stack);
    res.status(500).send('Error downloading file');
  }
});

// Initialize the server
(async () => {
  await initializeFromGoogleDrive();
  startGoogleDriveSync();
  app.listen(PORT, () => {
    console.log(`Server running on port ${PORT}`);
  });
})();
