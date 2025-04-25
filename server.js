const express = require('express');
const bodyParser = require('body-parser');
const ExcelJS = require('exceljs');
const path = require('path');
const fs = require('fs').promises;
const { google } = require('googleapis');
const disk = require('diskusage');

const app = express();
const PORT = process.env.PORT || 10000;
const LOCAL_EXCEL_FILE = path.join(__dirname, 'customers.xlsx');
const LOCAL_BACKUP_FILE = path.join(__dirname, 'customers_backup.xlsx');
const GOOGLE_DRIVE_FOLDER_ID = '1l4e6cq0LaFS2IFkJlWKLFJ_CVIEqPqTK';

// Use a promise-based lock to prevent concurrent file access
let fileLockPromise = Promise.resolve();

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
      await fs.chmod(filePath, 0o666);
      console.log(`Permissions fixed for ${filePath}`);
    }
  } catch (error) {
    console.error('Disk space or permission check failed:', error.message);
    throw error;
  }
}

// Validate that the workbook is readable
async function validateWorkbook(workbook) {
  try {
    const sheet = workbook.getWorksheet('Customers');
    if (!sheet) {
      throw new Error('Customers worksheet not found');
    }

    const headers = sheet.getRow(1).values;
    console.log('Worksheet headers:', headers);
    const expectedHeaders = ['Name', 'Email', 'Phone'];
    const actualHeaders = headers.slice(1, 4); // Ignore first empty cell, take first 3 headers
    const headersValid = expectedHeaders.every((header, index) => header === actualHeaders[index]);
    if (!headersValid) {
      throw new Error(`Invalid worksheet headers. Expected ${expectedHeaders}, got ${actualHeaders}`);
    }

    // Try accessing a few rows to ensure the file is not corrupted
    sheet.eachRow((row, rowNumber) => {
      if (rowNumber === 1) return;
      const name = row.getCell(1).value; // Use column index to avoid key-based issues
      const email = row.getCell(2).value;
      const phone = row.getCell(3).value;
      console.log(`Validated row ${rowNumber}:`, [name, email, phone]);
    });
    return true;
  } catch (error) {
    console.error('Workbook validation failed:', error.message, error.stack);
    return false;
  }
}

// Extract existing data from the workbook if possible
async function extractExistingData(workbook) {
  const data = [];
  try {
    const sheet = workbook.getWorksheet('Customers');
    if (!sheet) {
      console.log('No Customers worksheet found for data extraction');
      return data;
    }

    sheet.eachRow((row, rowNumber) => {
      if (rowNumber === 1) return; // Skip header row
      try {
        // Try accessing cells by key first
        const name = row.getCell('name').value;
        const email = row.getCell('email').value;
        const phone = row.getCell('phone').value;
        if (name && email && phone) {
          data.push([name, email, phone]);
        }
      } catch (error) {
        console.warn('Failed to extract row using keys, trying column indices:', error.message);
        // Fallback to column indices
        try {
          const name = row.getCell(1).value;
          const email = row.getCell(2).value;
          const phone = row.getCell(3).value;
          if (name && email && phone) {
            data.push([name, email, phone]);
          }
        } catch (indexError) {
          console.error('Failed to extract row using indices:', indexError.message, indexError.stack);
        }
      }
    });
    console.log('Extracted existing data:', data);
  } catch (error) {
    console.error('Failed to extract existing data:', error.message, error.stack);
  }
  return data;
}

// Load the local Excel file or initialize a new one
async function loadLocalExcel() {
  let workbook;
  try {
    await checkDiskSpaceAndPermissions(LOCAL_EXCEL_FILE);
    workbook = new ExcelJS.Workbook();
    await workbook.xlsx.readFile(LOCAL_EXCEL_FILE);
    console.log('Loaded local Excel file:', LOCAL_EXCEL_FILE);

    // Validate the workbook
    const isValid = await validateWorkbook(workbook);
    if (!isValid) {
      throw new Error('Workbook is corrupted or invalid');
    }
  } catch (error) {
    console.log('Local Excel file not found, inaccessible, or invalid, initializing new one:', error.message);
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

// Create a backup of the local Excel file
async function createBackup() {
  try {
    await checkDiskSpaceAndPermissions(LOCAL_EXCEL_FILE);
    await fs.copyFile(LOCAL_EXCEL_FILE, LOCAL_BACKUP_FILE);
    console.log('Created backup of Excel file:', LOCAL_BACKUP_FILE);
  } catch (error) {
    console.error('Failed to create backup:', error.message, error.stack);
  }
}

// Upload the local Excel file to Google Drive
async function uploadToGoogleDrive() {
  try {
    // Verify the local file exists and is not empty
    const stats = await fs.stat(LOCAL_EXCEL_FILE);
    if (stats.size === 0) {
      throw new Error('Local Excel file is empty');
    }
    console.log('Local Excel file size:', stats.size, 'bytes');

    // Create a backup before uploading
    await createBackup();

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
  }
}

// Periodic sync with Google Drive (every 5 minutes)
function startGoogleDriveSync() {
  setInterval(async () => {
    try {
      console.log('Starting periodic sync with Google Drive...');
      fileLockPromise = fileLockPromise.then(async () => {
        await uploadToGoogleDrive();
      });
      await fileLockPromise;
      console.log('Periodic sync completed.');
    } catch (error) {
      console.error('Periodic sync failed:', error.message, error.stack);
    }
  }, 5 * 60 * 1000);
}

// Download the Excel file from Google Drive on server start, or use local backup
async function initializeFromGoogleDrive() {
  let workbook;
  // First, try loading the local backup if it exists
  try {
    await checkDiskSpaceAndPermissions(LOCAL_BACKUP_FILE);
    workbook = new ExcelJS.Workbook();
    await workbook.xlsx.readFile(LOCAL_BACKUP_FILE);
    console.log('Loaded local backup file:', LOCAL_BACKUP_FILE);
    const isValid = await validateWorkbook(workbook);
    if (isValid) {
      await fs.copyFile(LOCAL_BACKUP_FILE, LOCAL_EXCEL_FILE);
      console.log('Restored local backup to main file:', LOCAL_EXCEL_FILE);
      return;
    } else {
      console.log('Local backup is corrupted, proceeding with Google Drive initialization');
    }
  } catch (error) {
    console.log('Local backup not found or invalid, proceeding with Google Drive initialization:', error.message);
  }

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

      // Validate the downloaded file
      workbook = new ExcelJS.Workbook();
      await workbook.xlsx.readFile(LOCAL_EXCEL_FILE);
      const isValid = await validateWorkbook(workbook);
      if (!isValid) {
        console.log('Downloaded file from Google Drive is corrupted, initializing new one');
        throw new Error('Downloaded file is corrupted');
      }
    } else {
      console.log('No Excel file found in Google Drive, initializing new one locally.');
      workbook = await initializeExcel();
      await checkDiskSpaceAndPermissions(LOCAL_EXCEL_FILE);
      await workbook.xlsx.writeFile(LOCAL_EXCEL_FILE);
    }
  } catch (error) {
    console.error('Error initializing from Google Drive:', error.message, error.stack);
    workbook = await initializeExcel();
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

  try {
    // Use the lock to ensure exclusive access to the file
    let submissionResult;
    fileLockPromise = fileLockPromise.then(async () => {
      try {
        let workbook;
        let sheet;
 odpowiednie let existingData = [];

        try {
          workbook = await loadLocalExcel();
          sheet = workbook.getWorksheet('Customers');
          // Extract existing data before proceeding
          existingData = await extractExistingData(workbook);
        } catch (loadError) {
          console.error('Failed to load Excel file, forcing recreation:', loadError.message, loadError.stack);
          // Force recreation of the file if loading fails
          workbook = await initializeExcel();
          sheet = workbook.getWorksheet('Customers');
          await checkDiskSpaceAndPermissions(LOCAL_EXCEL_FILE);
          await workbook.xlsx.writeFile(LOCAL_EXCEL_FILE);
          console.log('Forced recreation of Excel file:', LOCAL_EXCEL_FILE);
        }

        let emailExists = false;
        let phoneExists = false;
        try {
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
        } catch (rowError) {
          console.error('Error accessing rows, likely corrupted file:', rowError.message, rowError.stack);
          // If accessing rows fails (e.g., Out of bounds error), recreate the file
          // Re-extract existing data if possible
          existingData = await extractExistingData(workbook);
          workbook = await initializeExcel();
          sheet = workbook.getWorksheet('Customers');
          // Re-add existing data
          if (existingData.length > 0) {
            existingData.forEach(rowData => {
              const newRow = sheet.addRow(rowData);
              newRow.commit();
              console.log('Re-added existing row during recreation:', rowData);
            });
          } else {
            console.warn('No existing data could be extracted due to corruption; previous data may be lost');
          }
          await checkDiskSpaceAndPermissions(LOCAL_EXCEL_FILE);
          await workbook.xlsx.writeFile(LOCAL_EXCEL_FILE);
          console.log('Recreated Excel file due to row access error:', LOCAL_EXCEL_FILE);
          // Since the file was recreated, duplicate check is already done via existingData
          emailExists = existingData.some(row => row[1].toLowerCase() === email.toLowerCase());
          phoneExists = existingData.some(row => row[2].toString() === phone.toString());
        }

        if (emailExists || phoneExists) {
          console.log(`Duplicate found - Email exists: ${emailExists}, Phone exists: ${phoneExists}`);
          let errorMessage;
          if (emailExists && phoneExists) {
            errorMessage = 'Email and phone number already exist';
          } else if (emailExists) {
            errorMessage = 'Email already exists';
          } else {
            errorMessage = 'Phone number already exists';
          }
          submissionResult = { status: 400, body: { success: false, error: errorMessage } };
          return;
        }

        const newRow = sheet.addRow([name, email, phone]);
        newRow.commit();
        console.log('Added new row to worksheet:', [name, email, phone]);

        await checkDiskSpaceAndPermissions(LOCAL_EXCEL_FILE);
        try {
          await workbook.xlsx.writeFile(LOCAL_EXCEL_FILE);
          console.log('Data saved to local Excel file:', LOCAL_EXCEL_FILE);
        } catch (writeError) {
          console.error('Failed to write to Excel file, attempting to recreate:', writeError.message, writeError.stack);
          // If writing fails, recreate the file and retry
          existingData = await extractExistingData(workbook);
          workbook = await initializeExcel();
          const retrySheet = workbook.getWorksheet('Customers');
          // Re-add existing data
          if (existingData.length > 0) {
            existingData.forEach(rowData => {
              const newRow = retrySheet.addRow(rowData);
              newRow.commit();
              console.log('Re-added existing row after write failure:', rowData);
            });
          } else {
            console.warn('No existing data could be extracted due to corruption; previous data may be lost');
          }
          // Add the new row
          const newRetryRow = retrySheet.addRow([name, email, phone]);
          newRetryRow.commit();
          console.log('Added new row after write failure:', [name, email, phone]);
          await checkDiskSpaceAndPermissions(LOCAL_EXCEL_FILE);
          await workbook.xlsx.writeFile(LOCAL_EXCEL_FILE);
          console.log('Recreated and saved to Excel file after write failure:', LOCAL_EXCEL_FILE);
        }

        try {
          await uploadToGoogleDrive();
          submissionResult = { status: 200, body: { success: true, name } };
        } catch (syncError) {
          console.error('Google Drive sync failed after local write:', syncError.message, syncError.stack);
          submissionResult = { 
            status: 200, 
            body: { 
              success: true, 
              name, 
              warning: 'Data saved locally, but failed to sync to Google Drive. Please contact support.'
            }
          };
        }
      } catch (error) {
        throw error; // Re-throw to be caught by the outer catch
      }
    });

    await fileLockPromise;

    if (submissionResult) {
      res.status(submissionResult.status).json(submissionResult.body);
    } else {
      throw new Error('Submission result not set');
    }
  } catch (error) {
    console.error('Failed to save to local Excel:', error.message, error.stack);
    if (error.message.includes('Insufficient disk space')) {
      res.status(500).json({ success: false, error: 'Server disk space is full. Please contact support.' });
    } else if (error.message.includes('Permission denied')) {
      res.status(500).json({ success: false, error: 'File permission error. Please contact support.' });
    } else if (error.message.includes('Corrupt')) {
      console.log('Excel file appears to be corrupted, already recreated in main flow.');
      res.status(503).json({ success: false, error: 'File was corrupted, please try again.' });
    } else if (error.message.includes('Out of bounds')) {
      console.log('Excel file has invalid column structure, already recreated in main flow.');
      res.status(503).json({ success: false, error: 'File was corrupted, please try again.' });
    } else {
      res.status(500).json({ success: false, error: 'Unable to save your submission. Please try again later.' });
    }
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
