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
// Flag to indicate if a local change has been made but not yet synced to Google Drive
let localChangesPending = false;

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
  console.log('Initialized worksheet columns:', sheet.columns.map(col => ({ header: col.header, key: col.key })));
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
    const actualHeaders = headers.slice(1, 4);
    const headersValid = expectedHeaders.every((header, index) => header === actualHeaders[index]);
    if (!headersValid) {
      throw new Error(`Invalid worksheet headers. Expected ${expectedHeaders}, got ${actualHeaders}`);
    }

    sheet.eachRow((row, rowNumber) => {
      if (rowNumber === 1) return;
      const name = row.getCell(1).value;
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

// Extract existing data from the workbook
async function extractExistingData(workbook) {
  const data = [];
  try {
    const sheet = workbook.getWorksheet('Customers');
    if (!sheet) {
      console.log('No Customers worksheet found for data extraction');
      return data;
    }

    sheet.eachRow((row, rowNumber) => {
      if (rowNumber === 1) return;
      try {
        const name = row.getCell(1).value;
        const email = row.getCell(2).value;
        const phone = row.getCell(3).value;
        if (name && email && phone) {
          data.push([name, email, phone]);
        } else {
          console.warn(`Row ${rowNumber} has missing data:`, [name, email, phone]);
        }
      } catch (error) {
        console.error(`Failed to extract row ${rowNumber}:`, error.message, error.stack);
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
  let existingData = [];
  try {
    await checkDiskSpaceAndPermissions(LOCAL_EXCEL_FILE);
    workbook = new ExcelJS.Workbook();
    await workbook.xlsx.readFile(LOCAL_EXCEL_FILE);
    console.log('Loaded local Excel file:', LOCAL_EXCEL_FILE);

    const isValid = await validateWorkbook(workbook);
    if (!isValid) {
      console.log('Workbook validation failed, extracting data before recreation...');
      existingData = await extractExistingData(workbook);
      workbook = await initializeExcel();
      const sheet = workbook.getWorksheet('Customers');
      if (existingData.length > 0) {
        existingData.forEach(rowData => {
          const newRow = sheet.addRow(rowData);
          newRow.commit();
          console.log('Re-added existing row during load:', rowData);
        });
      }
      await checkDiskSpaceAndPermissions(LOCAL_EXCEL_FILE);
      await workbook.xlsx.writeFile(LOCAL_EXCEL_FILE);
      console.log('Recreated Excel file with existing data:', LOCAL_EXCEL_FILE);
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

// Download the Excel file from Google Drive
async function downloadFromGoogleDrive() {
  // Skip download if there are pending local changes
  if (localChangesPending) {
    console.log('Skipping Google Drive download due to pending local changes');
    return;
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

      const workbook = new ExcelJS.Workbook();
      await workbook.xlsx.readFile(LOCAL_EXCEL_FILE);
      const isValid = await validateWorkbook(workbook);
      if (!isValid) {
        console.log('Downloaded file from Google Drive is corrupted, extracting data before recreation...');
        const existingData = await extractExistingData(workbook);
        const newWorkbook = await initializeExcel();
        const sheet = newWorkbook.getWorksheet('Customers');
        if (existingData.length > 0) {
          existingData.forEach(rowData => {
            const newRow = sheet.addRow(rowData);
            newRow.commit();
            console.log('Re-added existing row after download:', rowData);
          });
        }
        await checkDiskSpaceAndPermissions(LOCAL_EXCEL_FILE);
        await newWorkbook.xlsx.writeFile(LOCAL_EXCEL_FILE);
        console.log('Recreated Excel file with existing data after download:', LOCAL_EXCEL_FILE);
      }

      const sheet = workbook.getWorksheet('Customers');
      console.log('File contents after sync:');
      console.log('Column keys:', sheet.columns.map(col => col.key));
      sheet.eachRow((row, rowNumber) => {
        if (rowNumber === 1) {
          console.log('Headers:', row.values);
        } else {
          try {
            const name = row.getCell(1).value || '';
            const email = row.getCell(2).value || '';
            const phone = row.getCell(3).value || '';
            console.log('Row ' + rowNumber + ':', [name, email, phone]);
          } catch (error) {
            console.error(`Failed to log row ${rowNumber}:`, error.message, error.stack);
          }
        }
      });

      await createBackup();
    } else {
      console.log('No Excel file found in Google Drive, initializing new one locally.');
      const workbook = await initializeExcel();
      await checkDiskSpaceAndPermissions(LOCAL_EXCEL_FILE);
      await workbook.xlsx.writeFile(LOCAL_EXCEL_FILE);
      await createBackup();
    }
  } catch (error) {
    console.error('Error downloading from Google Drive:', error.message, error.stack);
    const workbook = await initializeExcel();
    await checkDiskSpaceAndPermissions(LOCAL_EXCEL_FILE);
    await workbook.xlsx.writeFile(LOCAL_EXCEL_FILE);
    await createBackup();
  }
}

// Upload the local Excel file to Google Drive with retry
async function uploadToGoogleDrive(maxRetries = 3) {
  let retries = 0;
  while (retries < maxRetries) {
    try {
      const stats = await fs.stat(LOCAL_EXCEL_FILE);
      if (stats.size === 0) {
        throw new Error('Local Excel file is empty');
      }
      console.log('Local Excel file size:', stats.size, 'bytes');

      await createBackup();

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
      localChangesPending = false; // Reset flag after successful upload
      return true; // Success
    } catch (error) {
      retries++;
      console.error(`Failed to upload to Google Drive (attempt ${retries}/${maxRetries}):`, error.message, error.stack);
      if (retries === maxRetries) {
        throw error;
      }
      // Wait before retrying (exponential backoff)
      await new Promise(resolve => setTimeout(resolve, 1000 * retries));
    }
  }
  throw new Error('Failed to upload to Google Drive after maximum retries');
}

// Periodic sync with Google Drive (every 5 minutes) - Temporarily disabled
function startGoogleDriveSync() {
  console.log('Periodic Google Drive sync is disabled for debugging');
}

// Download the Excel file from Google Drive on server start, or use local backup
async function initializeFromGoogleDrive() {
  let workbook;
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
      console.log('Local backup is corrupted, extracting data before recreation...');
      const existingData = await extractExistingData(workbook);
      workbook = await initializeExcel();
      const sheet = workbook.getWorksheet('Customers');
      if (existingData.length > 0) {
        existingData.forEach(rowData => {
          const newRow = sheet.addRow(rowData);
          newRow.commit();
          console.log('Re-added existing row from backup:', rowData);
        });
      }
      await checkDiskSpaceAndPermissions(LOCAL_EXCEL_FILE);
      await workbook.xlsx.writeFile(LOCAL_EXCEL_FILE);
      console.log('Recreated Excel file from backup with existing data:', LOCAL_EXCEL_FILE);
    }
  } catch (error) {
    console.log('Local backup not found or invalid, proceeding with Google Drive initialization:', error.message);
  }

  await downloadFromGoogleDrive();
}

// Helper function to generate the duplicate error message
function getDuplicateErrorMessage(emailExists, phoneExists) {
  if (emailExists && phoneExists) {
    return 'Email and phone number already exist';
  } else if (emailExists) {
    return 'Email already exists';
  } else {
    return 'Phone number already exists';
  }
}

// Handle form submission
app.post('/submit', async (req, res) => {
  const { name, email, phone } = req.body;

  console.log('SUBMIT: Received submission:', { name, email, phone });

  // Input validation
  if (!name || !email || !phone) {
    console.log('SUBMIT: Validation failed: Missing required fields');
    return res.status(400).json({ success: false, error: 'Missing required fields' });
  }

  const emailRegex = /^[a-zA-Z0-9._%+-]+@[a-zA-Z0-9.-]+\.(com|org|net|edu|gov|co|io|me|biz)$/i;
  if (!emailRegex.test(email)) {
    console.log('SUBMIT: Validation failed: Invalid email address');
    return res.status(400).json({ success: false, error: 'Please check the email address' });
  }

  const domain = email.split('@')[1].toLowerCase();
  const commonMisspellings = ['gmil.com', 'gail.com', 'gmai.com', 'gnail.com'];
  if (commonMisspellings.includes(domain)) {
    console.log(`SUBMIT: Detected common email domain misspelling: ${domain}`);
    return res.status(400).json({ success: false, error: 'Please check the email address' });
  }

  if (!/^\d{10}$/.test(phone)) {
    console.log('SUBMIT: Validation failed: Invalid phone number');
    return res.status(400).json({ success: false, error: 'Invalid phone number (10 digits required)' });
  }

  try {
    let submissionResult;
    fileLockPromise = fileLockPromise.then(async () => {
      try {
        // Temporarily disable Google Drive sync to isolate local file operations
        console.log('SUBMIT: Google Drive sync disabled for debugging');

        let workbook;
        let sheet;

        // Step 1: Load the existing Excel file
        console.log('SUBMIT: Loading local Excel file...');
        try {
          workbook = await loadLocalExcel();
          sheet = workbook.getWorksheet('Customers');
          console.log('SUBMIT: Successfully loaded local Excel file:', LOCAL_EXCEL_FILE);
        } catch (loadError) {
          console.error('SUBMIT: Failed to load Excel file, forcing recreation:', loadError.message, loadError.stack);
          workbook = await initializeExcel();
          sheet = workbook.getWorksheet('Customers');
          await checkDiskSpaceAndPermissions(LOCAL_EXCEL_FILE);
          await workbook.xlsx.writeFile(LOCAL_EXCEL_FILE);
          console.log('SUBMIT: Forced recreation of Excel file:', LOCAL_EXCEL_FILE);
        }

        // Step 2: Collect all existing rows (excluding header)
        console.log('SUBMIT: Collecting existing rows before adding new row...');
        const existingRows = [];
        sheet.eachRow((row, rowNumber) => {
          if (rowNumber === 1) return; // Skip header
          const name = row.getCell(1).value || '';
          const email = row.getCell(2).value || '';
          const phone = row.getCell(3).value || '';
          existingRows.push([name, email, phone]);
          console.log(`SUBMIT: Existing Row ${rowNumber}:`, [name, email, phone]);
        });

        // Step 3: Check for duplicates
        console.log('SUBMIT: Checking for duplicates...');
        let emailExists = false;
        let phoneExists = false;
        const normalizedEmail = email.toLowerCase().trim();
        const normalizedPhone = phone.toString().trim();

        existingRows.forEach((row, index) => {
          const existingEmail = row[1];
          const existingPhone = row[2];

          const normalizedExistingEmail = existingEmail ? existingEmail.toString().toLowerCase().trim() : '';
          const normalizedExistingPhone = existingPhone ? existingPhone.toString().trim() : '';

          console.log(`SUBMIT: Existing Row ${index + 2} - Email: '${normalizedExistingEmail}', Phone: '${normalizedExistingPhone}'`);
          console.log(`SUBMIT: Comparing Email - Input: '${normalizedEmail}', Existing: '${normalizedExistingEmail}', Match: ${normalizedExistingEmail === normalizedEmail}`);
          console.log(`SUBMIT: Comparing Phone - Input: '${normalizedPhone}', Existing: '${normalizedExistingPhone}', Match: ${normalizedExistingPhone === normalizedPhone}`);

          if (normalizedExistingEmail && normalizedExistingEmail === normalizedEmail) {
            emailExists = true;
          }
          if (normalizedExistingPhone && normalizedExistingPhone === normalizedPhone) {
            phoneExists = true;
          }
        });

        if (emailExists || phoneExists) {
          console.log('SUBMIT: Duplicate check - Email exists:', emailExists, 'Phone exists:', phoneExists);
          const errorMessage = getDuplicateErrorMessage(emailExists, phoneExists);
          submissionResult = { status: 400, body: { success: false, error: errorMessage } };
          return;
        } else {
          console.log('SUBMIT: No duplicates found, proceeding to add new row.');
        }

        // Step 4: Add the new row to the list
        existingRows.push([name, email, phone]);
        console.log('SUBMIT: Added new row to list:', [name, email, phone]);

        // Step 5: Recreate the worksheet with contiguous rows
        console.log('SUBMIT: Recreating worksheet with contiguous rows...');
        workbook.removeWorksheet('Customers');
        const newSheet = workbook.addWorksheet('Customers');
        newSheet.columns = [
          { header: 'Name', key: 'name', width: 20 },
          { header: 'Email', key: 'email', width: 30 },
          { header: 'Phone', key: 'phone', width: 15 },
        ];

        existingRows.forEach((rowValues, index) => {
          const newRow = newSheet.addRow(rowValues);
          newRow.commit();
          console.log(`SUBMIT: Added row ${index + 2} to new worksheet:`, rowValues);
        });

        // Step 6: Save the file with retries
        console.log('SUBMIT: Checking disk space and permissions before saving...');
        await checkDiskSpaceAndPermissions(LOCAL_EXCEL_FILE);
        let writeAttempts = 0;
        const maxWriteAttempts = 3;
        let fileWritten = false;
        console.log('SUBMIT: Attempting to save Excel file...');
        while (writeAttempts < maxWriteAttempts && !fileWritten) {
          try {
            await workbook.xlsx.writeFile(LOCAL_EXCEL_FILE);
            console.log('SUBMIT: Data successfully saved to local Excel file:', LOCAL_EXCEL_FILE);
            fileWritten = true;
          } catch (writeError) {
            writeAttempts++;
            console.error(`SUBMIT: Failed to write to Excel file (attempt ${writeAttempts}/${maxWriteAttempts}):`, writeError.message, writeError.stack);
            if (writeAttempts === maxWriteAttempts) {
              throw new Error('SUBMIT: Failed to write to Excel file after maximum attempts');
            }
            // Wait before retrying
            await new Promise(resolve => setTimeout(resolve, 1000));
          }
        }

        // Step 7: Verify the file contents after saving
        console.log('SUBMIT: Verifying file contents after save...');
        const updatedWorkbook = new ExcelJS.Workbook();
        await updatedWorkbook.xlsx.readFile(LOCAL_EXCEL_FILE);
        const updatedSheet = updatedWorkbook.getWorksheet('Customers');
        console.log('SUBMIT: Local file contents after submission:');
        let newRowFound = false;
        updatedSheet.eachRow((row, rowNumber) => {
          const rowName = row.getCell(1).value || '';
          const rowEmail = row.getCell(2).value || '';
          const rowPhone = row.getCell(3).value || '';
          console.log(`SUBMIT: Row ${rowNumber}:`, [rowName, rowEmail, rowPhone]);
          if (rowName === name && rowEmail === email && rowPhone === phone) {
            newRowFound = true;
          }
        });

        if (!newRowFound) {
          console.error('SUBMIT: New row not found in file after save. Attempting to recreate file...');
          workbook = await initializeExcel();
          sheet = workbook.getWorksheet('Customers');
          const newRow = sheet.addRow([name, email, phone]);
          newRow.commit();
          await checkDiskSpaceAndPermissions(LOCAL_EXCEL_FILE);
          await workbook.xlsx.writeFile(LOCAL_EXCEL_FILE);
          console.log('SUBMIT: Recreated Excel file with only the new row:', LOCAL_EXCEL_FILE);
        } else {
          console.log('SUBMIT: New row successfully verified in file.');
        }

        // Step 8: Set response (Google Drive upload disabled)
        submissionResult = { status: 200, body: { success: true, name } };
      } catch (error) {
        throw error;
      }
    });

    await fileLockPromise;

    if (submissionResult) {
      console.log('SUBMIT: Sending response:', submissionResult.body);
      res.status(submissionResult.status).json(submissionResult.body);
    } else {
      throw new Error('SUBMIT: Submission result not set');
    }
  } catch (error) {
    console.error('SUBMIT: Failed to save to local Excel:', error.message, error.stack);
    if (error.message.includes('Insufficient disk space')) {
      res.status(500).json({ success: false, error: 'Server disk space is full. Please contact support.' });
    } else if (error.message.includes('Permission denied')) {
      res.status(500).json({ success: false, error: 'File permission error. Please contact support.' });
    } else if (error.message.includes('Corrupt')) {
      console.log('SUBMIT: Excel file appears to be corrupted, already recreated in main flow.');
      res.status(503).json({ success: false, error: 'File was corrupted, please try again.' });
    } else if (error.message.includes('Out of bounds')) {
      console.log('SUBMIT: Excel file has invalid column structure, already recreated in main flow.');
      res.status(503).json({ success: false, error: 'File was corrupted, please try again.' });
    } else {
      res.status(500).json({ success: false, error: 'Unable to save your submission. Please try again later.' });
    }
  }
});

// Handle row deletion
app.post('/delete', async (req, res) => {
  const { email, phone } = req.body;

  console.log('DELETE: Received delete request:', { email, phone });

  if (!email && !phone) {
    console.log('DELETE: Validation failed: Email or phone required for deletion');
    return res.status(400).json({ success: false, error: 'Email or phone required for deletion' });
  }

  try {
    let deletionResult;
    fileLockPromise = fileLockPromise.then(async () => {
      try {
        console.log('DELETE: Syncing with Google Drive before deletion...');
        await downloadFromGoogleDrive();

        let workbook;
        let sheet;

        console.log('DELETE: Loading local Excel file...');
        try {
          workbook = await loadLocalExcel();
          sheet = workbook.getWorksheet('Customers');
          console.log('DELETE: Successfully loaded local Excel file:', LOCAL_EXCEL_FILE);
        } catch (loadError) {
          console.error('DELETE: Failed to load Excel file for deletion, forcing recreation:', loadError.message, loadError.stack);
          workbook = await initializeExcel();
          sheet = workbook.getWorksheet('Customers');
          await checkDiskSpaceAndPermissions(LOCAL_EXCEL_FILE);
          await workbook.xlsx.writeFile(LOCAL_EXCEL_FILE);
          console.log('DELETE: Forced recreation of Excel file for deletion:', LOCAL_EXCEL_FILE);
          deletionResult = { status: 404, body: { success: false, error: 'No data to delete after file recreation' } };
          return;
        }

        console.log('DELETE: Checking for matching rows to delete...');
        let rowFound = false;
        const rowsToKeep = [];
        rowsToKeep.push(sheet.getRow(1).values);

        const normalizedEmail = email ? email.toLowerCase().trim() : null;
        const normalizedPhone = phone ? phone.toString().trim() : null;

        sheet.eachRow((row, rowNumber) => {
          if (rowNumber === 1) return;
          const existingEmail = row.getCell(2).value;
          const existingPhone = row.getCell(3).value;

          const normalizedExistingEmail = existingEmail ? existingEmail.toString().toLowerCase().trim() : '';
          const normalizedExistingPhone = existingPhone ? existingPhone.toString().trim() : '';

          const emailMatch = normalizedEmail && normalizedExistingEmail === normalizedEmail;
          const phoneMatch = normalizedPhone && normalizedExistingPhone === normalizedPhone;

          if (emailMatch || phoneMatch) {
            console.log(`DELETE: Found matching row ${rowNumber} to delete:`, [row.getCell(1).value, normalizedExistingEmail, normalizedExistingPhone]);
            rowFound = true;
          } else {
            rowsToKeep.push(row.values);
          }
        });

        if (!rowFound) {
          console.log('DELETE: No matching row found for deletion');
          deletionResult = { status: 404, body: { success: false, error: 'Customer not found' } };
          return;
        }

        console.log('DELETE: Rows to keep after deletion:', rowsToKeep);

        // Recreate the sheet with remaining rows
        console.log('DELETE: Recreating worksheet with remaining rows...');
        workbook.removeWorksheet('Customers');
        const newSheet = workbook.addWorksheet('Customers');
        newSheet.columns = [
          { header: 'Name', key: 'name', width: 20 },
          { header: 'Email', key: 'email', width: 30 },
          { header: 'Phone', key: 'phone', width: 15 },
        ];

        rowsToKeep.forEach((rowValues, index) => {
          const newRow = newSheet.addRow(rowValues);
          newRow.commit();
          console.log(`DELETE: Re-added row ${index + 1} after deletion:`, rowValues);
        });

        console.log('DELETE: Checking disk space and permissions before saving...');
        await checkDiskSpaceAndPermissions(LOCAL_EXCEL_FILE);
        let writeAttempts = 0;
        const maxWriteAttempts = 3;
        let fileWritten = false;
        console.log('DELETE: Attempting to save Excel file after deletion...');
        while (writeAttempts < maxWriteAttempts && !fileWritten) {
          try {
            await workbook.xlsx.writeFile(LOCAL_EXCEL_FILE);
            console.log('DELETE: Data successfully saved to local Excel file after deletion:', LOCAL_EXCEL_FILE);
            fileWritten = true;
          } catch (writeError) {
            writeAttempts++;
            console.error(`DELETE: Failed to write to Excel file after deletion (attempt ${writeAttempts}/${maxWriteAttempts}):`, writeError.message, writeError.stack);
            if (writeAttempts === maxWriteAttempts) {
              throw new Error('DELETE: Failed to write to Excel file after maximum attempts');
            }
            // Wait before retrying
            await new Promise(resolve => setTimeout(resolve, 1000));
          }
        }

        // Verify the local file contents after saving
        console.log('DELETE: Verifying file contents after deletion...');
        const updatedWorkbook = new ExcelJS.Workbook();
        await updatedWorkbook.xlsx.readFile(LOCAL_EXCEL_FILE);
        const updatedSheet = updatedWorkbook.getWorksheet('Customers');
        console.log('DELETE: Local file contents after deletion:');
        updatedSheet.eachRow((row, rowNumber) => {
          const name = row.getCell(1).value || '';
          const email = row.getCell(2).value || '';
          const phone = row.getCell(3).value || '';
          console.log(`DELETE: Row ${rowNumber}:`, [name, email, phone]);
        });

        localChangesPending = true; // Set flag to indicate local changes
        console.log('DELETE: Attempting to upload to Google Drive...');
        try {
          await uploadToGoogleDrive();
          console.log('DELETE: Successfully uploaded to Google Drive.');
          deletionResult = { status: 200, body: { success: true, message: 'Customer deleted successfully' } };
        } catch (syncError) {
          console.error('DELETE: Google Drive sync failed after deletion:', syncError.message, syncError.stack);
          deletionResult = { 
            status: 200, 
            body: { 
              success: true, 
              message: 'Customer deleted locally, but failed to sync to Google Drive. Please contact support.'
            }
          };
        }
      } catch (error) {
        throw error;
      }
    });

    await fileLockPromise;

    if (deletionResult) {
      console.log('DELETE: Sending response:', deletionResult.body);
      res.status(deletionResult.status).json(deletionResult.body);
    } else {
      throw new Error('DELETE: Deletion result not set');
    }
  } catch (error) {
    console.error('DELETE: Failed to delete customer:', error.message, error.stack);
    if (error.message.includes('Insufficient disk space')) {
      res.status(500).json({ success: false, error: 'Server disk space is full. Please contact support.' });
    } else if (error.message.includes('Permission denied')) {
      res.status(500).json({ success: false, error: 'File permission error. Please contact support.' });
    } else if (error.message.includes('Corrupt')) {
      console.log('DELETE: Excel file appears to be corrupted, already recreated in main flow.');
      res.status(503).json({ success: false, error: 'File was corrupted, please try again.' });
    } else {
      res.status(500).json({ success: false, error: 'Unable to delete customer. Please try again later.' });
    }
  }
});

// Handle file download
app.get('/download', async (req, res) => {
  try {
    const fileExists = await fs.access(LOCAL_EXCEL_FILE).then(() => true).catch(() => false);
    if (!fileExists) {
      console.log('DOWNLOAD: No customer data available yet');
      return res.status(404).send('No customer data available yet');
    }

    // Log the file contents before sending
    console.log('DOWNLOAD: Reading file contents before sending...');
    const workbook = new ExcelJS.Workbook();
    await workbook.xlsx.readFile(LOCAL_EXCEL_FILE);
    const sheet = workbook.getWorksheet('Customers');
    console.log('DOWNLOAD: File contents before download:');
    sheet.eachRow((row, rowNumber) => {
      const name = row.getCell(1).value || '';
      const email = row.getCell(2).value || '';
      const phone = row.getCell(3).value || '';
      console.log(`DOWNLOAD: Row ${rowNumber}:`, [name, email, phone]);
    });

    console.log('DOWNLOAD: Sending Excel file to client...');
    res.setHeader('Content-Disposition', 'attachment; filename=customers.xlsx');
    res.setHeader('Content-Type', 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet');
    const fileStream = require('fs').createReadStream(LOCAL_EXCEL_FILE);
    fileStream.pipe(res);
    console.log('DOWNLOAD: File stream initiated.');
  } catch (error) {
    console.error('DOWNLOAD: Error downloading local file:', error.message, error.stack);
    res.status(500).send('Error downloading file');
  }
});

// Initialize the server
(async () => {
  console.log('SERVER: Initializing server and syncing with Google Drive...');
  await initializeFromGoogleDrive();
  startGoogleDriveSync();
  app.listen(PORT, () => {
    console.log(`SERVER: Server running on port ${PORT}`);
  });
})();
