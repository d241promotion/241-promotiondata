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
const GOOGLE_DRIVE_FOLDER_ID = '1l4e6cq0LaFS2IFkJlWKLFJ_CVIEqPqTK';

// Use a promise-based lock to prevent concurrent file access
let fileLockPromise = Promise.resolve();
// Flag to indicate if a local change has been made but not yet synced to Google Drive
let localChangesPending = false;
// In-memory cache for the workbook
let cachedWorkbook = null;

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
  const customerSheet = workbook.addWorksheet('Customers');
  customerSheet.columns = [
    { header: 'Name', key: 'name', width: 20 },
    { header: 'Email', key: 'email', width: 30 },
    { header: 'Phone', key: 'phone', width: 15 },
    { header: 'Date of Birth', key: 'dob', width: 15 },
  ];
  console.log('Initialized Customers worksheet columns:', customerSheet.columns.map(col => ({ header: col.header, key: col.key })));

  const feedbackSheet = workbook.addWorksheet('Feedback');
  feedbackSheet.columns = [
    { header: 'Name', key: 'name', width: 20 },
    { header: 'Review', key: 'review', width: 50 },
    { header: 'Timestamp', key: 'timestamp', width: 20 },
  ];
  console.log('Initialized Feedback worksheet columns:', feedbackSheet.columns.map(col => ({ header: col.header, key: col.key })));

  return workbook;
}

// Check disk space and file permissions
async function checkDiskSpaceAndPermissions(filePath, isNewFile = false) {
  try {
    const { available, total } = await disk.check(path.dirname(filePath));
    const availableMB = available / (1024 * 1024);
    console.log(`Available disk space: ${availableMB.toFixed(2)} MB`);
    if (availableMB < 10) {
      throw new Error(`Insufficient disk space: ${availableMB.toFixed(2)} MB available`);
    }

    if (!isNewFile) {
      try {
        await fs.access(filePath, fs.constants.R_OK | fs.constants.W_OK);
        console.log(`File ${filePath} is readable and writable`);
      } catch (error) {
        console.log(`File ${filePath} not accessible, attempting to fix permissions`);
        await fs.chmod(filePath, 0o666);
        console.log(`Permissions fixed for ${filePath}`);
      }
    } else {
      const dirPath = path.dirname(filePath);
      await fs.access(dirPath, fs.constants.W_OK);
      console.log(`Directory ${dirPath} is writable for new file creation`);
    }
  } catch (error) {
    console.error('Disk space or permission check failed:', error.message);
    throw error;
  }
}

// Log file stats (timestamp and size)
async function logFileStats(filePath, label) {
  try {
    const stats = await fs.stat(filePath);
    console.log(`${label} - File stats for ${filePath}:`);
    console.log(`  Last modified: ${stats.mtime}`);
    console.log(`  Size: ${stats.size} bytes`);
  } catch (error) {
    console.error(`${label} - Failed to get file stats for ${filePath}:`, error.message);
  }
}

// Validate that the workbook is readable
async function validateWorkbook(workbook) {
  try {
    const customerSheet = workbook.getWorksheet('Customers');
    if (!customerSheet) {
      throw new Error('Customers worksheet not found');
    }

    const customerHeaders = customerSheet.getRow(1).values;
    console.log('Customers worksheet headers:', customerHeaders);
    const expectedCustomerHeaders = ['Name', 'Email', 'Phone', 'Date of Birth'];
    const actualCustomerHeaders = customerHeaders.slice(1, 5);
    const customerHeadersValid = expectedCustomerHeaders.every((header, index) => header === actualCustomerHeaders[index]);
    if (!customerHeadersValid) {
      throw new Error(`Invalid Customers worksheet headers. Expected ${expectedCustomerHeaders}, got ${actualCustomerHeaders}`);
    }

    const feedbackSheet = workbook.getWorksheet('Feedback');
    if (!feedbackSheet) {
      throw new Error('Feedback worksheet not found');
    }

    const feedbackHeaders = feedbackSheet.getRow(1).values;
    console.log('Feedback worksheet headers:', feedbackHeaders);
    const expectedFeedbackHeaders = ['Name', 'Review', 'Timestamp'];
    const actualFeedbackHeaders = feedbackHeaders.slice(1, 4);
    const feedbackHeadersValid = expectedFeedbackHeaders.every((header, index) => header === actualFeedbackHeaders[index]);
    if (!feedbackHeadersValid) {
      throw new Error(`Invalid Feedback worksheet headers. Expected ${expectedFeedbackHeaders}, got ${actualFeedbackHeaders}`);
    }

    let hasCustomerDataRows = false;
    customerSheet.eachRow((row, rowNumber) => {
      if (rowNumber === 1) return;
      const name = row.getCell(1).value;
      const email = row.getCell(2).value;
      const phone = row.getCell(3).value;
      const dob = row.getCell(4).value;
      if (name || email || phone || dob) {
        hasCustomerDataRows = true;
        console.log(`Validated Customers row ${rowNumber}:`, [name, email, phone, dob]);
      }
    });

    let hasFeedbackDataRows = false;
    feedbackSheet.eachRow((row, rowNumber) => {
      if (rowNumber === 1) return;
      const name = row.getCell(1).value;
      const review = row.getCell(2).value;
      const timestamp = row.getCell(3).value;
      if (name || review || timestamp) {
        hasFeedbackDataRows = true;
        console.log(`Validated Feedback row ${rowNumber}:`, [name, review, timestamp]);
      }
    });

    if (!hasCustomerDataRows) {
      console.log('Customers worksheet has no data rows, but header is valid');
    }
    if (!hasFeedbackDataRows) {
      console.log('Feedback worksheet has no data rows, but header is valid');
    }
    return true;
  } catch (error) {
    console.error('Workbook validation failed:', error.message, error.stack);
    return false;
  }
}

// Extract existing customer data from the workbook
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
        const dob = row.getCell(4).value;
        if (name && email && phone && name.toString().trim() && email.toString().trim() && phone.toString().trim()) {
          data.push([name, email, phone, dob]);
        } else {
          console.warn(`Customers row ${rowNumber} has missing or empty data, skipping:`, [name, email, phone, dob]);
        }
      } catch (error) {
        console.error(`Failed to extract Customers row ${rowNumber}:`, error.message, error.stack);
      }
    });
    console.log('Extracted existing Customers data:', data);
  } catch (error) {
    console.error('Failed to extract existing Customers data:', error.message, error.stack);
  }
  return data;
}

// Extract existing feedback data from the workbook
async function extractExistingFeedbackData(workbook) {
  const data = [];
  try {
    const sheet = workbook.getWorksheet('Feedback');
    if (!sheet) {
      console.log('No Feedback worksheet found for data extraction');
      return data;
    }

    sheet.eachRow((row, rowNumber) => {
      if (rowNumber === 1) return;
      try {
        const name = row.getCell(1).value;
        const review = row.getCell(2).value;
        const timestamp = row.getCell(3).value;
        if (name && review && name.toString().trim() && review.toString().trim()) {
          data.push([name, review, timestamp]);
        } else {
          console.warn(`Feedback row ${rowNumber} has missing or empty data, skipping:`, [name, review, timestamp]);
        }
      } catch (error) {
        console.error(`Failed to extract Feedback row ${rowNumber}:`, error.message, error.stack);
      }
    });
    console.log('Extracted existing Feedback data:', data);
  } catch (error) {
    console.error('Failed to extract existing Feedback data:', error.message, error.stack);
  }
  return data;
}

// Load the local Excel file or initialize a new one with retries
async function loadLocalExcel() {
  if (cachedWorkbook) {
    console.log('LOAD: Using cached workbook');
    console.log('LOAD: Cached workbook contents:');
    const customerSheet = cachedWorkbook.getWorksheet('Customers');
    customerSheet.eachRow((row, rowNumber) => {
      const name = row.getCell(1).value || '';
      const email = row.getCell(2).value || '';
      const phone = row.getCell(3).value || '';
      const dob = row.getCell(4).value || '';
      console.log(`LOAD: Customers Row ${rowNumber}:`, [name, email, phone, dob]);
    });
    const feedbackSheet = cachedWorkbook.getWorksheet('Feedback');
    feedbackSheet.eachRow((row, rowNumber) => {
      const name = row.getCell(1).value || '';
      const review = row.getCell(2).value || '';
      const timestamp = row.getCell(3).value || '';
      console.log(`LOAD: Feedback Row ${rowNumber}:`, [name, review, timestamp]);
    });
    const isValid = await validateWorkbook(cachedWorkbook);
    if (isValid) {
      return cachedWorkbook;
    } else {
      console.log('LOAD: Cached workbook is invalid, clearing cache and reloading from disk...');
      cachedWorkbook = null;
    }
  } else {
    console.log('LOAD: No cached workbook found');
  }

  let workbook;
  let existingCustomerData = [];
  let existingFeedbackData = [];
  try {
    const fileExists = await fs.access(LOCAL_EXCEL_FILE).then(() => true).catch(() => false);
    if (!fileExists) {
      console.log('LOAD: Local Excel file does not exist, downloading from Google Drive...');
      await downloadFromGoogleDrive();
    } else {
      console.log('LOAD: Local Excel file exists, loading from disk...');
    }

    await checkDiskSpaceAndPermissions(LOCAL_EXCEL_FILE);
    await logFileStats(LOCAL_EXCEL_FILE, 'LOAD');
    workbook = new ExcelJS.Workbook();
    await workbook.xlsx.readFile(LOCAL_EXCEL_FILE);
    console.log('Loaded local Excel file:', LOCAL_EXCEL_FILE);

    const isValid = await validateWorkbook(workbook);
    if (!isValid) {
      console.log('Workbook validation failed, extracting data before recreation...');
      existingCustomerData = await extractExistingData(workbook);
      existingFeedbackData = await extractExistingFeedbackData(workbook);
      workbook = await initializeExcel();
      const customerSheet = workbook.getWorksheet('Customers');
      if (existingCustomerData.length > 0) {
        existingCustomerData.forEach(rowData => {
          const newRow = customerSheet.addRow(rowData);
          newRow.commit();
          console.log('Re-added existing Customers row during load:', rowData);
        });
      }
      const feedbackSheet = workbook.getWorksheet('Feedback');
      if (existingFeedbackData.length > 0) {
        existingFeedbackData.forEach(rowData => {
          const newRow = feedbackSheet.addRow(rowData);
          newRow.commit();
          console.log('Re-added existing Feedback row during load:', rowData);
        });
      }
      await checkDiskSpaceAndPermissions(LOCAL_EXCEL_FILE, true);
      await workbook.xlsx.writeFile(LOCAL_EXCEL_FILE);
      await logFileStats(LOCAL_EXCEL_FILE, 'LOAD: After Recreation');
      console.log('Recreated Excel file with existing data:', LOCAL_EXCEL_FILE);

      const fileExists = await fs.access(LOCAL_EXCEL_FILE).then(() => true).catch(() => false);
      if (!fileExists) {
        throw new Error('Failed to create Excel file after recreation');
      }
      console.log('Verified: Excel file exists after recreation');
      cachedWorkbook = workbook;
    }
  } catch (error) {
    console.log('Local Excel file not found, inaccessible, or invalid, initializing new one:', error.message);
    workbook = await initializeExcel();
    let fileCreated = false;
    let writeAttempts = 0;
    const maxWriteAttempts = 3;

    while (writeAttempts < maxWriteAttempts && !fileCreated) {
      try {
        await checkDiskSpaceAndPermissions(LOCAL_EXCEL_FILE, true);
        await workbook.xlsx.writeFile(LOCAL_EXCEL_FILE);
        await logFileStats(LOCAL_EXCEL_FILE, 'LOAD: After Initialization');
        console.log('Initialized new Excel file:', LOCAL_EXCEL_FILE);

        const fileExists = await fs.access(LOCAL_EXCEL_FILE).then(() => true).catch(() => false);
        if (!fileExists) {
          throw new Error('Failed to create Excel file after initialization');
        }
        console.log('Verified: New Excel file exists after initialization');

        const verifyWorkbook = new ExcelJS.Workbook();
        await verifyWorkbook.xlsx.readFile(LOCAL_EXCEL_FILE);
        const customerSheet = verifyWorkbook.getWorksheet('Customers');
        const feedbackSheet = verifyWorkbook.getWorksheet('Feedback');
        if (!customerSheet || !feedbackSheet) {
          throw new Error('Customers or Feedback worksheet not found in newly created file');
        }
        console.log('Verified: New Excel file contains Customers and Feedback worksheets');
        fileCreated = true;
      } catch (writeError) {
        writeAttempts++;
        console.error(`Failed to initialize new Excel file (attempt ${writeAttempts}/${maxWriteAttempts}):`, writeError.message, writeError.stack);
        if (writeAttempts === maxWriteAttempts) {
          throw new Error(`Failed to initialize new Excel file after ${maxWriteAttempts} attempts: ${writeError.message}`);
        }
        await new Promise(resolve => setTimeout(resolve, 1000));
      }
    }
    cachedWorkbook = workbook;
  }

  cachedWorkbook = workbook;
  return workbook;
}

// Download the Excel file from Google Drive
async function downloadFromGoogleDrive(forceSync = false) {
  console.log(`DOWNLOAD: localChangesPending state: ${localChangesPending}, forceSync: ${forceSync}`);
  if (localChangesPending && !forceSync) {
    console.log('Skipping Google Drive download due to pending local changes');
    return;
  }

  try {
    console.log('DOWNLOAD: Querying Google Drive for customers.xlsx...');
    const response = await drive.files.list({
      q: `'${GOOGLE_DRIVE_FOLDER_ID}' in parents and name = 'customers.xlsx' and trashed = false`,
      fields: 'files(id, name, trashed)',
    });

    console.log('DOWNLOAD: Google Drive files found:', response.data.files);

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
      await logFileStats(LOCAL_EXCEL_FILE, 'DOWNLOAD: After Download');

      const workbook = new ExcelJS.Workbook();
      await workbook.xlsx.readFile(LOCAL_EXCEL_FILE);
      const isValid = await validateWorkbook(workbook);
      if (!isValid) {
        console.log('Downloaded file from Google Drive is corrupted, extracting data before recreation...');
        const existingCustomerData = await extractExistingData(workbook);
        const existingFeedbackData = await extractExistingFeedbackData(workbook);
        const newWorkbook = await initializeExcel();
        const customerSheet = newWorkbook.getWorksheet('Customers');
        const feedbackSheet = newWorkbook.getWorksheet('Feedback');
        if (existingCustomerData.length > 0) {
          existingCustomerData.forEach(rowData => {
            const newRow = customerSheet.addRow(rowData);
            newRow.commit();
            console.log('Re-added existing Customers row after download:', rowData);
          });
        }
        if (existingFeedbackData.length > 0) {
          existingFeedbackData.forEach(rowData => {
            const newRow = feedbackSheet.addRow(rowData);
            newRow.commit();
            console.log('Re-added existing Feedback row after download:', rowData);
          });
        }
        await checkDiskSpaceAndPermissions(LOCAL_EXCEL_FILE, true);
        await newWorkbook.xlsx.writeFile(LOCAL_EXCEL_FILE);
        await logFileStats(LOCAL_EXCEL_FILE, 'DOWNLOAD: After Recreation');
        console.log('Recreated Excel file with existing data after download:', LOCAL_EXCEL_FILE);
        cachedWorkbook = newWorkbook;
      } else {
        cachedWorkbook = workbook;
      }

      const customerSheet = cachedWorkbook.getWorksheet('Customers');
      console.log('Customers sheet contents after sync:');
      console.log('Column keys:', customerSheet.columns.map(col => col.key));
      customerSheet.eachRow((row, rowNumber) => {
        if (rowNumber === 1) {
          console.log('Headers:', row.values);
        } else {
          try {
            const name = row.getCell(1).value || '';
            const email = row.getCell(2).value || '';
            const phone = row.getCell(3).value || '';
            const dob = row.getCell(4).value || '';
            console.log('Row ' + rowNumber + ':', [name, email, phone, dob]);
          } catch (error) {
            console.error(`Failed to log Customers row ${rowNumber}:`, error.message, error.stack);
          }
        }
      });

      const feedbackSheet = cachedWorkbook.getWorksheet('Feedback');
      console.log('Feedback sheet contents after sync:');
      console.log('Column keys:', feedbackSheet.columns.map(col => col.key));
      feedbackSheet.eachRow((row, rowNumber) => {
        if (rowNumber === 1) {
          console.log('Headers:', row.values);
        } else {
          try {
            const name = row.getCell(1).value || '';
            const review = row.getCell(2).value || '';
            const timestamp = row.getCell(3).value || '';
            console.log('Row ' + rowNumber + ':', [name, review, timestamp]);
          } catch (error) {
            console.error(`Failed to log Feedback row ${rowNumber}:`, error.message, error.stack);
          }
        }
      });
    } else {
      console.log('No Excel file found in Google Drive, initializing new one locally.');
      const workbook = await initializeExcel();
      await checkDiskSpaceAndPermissions(LOCAL_EXCEL_FILE, true);
      await workbook.xlsx.writeFile(LOCAL_EXCEL_FILE);
      await logFileStats(LOCAL_EXCEL_FILE, 'DOWNLOAD: After Initialization');
      cachedWorkbook = workbook;

      const customerSheet = workbook.getWorksheet('Customers');
      console.log('Newly initialized Customers sheet contents:');
      customerSheet.eachRow((row, rowNumber) => {
        const name = row.getCell(1).value || '';
        const email = row.getCell(2).value || '';
        const phone = row.getCell(3).value || '';
        const dob = row.getCell(4).value || '';
        console.log(`DOWNLOAD: Customers Row ${rowNumber}:`, [name, email, phone, dob]);
      });

      const feedbackSheet = workbook.getWorksheet('Feedback');
      console.log('Newly initialized Feedback sheet contents:');
      feedbackSheet.eachRow((row, rowNumber) => {
        const name = row.getCell(1).value || '';
        const review = row.getCell(2).value || '';
        const timestamp = row.getCell(3).value || '';
        console.log(`DOWNLOAD: Feedback Row ${rowNumber}:`, [name, review, timestamp]);
      });
    }

    // Reset localChangesPending after successful sync
    localChangesPending = false;
    console.log('DOWNLOAD: Reset localChangesPending to false after sync');
  } catch (error) {
    console.error('Error downloading from Google Drive:', error.message, error.stack);
    const workbook = await initializeExcel();
    await checkDiskSpaceAndPermissions(LOCAL_EXCEL_FILE, true);
    await workbook.xlsx.writeFile(LOCAL_EXCEL_FILE);
    await logFileStats(LOCAL_EXCEL_FILE, 'DOWNLOAD: After Error Recovery');
    cachedWorkbook = workbook;

    const customerSheet = workbook.getWorksheet('Customers');
    console.log('Customers sheet contents after error recovery:');
    customerSheet.eachRow((row, rowNumber) => {
      const name = row.getCell(1).value || '';
      const email = row.getCell(2).value || '';
      const phone = row.getCell(3).value || '';
      const dob = row.getCell(4).value || '';
      console.log(`DOWNLOAD: Customers Row ${rowNumber}:`, [name, email, phone, dob]);
    });

    const feedbackSheet = workbook.getWorksheet('Feedback');
    console.log('Feedback sheet contents after error recovery:');
    feedbackSheet.eachRow((row, rowNumber) => {
      const name = row.getCell(1).value || '';
      const review = row.getCell(2).value || '';
      const timestamp = row.getCell(3).value || '';
      console.log(`DOWNLOAD: Feedback Row ${rowNumber}:`, [name, review, timestamp]);
    });

    // Reset localChangesPending on error recovery
    localChangesPending = false;
    console.log('DOWNLOAD: Reset localChangesPending to false after error recovery');
  }
}

// Upload the local Excel file to Google Drive with retry
async function uploadToGoogleDrive(maxRetries = 3) {
  const startTime = Date.now();
  console.log(`UPLOAD: Starting upload at ${new Date(startTime).toISOString()}`);
  let retries = 0;
  while (retries < maxRetries) {
    try {
      console.log('UPLOAD: Checking local Excel file...');
      const stats = await fs.stat(LOCAL_EXCEL_FILE);
      if (stats.size === 0) {
        throw new Error('Local Excel file is empty');
      }
      console.log('UPLOAD: Local Excel file size:', stats.size, 'bytes');
      await logFileStats(LOCAL_EXCEL_FILE, 'UPLOAD: Before');

      console.log('UPLOAD: Authenticating with Google Drive...');
      const authClient = await auth.getClient();
      console.log('UPLOAD: Google Drive authentication successful:', !!authClient);
      if (!authClient) {
        throw new Error('Google Drive authentication failed');
      }

      console.log('UPLOAD: Checking folder access...');
      await drive.files.get({ fileId: GOOGLE_DRIVE_FOLDER_ID, fields: 'id, name' });
      console.log('UPLOAD: Folder access confirmed, ID:', GOOGLE_DRIVE_FOLDER_ID);

      console.log('UPLOAD: Listing existing files in Google Drive...');
      const existingFiles = await drive.files.list({
        q: `'${GOOGLE_DRIVE_FOLDER_ID}' in parents and name = 'customers.xlsx' and trashed = false`,
        fields: 'files(id, name)',
      });
      console.log('UPLOAD: Existing files found:', existingFiles.data.files);

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
        console.log('UPLOAD: Updating existing file, ID:', fileId);
        file = await drive.files.update({
          fileId: fileId,
          media: media,
          fields: 'id',
        });
        console.log('UPLOAD: Updated file in Google Drive, ID:', file.data.id);
      } else {
        console.log('UPLOAD: Creating new file in Google Drive...');
        file = await drive.files.create({
          resource: fileMetadata,
          media: media,
          fields: 'id',
        });
        console.log('UPLOAD: Created new file in Google Drive, ID:', file.data.id);
      }
      localChangesPending = false;
      console.log('UPLOAD: Upload successful, localChangesPending set to false');
      const endTime = Date.now();
      console.log(`UPLOAD: Upload completed at ${new Date(endTime).toISOString()}, took ${(endTime - startTime) / 1000} seconds`);
      return true;
    } catch (error) {
      retries++;
      if (error.code === 429) {
        console.error(`UPLOAD: Rate limit exceeded (attempt ${retries}/${maxRetries}):`, error.message);
      } else {
        console.error(`UPLOAD: Failed to upload to Google Drive (attempt ${retries}/${maxRetries}):`, error.message, error.stack);
      }
      console.error('UPLOAD: Error details:', JSON.stringify(error, null, 2));
      if (retries === maxRetries) {
        throw new Error(`Failed to upload to Google Drive after ${maxRetries} attempts: ${error.message}`);
      }
      const delay = 5000; // 5 seconds delay between retries
      console.log(`UPLOAD: Retrying after ${delay / 1000} seconds...`);
      await new Promise(resolve => setTimeout(resolve, delay));
    }
  }
  throw new Error('Failed to upload to Google Drive after maximum retries');
}

// Periodic sync with Google Drive (every 5 minutes)
function startGoogleDriveSync() {
  console.log('Starting periodic Google Drive sync (every 5 minutes)...');
  setInterval(async () => {
    const syncStartTime = Date.now();
    console.log(`SYNC: Starting sync at ${new Date(syncStartTime).toISOString()}`);
    try {
      console.log(`SYNC: Checking localChangesPending: ${localChangesPending}`);
      if (localChangesPending) {
        console.log('Periodic sync: Local changes detected, uploading to Google Drive...');
        await uploadToGoogleDrive();
        console.log('Periodic sync: Successfully uploaded to Google Drive.');
      } else {
        console.log('Periodic sync: No local changes to sync.');
      }
    } catch (error) {
      console.error('Periodic sync: Failed to sync with Google Drive:', error.message, error.stack);
      console.error('Periodic sync: Error details:', JSON.stringify(error, null, 2));
    }
    const syncEndTime = Date.now();
    console.log(`SYNC: Sync completed at ${new Date(syncEndTime).toISOString()}, took ${(syncEndTime - syncStartTime) / 1000} seconds`);
  }, 5 * 60 * 1000);
}

// Download the Excel file from Google Drive on server start
async function initializeFromGoogleDrive() {
  console.log('INIT: Initializing server with data from Google Drive...');
  // Force sync even if localChangesPending is true
  await downloadFromGoogleDrive(true);

  const fileExists = await fs.access(LOCAL_EXCEL_FILE).then(() => true).catch(() => false);
  if (!fileExists) {
    console.log('INIT: Main Excel file not found locally after download, creating a new one...');
    const workbook = await initializeExcel();
    let fileCreated = false;
    let writeAttempts = 0;
    const maxWriteAttempts = 3;

    while (writeAttempts < maxWriteAttempts && !fileCreated) {
      try {
        await checkDiskSpaceAndPermissions(LOCAL_EXCEL_FILE, true);
        await workbook.xlsx.writeFile(LOCAL_EXCEL_FILE);
        await logFileStats(LOCAL_EXCEL_FILE, 'INIT: After New File Creation');
        console.log('INIT: Initialized new Excel file on server start:', LOCAL_EXCEL_FILE);

        const fileExists = await fs.access(LOCAL_EXCEL_FILE).then(() => true).catch(() => false);
        if (!fileExists) {
          throw new Error('Failed to create Excel file on server start');
        }
        console.log('INIT: Verified: New Excel file exists on server start');
        fileCreated = true;
      } catch (writeError) {
        writeAttempts++;
        console.error(`INIT: Failed to initialize new Excel file on server start (attempt ${writeAttempts}/${maxWriteAttempts}):`, writeError.message, writeError.stack);
        if (writeAttempts === maxWriteAttempts) {
          throw new Error(`Failed to initialize new Excel file on server start after ${maxWriteAttempts} attempts: ${writeError.message}`);
        }
        await new Promise(resolve => setTimeout(resolve, 1000));
      }
    }
    cachedWorkbook = workbook;
  }

  const workbook = new ExcelJS.Workbook();
  await workbook.xlsx.readFile(LOCAL_EXCEL_FILE);
  const customerSheet = workbook.getWorksheet('Customers');
  console.log('INIT: Customers sheet contents after initialization:');
  customerSheet.eachRow((row, rowNumber) => {
    const name = row.getCell(1).value || '';
    const email = row.getCell(2).value || '';
    const phone = row.getCell(3).value || '';
    const dob = row.getCell(4).value || '';
    console.log(`INIT: Customers Row ${rowNumber}:`, [name, email, phone, dob]);
  });
  const feedbackSheet = workbook.getWorksheet('Feedback');
  console.log('INIT: Feedback sheet contents after initialization:');
  feedbackSheet.eachRow((row, rowNumber) => {
    const name = row.getCell(1).value || '';
    const review = row.getCell(2).value || '';
    const timestamp = row.getCell(3).value || '';
    console.log(`INIT: Feedback Row ${rowNumber}:`, [name, review, timestamp]);
  });
  cachedWorkbook = workbook;
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

// Handle feedback submission
app.post('/submit-feedback', async (req, res) => {
  const feedbackStartTime = Date.now();
  console.log(`FEEDBACK: Received submission at ${new Date(feedbackStartTime).toISOString()}:`, req.body);

  const { name, review } = req.body;

  if (!name || !review) {
    console.log('FEEDBACK: Validation failed: Missing required fields');
    return res.status(400).json({ success: false, error: 'Missing required fields' });
  }

  const trimmedReview = review.trim();
  if (trimmedReview.length === 0) {
    console.log('FEEDBACK: Validation failed: Review cannot be empty');
    return res.status(400).json({ success: false, error: 'Review cannot be empty' });
  }

  if (trimmedReview.length > 500) {
    console.log('FEEDBACK: Validation failed: Review exceeds 500 characters');
    return res.status(400).json({ success: false, error: 'Review cannot exceed 500 characters' });
  }

  try {
    let feedbackResult;
    fileLockPromise = fileLockPromise.then(async () => {
      try {
        let workbook;
        let feedbackSheet;

        console.log('FEEDBACK: Loading local Excel file...');
        try {
          workbook = await loadLocalExcel();
          feedbackSheet = workbook.getWorksheet('Feedback');
          console.log('FEEDBACK: Successfully loaded local Excel file:', LOCAL_EXCEL_FILE);

          console.log('FEEDBACK: Feedback sheet contents before submission:');
          feedbackSheet.eachRow((row, rowNumber) => {
            const name = row.getCell(1).value || '';
            const review = row.getCell(2).value || '';
            const timestamp = row.getCell(3).value || '';
            console.log(`FEEDBACK: Row ${rowNumber}:`, [name, review, timestamp]);
          });
        } catch (loadError) {
          console.error('FEEDBACK: Failed to load Excel file, forcing recreation:', loadError.message, loadError.stack);
          workbook = await initializeExcel();
          feedbackSheet = workbook.getWorksheet('Feedback');
          let fileCreated = false;
          let writeAttempts = 0;
          const maxWriteAttempts = 3;

          while (writeAttempts < maxWriteAttempts && !fileCreated) {
            try {
              await checkDiskSpaceAndPermissions(LOCAL_EXCEL_FILE, true);
              await workbook.xlsx.writeFile(LOCAL_EXCEL_FILE);
              await logFileStats(LOCAL_EXCEL_FILE, 'FEEDBACK: After Forced Recreation');
              console.log('FEEDBACK: Forced recreation of Excel file:', LOCAL_EXCEL_FILE);

              const fileExists = await fs.access(LOCAL_EXCEL_FILE).then(() => true).catch(() => false);
              if (!fileExists) {
                throw new Error('Failed to create Excel file during forced recreation');
              }
              console.log('FEEDBACK: Verified: Excel file exists after forced recreation');
              fileCreated = true;
            } catch (writeError) {
              writeAttempts++;
              console.error(`FEEDBACK: Failed to recreate Excel file (attempt ${writeAttempts}/${maxWriteAttempts}):`, writeError.message, writeError.stack);
              if (writeAttempts === maxWriteAttempts) {
                throw new Error(`FEEDBACK: Failed to recreate Excel file after ${maxWriteAttempts} attempts: ${writeError.message}`);
              }
              await new Promise(resolve => setTimeout(resolve, 1000));
            }
          }
          cachedWorkbook = workbook;
        }

        console.log('FEEDBACK: Adding new feedback row...');
        const timestamp = new Date().toISOString();
        const newRow = feedbackSheet.addRow([name, trimmedReview, timestamp]);
        newRow.commit();
        console.log('FEEDBACK: Added feedback row:', [name, trimmedReview, timestamp]);

        console.log('FEEDBACK: Checking disk space and permissions before saving...');
        await checkDiskSpaceAndPermissions(LOCAL_EXCEL_FILE);
        let writeAttempts = 0;
        const maxWriteAttempts = 3;
        let fileWritten = false;
        console.log('FEEDBACK: Attempting to save Excel file...');
        await logFileStats(LOCAL_EXCEL_FILE, 'FEEDBACK: Before Save');
        while (writeAttempts < maxWriteAttempts && !fileWritten) {
          try {
            await workbook.xlsx.writeFile(LOCAL_EXCEL_FILE);
            console.log('FEEDBACK: Data successfully saved to local Excel file:', LOCAL_EXCEL_FILE);
            fileWritten = true;
          } catch (writeError) {
            writeAttempts++;
            console.error(`FEEDBACK: Failed to write to Excel file (attempt ${writeAttempts}/${maxWriteAttempts}):`, writeError.message, writeError.stack);
            if (writeAttempts === maxWriteAttempts) {
              throw new Error('FEEDBACK: Failed to write to Excel file after maximum attempts');
            }
            await new Promise(resolve => setTimeout(resolve, 1000));
          }
        }
        await logFileStats(LOCAL_EXCEL_FILE, 'FEEDBACK: After Save');

        cachedWorkbook = workbook;

        localChangesPending = true;
        console.log('FEEDBACK: Local changes marked for immediate sync to Google Drive.');

        // Perform immediate sync to Google Drive
        try {
          console.log('FEEDBACK: Starting immediate sync to Google Drive...');
          await uploadToGoogleDrive();
          console.log('FEEDBACK: Immediate sync to Google Drive completed successfully.');
        } catch (syncError) {
          console.error('FEEDBACK: Immediate sync to Google Drive failed:', syncError.message, syncError.stack);
          console.error('FEEDBACK: Immediate sync error details:', JSON.stringify(syncError, null, 2));
          console.log('FEEDBACK: Changes will be synced during the next periodic sync.');
        }

        feedbackResult = { status: 200, body: { success: true, message: 'Feedback submitted successfully' } };
      } catch (error) {
        throw error;
      }
    });

    await fileLockPromise;

    if (feedbackResult) {
      console.log('FEEDBACK: Sending response:', feedbackResult.body);
      const feedbackEndTime = Date.now();
      console.log(`FEEDBACK: Submission completed at ${new Date(feedbackEndTime).toISOString()}, took ${(feedbackEndTime - feedbackStartTime) / 1000} seconds`);
      res.status(feedbackResult.status).json(feedbackResult.body);
    } else {
      throw new Error('FEEDBACK: Feedback result not set');
    }
  } catch (error) {
    console.error('FEEDBACK: Failed to save feedback to local Excel:', error.message, error.stack);
    if (error.message.includes('Insufficient disk space')) {
      res.status(500).json({ success: false, error: 'Server disk space is full. Please contact support.' });
    } else if (error.message.includes('Permission denied')) {
      res.status(500).json({ success: false, error: 'File permission error. Please contact support.' });
    } else if (error.message.includes('Corrupt')) {
      console.log('FEEDBACK: Excel file appears to be corrupted, already recreated in main flow.');
      res.status(503).json({ success: false, error: 'File was corrupted, please try again.' });
    } else if (error.message.includes('Failed to initialize new Excel file')) {
      res.status(500).json({ success: false, error: 'Failed to create Excel file. Please contact support.' });
    } else if (error.message.includes('Failed to recreate Excel file')) {
      res.status(500).json({ success: false, error: 'Failed to recreate Excel file. Please contact support.' });
    } else {
      res.status(500).json({ success: false, error: 'Unable to save your feedback. Please try again later.' });
    }
  }
});

// Handle form submission
app.post('/submit', async (req, res) => {
  const submitStartTime = Date.now();
  console.log(`SUBMIT: Received submission at ${new Date(submitStartTime).toISOString()}:`, req.body);

  const { name, email, phone, dob } = req.body;

  if (!name || !email || !phone || !dob) {
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

  const dobDate = new Date(dob);
  const today = new Date();
  if (isNaN(dobDate.getTime())) {
    console.log('SUBMIT: Validation failed: Invalid date of birth');
    return res.status(400).json({ success: false, error: 'Invalid date of birth' });
  }
  if (dobDate > today) {
    console.log('SUBMIT: Validation failed: Date of birth is in the future');
    return res.status(400).json({ success: false, error: 'Date of birth cannot be in the future' });
  }

  try {
    let submissionResult;
    fileLockPromise = fileLockPromise.then(async () => {
      try {
        let workbook;
        let sheet;

        console.log('SUBMIT: Syncing with Google Drive before duplicate check...');
        await downloadFromGoogleDrive(); // Ensure latest data before duplicate check

        console.log('SUBMIT: Loading local Excel file...');
        try {
          workbook = await loadLocalExcel();
          sheet = workbook.getWorksheet('Customers');
          console.log('SUBMIT: Successfully loaded local Excel file:', LOCAL_EXCEL_FILE);

          console.log('SUBMIT: File contents before submission:');
          sheet.eachRow((row, rowNumber) => {
            const name = row.getCell(1).value || '';
            const email = row.getCell(2).value || '';
            const phone = row.getCell(3).value || '';
            const dob = row.getCell(4).value || '';
            console.log(`SUBMIT: Row ${rowNumber}:`, [name, email, phone, dob]);
          });
        } catch (loadError) {
          console.error('SUBMIT: Failed to load Excel file, forcing recreation:', loadError.message, loadError.stack);
          workbook = await initializeExcel();
          sheet = workbook.getWorksheet('Customers');
          let fileCreated = false;
          let writeAttempts = 0;
          const maxWriteAttempts = 3;

          while (writeAttempts < maxWriteAttempts && !fileCreated) {
            try {
              await checkDiskSpaceAndPermissions(LOCAL_EXCEL_FILE, true);
              await workbook.xlsx.writeFile(LOCAL_EXCEL_FILE);
              await logFileStats(LOCAL_EXCEL_FILE, 'SUBMIT: After Forced Recreation');
              console.log('SUBMIT: Forced recreation of Excel file:', LOCAL_EXCEL_FILE);

              const fileExists = await fs.access(LOCAL_EXCEL_FILE).then(() => true).catch(() => false);
              if (!fileExists) {
                throw new Error('Failed to create Excel file during forced recreation');
              }
              console.log('SUBMIT: Verified: Excel file exists after forced recreation');
              fileCreated = true;
            } catch (writeError) {
              writeAttempts++;
              console.error(`SUBMIT: Failed to recreate Excel file (attempt ${writeAttempts}/${maxWriteAttempts}):`, writeError.message, writeError.stack);
              if (writeAttempts === maxWriteAttempts) {
                throw new Error(`SUBMIT: Failed to recreate Excel file after ${maxWriteAttempts} attempts: ${writeError.message}`);
              }
              await new Promise(resolve => setTimeout(resolve, 1000));
            }
          }
          cachedWorkbook = workbook;
        }

        console.log('SUBMIT: Collecting existing rows before adding new row...');
        const existingRows = [];
        sheet.eachRow((row, rowNumber) => {
          if (rowNumber === 1) return;
          const name = row.getCell(1).value || '';
          const email = row.getCell(2).value || '';
          const phone = row.getCell(3).value || '';
          const dob = row.getCell(4).value || '';
          if (name && email && phone) {
            existingRows.push([name, email, phone, dob]);
            console.log(`SUBMIT: Existing Row ${rowNumber}:`, [name, email, phone, dob]);
          }
        });

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

        existingRows.push([name, email, phone, dob]);
        console.log('SUBMIT: Added new row to list:', [name, email, phone, dob]);

        console.log('SUBMIT: Recreating worksheet with contiguous rows...');
        workbook.removeWorksheet('Customers');
        const newSheet = workbook.addWorksheet('Customers');
        newSheet.columns = [
          { header: 'Name', key: 'name', width: 20 },
          { header: 'Email', key: 'email', width: 30 },
          { header: 'Phone', key: 'phone', width: 15 },
          { header: 'Date of Birth', key: 'dob', width: 15 },
        ];

        existingRows.forEach((rowValues, index) => {
          const newRow = newSheet.addRow(rowValues);
          newRow.commit();
          console.log(`SUBMIT: Added row ${index + 2} to new worksheet:`, rowValues);
        });

        console.log('SUBMIT: Checking disk space and permissions before saving...');
        await checkDiskSpaceAndPermissions(LOCAL_EXCEL_FILE);
        let writeAttempts = 0;
        const maxWriteAttempts = 3;
        let fileWritten = false;
        console.log('SUBMIT: Attempting to save Excel file...');
        await logFileStats(LOCAL_EXCEL_FILE, 'SUBMIT: Before Save');
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
            await new Promise(resolve => setTimeout(resolve, 1000));
          }
        }
        await logFileStats(LOCAL_EXCEL_FILE, 'SUBMIT: After Save');

        cachedWorkbook = workbook;

        localChangesPending = true;
        console.log('SUBMIT: Local changes marked for immediate sync to Google Drive.');

        // Perform immediate sync to Google Drive
        try {
          console.log('SUBMIT: Starting immediate sync to Google Drive...');
          await uploadToGoogleDrive();
          console.log('SUBMIT: Immediate sync to Google Drive completed successfully.');
        } catch (syncError) {
          console.error('SUBMIT: Immediate sync to Google Drive failed:', syncError.message, syncError.stack);
          console.error('SUBMIT: Immediate sync error details:', JSON.stringify(syncError, null, 2));
          console.log('SUBMIT: Changes will be synced during the next periodic sync.');
        }

        submissionResult = { status: 200, body: { success: true, name } };
      } catch (error) {
        throw error;
      }
    });

    await fileLockPromise;

    if (submissionResult) {
      console.log('SUBMIT: Sending response:', submissionResult.body);
      const submitEndTime = Date.now();
      console.log(`SUBMIT: Submission completed at ${new Date(submitEndTime).toISOString()}, took ${(submitEndTime - submitStartTime) / 1000} seconds`);
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
    } else if (error.message.includes('Failed to initialize new Excel file')) {
      res.status(500).json({ success: false, error: 'Failed to create Excel file. Please contact support.' });
    } else if (error.message.includes('Failed to recreate Excel file')) {
      res.status(500).json({ success: false, error: 'Failed to recreate Excel file. Please contact support.' });
    } else {
      res.status(500).json({ success: false, error: 'Unable to save your submission. Please try again later.' });
    }
  }
});

// Handle row deletion
app.post('/delete', async (req, res) => {
  const deleteStartTime = Date.now();
  console.log(`DELETE: Received delete request at ${new Date(deleteStartTime).toISOString()}:`, req.body);

  const { email, phone } = req.body;

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

          console.log('DELETE: File contents before deletion:');
          sheet.eachRow((row, rowNumber) => {
            const name = row.getCell(1).value || '';
            const email = row.getCell(2).value || '';
            const phone = row.getCell(3).value || '';
            const dob = row.getCell(4).value || '';
            console.log(`DELETE: Row ${rowNumber}:`, [name, email, phone, dob]);
          });
        } catch (loadError) {
          console.error('DELETE: Failed to load Excel file for deletion, forcing recreation:', loadError.message, loadError.stack);
          workbook = await initializeExcel();
          sheet = workbook.getWorksheet('Customers');
          await checkDiskSpaceAndPermissions(LOCAL_EXCEL_FILE, true);
          await workbook.xlsx.writeFile(LOCAL_EXCEL_FILE);
          await logFileStats(LOCAL_EXCEL_FILE, 'DELETE: After Forced Recreation');
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
            console.log(`DELETE: Found matching row ${rowNumber} to delete:`, [row.getCell(1).value, normalizedExistingEmail, normalizedExistingPhone, row.getCell(4).value]);
            rowFound = true;
          } else {
            rowsToKeep.push([row.getCell(1).value, row.getCell(2).value, row.getCell(3).value, row.getCell(4).value]);
          }
        });

        if (!rowFound) {
          console.log('DELETE: No matching row found for deletion');
          deletionResult = { status: 404, body: { success: false, error: 'Customer not found' } };
          return;
        }

        console.log('DELETE: Rows to keep after deletion:', rowsToKeep);

        workbook.removeWorksheet('Customers');
        const newSheet = workbook.addWorksheet('Customers');
        newSheet.columns = [
          { header: 'Name', key: 'name', width: 20 },
          { header: 'Email', key: 'email', width: 30 },
          { header: 'Phone', key: 'phone', width: 15 },
          { header: 'Date of Birth', key: 'dob', width: 15 },
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
        await logFileStats(LOCAL_EXCEL_FILE, 'DELETE: Before Save');
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
            await new Promise(resolve => setTimeout(resolve, 1000));
          }
        }
        await logFileStats(LOCAL_EXCEL_FILE, 'DELETE: After Save');

        cachedWorkbook = workbook;

        localChangesPending = true;
        console.log('DELETE: Local changes marked for immediate sync to Google Drive.');

        // Perform immediate sync to Google Drive
        try {
          console.log('DELETE: Starting immediate sync to Google Drive...');
          await uploadToGoogleDrive();
          console.log('DELETE: Immediate sync to Google Drive completed successfully.');
        } catch (syncError) {
          console.error('DELETE: Immediate sync to Google Drive failed:', syncError.message, syncError.stack);
          console.error('DELETE: Immediate sync error details:', JSON.stringify(syncError, null, 2));
          console.log('DELETE: Changes will be synced during the next periodic sync.');
        }

        deletionResult = { status: 200, body: { success: true, message: 'Customer deleted successfully' } };
      } catch (error) {
        throw error;
      }
    });

    await fileLockPromise;

    if (deletionResult) {
      console.log('DELETE: Sending response:', deletionResult.body);
      const deleteEndTime = Date.now();
      console.log(`DELETE: Deletion completed at ${new Date(deleteEndTime).toISOString()}, took ${(deleteEndTime - deleteStartTime) / 1000} seconds`);
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
    let downloadResult;
    fileLockPromise = fileLockPromise.then(async () => {
      try {
        const fileExists = await fs.access(LOCAL_EXCEL_FILE).then(() => true).catch(() => false);
        if (!fileExists) {
          console.log('DOWNLOAD: No customer data available yet');
          downloadResult = { status: 404, body: 'No customer data available yet' };
          return;
        }

        console.log('DOWNLOAD: Reading file contents before sending...');
        const workbook = new ExcelJS.Workbook();
        await workbook.xlsx.readFile(LOCAL_EXCEL_FILE);
        const customerSheet = workbook.getWorksheet('Customers');
        console.log('DOWNLOAD: Customers sheet contents before download:');
        customerSheet.eachRow((row, rowNumber) => {
          const name = row.getCell(1).value || '';
          const email = row.getCell(2).value || '';
          const phone = row.getCell(3).value || '';
          const dob = row.getCell(4).value || '';
          console.log(`DOWNLOAD: Customers Row ${rowNumber}:`, [name, email, phone, dob]);
        });
        const feedbackSheet = workbook.getWorksheet('Feedback');
        console.log('DOWNLOAD: Feedback sheet contents before download:');
        feedbackSheet.eachRow((row, rowNumber) => {
          const name = row.getCell(1).value || '';
          const review = row.getCell(2).value || '';
          const timestamp = row.getCell(3).value || '';
          console.log(`DOWNLOAD: Feedback Row ${rowNumber}:`, [name, review, timestamp]);
        });

        console.log('DOWNLOAD: Sending Excel file to client...');
        res.setHeader('Content-Disposition', 'attachment; filename=customers.xlsx');
        res.setHeader('Content-Type', 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet');
        const fileStream = require('fs').createReadStream(LOCAL_EXCEL_FILE);
        fileStream.pipe(res);
        console.log('DOWNLOAD: File stream initiated.');
      } catch (error) {
        throw error;
      }
    });

    await fileLockPromise;

    if (downloadResult) {
      res.status(downloadResult.status).send(downloadResult.body);
    }
  } catch (error) {
    console.error('DOWNLOAD: Error downloading local file:', error.message, error.stack);
    res.status(500).send('Error downloading file');
  }
});

// Initialize the server
(async () => {
  console.log('SERVER: Initializing server...');
  await initializeFromGoogleDrive();
  startGoogleDriveSync();
  app.listen(PORT, () => {
    console.log(`SERVER: Server running on port ${PORT}`);
  });
})();
