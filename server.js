const express = require('express');
const bodyParser = require('body-parser');
const ExcelJS = require('exceljs');
const path = require('path');
const fs = require('fs').promises;
const { google } = require('googleapis');

const app = express();
const PORT = process.env.PORT || 3000;
const LOCAL_EXCEL_FILE = path.join(__dirname, 'customers.xlsx');
const GOOGLE_DRIVE_FOLDER_ID = '1Vila0sI9fAaAxp17_IZmbbkcegOPGJLD';

let isFileWriting = false;

const auth = new google.auth.GoogleAuth({
  credentials: JSON.parse(process.env.GOOGLE_SERVICE_ACCOUNT),
  scopes: ['https://www.googleapis.com/auth/drive'],
});
const drive = google.drive({ version: 'v3', auth });

app.use(bodyParser.json());
app.use(express.static(__dirname));

async function initializeExcel() {
  const workbook = new ExcelJS.Workbook();
  const sheet = workbook.addWorksheet('Customers', {
    properties: { defaultColWidth: 20 }
  });
  sheet.columns = [
    { header: 'Name', key: 'name', width: 20 },
    { header: 'Email', key: 'email', width: 30 },
    { header: 'Phone', key: 'phone', width: 15 },
    { header: 'Date', key: 'date', width: 15 }, // New Date column
  ];
  sheet.addRow(['Name', 'Email', 'Phone', 'Date']);
  return workbook;
}

async function loadLocalExcel() {
  let workbook;
  try {
    const fileExists = await fs.access(LOCAL_EXCEL_FILE).then(() => true).catch(() => false);
    if (!fileExists) {
      throw new Error('Excel file does not exist. Creating a new one.');
    }

    const fileStats = await fs.stat(LOCAL_EXCEL_FILE);
    if (fileStats.size < 1000) {
      throw new Error('Excel file is too small or empty. Recreating the file.');
    }

    workbook = new ExcelJS.Workbook();
    await workbook.xlsx.readFile(LOCAL_EXCEL_FILE);
    console.log('Loaded local Excel file:', LOCAL_EXCEL_FILE);

    const sheet = workbook.getWorksheet('Customers');
    if (!sheet) {
      throw new Error('Worksheet "Customers" not found in the Excel file.');
    }

    if (sheet.rowCount === 0) {
      throw new Error('Worksheet is empty. Recreating the file.');
    }

    if (sheet.columnCount > 16384) {
      throw new Error('Excel file has too many columns. Recreating the file.');
    }

    const expectedColumns = ['Name', 'Email', 'Phone', 'Date'];
    const actualColumns = sheet.getRow(1).values?.slice(1) || [];
    if (!expectedColumns.every((col, idx) => actualColumns[idx] === col)) {
      console.log('Invalid column structure detected:', actualColumns);
      throw new Error('Invalid column structure.');
    }

    let rowData = [];
    sheet.eachRow((row, rowNumber) => {
      if (rowNumber > 1) {
        const name = row.getCell(1)?.value?.toString().trim();
        const email = row.getCell(2)?.value?.toString().trim();
        const phone = row.getCell(3)?.value?.toString().trim();
        const date = row.getCell(4)?.value; // Date might be a string or Date object
        if (!name || !email || !phone) {
          console.log(`Invalid data in row ${rowNumber}:`, { name, email, phone, date });
          throw new Error(`Invalid data in row ${rowNumber}`);
        }
        rowData.push({ name, email, phone, date });
      }
    });

    return { workbook, rowData };
  } catch (error) {
    console.log('Error loading Excel file, initializing a new one:', error.message);
    workbook = await initializeExcel();
    await workbook.xlsx.writeFile(LOCAL_EXCEL_FILE);
    console.log('Created new Excel file:', LOCAL_EXCEL_FILE);
    return { workbook, rowData: [] };
  }
}

async function uploadToGoogleDrive() {
  if (isFileWriting) {
    console.log('File is being written, skipping Google Drive sync.');
    return;
  }

  try {
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
    console.error('Failed to upload to Google Drive:', error.message);
    throw error;
  }
}

function startGoogleDriveSync() {
  setInterval(async () => {
    try {
      console.log('Starting periodic sync with Google Drive...');
      await uploadToGoogleDrive();
      console.log('Periodic sync completed.');
    } catch (error) {
      console.error('Periodic sync failed:', error.message);
    }
  }, 5 * 60 * 1000);
}

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
      await workbook.xlsx.writeFile(LOCAL_EXCEL_FILE);
    }
  } catch (error) {
    console.error('Error initializing from Google Drive:', error.message);
    const workbook = await initializeExcel();
    await workbook.xlsx.writeFile(LOCAL_EXCEL_FILE);
  }
}

async function checkDuplicates(email, phone) {
  const { workbook, rowData } = await loadLocalExcel();
  let duplicateField = null;

  const normalizedEmail = email.toString().trim().toLowerCase();
  const normalizedPhone = phone.toString().trim();

  console.log('Checking for duplicates with:', { email: normalizedEmail, phone: normalizedPhone });

  for (const row of rowData) {
    const existingEmail = row.email.toLowerCase();
    const existingPhone = row.phone;

    console.log('Existing row:', { existingEmail, existingPhone });

    if (existingEmail === normalizedEmail) {
      duplicateField = 'email';
      break;
    } else if (existingPhone === normalizedPhone) {
      duplicateField = 'phone';
      break;
    }
  }

  return { duplicateField, workbook };
}

const delay = (ms) => new Promise(resolve => setTimeout(resolve, ms));

app.post('/submit', async (req, res) => {
  let responseSent = false;
  try {
    const { name, email, phone } = req.body;
    console.log('Received submission:', { name, email, phone });

    // Validate input
    if (!name || !email || !phone) {
      console.log('Validation failed: Missing required fields');
      responseSent = true;
      return res.status(400).json({ success: false, error: 'Missing required fields' });
    }

    if (!/^\d{10}$/.test(phone)) {
      console.log('Validation failed: Invalid phone number');
      responseSent = true;
      return res.status(400).json({ success: false, error: 'Invalid phone number (10 digits required)' });
    }

    // Check for duplicates
    let { duplicateField, workbook } = await checkDuplicates(email, phone);
    if (duplicateField) {
      console.log(`Duplicate ${duplicateField} detected:`, duplicateField === 'email' ? email : phone);
      responseSent = true;
      return res.status(409).json({
        success: false,
        error: `Details already exist with this ${duplicateField}. One entry per customer!`
      });
    }

    // Prepare and add new row with current date
    let sheet = workbook.getWorksheet('Customers');
    const nameStr = String(name).trim();
    const emailStr = String(email).trim();
    const phoneStr = String(phone).trim();
    const dateStr = new Date().toLocaleDateString(); // e.g., "4/06/2025"
    const newRow = sheet.addRow({ name: nameStr, email: emailStr, phone: phoneStr, date: dateStr });
    console.log('Added new row:', { name: nameStr, email: emailStr, phone: phoneStr, date: dateStr, rowNumber: newRow.number });

    // Write to file
    isFileWriting = true;
    const buffer = await workbook.xlsx.writeBuffer();
    await fs.writeFile(LOCAL_EXCEL_FILE, buffer);
    console.log('Data written directly to:', LOCAL_EXCEL_FILE);

    // Validate the written data
    const validationWorkbook = new ExcelJS.Workbook();
    await validationWorkbook.xlsx.readFile(LOCAL_EXCEL_FILE);
    const validationSheet = validationWorkbook.getWorksheet('Customers');
    const writtenRow = validationSheet.getRow(newRow.number);
    if (writtenRow) {
      const lastName = String(writtenRow.getCell(1)?.value || '').trim();
      const lastEmail = String(writtenRow.getCell(2)?.value || '').trim();
      const lastPhone = String(writtenRow.getCell(3)?.value || '').trim();
      const lastDate = String(writtenRow.getCell(4)?.value || '').trim();
      console.log('Validated written row:', { lastName, lastEmail, lastPhone, lastDate });
      if (lastName !== nameStr || lastEmail !== emailStr || lastPhone !== phoneStr || lastDate !== dateStr) {
        console.log('Mismatch detected - Submitted:', { name: nameStr, email: emailStr, phone: phoneStr, date: dateStr });
        console.log('Mismatch detected - Written:', { lastName, lastEmail, lastPhone, lastDate });
        throw new Error('Last row does not match the submitted data.');
      }
    } else {
      throw new Error('Written row not found after write.');
    }

    // Sync with Google Drive
    await delay(3000);
    await uploadToGoogleDrive();

    responseSent = true;
    res.status(200).json({ success: true, name });
  } catch (error) {
    console.error('Failed to process submission:', error.message);
    if (!responseSent) {
      responseSent = true;
      res.status(500).json({ success: false, error: `Failed to save data: ${error.message}` });
    }
  } finally {
    isFileWriting = false;
  }
});

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
    console.error('Error downloading local file:', error.message);
    res.status(500).send('Error downloading file');
  }
});

(async () => {
  await initializeFromGoogleDrive();
  startGoogleDriveSync();

  app.listen(PORT, () => {
    console.log(`Server running on port ${PORT}`);
  });
})();
