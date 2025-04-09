const express = require('express');
const bodyParser = require('body-parser');
const ExcelJS = require('exceljs');
const path = require('path');
const fs = require('fs').promises;
const { google } = require('googleapis');

const app = express();
const PORT = process.env.PORT || 10000;
const LOCAL_EXCEL_FILE = path.join(__dirname, 'customers.xlsx');
const GOOGLE_DRIVE_FOLDER_ID = '1nSJaqKO2q3bj3QhGq0Qf_nEO1sN-kqwW';

let isFileWriting = false;

const auth = new google.auth.GoogleAuth({
  credentials: JSON.parse(process.env.GOOGLE_SERVICE_ACCOUNT || '{}'),
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
    { header: 'Date', key: 'date', width: 15 },
  ];
  await workbook.xlsx.writeFile(LOCAL_EXCEL_FILE);
  console.log('Initialized fresh Excel file:', LOCAL_EXCEL_FILE);
  return { workbook, rowData: [] };
}

async function loadLocalExcel() {
  let workbook;
  try {
    const fileExists = await fs.access(LOCAL_EXCEL_FILE).then(() => true).catch(() => false);
    if (!fileExists) {
      console.log('Excel file not found locally. Creating new one.');
      return await initializeExcel();
    }

    workbook = new ExcelJS.Workbook();
    await workbook.xlsx.readFile(LOCAL_EXCEL_FILE);
    console.log('Loaded local Excel file:', LOCAL_EXCEL_FILE);

    const sheet = workbook.getWorksheet('Customers');
    if (!sheet) {
      console.log('Worksheet "Customers" not found, reinitializing.');
      return await initializeExcel();
    }

    const expectedColumns = ['Name', 'Email', 'Phone', 'Date'];
    const actualColumns = sheet.getRow(1).values?.slice(1) || [];
    if (!expectedColumns.every((col, idx) => actualColumns[idx] === col)) {
      console.log('Invalid column structure, reinitializing file.');
      return await initializeExcel();
    }

    let rowData = [];
    sheet.eachRow((row, rowNumber) => {
      if (rowNumber === 1) return; // Skip header row
      const name = String(row.getCell(1)?.value || '').trim();
      const email = String(row.getCell(2)?.value || '').trim();
      const phone = String(row.getCell(3)?.value || '').trim();
      const date = String(row.getCell(4)?.value || '').trim();
      // Skip any row that looks like a duplicate header
      if (name === 'Name' && email === 'Email' && phone === 'Phone' && date === 'Date') {
        console.log('Skipping duplicate header row at line', rowNumber);
        return;
      }
      if (name && email && phone) {
        rowData.push({ name, email, phone, date });
      }
    });

    return { workbook, rowData };
  } catch (error) {
    console.error('Error loading Excel file:', error.message);
    return await initializeExcel();
  }
}

async function uploadToGoogleDrive() {
  if (isFileWriting) {
    console.log('File is being written, skipping Google Drive sync.');
    return;
  }

  try {
    const fileExists = await fs.access(LOCAL_EXCEL_FILE).then(() => true).catch(() => false);
    if (!fileExists) {
      console.error('Local Excel file not found for upload:', LOCAL_EXCEL_FILE);
      throw new Error('Local Excel file does not exist.');
    }

    const fileStats = await fs.stat(LOCAL_EXCEL_FILE);
    if (fileStats.size < 100) {
      console.error('Local Excel file is empty or too small:', LOCAL_EXCEL_FILE);
      throw new Error('Local Excel file is empty.');
    }

    console.log('Uploading to Google Drive...');
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
    console.error('Google Drive upload failed:', error.message);
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
  }, 5 * 60 * 1000); // 5-minute backup sync
}

async function initializeFromGoogleDrive() {
  try {
    await fs.unlink(LOCAL_EXCEL_FILE).catch(() => console.log('No local file to delete or already deleted:', LOCAL_EXCEL_FILE));

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
      console.log('Downloaded Excel file from Google Drive:', LOCAL_EXCEL_FILE);
    } else {
      console.log('No Excel file in Google Drive, initializing fresh locally.');
      await initializeExcel();
    }
  } catch (error) {
    console.error('Error during initialization from Google Drive:', error.message);
    await initializeExcel();
  }
}

async function checkDuplicates(email, phone) {
  const { workbook, rowData } = await loadLocalExcel();
  let duplicateField = null;

  const normalizedEmail = email.toString().trim().toLowerCase();
  const normalizedPhone = phone.toString().trim();

  for (const row of rowData) {
    if (row.email.toLowerCase() === normalizedEmail) {
      duplicateField = 'email';
      break;
    } else if (row.phone === normalizedPhone) {
      duplicateField = 'phone';
      break;
    }
  }

  return { duplicateField, workbook };
}

app.post('/submit', async (req, res) => {
  let responseSent = false;
  try {
    const { name, email, phone } = req.body;
    console.log('Received submission:', { name, email, phone });

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

    let { duplicateField, workbook } = await checkDuplicates(email, phone);
    if (duplicateField) {
      console.log(`Duplicate ${duplicateField} detected:`, duplicateField === 'email' ? email : phone);
      responseSent = true;
      return res.status(409).json({
        success: false,
        error: `Details already exist with this ${duplicateField}. One entry per customer!`
      });
    }

    let sheet = workbook.getWorksheet('Customers');
    const nameStr = String(name).trim();
    const emailStr = String(email).trim();
    const phoneStr = String(phone).trim();
    const dateStr = new Date().toISOString().split('T')[0]; // e.g., "2025-04-07"
    sheet.addRow({ name: nameStr, email: emailStr, phone: phoneStr, date: dateStr });
    console.log('Added new row:', { name: nameStr, email: emailStr, phone: phoneStr, date: dateStr });

    isFileWriting = true;
    await workbook.xlsx.writeFile(LOCAL_EXCEL_FILE);
    console.log('Data written to local file:', LOCAL_EXCEL_FILE);
    isFileWriting = false;

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
