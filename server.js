const express = require('express');
const bodyParser = require('body-parser');
const ExcelJS = require('exceljs');
const path = require('path');
const fs = require('fs').promises;
const { google } = require('googleapis');
const { promisify } = require('util');
const lockfile = require('lockfile');
const compression = require('compression');

const app = express();
const PORT = process.env.PORT || 10000;
const LOCAL_EXCEL_FILE = path.join(__dirname, 'customers.xlsx');
const TEMP_EXCEL_FILE = path.join(__dirname, 'customers_temp.xlsx');
const LOCK_FILE = path.join(__dirname, 'file.lock');
const GOOGLE_DRIVE_FOLDER_ID = '1l4e6cq0LaFS2IFkJlWKLFJ_CVIEqPqTK';

const auth = new google.auth.GoogleAuth({
  credentials: JSON.parse(process.env.GOOGLE_SERVICE_ACCOUNT || '{}'),
  scopes: ['https://www.googleapis.com/auth/drive'],
});
const drive = google.drive({ version: 'v3', auth });

app.use(compression());
app.use(bodyParser.json());
app.use(express.static(__dirname, {
  maxAge: '1d',
  extensions: ['html'],
  setHeaders: (res, path) => {
    console.log(`Serving static file: ${path}`);
  }
}));

app.use((req, res, next) => {
  console.log(`[${new Date().toISOString()}] ${req.method} ${req.url}`);
  next();
});

app.get('/health', (req, res) => {
  console.log('Health check requested');
  res.status(200).send('Server is running');
});

let cachedWorkbook = null;

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
    { header: 'Prize', key: 'prize', width: 20 },
  ];
  await workbook.xlsx.writeFile(LOCAL_EXCEL_FILE);
  console.log('Initialized fresh Excel file:', LOCAL_EXCEL_FILE);
  return workbook;
}

async function loadFromGoogleDrive() {
  if (cachedWorkbook) {
    console.log('Using cached workbook');
    return cachedWorkbook;
  }

  try {
    const response = await drive.files.list({
      q: `'${GOOGLE_DRIVE_FOLDER_ID}' in parents and name = 'customers.xlsx' and trashed = false`,
      fields: 'files(id)',
    });

    let workbook;
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

      workbook = new ExcelJS.Workbook();
      try {
        await workbook.xlsx.readFile(LOCAL_EXCEL_FILE);
      } catch (error) {
        console.error('Corrupted Excel file, reinitializing:', error.message);
        workbook = await initializeExcel();
      }

      const sheet = workbook.getWorksheet('Customers');
      if (!sheet || !sheet.columns.length) {
        console.warn('Invalid Customers sheet, reinitializing');
        workbook = await initializeExcel();
      } else {
        const expectedColumns = ['Name', 'Email', 'Phone', 'Date', 'Prize'];
        const actualColumns = sheet.getRow(1).values.slice(1);
        if (!expectedColumns.every((col, i) => actualColumns[i] === col)) {
          console.warn('Invalid column headers, reinitializing');
          workbook = await initializeExcel();
        }
      }
    } else {
      console.log('No Excel file in Google Drive, initializing fresh.');
      workbook = await initializeExcel();
    }

    cachedWorkbook = workbook;
    return workbook;
  } catch (error) {
    console.error('Error loading from Google Drive:', error.stack);
    cachedWorkbook = await initializeExcel();
    return cachedWorkbook;
  }
}

async function uploadToGoogleDrive() {
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
        fields: 'id, size',
      });
      console.log('Updated file in Google Drive, ID:', file.data.id, 'Size:', file.data.size);
    } else {
      file = await drive.files.create({
        resource: fileMetadata,
        media: media,
        fields: 'id, size',
      });
      console.log('Created new file in Google Drive, ID:', file.data.id, 'Size:', file.data.size);
    }
  } catch (error) {
    console.error('Google Drive upload failed:', error.stack);
    throw error;
  }
}

async function withFileLock(operation) {
  try {
    await promisify(lockfile.lock)(LOCK_FILE, { wait: 10000 });
    return await operation();
  } finally {
    await promisify(lockfile.unlock)(LOCK_FILE);
  }
}

async function saveWorkbook() {
  if (!cachedWorkbook) return;
  await withFileLock(async () => {
    await cachedWorkbook.xlsx.writeFile(TEMP_EXCEL_FILE);
    await fs.rename(TEMP_EXCEL_FILE, LOCAL_EXCEL_FILE);
    await uploadToGoogleDrive();
  });
}

const saveQueue = require('async').queue(async (task, callback) => {
  try {
    await task();
    callback();
  } catch (error) {
    console.error('Queue error:', error);
    callback(error);
  }
}, 1);

async function checkDuplicates(email, phone, workbook) {
  const sheet = workbook.getWorksheet('Customers');
  let duplicateField = null;

  const normalizedEmail = email.toString().trim().toLowerCase().replace(/\s+/g, '');
  const normalizedPhone = phone.toString().trim().replace(/[-\s]/g, '');

  sheet.eachRow((row, rowNumber) => {
    if (rowNumber === 1) return;
    const rowEmail = String(row.getCell(2)?.value || '').trim().toLowerCase().replace(/\s+/g, '');
    const rowPhone = String(row.getCell(3)?.value || '').trim().replace(/[-\s]/g, '');
    if (rowEmail && rowEmail === normalizedEmail) {
      duplicateField = 'email';
    } else if (rowPhone && rowPhone === normalizedPhone) {
      duplicateField = 'phone';
    }
  });

  return duplicateField;
}

app.post('/submit', async (req, res) => {
  const start = Date.now();
  let responseSent = false;
  try {
    console.log('Raw request body:', req.body);
    const { name, email, phone } = req.body;

    if (!name || !email || !phone) {
      console.log('Validation failed: Missing required fields');
      responseSent = true;
      return res.status(400).json({ success: false, error: 'Missing required fields' });
    }

    console.log('Validating email:', email);
    if (!/^[a-zA-Z0-9._%+-]+@[a-zA-Z0-9.-]+\.[a-zA-Z]{2,}$/.test(email)) {
      console.log('Validation failed: Invalid email format:', email);
      responseSent = true;
      return res.status(400).json({ success: false, error: 'Invalid email format' });
    }

    console.log('Validating phone:', phone);
    if (!/^\d{10}$/.test(phone)) {
      console.log('Validation failed: Invalid phone number:', phone);
      responseSent = true;
      return res.status(400).json({ success: false, error: 'Invalid phone number (10 digits required)' });
    }

    const result = await withFileLock(async () => {
      const workbook = await loadFromGoogleDrive();
      const sheet = workbook.getWorksheet('Customers');

      const duplicateField = await checkDuplicates(email, phone, workbook);
      if (duplicateField) {
        console.log(`Duplicate ${duplicateField} detected:`, duplicateField === 'email' ? email : phone);
        throw new Error(`Details already exist with this ${duplicateField}. One entry per customer!`);
      }

      const nameStr = String(name).trim();
      const emailStr = String(email).trim();
      const phoneStr = String(phone).trim();
      const dateStr = new Date().toISOString().split('T')[0];

      if (!sheet.columns.length) {
        sheet.columns = [
          { header: 'Name', key: 'name', width: 20 },
          { header: 'Email', key: 'email', width: 30 },
          { header: 'Phone', key: 'phone', width: 15 },
          { header: 'Date', key: 'date', width: 15 },
          { header: 'Prize', key: 'prize', width: 20 },
        ];
      }

      console.log('Adding new row to sheet');
      const newRow = sheet.addRow();
      newRow.getCell(1).value = nameStr;
      newRow.getCell(2).value = emailStr;
      newRow.getCell(3).value = phoneStr;
      newRow.getCell(4).value = dateStr;
      newRow.commit();
      console.log('Added row values:', [newRow.getCell(1).value, newRow.getCell(2).value, newRow.getCell(3).value, newRow.getCell(4).value]);

      saveQueue.push(saveWorkbook);
      return { name: nameStr };
    });

    responseSent = true;
    res.status(200).json({ success: true, name: result.name });
    console.log(`Submit took ${Date.now() - start}ms`);
  } catch (error) {
    console.error('Failed to process submission:', error.stack);
    if (!responseSent) {
      responseSent = true;
      res.status(error.message.includes('already exist') ? 409 : 500).json({
        success: false,
        error: error.message,
      });
    }
    console.log(`Submit took ${Date.now() - start}ms`);
  }
});

app.post('/save-prize', async (req, res) => {
  const start = Date.now();
  let responseSent = false;
  try {
    console.log('Save-prize endpoint hit with body:', req.body);
    const { email, prize } = req.body;

    if (!email || !prize) {
      console.log('Validation failed: Missing email or prize');
      responseSent = true;
      return res.status(400).json({ success: false, error: 'Missing email or prize' });
    }

    const validPrizes = ['Free Dip', 'Free Cookie', 'Free Can', 'Free Chipsbag'];
    if (!validPrizes.includes(prize)) {
      console.log('Validation failed: Invalid prize:', prize);
      responseSent = true;
      return res.status(400).json({ success: false, error: `Invalid prize: ${prize}` });
    }

    await withFileLock(async () => {
      const workbook = await loadFromGoogleDrive();
      const sheet = workbook.getWorksheet('Customers');

      let found = false;
      sheet.eachRow((row, rowNumber) => {
        if (rowNumber === 1) return;
        if (String(row.getCell(2).value).trim().toLowerCase() === email.toLowerCase()) {
          row.getCell(5).value = prize;
          row.commit();
          found = true;
          console.log(`Updated prize for email ${email}: ${prize}`);
        }
      });

      if (!found) {
        console.log('Email not found:', email);
        throw new Error('Email not found');
      }

      saveQueue.push(saveWorkbook);
    });

    responseSent = true;
    res.status(200).json({ success: true });
    console.log(`Save-prize took ${Date.now() - start}ms`);
  } catch (error) {
    console.error('Failed to save prize:', error.stack);
    if (!responseSent) {
      responseSent = true;
      res.status(error.message === 'Email not found' ? 404 : 500).json({
        success: false,
        error: error.message,
      });
    }
    console.log(`Save-prize took ${Date.now() - start}ms`);
  }
});

app.get('/download', async (req, res) => {
  try {
    const workbook = await loadFromGoogleDrive();
    await withFileLock(async () => {
      await workbook.xlsx.writeFile(LOCAL_EXCEL_FILE);
    });
    const fileExists = await fs.access(LOCAL_EXCEL_FILE).then(() => true).catch(() => false);
    if (!fileExists) {
      return res.status(404).send('No customer data available yet');
    }

    res.setHeader('Content-Disposition', 'attachment; filename=customers.xlsx');
    res.setHeader('Content-Type', 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet');
    const fileStream = require('fs').createReadStream(LOCAL_EXCEL_FILE);
    fileStream.pipe(res);
  } catch (error) {
    console.error('Error downloading local file:', error.stack);
    res.status(500).send('Error downloading file');
  }
});

app.use((req, res, next) => {
  console.log(`404 Not Found: ${req.method} ${req.url}`);
  res.status(404).send('Not Found');
});

process.on('uncaughtException', (error) => {
  console.error('Uncaught Exception:', error.stack);
});

process.on('unhandledRejection', (reason, promise) => {
  console.error('Unhandled rejection at:', promise, 'reason:', reason.stack || reason);
});

(async () => {
  try {
    await loadFromGoogleDrive();
    app.listen(PORT, () => {
      console.log(`Server running on port ${PORT}`);
    });
  } catch (error) {
    console.error('Startup failed:', error.stack);
    process.exit(1);
  }
})();
