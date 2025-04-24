const express = require('express');
const bodyParser = require('body-parser');
const ExcelJS = require('exceljs');
const path = require('path');
const fs = require('fs').promises;
const { google } = require('googleapis');
const { promisify } = require('util');
const lockfile = require('lockfile');
const compression = require('compression');
const cors = require('cors');

const app = express();
const PORT = process.env.PORT || 10000;
const LOCAL_EXCEL_FILE = path.join(__dirname, 'customers.xlsx');
const TEMP_EXCEL_FILE = path.join(__dirname, 'customers_temp.xlsx');
const LOCK_FILE = path.join(__dirname, 'file.lock');
const GOOGLE_DRIVE_FOLDER_ID = '1l4e6cq0LaFS2IFkJlWKLFJ_CVIEqPqTK';

const auth = new google.auth.GoogleAuth({
  credentials: JSON.parse(process.env.GOOGLE_CREDENTIALS || '{}'),
  scopes: ['https://www.googleapis.com/auth/drive'],
});
const drive = google.drive({ version: 'v3', auth });

app.use(compression());
app.use(cors());
app.use(bodyParser.json());

app.use((req, res, next) => {
  console.log(`[${new Date().toISOString()}] ${req.method} ${req.url}`);
  next();
});

app.get('/health', (req, res) => {
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
    { header: 'Date', key: 'date', width: 15 }
  ];
  await workbook.xlsx.writeFile(LOCAL_EXCEL_FILE);
  console.log('Initialized fresh Excel file:', LOCAL_EXCEL_FILE);
  return workbook;
}

async function loadFromGoogleDrive() {
  if (cachedWorkbook) return cachedWorkbook;
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
        file.data.pipe(dest).on('finish', resolve).on('error', reject);
      });
      workbook = new ExcelJS.Workbook();
      await workbook.xlsx.readFile(LOCAL_EXCEL_FILE);
    } else {
      workbook = await initializeExcel();
    }
    cachedWorkbook = workbook;
    return workbook;
  } catch (error) {
    console.error('Error loading from Drive:', error.stack);
    cachedWorkbook = await initializeExcel();
    return cachedWorkbook;
  }
}

async function uploadToGoogleDrive() {
  try {
    const fileStats = await fs.stat(LOCAL_EXCEL_FILE);
    if (fileStats.size < 100) throw new Error('Local Excel file is empty.');
    const files = await drive.files.list({
      q: `'${GOOGLE_DRIVE_FOLDER_ID}' in parents and name = 'customers.xlsx'`,
      fields: 'files(id)',
    });
    const media = {
      mimeType: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
      body: require('fs').createReadStream(LOCAL_EXCEL_FILE),
    };
    if (files.data.files.length > 0) {
      await drive.files.update({ fileId: files.data.files[0].id, media });
    } else {
      await drive.files.create({
        resource: { name: 'customers.xlsx', parents: [GOOGLE_DRIVE_FOLDER_ID] },
        media,
      });
    }
  } catch (error) {
    console.error('Drive upload failed:', error.stack);
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

const saveQueue = require('async').queue(async (task, cb) => {
  try {
    await task();
    cb();
  } catch (err) {
    console.error('Queue error:', err);
    cb(err);
  }
}, 1);

function normalize(str) {
  return String(str).trim().toLowerCase().replace(/[-\s]/g, '');
}

async function checkDuplicates(email, phone, workbook) {
  const sheet = workbook.getWorksheet('Customers');
  let duplicate = null;
  sheet.eachRow((row, i) => {
    if (i === 1) return;
    if (normalize(row.getCell(2).value) === normalize(email)) duplicate = 'email';
    if (normalize(row.getCell(3).value) === normalize(phone)) duplicate = 'phone';
  });
  return duplicate;
}

app.post('/submit', async (req, res) => {
  try {
    const { name, email, phone } = req.body;
    if (!name || !email || !phone) return res.status(400).json({ success: false, error: 'Missing fields' });
    const workbook = await withFileLock(loadFromGoogleDrive);
    const dup = await checkDuplicates(email, phone, workbook);
    if (dup) return res.status(409).json({ success: false, error: `Duplicate ${dup}` });
    const sheet = workbook.getWorksheet('Customers');
    sheet.addRow({ name, email, phone, date: new Date().toISOString().split('T')[0] }).commit();
    saveQueue.push(saveWorkbook);
    res.status(200).json({ success: true, name });
  } catch (err) {
    res.status(500).json({ success: false, error: err.message });
  }
});

app.get('/download', async (req, res) => {
  try {
    const workbook = await withFileLock(loadFromGoogleDrive);
    await workbook.xlsx.writeFile(LOCAL_EXCEL_FILE);
    res.setHeader('Content-Disposition', 'attachment; filename=customers.xlsx');
    res.setHeader('Content-Type', 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet');
    fs.createReadStream(LOCAL_EXCEL_FILE).pipe(res);
  } catch (err) {
    res.status(500).send('Download failed');
  }
});

app.use((req, res) => {
  res.status(404).send('Not Found');
});

process.on('uncaughtException', err => console.error('Uncaught Exception:', err));
process.on('unhandledRejection', (reason, p) => console.error('Unhandled Rejection:', reason));

(async () => {
  try {
    await loadFromGoogleDrive();
    app.listen(PORT, () => console.log(`Server running on port ${PORT}`));
  } catch (err) {
    console.error('Startup failed:', err);
    process.exit(1);
  }
})();
