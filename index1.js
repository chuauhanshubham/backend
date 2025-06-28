/** ✅ server.js - FULL WORKING VERSION with Summary Sheet and Frontend Summary Support */

const express = require('express');
const multer = require('multer');
const XLSX = require('xlsx');
const fs = require('fs-extra');
const path = require('path');
const bodyParser = require('body-parser');
const cors = require('cors');

const app = express();
const PORT = 5000;

const uploadDir = path.join(__dirname, 'uploads');
const outputDir = path.join(__dirname, 'output');
fs.ensureDirSync(uploadDir);
fs.ensureDirSync(outputDir);

app.use(bodyParser.json());
app.use(cors());
app.use(express.static('public'));
app.use('/output', express.static(outputDir)); // ✅ Serve Excel file

let globalData = [];

const storage = multer.diskStorage({
  destination: uploadDir,
  filename: (req, file, cb) => cb(null, 'input.xlsx')
});
const upload = multer({ storage });

function extractDateOnly(value) {
  if (!value) return '';
  if (typeof value === 'number') {
    const excelEpoch = new Date(Date.UTC(1899, 11, 30));
    return new Date(excelEpoch.getTime() + (value - 1) * 86400000).toISOString().slice(0, 10);
  }
  if (typeof value === 'string' && /^\d{2}-\d{2}-\d{4}/.test(value)) {
    const [d, m, y] = value.split(' ')[0].split('-');
    const date = new Date(`${y}-${m}-${d}`);
    return isNaN(date.getTime()) ? '' : date.toISOString().slice(0, 10);
  }
  const parsed = new Date(value);
  return isNaN(parsed.getTime()) ? '' : parsed.toISOString().slice(0, 10);
}

app.post('/upload', upload.single('file'), (req, res) => {
  const filePath = req.file.path;
  const workbook = XLSX.readFile(filePath);
  const sheet = workbook.Sheets[workbook.SheetNames[0]];
  const rawData = XLSX.utils.sheet_to_json(sheet);

  globalData = rawData.map(row => ({ ...row, DateOnly: extractDateOnly(row['Date']) }));
  const merchants = [...new Set(globalData.map(row => row['Merchant Name']).filter(Boolean))];
  res.json({ merchants });
});

app.post('/generate', async (req, res) => {
  const { selectedMerchants, percentage, startDate, endDate } = req.body;

  if (!selectedMerchants || !startDate || !endDate || !percentage) {
    return res.status(400).json({ error: 'Missing fields' });
  }

  const percent = parseFloat(percentage);
  if (isNaN(percent)) return res.status(400).json({ error: 'Invalid percentage' });

  const normalizedStart = new Date(startDate).toISOString().slice(0, 10);
  const normalizedEnd = new Date(endDate).toISOString().slice(0, 10);

  const dateFiltered = globalData.filter(row =>
    row.DateOnly >= normalizedStart && row.DateOnly <= normalizedEnd
  );

  const summaryData = [];
  const filteredData = [];

  let grandW = 0, grandF = 0, grandP = 0;

  for (const merchant of selectedMerchants) {
    const rows = dateFiltered.filter(row => row['Merchant Name'] === merchant);

    let totalW = 0, totalF = 0, totalP = 0;

    rows.forEach(row => {
      const withdrawal = parseFloat(row['Withdrawal Amount'] || 0);
      const fee = parseFloat(row['Withdrawal Fees'] || 0);
      const percentAmount = withdrawal * percent / 100;

      row[`${percentage}% Amount`] = percentAmount.toFixed(2);
      totalW += withdrawal;
      totalF += fee;
      totalP += percentAmount;
      filteredData.push(row);
    });

    summaryData.push({
      'Merchant': merchant,
      'Total Withdrawal Amount': totalW.toFixed(2),
      'Total Withdrawal Fees': totalF.toFixed(2),
      [`${percentage}% Amount`]: totalP.toFixed(2)
    });

    grandW += totalW;
    grandF += totalF;
    grandP += totalP;
  }

  summaryData.push({
    'Merchant': 'TOTAL',
    'Total Withdrawal Amount': grandW.toFixed(2),
    'Total Withdrawal Fees': grandF.toFixed(2),
    [`${percentage}% Amount`]: grandP.toFixed(2)
  });

  if (filteredData.length === 0 && summaryData.length === 0) {
    return res.status(404).json({ error: 'No data found in range' });
  }

  const wb = XLSX.utils.book_new();
  const wsData = XLSX.utils.json_to_sheet(filteredData);
  XLSX.utils.book_append_sheet(wb, wsData, 'Filtered Data');

  const wsSummary = XLSX.utils.json_to_sheet(summaryData);
  XLSX.utils.book_append_sheet(wb, wsSummary, 'Summary');

  const outputPath = path.join(outputDir, 'filtered_output.xlsx');
  XLSX.writeFile(wb, outputPath);

  res.json({
    summary: summaryData,
    downloadUrl: '/output/filtered_output.xlsx',
    dateRange: `${normalizedStart} to ${normalizedEnd}`
  });
});

app.listen(PORT, () => {
  console.log(`✅ Server running at http://localhost:${PORT}`);
});
