const express = require('express');
const multer = require('multer');
const xlsx = require('xlsx');
const cors = require('cors');
const path = require('path');
const fs = require('fs-extra');
const { v4: uuidv4 } = require('uuid');

const app = express();
const PORT = process.env.PORT || 5000;

// Configure CORS
app.use(cors());
app.use(express.json());

// Set up temporary directories
const uploadDir = path.join(__dirname, 'temp', 'uploads');
const outputDir = path.join(__dirname, 'temp', 'output');

fs.ensureDirSync(uploadDir);
fs.ensureDirSync(outputDir);

// Configure multer for file uploads
const storage = multer.diskStorage({
  destination: uploadDir,
  filename: (req, file, cb) => {
    const uniqueName = `${uuidv4()}${path.extname(file.originalname)}`;
    cb(null, uniqueName);
  }
});

const upload = multer({ 
  storage,
  fileFilter: (req, file, cb) => {
    if (file.mimetype.includes('excel') || file.mimetype.includes('spreadsheet') || 
        path.extname(file.originalname).match(/\.(xlsx|xls)$/)) {
      cb(null, true);
    } else {
      cb(new Error('Only Excel files are allowed!'), false);
    }
  },
  limits: { fileSize: 10 * 1024 * 1024 } // 10MB limit
});

// Clean up temp files on server start
async function cleanTempFiles() {
  try {
    await fs.emptyDir(uploadDir);
    await fs.emptyDir(outputDir);
    console.log('Temporary directories cleaned');
  } catch (err) {
    console.error('Error cleaning temp directories:', err);
  }
}

cleanTempFiles();

// Helper function to parse dates from Excel
function parseExcelDate(excelDate) {
  try {
    if (typeof excelDate === 'number') {
      // Excel dates are numbers where 1 = 1900-01-01
      const excelEpoch = new Date(1899, 11, 30);
      const date = new Date(excelEpoch.getTime() + (excelDate - 1) * 86400000);
      return date.toISOString().split('T')[0];
    } else if (typeof excelDate === 'string') {
      // Handle string dates in format "DD-MM-YYYY"
      const parts = excelDate.split(/[-/]/);
      if (parts.length === 3) {
        const date = new Date(`${parts[2]}-${parts[1]}-${parts[0]}`);
        return date.toISOString().split('T')[0];
      }
    }
    // Try parsing as ISO date
    const date = new Date(excelDate);
    return date.toISOString().split('T')[0];
  } catch (err) {
    return null;
  }
}

// API Endpoints

// Upload Excel file and extract merchants
app.post('/upload', upload.single('file'), async (req, res) => {
  try {
    if (!req.file) {
      return res.status(400).json({ error: 'No file uploaded' });
    }

    const filePath = req.file.path;
    const workbook = xlsx.readFile(filePath);
    const sheetName = workbook.SheetNames[0];
    const worksheet = workbook.Sheets[sheetName];
    const rawData = xlsx.utils.sheet_to_json(worksheet);

    // Process data and add formatted dates
    const processedData = rawData.map(row => {
      const dateOnly = parseExcelDate(row.Date);
      return { ...row, DateOnly: dateOnly };
    });

    // Extract unique merchant names
    const merchants = [...new Set(
      processedData
        .map(row => row['Merchant Name'])
        .filter(name => name && name.trim() !== '')
    )];

    // Store processed data in memory (in production, use a database)
    req.app.locals.processedData = processedData;

    // Clean up the uploaded file
    await fs.unlink(filePath);

    res.json({ merchants });
  } catch (err) {
    console.error('Upload error:', err);
    res.status(500).json({ error: 'Failed to process file', details: err.message });
  }
});

// Generate summary and Excel file
app.post('/generate', async (req, res) => {
  try {
    const { selectedMerchants, percentage, startDate, endDate } = req.body;

    // Validate input
    if (!selectedMerchants || !selectedMerchants.length) {
      return res.status(400).json({ error: 'No merchants selected' });
    }
    if (!percentage || isNaN(parseFloat(percentage))) {
      return res.status(400).json({ error: 'Invalid percentage value' });
    }
    if (!startDate || !endDate) {
      return res.status(400).json({ error: 'Date range not specified' });
    }

    const percent = parseFloat(percentage);
    const processedData = req.app.locals.processedData;

    if (!processedData || !processedData.length) {
      return res.status(400).json({ error: 'No data available. Please upload a file first.' });
    }

    // Filter data by date range and merchants
    const filteredData = processedData.filter(row => {
      return (
        row.DateOnly && 
        row.DateOnly >= startDate && 
        row.DateOnly <= endDate &&
        selectedMerchants.includes(row['Merchant Name'])
      );
    });

    if (!filteredData.length) {
      return res.status(404).json({ error: 'No transactions found for the selected criteria' });
    }

    // Calculate summary
    const summaryMap = new Map();

    filteredData.forEach(row => {
      const merchant = row['Merchant Name'];
      const withdrawal = parseFloat(row['Withdrawal Amount']) || 0;
      const fee = parseFloat(row['Withdrawal Fees']) || 0;
      const percentAmount = withdrawal * percent / 100;

      if (!summaryMap.has(merchant)) {
        summaryMap.set(merchant, {
          withdrawal: 0,
          fee: 0,
          percentAmount: 0
        });
      }

      const current = summaryMap.get(merchant);
      current.withdrawal += withdrawal;
      current.fee += fee;
      current.percentAmount += percentAmount;
    });

    // Convert to array format
    const summaryData = Array.from(summaryMap.entries()).map(([merchant, values]) => ({
      Merchant: merchant,
      'Total Withdrawal Amount': values.withdrawal.toFixed(2),
      'Total Withdrawal Fees': values.fee.toFixed(2),
      [`${percent}% Amount`]: values.percentAmount.toFixed(2)
    }));

    // Add totals row
    const totals = {
      Merchant: 'TOTAL',
      'Total Withdrawal Amount': summaryData.reduce((sum, row) => sum + parseFloat(row['Total Withdrawal Amount']), 0).toFixed(2),
      'Total Withdrawal Fees': summaryData.reduce((sum, row) => sum + parseFloat(row['Total Withdrawal Fees']), 0).toFixed(2),
      [`${percent}% Amount`]: summaryData.reduce((sum, row) => sum + parseFloat(row[`${percent}% Amount`]), 0).toFixed(2)
    };
    summaryData.push(totals);

    // Generate Excel file
    const outputFilename = `summary_${Date.now()}.xlsx`;
    const outputPath = path.join(outputDir, outputFilename);

    const wb = xlsx.utils.book_new();
    xlsx.utils.book_append_sheet(wb, xlsx.utils.json_to_sheet(filteredData), 'Transactions');
    xlsx.utils.book_append_sheet(wb, xlsx.utils.json_to_sheet(summaryData), 'Summary');
    xlsx.writeFile(wb, outputPath);

    // Set download URL (valid for 1 hour)
    const downloadUrl = `/download/${outputFilename}`;

    res.json({
      summary: summaryData,
      downloadUrl,
      dateRange: `${startDate} to ${endDate}`
    });
  } catch (err) {
    console.error('Generate error:', err);
    res.status(500).json({ error: 'Failed to generate summary', details: err.message });
  }
});

// Download generated Excel file
app.get('/download/:filename', async (req, res) => {
  try {
    const filename = req.params.filename;
    const filePath = path.join(outputDir, filename);

    if (!fs.existsSync(filePath)) {
      return res.status(404).json({ error: 'File not found or expired' });
    }

    res.download(filePath, filename, (err) => {
      if (err) {
        console.error('Download error:', err);
      }
      // Clean up the file after download
      fs.unlink(filePath).catch(console.error);
    });
  } catch (err) {
    console.error('Download error:', err);
    res.status(500).json({ error: 'Failed to download file' });
  }
});

// Error handling middleware
app.use((err, req, res, next) => {
  console.error('Server error:', err);
  res.status(500).json({ error: 'Internal server error', details: err.message });
});

// Start server
app.listen(PORT, () => {
  console.log(`Server running on port ${PORT}`);
  console.log(`Upload directory: ${uploadDir}`);
  console.log(`Output directory: ${outputDir}`);
});

// Clean up on exit
process.on('SIGINT', async () => {
  try {
    await fs.remove(uploadDir);
    await fs.remove(outputDir);
    console.log('Temporary directories removed');
    process.exit(0);
  } catch (err) {
    console.error('Cleanup error:', err);
    process.exit(1);
  }
});