require('dotenv').config();
const express = require('express');
const multer = require('multer');
const xlsx = require('xlsx');
const cors = require('cors');
const path = require('path');
const fs = require('fs-extra');
const { v4: uuidv4 } = require('uuid');
const rateLimit = require('express-rate-limit');

const app = express();
const PORT = process.env.PORT || 5000;

// Enhanced CORS configuration
app.use(cors({
  origin: process.env.ALLOWED_ORIGINS?.split(',') || '*',
  methods: ['GET', 'POST', 'OPTIONS'],
  allowedHeaders: ['Content-Type', 'Authorization'],
  exposedHeaders: ['Content-Disposition']
}));

app.use(express.json());
app.use(express.urlencoded({ extended: true }));

// Rate limiting
const limiter = rateLimit({
  windowMs: 15 * 60 * 1000, // 15 minutes
  max: 100, // limit each IP to 100 requests per windowMs
  message: 'Too many requests from this IP, please try again later'
});
app.use(limiter);

// Configure temporary directories
const tempDir = path.join(__dirname, 'temp');
const uploadDir = path.join(tempDir, 'uploads');
const outputDir = path.join(tempDir, 'output');

// Ensure directories exist and are clean
async function initializeDirectories() {
  try {
    await fs.ensureDir(uploadDir);
    await fs.ensureDir(outputDir);
    await fs.emptyDir(uploadDir);
    await fs.emptyDir(outputDir);
    console.log('Directories initialized successfully');
  } catch (err) {
    console.error('Directory initialization failed:', err);
    process.exit(1);
  }
}

// Configure multer with enhanced file handling
const storage = multer.diskStorage({
  destination: (req, file, cb) => {
    cb(null, uploadDir);
  },
  filename: (req, file, cb) => {
    const uniqueName = `${uuidv4()}${path.extname(file.originalname)}`;
    cb(null, uniqueName);
  }
});

const fileFilter = (req, file, cb) => {
  const filetypes = /xlsx|xls|csv/;
  const extname = filetypes.test(path.extname(file.originalname).toLowerCase());
  const mimetype = filetypes.test(file.mimetype);

  if (mimetype && extname) {
    cb(null, true);
  } else {
    cb(new Error('Only Excel/CSV files are allowed!'), false);
  }
};

const upload = multer({
  storage,
  fileFilter,
  limits: {
    fileSize: 50 * 1024 * 1024, // 50MB
    files: 1
  }
});

// Enhanced Excel date parser
function parseExcelDate(value) {
  if (!value) return null;

  try {
    // Handle Excel numeric dates (days since 1900-01-01)
    if (typeof value === 'number') {
      const excelEpoch = new Date(1899, 11, 30);
      const date = new Date(excelEpoch.getTime() + (value - 1) * 86400000);
      return date.toISOString().split('T')[0];
    }

    // Handle string dates (DD-MM-YYYY or MM/DD/YYYY)
    if (typeof value === 'string') {
      // Try common separators
      const separators = ['-', '/', '.'];
      for (const sep of separators) {
        const parts = value.split(sep);
        if (parts.length === 3) {
          // Try both day-first and month-first formats
          const formats = [
            `${parts[2]}-${parts[1]}-${parts[0]}`, // DD-MM-YYYY
            `${parts[2]}-${parts[0]}-${parts[1]}`  // MM-DD-YYYY
          ];
          
          for (const format of formats) {
            const date = new Date(format);
            if (!isNaN(date.getTime())) {
              return date.toISOString().split('T')[0];
            }
          }
        }
      }
    }

    // Fallback to native Date parsing
    const date = new Date(value);
    if (!isNaN(date.getTime())) {
      return date.toISOString().split('T')[0];
    }

    return null;
  } catch (err) {
    console.error('Date parsing error:', { value, error: err.message });
    return null;
  }
}

// API Endpoints

/**
 * @api {post} /upload Upload Excel File
 * @apiName UploadFile
 * @apiGroup File
 * 
 * @apiParam {File} file Excel file to upload
 * 
 * @apiSuccess {String[]} merchants List of merchant names
 * @apiSuccess {Number} count Number of transactions processed
 */
app.post('/upload', upload.single('file'), async (req, res) => {
  try {
    // Validate file
    if (!req.file) {
      return res.status(400).json({
        success: false,
        error: 'NO_FILE',
        message: 'No file was uploaded',
        solution: 'Please select an Excel file to upload'
      });
    }

    // Verify file exists
    try {
      await fs.access(req.file.path);
    } catch (err) {
      return res.status(500).json({
        success: false,
        error: 'FILE_ACCESS_ERROR',
        message: 'Uploaded file could not be accessed',
        details: err.message
      });
    }

    // Process Excel file
    let workbook;
    try {
      workbook = xlsx.readFile(req.file.path, { cellDates: true });
    } catch (err) {
      await fs.unlink(req.file.path);
      return res.status(400).json({
        success: false,
        error: 'INVALID_EXCEL',
        message: 'The file could not be read as an Excel file',
        details: err.message
      });
    }

    // Validate worksheet
    if (workbook.SheetNames.length === 0) {
      await fs.unlink(req.file.path);
      return res.status(400).json({
        success: false,
        error: 'NO_SHEETS',
        message: 'Excel file contains no worksheets'
      });
    }

    const sheetName = workbook.SheetNames[0];
    const worksheet = workbook.Sheets[sheetName];
    const rawData = xlsx.utils.sheet_to_json(worksheet, { defval: null });

    // Validate data
    if (!rawData || rawData.length === 0) {
      await fs.unlink(req.file.path);
      return res.status(400).json({
        success: false,
        error: 'NO_DATA',
        message: 'Excel worksheet contains no data'
      });
    }

    // Process data
    const processedData = rawData.map((row, index) => {
      try {
        return {
          ...row,
          RowIndex: index + 2, // Excel rows start at 1, header at 1
          DateOnly: parseExcelDate(row.Date)
        };
      } catch (err) {
        console.error(`Error processing row ${index}:`, err);
        return null;
      }
    }).filter(row => row !== null);

    // Extract merchants
    const merchants = [...new Set(
      processedData
        .map(row => row['Merchant Name'])
        .filter(name => name && typeof name === 'string' && name.trim() !== '')
    )].sort();

    // Store data in memory (for demo - use DB in production)
    req.app.locals.fileData = {
      originalName: req.file.originalname,
      processedData,
      merchants,
      uploadTime: new Date()
    };

    // Cleanup
    await fs.unlink(req.file.path);

    // Response
    res.json({
      success: true,
      merchants,
      count: processedData.length,
      firstFew: processedData.slice(0, 3) // For debugging
    });

  } catch (err) {
    console.error('Upload processing error:', err);

    // Cleanup if file exists
    if (req.file) {
      await fs.unlink(req.file.path).catch(console.error);
    }

    res.status(500).json({
      success: false,
      error: 'PROCESSING_ERROR',
      message: 'An error occurred while processing the file',
      details: process.env.NODE_ENV === 'development' ? err.message : undefined
    });
  }
});

/**
 * @api {post} /generate Generate Summary
 * @apiName GenerateSummary
 * @apiGroup Analysis
 * 
 * @apiParam {String[]} selectedMerchants Array of merchant names
 * @apiParam {Number} percentage Percentage to calculate
 * @apiParam {String} startDate Start date (YYYY-MM-DD)
 * @apiParam {String} endDate End date (YYYY-MM-DD)
 * 
 * @apiSuccess {Object[]} summary Generated summary data
 * @apiSuccess {String} downloadUrl URL to download Excel file
 * @apiSuccess {String} dateRange Formatted date range
 */
app.post('/generate', async (req, res) => {
  try {
    // Validate input
    const { selectedMerchants, percentage, startDate, endDate } = req.body;

    if (!selectedMerchants || !Array.isArray(selectedMerchants) || selectedMerchants.length === 0) {
      return res.status(400).json({
        success: false,
        error: 'NO_MERCHANTS',
        message: 'No merchants were selected'
      });
    }

    const percent = parseFloat(percentage);
    if (isNaN(percent) || percent < 0 || percent > 100) {
      return res.status(400).json({
        success: false,
        error: 'INVALID_PERCENTAGE',
        message: 'Percentage must be a number between 0 and 100'
      });
    }

    if (!startDate || !endDate || !Date.parse(startDate) || !Date.parse(endDate)) {
      return res.status(400).json({
        success: false,
        error: 'INVALID_DATE',
        message: 'Invalid date range provided'
      });
    }

    // Get data from memory (from upload)
    const fileData = req.app.locals.fileData;
    if (!fileData || !fileData.processedData) {
      return res.status(400).json({
        success: false,
        error: 'NO_DATA',
        message: 'No data available. Please upload a file first.'
      });
    }

    // Filter data
    const filteredData = fileData.processedData.filter(row => {
      return (
        row.DateOnly &&
        row.DateOnly >= startDate &&
        row.DateOnly <= endDate &&
        selectedMerchants.includes(row['Merchant Name'])
      );
    });

    if (filteredData.length === 0) {
      return res.status(404).json({
        success: false,
        error: 'NO_MATCHING_DATA',
        message: 'No transactions found for the selected criteria',
        details: {
          dateRange: `${startDate} to ${endDate}`,
          merchants: selectedMerchants
        }
      });
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
          percentAmount: 0,
          count: 0
        });
      }

      const current = summaryMap.get(merchant);
      current.withdrawal += withdrawal;
      current.fee += fee;
      current.percentAmount += percentAmount;
      current.count += 1;
    });

    // Convert to array format
    const summaryData = Array.from(summaryMap.entries()).map(([merchant, values]) => ({
      Merchant: merchant,
      'Total Withdrawal Amount': values.withdrawal.toFixed(2),
      'Total Withdrawal Fees': values.fee.toFixed(2),
      [`${percent}% Amount`]: values.percentAmount.toFixed(2),
      'Transaction Count': values.count
    }));

    // Add totals row
    const totals = {
      Merchant: 'TOTAL',
      'Total Withdrawal Amount': summaryData.reduce((sum, row) => sum + parseFloat(row['Total Withdrawal Amount']), 0).toFixed(2),
      'Total Withdrawal Fees': summaryData.reduce((sum, row) => sum + parseFloat(row['Total Withdrawal Fees']), 0).toFixed(2),
      [`${percent}% Amount`]: summaryData.reduce((sum, row) => sum + parseFloat(row[`${percent}% Amount`]), 0).toFixed(2),
      'Transaction Count': summaryData.reduce((sum, row) => sum + row['Transaction Count'], 0)
    };
    summaryData.push(totals);

    // Generate Excel file
    const outputFilename = `summary_${Date.now()}.xlsx`;
    const outputPath = path.join(outputDir, outputFilename);

    const wb = xlsx.utils.book_new();
    
    // Add raw data sheet
    const wsData = xlsx.utils.json_to_sheet(filteredData.map(row => {
      const { RowIndex, DateOnly, ...rest } = row;
      return rest;
    }));
    xlsx.utils.book_append_sheet(wb, wsData, 'Transactions');
    
    // Add summary sheet
    const wsSummary = xlsx.utils.json_to_sheet(summaryData);
    xlsx.utils.book_append_sheet(wb, wsSummary, 'Summary');
    
    // Write file
    xlsx.writeFile(wb, outputPath);

    // Set download URL (valid until server restarts)
    const downloadUrl = `/download/${outputFilename}`;

    // Response
    res.json({
      success: true,
      summary: summaryData,
      downloadUrl,
      dateRange: `${startDate} to ${endDate}`,
      stats: {
        transactionCount: filteredData.length,
        merchantCount: summaryMap.size
      }
    });

  } catch (err) {
    console.error('Summary generation error:', err);
    res.status(500).json({
      success: false,
      error: 'GENERATION_ERROR',
      message: 'Failed to generate summary',
      details: process.env.NODE_ENV === 'development' ? err.message : undefined
    });
  }
});

/**
 * @api {get} /download/:filename Download Excel File
 * @apiName DownloadFile
 * @apiGroup File
 * 
 * @apiParam {String} filename Filename to download
 */
app.get('/download/:filename', async (req, res) => {
  try {
    const filename = req.params.filename;
    if (!filename || !filename.match(/^summary_\d+\.xlsx$/)) {
      return res.status(400).json({
        success: false,
        error: 'INVALID_FILENAME',
        message: 'Invalid filename requested'
      });
    }

    const filePath = path.join(outputDir, filename);
    
    // Verify file exists
    try {
      await fs.access(filePath);
    } catch (err) {
      return res.status(404).json({
        success: false,
        error: 'FILE_NOT_FOUND',
        message: 'The requested file does not exist or has expired',
        details: 'Generated files are temporary and removed after download'
      });
    }

    // Set headers
    res.setHeader('Content-Type', 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet');
    res.setHeader('Content-Disposition', `attachment; filename=${filename}`);

    // Stream file
    const fileStream = fs.createReadStream(filePath);
    fileStream.pipe(res);

    // Cleanup after download completes
    fileStream.on('end', () => {
      fs.unlink(filePath).catch(err => {
        console.error('Error deleting file after download:', err);
      });
    });

    fileStream.on('error', (err) => {
      console.error('File stream error:', err);
      fs.unlink(filePath).catch(console.error);
      res.status(500).end();
    });

  } catch (err) {
    console.error('Download error:', err);
    res.status(500).json({
      success: false,
      error: 'DOWNLOAD_ERROR',
      message: 'Failed to download file'
    });
  }
});

// Health check endpoint
app.get('/health', (req, res) => {
  res.json({
    status: 'OK',
    uptime: process.uptime(),
    timestamp: new Date(),
    memoryUsage: process.memoryUsage()
  });
});

// Error handling middleware
app.use((err, req, res, next) => {
  console.error('Unhandled error:', err);
  res.status(500).json({
    success: false,
    error: 'SERVER_ERROR',
    message: 'Internal server error'
  });
});

// Initialize and start server
initializeDirectories().then(() => {
  app.listen(PORT, () => {
    console.log(`Server running on port ${PORT}`);
    console.log(`Upload directory: ${uploadDir}`);
    console.log(`Output directory: ${outputDir}`);
    console.log(`Environment: ${process.env.NODE_ENV || 'development'}`);
  });
});

// Cleanup on exit
process.on('SIGINT', async () => {
  console.log('Shutting down server...');
  try {
    await fs.remove(tempDir);
    console.log('Temporary files cleaned up');
    process.exit(0);
  } catch (err) {
    console.error('Cleanup error:', err);
    process.exit(1);
  }
});