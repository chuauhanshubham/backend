const express = require('express');
const multer = require('multer');
const XLSX = require('xlsx');
const fs = require('fs-extra');
const path = require('path');

const app = express();
const PORT = 5000;

const uploadDir = path.join(__dirname, 'uploads');
const outputDir = path.join(__dirname, 'output');
fs.ensureDirSync(uploadDir);
fs.ensureDirSync(outputDir);

const storage = multer.diskStorage({
    destination: uploadDir,
    filename: (req, file, cb) => {
        cb(null, 'input.xlsx');
    }
});
const upload = multer({ storage });

app.post('/upload', upload.single('file'), (req, res) => {
    const filePath = req.file.path;

    const workbook = XLSX.readFile(filePath);
    const sheetName = workbook.SheetNames[0];
    const worksheet = workbook.Sheets[sheetName];
    const data = XLSX.utils.sheet_to_json(worksheet);

    const grouped = {};
    data.forEach(row => {
        const merchant = row['Merchant Name'] || 'Unknown';

        const withdrawalAmount = parseFloat(row['Withdrawal Amount']);
        if (!isNaN(withdrawalAmount)) {
            row['3.5% Amount'] = (withdrawalAmount * 0.035).toFixed(2);
        } else {
            row['3.5% Amount'] = '';
        }

        if (!grouped[merchant]) grouped[merchant] = [];
        grouped[merchant].push(row);
    });

    const combinedData = [];

    Object.keys(grouped).forEach(merchant => {
        const merchantData = grouped[merchant];

        let totalWithdrawal = 0;
        let totalPercent = 0;
        let totalFees = 0;

        merchantData.forEach(row => {
            const w = parseFloat(row['Withdrawal Amount']);
            const p = parseFloat(row['3.5% Amount']);
            const f = parseFloat(row['Withdrawal Fees']);
            if (!isNaN(w)) totalWithdrawal += w;
            if (!isNaN(p)) totalPercent += p;
            if (!isNaN(f)) totalFees += f;

            combinedData.push(row);
        });

        const totalRow = {
            'Merchant Name': `${merchant} TOTAL`,
            'Withdrawal Amount': totalWithdrawal.toFixed(2),
            '3.5% Amount': totalPercent.toFixed(2),
            'Withdrawal Fees': totalFees.toFixed(2)
        };
        combinedData.push(totalRow);

        // Empty row to separate next merchant
        combinedData.push({});
    });

    fs.emptyDirSync(outputDir);

    const newWb = XLSX.utils.book_new();
    const newWs = XLSX.utils.json_to_sheet(combinedData);
    XLSX.utils.book_append_sheet(newWb, newWs, 'All Merchants');

    const outputFilePath = path.join(outputDir, 'merged_merchants.xlsx');
    XLSX.writeFile(newWb, outputFilePath);

    res.download(outputFilePath);
});

app.get('/', (req, res) => {
    res.send(`
        <form action="/upload" method="post" enctype="multipart/form-data">
            <input type="file" name="file" />
            <button type="submit">Upload & Generate</button>
        </form>
    `);
});

app.listen(PORT, () => {
    console.log(`✅ Server running at http://localhost:${PORT}`);
});
