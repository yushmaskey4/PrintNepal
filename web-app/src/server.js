const express = require('express');
const cors = require('cors');
const multer = require('multer');
const ExcelJS = require('exceljs');
const fs = require('fs');
const path = require('path');
const docxConverter = require('docx-pdf');
require('dotenv').config();

const app = express();
const PORT = process.env.PORT || 5000;

// Ensure uploads folder exists
const uploadsDir = path.join(__dirname, '../uploads');
if (!fs.existsSync(uploadsDir)) {
  fs.mkdirSync(uploadsDir, { recursive: true });
}

app.use(cors());
app.use(express.json());

// Serve static files
app.use('/uploads', express.static(uploadsDir));
app.use(express.static(path.join(__dirname, '../public')));

// In-Memory Database for active orders
let orders = [];
let orderIdCounter = 101;

// Multer storage setup
const storage = multer.diskStorage({
  destination: (req, file, cb) => cb(null, uploadsDir),
  filename: (req, file, cb) => cb(null, `${Date.now()}-${file.originalname}`)
});
const upload = multer({ storage });

// Helper function to count PDF pages from buffer
function countPdfPages(dataBuffer) {
  const pdfText = dataBuffer.toString('latin1');
  const pageMatches = pdfText.match(/\/Type\s*\/Page\b/g);
  if (pageMatches) return pageMatches.length;

  const countMatch = pdfText.match(/\/Count\s+(\d+)/);
  if (countMatch && countMatch[1]) {
    return parseInt(countMatch[1], 10);
  }
  return 1;
}

// 1. Universal Upload Endpoint (.pdf, .docx, .doc)
app.post('/api/upload', upload.single('file'), async (req, res) => {
  try {
    if (!req.file) return res.status(400).json({ error: 'No file uploaded' });

    let finalFilePath = `/uploads/${req.file.filename}`;
    let finalDiskPath = req.file.path;
    const fileExt = path.extname(req.file.originalname).toLowerCase();

    if (fileExt === '.docx' || fileExt === '.doc') {
      const pdfFilename = `${Date.now()}-converted.pdf`;
      const convertedPdfPath = path.join(uploadsDir, pdfFilename);

      // Convert DOCX to PDF asynchronously
      await new Promise((resolve, reject) => {
        docxConverter(req.file.path, convertedPdfPath, (err, result) => {
          if (err) reject(err);
          else resolve(result);
        });
      });

      finalFilePath = `/uploads/${pdfFilename}`;
      finalDiskPath = convertedPdfPath;
    }

    const pdfBuffer = fs.readFileSync(finalDiskPath);
    const totalPages = countPdfPages(pdfBuffer);

    res.json({
      filePath: finalFilePath,
      fileName: req.file.originalname,
      totalPages: totalPages
    });
  } catch (err) {
    console.error('Upload / Conversion Error:', err);
    res.status(500).json({ error: 'Failed to process document file' });
  }
});

// 2. Create Order Endpoint
app.post('/api/orders', (req, res) => {
  const { customerName, phoneNumber, filePath, fileName, totalPages, copies, colorMode, duplex, totalPrice } = req.body;

  const newOrder = {
    id: `PN-${orderIdCounter++}`,
    customerName,
    phoneNumber,
    filePath,
    fileName,
    totalPages,
    copies,
    colorMode,
    duplex,
    totalPrice,
    status: 'pending',
    createdAt: new Date()
  };

  orders.push(newOrder);
  res.json({ success: true, order: newOrder });
});

// 3. Public Queue Endpoint
app.get('/api/queue', (req, res) => {
  const publicQueue = orders.map(o => ({
    id: o.id,
    customerName: o.customerName,
    fileName: o.fileName,
    specs: `${o.totalPages} Pages | ${o.copies} Copies | ${o.colorMode.toUpperCase()} | ${o.duplex ? 'Double-Sided' : 'Single-Sided'}`,
    status: o.status
  }));
  res.json(publicQueue);
});

// 4. Agent Endpoint: Get Pending Jobs
app.get('/api/agent/pending', (req, res) => {
  const pendingJobs = orders.filter(o => o.status === 'pending');
  res.json(pendingJobs);
});

// 5. Agent Endpoint: Update Job Status
app.post('/api/agent/update-status', async (req, res) => {
  const { id, status } = req.body;
  const order = orders.find(o => o.id === id);

  if (order) {
    order.status = status;
    if (status === 'ready') {
      await updateDailyExcelReport(order);
    }
    return res.json({ success: true });
  }
  res.status(404).json({ error: 'Order not found' });
});

// Helper Function: Update Daily Excel Report
async function updateDailyExcelReport(order) {
  const today = new Date().toISOString().split('T')[0];
  const excelPath = path.join(__dirname, `../daily_log_${today}.xlsx`);

  let workbook = new ExcelJS.Workbook();
  let sheet;

  if (fs.existsSync(excelPath)) {
    await workbook.xlsx.readFile(excelPath);
    sheet = workbook.getWorksheet('Daily Orders');
  } else {
    sheet = workbook.addWorksheet('Daily Orders');
    sheet.columns = [
      { header: 'Time', key: 'time', width: 15 },
      { header: 'Order ID', key: 'id', width: 15 },
      { header: 'Customer Name', key: 'customerName', width: 20 },
      { header: 'Phone Number', key: 'phoneNumber', width: 18 },
      { header: 'File Name', key: 'fileName', width: 25 },
      { header: 'Pages', key: 'totalPages', width: 10 },
      { header: 'Copies', key: 'copies', width: 10 },
      { header: 'Total Price (NPR)', key: 'totalPrice', width: 18 }
    ];
  }

  sheet.addRow({
    time: new Date().toLocaleTimeString(),
    id: order.id,
    customerName: order.customerName,
    phoneNumber: order.phoneNumber,
    fileName: order.fileName,
    totalPages: order.totalPages,
    copies: order.copies,
    totalPrice: order.totalPrice
  });

  await workbook.xlsx.writeFile(excelPath);
}

app.listen(PORT, () => console.log(`PrintNepal Server running on http://localhost:${PORT}`));
