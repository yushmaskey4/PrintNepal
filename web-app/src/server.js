const express = require('express');
const cors = require('cors');
const multer = require('multer');
const ExcelJS = require('exceljs');
const fs = require('fs');
const path = require('path');
const sharp = require('sharp');
const { print } = require('pdf-to-printer');
require('dotenv').config();

const app = express();
const PORT = process.env.PORT || 5000;

// Admin Credentials
const ADMIN_CREDENTIALS = {
  username: process.env.ADMIN_USER,
  password: process.env.ADMIN_PASSWORD
};

// Ensure uploads folder exists
const uploadsDir = path.join(__dirname, '../uploads');
if (!fs.existsSync(uploadsDir)) {
  fs.mkdirSync(uploadsDir, { recursive: true });
}

app.use(cors());
app.use(express.json());

// Serve static uploaded files and public portal files
app.use('/uploads', express.static(uploadsDir));
app.use(express.static(path.join(__dirname, '../public')));

// In-Memory Database for active orders
let orders = [];
let orderIdCounter = 101;

// Multer Storage Setup: standard disk storage for documents and screenshots
const storage = multer.diskStorage({
  destination: (req, file, cb) => cb(null, uploadsDir),
  filename: (req, file, cb) => {
    const uniqueSuffix = Date.now() + '-' + Math.round(Math.random() * 1e9);
    cb(null, `${file.fieldname}-${uniqueSuffix}${path.extname(file.originalname)}`);
  }
});

const upload = multer({
  storage,
  fileFilter: (req, file, cb) => {
    const ext = path.extname(file.originalname).toLowerCase();
    if (file.fieldname === 'document' && ext !== '.pdf') {
      return cb(new Error('Only PDF documents are allowed!'));
    }
    if (file.fieldname === 'paymentProof' && !['.png', '.jpg', '.jpeg', '.webp'].includes(ext)) {
      return cb(new Error('Payment proof must be an image (.png, .jpg, .jpeg)!'));
    }
    cb(null, true);
  }
});

// Helper function: native PDF page counting buffer reader
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

// Helper function: Send PDF directly to local default printer
async function autoPrintDocument(filePath, copies) {
  try {
    const fullPath = path.join(__dirname, '..', filePath);
    console.log(`[AutoPrint] Sending ${filePath} to local printer (${copies} copies)...`);

    await print(fullPath, {
      copies: copies
      // printer: "HP_LaserJet_1020" // Optional: specify custom printer name
    });

    console.log(`[AutoPrint] Successfully printed ${filePath}`);
  } catch (err) {
    console.error('[AutoPrint Error]:', err.message);
  }
}

// Helper function: Compress Payment Screenshot Image using Sharp
async function compressImage(filePath) {
  try {
    const absolutePath = path.join(__dirname, '..', filePath);
    const tempCompressedPath = absolutePath + '-compressed.jpg';

    // Resize max dimensions to 800px and convert to 60% quality JPEG
    await sharp(absolutePath)
      .resize({ width: 800, height: 800, fit: 'inside', withoutEnlargement: true })
      .jpeg({ quality: 60 })
      .toFile(tempCompressedPath);

    // Overwrite original file with compressed file
    fs.renameSync(tempCompressedPath, absolutePath);
    console.log(`[Storage] Compressed image screenshot: ${filePath}`);
  } catch (err) {
    console.error('[Compression Error]:', err.message);
  }
}

// Helper function: Automatically purge files older than 24 hours (86,400,000 ms)
function autoDeleteOldFiles() {
  const TWENTY_FOUR_HOURS = 24 * 60 * 60 * 1000;
  const now = Date.now();

  fs.readdir(uploadsDir, (err, files) => {
    if (err) return console.error('[Storage Cleanup Error]:', err);

    files.forEach(file => {
      const filePath = path.join(uploadsDir, file);
      fs.stat(filePath, (statErr, stats) => {
        if (statErr) return;

        // If file is older than 24 hours, delete it from disk
        if (now - stats.mtimeMs > TWENTY_FOUR_HOURS) {
          fs.unlink(filePath, unlinkErr => {
            if (!unlinkErr) {
              console.log(`[AutoCleanup] Deleted 24h+ file: ${file}`);
            }
          });
        }
      });
    });
  });
}

// Run storage cleanup every hour (3,600,000 ms)
setInterval(autoDeleteOldFiles, 60 * 60 * 1000);

// Helper function: Append completed orders to Daily Excel Log
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
      { header: 'Total Price (NPR)', key: 'totalPrice', width: 18 },
      { header: 'Payment Screenshot', key: 'paymentProof', width: 35 }
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
    totalPrice: order.totalPrice,
    paymentProof: order.paymentProofPath || 'N/A'
  });

  await workbook.xlsx.writeFile(excelPath);
}

// ------------------- PUBLIC ENDPOINTS -------------------

// 1. Universal Upload Endpoint (Document + Payment Screenshot with Auto-Compression)
app.post('/api/upload-with-payment', (req, res) => {
  upload.fields([
    { name: 'document', maxCount: 1 },
    { name: 'paymentProof', maxCount: 1 }
  ])(req, res, async (err) => {
    if (err) {
      return res.status(400).json({ error: err.message || 'File upload failed' });
    }

    if (!req.files || !req.files.document) {
      return res.status(400).json({ error: 'PDF document file is required.' });
    }

    try {
      const docFile = req.files.document[0];
      const pdfBuffer = fs.readFileSync(docFile.path);
      const totalPages = countPdfPages(pdfBuffer);

      let paymentProofPath = null;
      if (req.files.paymentProof && req.files.paymentProof[0]) {
        paymentProofPath = `/uploads/${req.files.paymentProof[0].filename}`;
        // Compress payment proof image asynchronously
        await compressImage(paymentProofPath);
      }

      res.json({
        filePath: `/uploads/${docFile.filename}`,
        fileName: docFile.originalname,
        totalPages: totalPages,
        paymentProofPath: paymentProofPath
      });
    } catch (parseErr) {
      console.error('PDF Read Error:', parseErr);
      res.status(500).json({ error: 'Failed to parse PDF document pages' });
    }
  });
});

// 2. Create Order Endpoint
app.post('/api/orders', (req, res) => {
  const { 
    customerName, 
    phoneNumber, 
    filePath, 
    fileName, 
    totalPages, 
    copies, 
    colorMode, 
    totalPrice,
    paymentProofPath 
  } = req.body;

  const newOrder = {
    id: `PN-${orderIdCounter++}`,
    customerName,
    phoneNumber,
    filePath,
    fileName,
    totalPages,
    copies,
    colorMode,
    totalPrice,
    paymentProofPath: paymentProofPath || null,
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
    specs: `${o.totalPages} Pages | ${o.copies} Copies | ${o.colorMode.toUpperCase()}`,
    status: o.status,
    hasPaymentProof: !!o.paymentProofPath
  }));
  res.json(publicQueue);
});

// ------------------- ADMIN / OPERATOR ENDPOINTS -------------------

// 4. Admin Authentication Login
app.post('/api/admin/login', (req, res) => {
  const { username, password } = req.body;
  if (username === ADMIN_CREDENTIALS.username && password === ADMIN_CREDENTIALS.password) {
    return res.json({ success: true, token: 'admin-authenticated-session-token' });
  }
  res.status(401).json({ error: 'Invalid admin username or password' });
});

// 5. Admin Get All Orders
app.get('/api/admin/orders', (req, res) => {
  res.json(orders);
});

// 6. Admin Delete/Cancel Order
app.delete('/api/admin/orders/:id', (req, res) => {
  const { id } = req.params;
  const index = orders.findIndex(o => o.id === id);
  if (index !== -1) {
    orders.splice(index, 1);
    return res.json({ success: true });
  }
  res.status(404).json({ error: 'Order not found' });
});

// 7. Admin Download Daily Excel Log
app.get('/api/admin/download-excel', (req, res) => {
  const today = new Date().toISOString().split('T')[0];
  const excelPath = path.join(__dirname, `../daily_log_${today}.xlsx`);

  if (fs.existsSync(excelPath)) {
    res.download(excelPath, `PrintNepal_DailyLog_${today}.xlsx`);
  } else {
    res.status(404).json({ error: 'No order log found for today yet.' });
  }
});

// 8. Operator Endpoint: Get Pending Jobs
app.get('/api/agent/pending', (req, res) => {
  const pendingJobs = orders.filter(o => o.status === 'pending');
  res.json(pendingJobs);
});

// 9. Operator Endpoint: Update Status & Trigger Physical Auto-Print
app.post('/api/agent/update-status', async (req, res) => {
  const { id, status } = req.body;
  const order = orders.find(o => o.id === id);

  if (order) {
    order.status = status;

    // Automatically trigger local physical printer when marked 'printing'
    if (status === 'printing') {
      autoPrintDocument(order.filePath, order.copies);
    }

    // Append to daily excel log when marked 'ready'
    if (status === 'ready') {
      await updateDailyExcelReport(order);
    }

    return res.json({ success: true, order });
  }
  res.status(404).json({ error: 'Order not found' });
});

// Trigger initial cleanup check on startup
autoDeleteOldFiles();

app.listen(PORT, () => console.log(`PrintNepal Server running on http://localhost:${PORT}`));