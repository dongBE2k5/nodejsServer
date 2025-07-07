const express = require('express');
const cors = require('cors');
const fs = require('fs');
const path = require('path');

const generateDocx = require('./export-docx');
const extractVariablesRouter = require('./extractVariablesRouter');

const app = express();
app.use(cors());
app.use(express.json());

// API tạo file docx từ dữ liệu
app.post('/generate-docx', (req, res) => {
  const data = req.body;

  try {
    generateDocx(data);
    const filePath = path.resolve(__dirname, 'output.docx');
    res.download(filePath, 'download.docx');
  } catch (err) {
    console.error(err);
    res.status(500).json({ error: 'Lỗi tạo file docx' });
  }
});

// API phân tích biến trong file .docx (dùng extractVariablesRouter)
app.use('/', extractVariablesRouter);

// Khởi động server
const PORT = 4000;
app.listen(PORT, () => {
  console.log(`✅ Server đang chạy tại http://localhost:${PORT}`);
});
