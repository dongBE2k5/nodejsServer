

// server/index.js

const bodyParser = require('body-parser');
const cors = require('cors'); // 👈 Thêm dòng này
const generateDocx = require('./export-docx');
const express = require('express');
const multer = require('multer');

const path = require('path');
const { exec } = require('child_process');
const fs = require('fs');

const upload = multer({ dest: 'uploads/' });

const app = express();

// 👇 Cho phép tất cả các origin (hoặc cấu hình cụ thể hơn bên dưới)
app.use(cors());

app.use(bodyParser.json());

app.post('/generate', (req, res) => {
  try {
    generateDocx(req.body);
    res.download('output.docx');
  } catch (err) {
    console.error(err);
    res.status(500).send('Có lỗi xảy ra khi tạo file Word');
  }
});

app.listen(3001, () => {
  console.log('Server running on http://localhost:3001');
});



app.post('/convert', upload.single('docx'), (req, res) => {
  const file = req.file;
  const outputDir = path.join(__dirname, 'converted');
  const outputFile = path.join(outputDir, file.filename + '.html');

  // Tạo thư mục nếu chưa có
  if (!fs.existsSync(outputDir)) {
    fs.mkdirSync(outputDir);
  }

  // Chuyển file docx → HTML bằng libreoffice CLI
  const cmd = `libreoffice --headless --convert-to html:"HTML (StarWriter)" "${file.path}" --outdir "${outputDir}"`;
  exec(cmd, (error, stdout, stderr) => {
    if (error) {
      console.error(stderr);
      return res.status(500).send('Lỗi chuyển đổi file');
    }

    // Trả về HTML file nội dung
    fs.readFile(outputFile, 'utf8', (err, html) => {
      if (err) {
        return res.status(500).send('Lỗi đọc file HTML');
      }
      res.send({ html });
      // Cleanup nếu cần
      fs.unlinkSync(file.path);
      fs.unlinkSync(outputFile);
    });
  });
});

app.listen(3001, () => {
  console.log('Server running on http://localhost:3001');
});