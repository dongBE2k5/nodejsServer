// extractVariablesRouter.js
const express = require('express');
const multer = require('multer');
const fs = require('fs');
const path = require('path');
const PizZip = require('pizzip');
const Docxtemplater = require('docxtemplater');

const router = express.Router();
const upload = multer({ dest: 'uploads/' });

router.post('/extract-variables', upload.single('file'), (req, res) => {
  try {
    const content = fs.readFileSync(req.file.path, 'binary');
    const zip = new PizZip(content);
    const doc = new Docxtemplater(zip);
    const tags = doc.getFullText();

    const regex = /{(.*?)}/g; // hoặc /{(.*?)}/g nếu bạn dùng {name}
    const matches = [];
    let match;
    while ((match = regex.exec(tags)) !== null) {
      matches.push(match[1]);
    }

    fs.unlinkSync(req.file.path); // Xoá file tạm
    res.json({ variables: matches });
  } catch (error) {
    console.error(error);
    res.status(500).json({ error: 'Không thể phân tích file' });
  }
});

module.exports = router;
