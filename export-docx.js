const fs = require('fs');
const path = require('path');
const PizZip = require('pizzip');
const Docxtemplater = require('docxtemplater');

function generateDocx(data) {
  const content = fs.readFileSync(path.resolve(__dirname, 'document.docx'), 'binary');
  const zip = new PizZip(content);
  const doc = new Docxtemplater(zip, { paragraphLoop: true, linebreaks: true });

try {
  doc.render(data);
} catch (error) {
  const e = {
    message: error.message,
    name: error.name,
    stack: error.stack,
    properties: error.properties,
  };
  console.log(JSON.stringify({ error: e }, null, 2));
  throw error;
}

  const buf = doc.getZip().generate({ type: 'nodebuffer' });
  fs.writeFileSync(path.resolve(__dirname, 'output.docx'), buf);
}

// 👇 Export hàm
module.exports = generateDocx;