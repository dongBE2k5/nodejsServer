
// server.js
// server.js
// server.js
// server.js
const express = require('express');
const multer = require('multer');
const JSZip = require('jszip');
const { DOMParser } = require('xmldom');
const cors = require('cors');

const app = express();
const port = 8000;

app.use(cors());
app.use(express.json());

const upload = multer({ storage: multer.memoryStorage() });

// --- CÁC HÀM HỖ TRỢ CHUNG ---

const TWIPS_TO_PX = 96 / 1440;
const twipsToPx = (twips) => parseFloat(twips) * TWIPS_TO_PX;
const halfPointsToPx = (halfPoints) => (parseFloat(halfPoints) / 2) * (96 / 72);

const getChildElementByLocalName = (parent, localName) => {
    if (!parent || !parent.childNodes) return null;
    for (let i = 0; i < parent.childNodes.length; i++) {
        const child = parent.childNodes[i];
        if (child.nodeType === 1 && child.localName === localName) { // nodeType 1 is ELEMENT_NODE
            return child;
        }
    }
    return null;
};

// --- API ENDPOINTS ---

app.post('/api/convert-docx-to-html', upload.single('docx'), async (req, res) => {
    try {
        if (!req.file) {
            return res.status(400).json({ error: 'Chưa tải file DOCX.' });
        }
        const docxBuffer = req.file.buffer;
        console.log(`Đã nhận file DOCX. Kích thước: ${docxBuffer.length} bytes`);

        const zip = await JSZip.loadAsync(docxBuffer);
        const documentXmlFile = zip.file('word/document.xml');
        if (!documentXmlFile) {
            return res.status(400).json({ error: 'Không tìm thấy document.xml trong file DOCX.' });
        }
        const documentXmlString = await documentXmlFile.async('string');

        const numberingXmlFile = zip.file('word/numbering.xml');
        let numberingDefinitions = {};
        if (numberingXmlFile) {
            const numberingXmlString = await numberingXmlFile.async('string');
            numberingDefinitions = parseNumberingXml(numberingXmlString);
        }

        const parser = new DOMParser();
        const docXml = parser.parseFromString(documentXmlString, 'text/xml');
        const body = getChildElementByLocalName(docXml.documentElement, 'body');

        if (!body) {
            return res.status(400).json({ error: 'Không tìm thấy phần tử w:body trong document.xml.' });
        }

        const htmlOutput = await processDocxNodes(body.childNodes, numberingDefinitions, zip);
        console.log('Đã hoàn tất chuyển đổi DOCX sang HTML.');
        res.json({ html: htmlOutput });
    } catch (error) {
        console.error('Lỗi nghiêm trọng trong quá trình xử lý DOCX:', error);
        res.status(500).json({ error: 'Không thể xử lý file DOCX: ' + error.message });
    }
});

app.listen(port, () => {
    console.log(`API đang lắng nghe tại http://localhost:${port}`);
});

// --- CÁC HÀM XỬ LÝ NODE ---

// Thêm case 'sdt' vào hàm processDocxNodes của bạn
async function processDocxNodes(nodes, numberingDefinitions, zip) {
    let html = '';
    for (let i = 0; i < nodes.length; i++) {
        const node = nodes[i];
        if (!node || node.nodeType !== 1) continue;

        switch (node.localName) {
            case 'p':
                // Giữ nguyên logic xử lý <p> của bạn
                html += (await processParagraphNode(node, numberingDefinitions, zip)).pHtml;
                break;
            case 'tbl':
                html += await processTableNode(node, numberingDefinitions, zip);
                break;
            
            // === THÊM CASE MỚI ĐỂ XỬ LÝ CHECKBOX VÀ CÁC NỘI DUNG KHÁC ===
            case 'sdt': // Structured Document Tag (Content Control)
                const sdtContent = getChildElementByLocalName(node, 'sdtContent');
                if (sdtContent) {
                    // Đệ quy để xử lý nội dung bên trong sdtContent (thường là các thẻ <p>)
                    html += await processDocxNodes(sdtContent.childNodes, numberingDefinitions, zip);
                }
                break;
            // === KẾT THÚC PHẦN THÊM MỚI ===
        }
    }
    // Logic xử lý list của bạn có thể cần giữ lại ở đây nếu có
    return html;
}

// =================================================================
// === HÀM XỬ LÝ ĐOẠN VĂN (PARAGRAPH) - NÂNG CẤP TAB LEADER ===
// =================================================================
// Thay thế hàm processParagraphNode hiện tại của bạn bằng hàm này
async function processParagraphNode(paragraphNode, numberingDefinitions, zip) {
    const pPr = getChildElementByLocalName(paragraphNode, 'pPr');
    let paragraphInlineStyles = ['margin: 0;', 'padding: 0;'];
    let hasDotLeader = false;

    if (pPr) {
        // Căn lề
        const jc = getChildElementByLocalName(pPr, 'jc');
        if (jc) {
            const alignVal = jc.getAttribute('w:val');
            paragraphInlineStyles.push(`text-align: ${alignVal === 'both' ? 'justify' : alignVal};`);
        }

        // Khoảng cách dòng
        const spacing = getChildElementByLocalName(pPr, 'spacing');
        if (spacing) {
            const before = spacing.getAttribute('w:before');
            if (before) paragraphInlineStyles.push(`margin-top: ${twipsToPx(before)}px;`);
            const after = spacing.getAttribute('w:after');
            if (after) paragraphInlineStyles.push(`margin-bottom: ${twipsToPx(after)}px;`);
            const line = spacing.getAttribute('w:line');
            if (line) paragraphInlineStyles.push(`line-height: ${spacing.getAttribute('w:lineRule') === 'auto' ? parseFloat(line) / 240 : twipsToPx(line) + 'px'};`);
        }
    }

    let paragraphHtmlContent = '';
    for (let i = 0; i < paragraphNode.childNodes.length; i++) {
        const childNode = paragraphNode.childNodes[i];
        if (childNode.nodeType === 1) { // Chỉ xử lý Element nodes
             // Truyền pPr xuống để processRunNode có thể biết ngữ cảnh của đoạn văn
             paragraphHtmlContent += await processRunNode(childNode, pPr);
        }
    }
    
    // Nếu nội dung có chứa tab-leader, chúng ta cần display: flex để nó hoạt động
    // Cách này vẫn có thể ảnh hưởng bảng, nhưng ít hơn.
    if (paragraphHtmlContent.includes('class="tab-leader"')) {
        paragraphInlineStyles.push('display: flex; align-items: baseline;');
    }

    return {
        pHtml: paragraphHtmlContent.trim() ? `<p style="${paragraphInlineStyles.join(' ')}">${paragraphHtmlContent}</p>` : '',
        listInfo: { isListItem: false } // Tạm thời chưa xử lý list trong phiên bản này
    };
}


// Thay thế hàm processRunNode hiện tại của bạn bằng hàm này
async function processRunNode(runNode, pPr) { // Nhận thêm pPr
    let runHtml = '';
    let textContent = '';
    const runInlineStyles = [];

    const rPr = getChildElementByLocalName(runNode, 'rPr');
    if (rPr) {
        if (getChildElementByLocalName(rPr, 'b')) runInlineStyles.push('font-weight: bold;');
        if (getChildElementByLocalName(rPr, 'i')) runInlineStyles.push('font-style: italic;');
        if (getChildElementByLocalName(rPr, 'u')) runInlineStyles.push('text-decoration: underline;');
        const sz = getChildElementByLocalName(rPr, 'sz');
        if (sz) runInlineStyles.push(`font-size: ${halfPointsToPx(sz.getAttribute('w:val'))}px;`);
        const color = getChildElementByLocalName(rPr, 'color');
        if (color && color.getAttribute('w:val') !== 'auto') runInlineStyles.push(`color: #${color.getAttribute('w:val')};`);
    }

    for (let k = 0; k < runNode.childNodes.length; k++) {
        const childNode = runNode.childNodes[k];
        if (childNode.nodeType === 3) {
            textContent += childNode.nodeValue;
        } else if (childNode.nodeType === 1) {
            if (childNode.localName === 't') {
                textContent += childNode.textContent;
            } else if (childNode.localName === 'br') {
                runHtml += '<br/>';
            } else if (childNode.localName === 'tab') {
                let hasDotLeader = false;
                if (pPr) { // Kiểm tra định nghĩa tab trong pPr
                    const tabs = getChildElementByLocalName(pPr, 'tabs');
                    if (tabs) {
                        const tabStops = tabs.getElementsByTagName('w:tab');
                        for (let i = 0; i < tabStops.length; i++) {
                            if (tabStops[i].getAttribute('w:leader') === 'dot') {
                                hasDotLeader = true;
                                break;
                            }
                        }
                    }
                }
                
                // Giải pháp đơn giản hơn: chỉ chèn một span chứa các dấu chấm
                // Nó sẽ không tự co giãn hoàn hảo nhưng sẽ không phá vỡ layout bảng
                if (hasDotLeader) {
                     runHtml += `<span class="tab-leader" style="padding: 0 0.5em;">........................................................</span>`;
                } else {
                    textContent += '\u00A0\u00A0\u00A0\u00A0';
                }
            }
        }
    }
    
    if (textContent) {
        runHtml += `<span style="${runInlineStyles.join(' ')}">${textContent.replace(/</g, '&lt;').replace(/>/g, '&gt;')}</span>`;
    }
    return runHtml;
}

async function processTableNode(tableNode, numberingDefinitions, zip) {
    const tableInlineStyles = ['border-collapse: collapse;', 'width: 100%;'];
    let tableHtml = `<table style="${tableInlineStyles.join(' ')}">`;
    const rowNodes = tableNode.getElementsByTagName('w:tr');
    for (let i = 0; i < rowNodes.length; i++) {
        tableHtml += '<tr>';
        const cellNodes = rowNodes[i].getElementsByTagName('w:tc');
        for (let j = 0; j < cellNodes.length; j++) {
            const cellNode = cellNodes[j];
            const cellAttrs = [];
            const cellInlineStyles = ['padding: 5px;'];
            const tcPr = getChildElementByLocalName(cellNode, 'tcPr');
            if (tcPr) {
                const gridSpan = getChildElementByLocalName(tcPr, 'gridSpan');
                if (gridSpan) cellAttrs.push(`colspan="${gridSpan.getAttribute('w:val')}"`);
                const vMerge = getChildElementByLocalName(tcPr, 'vMerge');
                if (vMerge && !vMerge.getAttribute('w:val')) continue;
                const shd = getChildElementByLocalName(tcPr, 'shd');
                if (shd && shd.getAttribute('w:fill') !== 'auto') cellInlineStyles.push(`background-color: #${shd.getAttribute('w:fill')};`);
                const vAlign = getChildElementByLocalName(tcPr, 'vAlign');
                if(vAlign) cellInlineStyles.push(`vertical-align: ${vAlign.getAttribute('w:val')};`);
            }
            let cellContentHtml = await processDocxNodes(cellNode.childNodes, numberingDefinitions, zip);
            tableHtml += `<td ${cellAttrs.join(' ')} style="${cellInlineStyles.join(' ')}">${cellContentHtml}</td>`;
        }
        tableHtml += '</tr>';
    }
    tableHtml += '</table>';
    return tableHtml;
}

function parseNumberingXml(numberingXmlString) {
    // This function remains the same as the previous version.
    const parser = new DOMParser();
    const docXml = parser.parseFromString(numberingXmlString, 'text/xml');
    const numberingDefinitions = {};
    const abstractNums = docXml.getElementsByTagName('w:abstractNum');
    for (let i = 0; i < abstractNums.length; i++) {
        const abstractNumId = abstractNums[i].getAttribute('w:abstractNumId');
        if (abstractNumId === null) continue;
        const levels = {};
        const lvlNodes = abstractNums[i].getElementsByTagName('w:lvl');
        for (let j = 0; j < lvlNodes.length; j++) {
            const ilvl = lvlNodes[j].getAttribute('w:ilvl');
            if(ilvl === null) continue;
            const numFmtNode = getChildElementByLocalName(lvlNodes[j], 'numFmt');
            levels[ilvl] = { fmt: numFmtNode ? numFmtNode.getAttribute('w:val') : 'bullet' };
        }
        numberingDefinitions[abstractNumId] = { type: 'abstract', levels };
    }
    const nums = docXml.getElementsByTagName('w:num');
    for (let i = 0; i < nums.length; i++) {
        const numId = nums[i].getAttribute('w:numId');
        if (numId === null) continue;
        const abstractNumIdRef = getChildElementByLocalName(nums[i], 'abstractNumId');
        if (abstractNumIdRef) {
            const abstractNumIdVal = abstractNumIdRef.getAttribute('w:val');
            if(abstractNumIdVal !== null) {
                 numberingDefinitions[numId] = { type: 'num', abstractNumId: abstractNumIdVal };
            }
        }
    }
    return numberingDefinitions;
}