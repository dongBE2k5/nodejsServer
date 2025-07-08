
// // server.js
// // server.js
// // server.js
// // server.js
// const express = require('express');
// const multer = require('multer');
// const JSZip = require('jszip');
// const { DOMParser } = require('xmldom');
// const cors = require('cors');

// const app = express();
// const port = 8000;

// app.use(cors());
// app.use(express.json());

// const upload = multer({ storage: multer.memoryStorage() });

// // --- CÁC HÀM HỖ TRỢ CHUNG ---

// const TWIPS_TO_PX = 96 / 1440;
// const twipsToPx = (twips) => parseFloat(twips) * TWIPS_TO_PX;
// const halfPointsToPx = (halfPoints) => (parseFloat(halfPoints) / 2) * (96 / 72);

// const getChildElementByLocalName = (parent, localName) => {
//     if (!parent || !parent.childNodes) return null;
//     for (let i = 0; i < parent.childNodes.length; i++) {
//         const child = parent.childNodes[i];
//         if (child.nodeType === 1 && child.localName === localName) { // nodeType 1 is ELEMENT_NODE
//             return child;
//         }
//     }
//     return null;
// };

// // --- API ENDPOINTS ---

// app.post('/api/convert-docx-to-html', upload.single('docx'), async (req, res) => {
//     try {
//         if (!req.file) {
//             return res.status(400).json({ error: 'Chưa tải file DOCX.' });
//         }
//         const docxBuffer = req.file.buffer;
//         console.log(`Đã nhận file DOCX. Kích thước: ${docxBuffer.length} bytes`);

//         const zip = await JSZip.loadAsync(docxBuffer);
//         const documentXmlFile = zip.file('word/document.xml');
//         if (!documentXmlFile) {
//             return res.status(400).json({ error: 'Không tìm thấy document.xml trong file DOCX.' });
//         }
//         const documentXmlString = await documentXmlFile.async('string');

//         const numberingXmlFile = zip.file('word/numbering.xml');
//         let numberingDefinitions = {};
//         if (numberingXmlFile) {
//             const numberingXmlString = await numberingXmlFile.async('string');
//             numberingDefinitions = parseNumberingXml(numberingXmlString);
//         }

//         const parser = new DOMParser();
//         const docXml = parser.parseFromString(documentXmlString, 'text/xml');
//         const body = getChildElementByLocalName(docXml.documentElement, 'body');

//         if (!body) {
//             return res.status(400).json({ error: 'Không tìm thấy phần tử w:body trong document.xml.' });
//         }

//         const htmlOutput = await processDocxNodes(body.childNodes, numberingDefinitions, zip);
//         console.log('Đã hoàn tất chuyển đổi DOCX sang HTML.');
//         res.json({ html: htmlOutput });
//     } catch (error) {
//         console.error('Lỗi nghiêm trọng trong quá trình xử lý DOCX:', error);
//         res.status(500).json({ error: 'Không thể xử lý file DOCX: ' + error.message });
//     }
// });

// app.listen(port, () => {
//     console.log(`API đang lắng nghe tại http://localhost:${port}`);
// });

// // --- CÁC HÀM XỬ LÝ NODE ---

// // Thêm case 'sdt' vào hàm processDocxNodes của bạn
// async function processDocxNodes(nodes, numberingDefinitions, zip) {
//     let html = '';
//     for (let i = 0; i < nodes.length; i++) {
//         const node = nodes[i];
//         if (!node || node.nodeType !== 1) continue;

//         switch (node.localName) {
//             case 'p':
//                 // Giữ nguyên logic xử lý <p> của bạn
//                 html += (await processParagraphNode(node, numberingDefinitions, zip)).pHtml;
//                 break;
//             case 'tbl':
//                 html += await processTableNode(node, numberingDefinitions, zip);
//                 break;

//             // === THÊM CASE MỚI ĐỂ XỬ LÝ CHECKBOX VÀ CÁC NỘI DUNG KHÁC ===
//             case 'sdt': // Structured Document Tag (Content Control)
//                 const sdtContent = getChildElementByLocalName(node, 'sdtContent');
//                 if (sdtContent) {
//                     // Đệ quy để xử lý nội dung bên trong sdtContent (thường là các thẻ <p>)
//                     html += await processDocxNodes(sdtContent.childNodes, numberingDefinitions, zip);
//                 }
//                 break;
//             // === KẾT THÚC PHẦN THÊM MỚI ===
//         }
//     }
//     // Logic xử lý list của bạn có thể cần giữ lại ở đây nếu có
//     return html;
// }

// // =================================================================
// // === HÀM XỬ LÝ ĐOẠN VĂN (PARAGRAPH) - NÂNG CẤP TAB LEADER ===
// // =================================================================
// // Thay thế hàm processParagraphNode hiện tại của bạn bằng hàm này
// async function processParagraphNode(paragraphNode, numberingDefinitions, zip) {
//     const pPr = getChildElementByLocalName(paragraphNode, 'pPr');
//     let paragraphInlineStyles = ['margin: 0;', 'padding: 0;'];
//     let hasDotLeader = false;

//     if (pPr) {
//         // Căn lề
//         const jc = getChildElementByLocalName(pPr, 'jc');
//         if (jc) {
//             const alignVal = jc.getAttribute('w:val');
//             paragraphInlineStyles.push(`text-align: ${alignVal === 'both' ? 'justify' : alignVal};`);
//         }

//         // Khoảng cách dòng
//         const spacing = getChildElementByLocalName(pPr, 'spacing');
//         if (spacing) {
//             const before = spacing.getAttribute('w:before');
//             if (before) paragraphInlineStyles.push(`margin-top: ${twipsToPx(before)}px;`);
//             const after = spacing.getAttribute('w:after');
//             if (after) paragraphInlineStyles.push(`margin-bottom: ${twipsToPx(after)}px;`);
//             const line = spacing.getAttribute('w:line');
//             if (line) paragraphInlineStyles.push(`line-height: ${spacing.getAttribute('w:lineRule') === 'auto' ? parseFloat(line) / 240 : twipsToPx(line) + 'px'};`);
//         }
//     }

//     let paragraphHtmlContent = '';
//     for (let i = 0; i < paragraphNode.childNodes.length; i++) {
//         const childNode = paragraphNode.childNodes[i];
//         if (childNode.nodeType === 1) { // Chỉ xử lý Element nodes
//              // Truyền pPr xuống để processRunNode có thể biết ngữ cảnh của đoạn văn
//              paragraphHtmlContent += await processRunNode(childNode, pPr);
//         }
//     }

//     // Nếu nội dung có chứa tab-leader, chúng ta cần display: flex để nó hoạt động
//     // Cách này vẫn có thể ảnh hưởng bảng, nhưng ít hơn.
//     if (paragraphHtmlContent.includes('class="tab-leader"')) {
//         paragraphInlineStyles.push('display: flex; align-items: baseline;');
//     }

//     return {
//         pHtml: paragraphHtmlContent.trim() ? `<p style="${paragraphInlineStyles.join(' ')}">${paragraphHtmlContent}</p>` : '',
//         listInfo: { isListItem: false } // Tạm thời chưa xử lý list trong phiên bản này
//     };
// }


// // Thay thế hàm processRunNode hiện tại của bạn bằng hàm này
// async function processRunNode(runNode, pPr) { // Nhận thêm pPr
//     let runHtml = '';
//     let textContent = '';
//     const runInlineStyles = [];

//     const rPr = getChildElementByLocalName(runNode, 'rPr');
//     if (rPr) {
//         if (getChildElementByLocalName(rPr, 'b')) runInlineStyles.push('font-weight: bold;');
//         if (getChildElementByLocalName(rPr, 'i')) runInlineStyles.push('font-style: italic;');
//         if (getChildElementByLocalName(rPr, 'u')) runInlineStyles.push('text-decoration: underline;');
//         const sz = getChildElementByLocalName(rPr, 'sz');
//         if (sz) runInlineStyles.push(`font-size: ${halfPointsToPx(sz.getAttribute('w:val'))}px;`);
//         const color = getChildElementByLocalName(rPr, 'color');
//         if (color && color.getAttribute('w:val') !== 'auto') runInlineStyles.push(`color: #${color.getAttribute('w:val')};`);
//     }

//     for (let k = 0; k < runNode.childNodes.length; k++) {
//         const childNode = runNode.childNodes[k];
//         if (childNode.nodeType === 3) {
//             textContent += childNode.nodeValue;
//         } else if (childNode.nodeType === 1) {
//             if (childNode.localName === 't') {
//                 textContent += childNode.textContent;
//             } else if (childNode.localName === 'br') {
//                 runHtml += '<br/>';
//             } else if (childNode.localName === 'tab') {
//                 let hasDotLeader = false;
//                 if (pPr) { // Kiểm tra định nghĩa tab trong pPr
//                     const tabs = getChildElementByLocalName(pPr, 'tabs');
//                     if (tabs) {
//                         const tabStops = tabs.getElementsByTagName('w:tab');
//                         for (let i = 0; i < tabStops.length; i++) {
//                             if (tabStops[i].getAttribute('w:leader') === 'dot') {
//                                 hasDotLeader = true;
//                                 break;
//                             }
//                         }
//                     }
//                 }

//                 // Giải pháp đơn giản hơn: chỉ chèn một span chứa các dấu chấm
//                 // Nó sẽ không tự co giãn hoàn hảo nhưng sẽ không phá vỡ layout bảng
//                 if (hasDotLeader) {
//                      runHtml += `<span class="tab-leader" style="padding: 0 0.5em;">........................................................</span>`;
//                 } else {
//                     textContent += '\u00A0\u00A0\u00A0\u00A0';
//                 }
//             }
//         }
//     }

//     if (textContent) {
//         runHtml += `<span style="${runInlineStyles.join(' ')}">${textContent.replace(/</g, '&lt;').replace(/>/g, '&gt;')}</span>`;
//     }
//     return runHtml;
// }

// async function processTableNode(tableNode, numberingDefinitions, zip) {
//     const tableInlineStyles = ['border-collapse: collapse;', 'width: 100%;'];
//     let tableHtml = `<table style="${tableInlineStyles.join(' ')}">`;
//     const rowNodes = tableNode.getElementsByTagName('w:tr');
//     for (let i = 0; i < rowNodes.length; i++) {
//         tableHtml += '<tr>';
//         const cellNodes = rowNodes[i].getElementsByTagName('w:tc');
//         for (let j = 0; j < cellNodes.length; j++) {
//             const cellNode = cellNodes[j];
//             const cellAttrs = [];
//             const cellInlineStyles = ['padding: 5px;'];
//             const tcPr = getChildElementByLocalName(cellNode, 'tcPr');
//             if (tcPr) {
//                 const gridSpan = getChildElementByLocalName(tcPr, 'gridSpan');
//                 if (gridSpan) cellAttrs.push(`colspan="${gridSpan.getAttribute('w:val')}"`);
//                 const vMerge = getChildElementByLocalName(tcPr, 'vMerge');
//                 if (vMerge && !vMerge.getAttribute('w:val')) continue;
//                 const shd = getChildElementByLocalName(tcPr, 'shd');
//                 if (shd && shd.getAttribute('w:fill') !== 'auto') cellInlineStyles.push(`background-color: #${shd.getAttribute('w:fill')};`);
//                 const vAlign = getChildElementByLocalName(tcPr, 'vAlign');
//                 if(vAlign) cellInlineStyles.push(`vertical-align: ${vAlign.getAttribute('w:val')};`);
//             }
//             let cellContentHtml = await processDocxNodes(cellNode.childNodes, numberingDefinitions, zip);
//             tableHtml += `<td ${cellAttrs.join(' ')} style="${cellInlineStyles.join(' ')}">${cellContentHtml}</td>`;
//         }
//         tableHtml += '</tr>';
//     }
//     tableHtml += '</table>';
//     return tableHtml;
// }

// function parseNumberingXml(numberingXmlString) {
//     // This function remains the same as the previous version.
//     const parser = new DOMParser();
//     const docXml = parser.parseFromString(numberingXmlString, 'text/xml');
//     const numberingDefinitions = {};
//     const abstractNums = docXml.getElementsByTagName('w:abstractNum');
//     for (let i = 0; i < abstractNums.length; i++) {
//         const abstractNumId = abstractNums[i].getAttribute('w:abstractNumId');
//         if (abstractNumId === null) continue;
//         const levels = {};
//         const lvlNodes = abstractNums[i].getElementsByTagName('w:lvl');
//         for (let j = 0; j < lvlNodes.length; j++) {
//             const ilvl = lvlNodes[j].getAttribute('w:ilvl');
//             if(ilvl === null) continue;
//             const numFmtNode = getChildElementByLocalName(lvlNodes[j], 'numFmt');
//             levels[ilvl] = { fmt: numFmtNode ? numFmtNode.getAttribute('w:val') : 'bullet' };
//         }
//         numberingDefinitions[abstractNumId] = { type: 'abstract', levels };
//     }
//     const nums = docXml.getElementsByTagName('w:num');
//     for (let i = 0; i < nums.length; i++) {
//         const numId = nums[i].getAttribute('w:numId');
//         if (numId === null) continue;
//         const abstractNumIdRef = getChildElementByLocalName(nums[i], 'abstractNumId');
//         if (abstractNumIdRef) {
//             const abstractNumIdVal = abstractNumIdRef.getAttribute('w:val');
//             if(abstractNumIdVal !== null) {
//                  numberingDefinitions[numId] = { type: 'num', abstractNumId: abstractNumIdVal };
//             }
//         }
//     }
//     return numberingDefinitions;
// }



const JSZip = require('jszip');
const { DOMParser } = require('xmldom');

// --- CÁC HÀM HỖ TRỢ CHUNG ---

// DPI mặc định cho web là 96
const TWIPS_PER_INCH = 1440; // 1 inch = 1440 twips
const POINTS_PER_INCH = 72; // 1 inch = 72 points
const PIXELS_PER_INCH = 96; // 1 inch = 96 pixels (standard web DPI)

const twipsToPx = (twips) => parseFloat(twips) * (PIXELS_PER_INCH / TWIPS_PER_INCH); // 1440 twips = 1 inch = 96px => 1 twip = 96/1440 px
const halfPointsToPx = (halfPoints) => parseFloat(halfPoints) * (PIXELS_PER_INCH / (POINTS_PER_INCH * 2)); // half-points = point / 2 => (96/72) / 2 px

const getChildElementByLocalName = (parent, localName) => {
    if (!parent || !parent.childNodes) return null;
    for (let i = 0; i < parent.childNodes.length; i++) {
        const child = parent.childNodes[i];
        if (child.nodeType === 1 && child.localName === localName) {
            return child;
        }
    }
    return null;
};

// Hàm hợp nhất các thuộc tính style, các thuộc tính sau ghi đè các thuộc tính trước
const mergeProperties = (baseProps, newProps) => {
    const merged = { ...baseProps };
    for (const key in newProps) {
        if (newProps[key] !== undefined) {
            merged[key] = newProps[key];
        }
    }
    return merged;
};

// Hàm chuyển đổi object style props sang chuỗi CSS inline
const propsToCssString = (props) => {
    const css = [];
    for (const key in props) {
        let value = props[key];
        let cssKey = key.replace(/([A-Z])/g, '-$1').toLowerCase(); // camelCase to kebab-case

        // Add 'px' for numeric pixel values if unit is not already specified
        if (typeof value === 'number' && key !== 'lineHeight') { // lineHeight can be a unitless factor
            value = `${value}px`;
        }
        
        if (value !== undefined && value !== null && value !== '') {
            css.push(`${cssKey}: ${value};`);
        }
    }
    return css.join(' ');
};

// --- HÀM XỬ LÝ CÁC LOẠI NODE CỤ THỂ ---

async function processDocxNodes(nodes, numberingDefinitions, zip, documentStyles) {
    let html = '';
    // Theo dõi trạng thái của danh sách hiện tại để đóng/mở thẻ ul/ol
    let currentListState = {
        isListItem: false,
        level: -1, // -1 means not in a list
        numId: null, // To track the specific numbering instance
        abstractNumId: null // To get list format (bullet/decimal)
    };

    for (let i = 0; i < nodes.length; i++) {
        const node = nodes[i];
        if (!node || node.nodeType !== 1) continue;

        if (node.localName === 'p') {
            const { pContentHtml, pInlineStyles, listProperties } = await processParagraphNode(node, numberingDefinitions, zip, documentStyles);

            // Xử lý danh sách (<ul>/<ol> và <li>)
            if (listProperties.isListItem) {
                // Determine the correct parent list tag (ul or ol)
                const abstractNumDef = numberingDefinitions[listProperties.abstractNumId];
                const levelDef = abstractNumDef?.levels[listProperties.level];
                const numFmt = levelDef?.fmt || 'decimal'; // default to decimal
                const lvlText = levelDef?.lvlText; // Custom bullet/number text
                
                const listContainerTag = (numFmt === 'bullet' || numFmt === 'hybridMultilevel' || numFmt === 'none') ? 'ul' : 'ol'; // 'none' for custom bullet via lvlText

                // Close existing lists if level decreases or list type/id changes
                while (currentListState.isListItem &&
                       (listProperties.level < currentListState.level || // Level decreased
                        (listProperties.level === currentListState.level && listProperties.numId !== currentListState.numId) // Same level, different list
                       )) {
                    if (currentListState.level >= 0) { // Only close if a list was actually open
                        const currentAbstractNumDef = numberingDefinitions[currentListState.abstractNumId];
                        const currentNumFmt = currentAbstractNumDef?.levels[currentListState.level]?.fmt || 'decimal';
                        const currentListContainerTag = (currentNumFmt === 'bullet' || currentNumFmt === 'hybridMultilevel' || currentNumFmt === 'none') ? 'ul' : 'ol';
                        html += `</${currentListContainerTag}>`;
                    }
                    currentListState.level--; // Decrement current level to match the new item's level or exit loop
                }
                
                // Open new lists if level increases or a new list starts
                if (!currentListState.isListItem || listProperties.level > currentListState.level || listProperties.numId !== currentListState.numId) {
                    // Apply indentation from the numbering definition
                    const levelIndentPx = levelDef?.indent?.left || 0;
                    html += `<${listContainerTag} style="margin-left: ${levelIndentPx}px; padding-left: 0;">`; // Let the list item handle its own prefix/indent
                }
                
                let listItemStyle = pInlineStyles.join(' ');
                let bulletOrNumberPrefix = '';

                if (lvlText) {
                    // For custom bullet/number text, we'll try to put it in a pseudo-element via CSS
                    // or directly if it's a simple character and `list-style-type: none;` is used.
                    // For complex cases like "%1.", it typically relies on CSS counters which
                    // are best handled by external CSS. For inline HTML, we can try to render it directly.
                    bulletOrNumberPrefix = `<span style="display: inline-block; width: 20px; margin-left: -20px; text-align: right; margin-right: 5px;">${lvlText.replace(/%(\d)/g, '')}</span>`; // Remove %1, %2 etc. for simplicity
                    // Also, enforce list-style-type: none for custom prefixes
                    listItemStyle += ' list-style-type: none;';
                } else {
                    // For standard bullets/numbers, let browser handle via list-style-type
                    listItemStyle += ` list-style-type: ${numFmt === 'bullet' ? 'disc' : 'decimal'};`;
                }

                html += `<li style="${listItemStyle}">${bulletOrNumberPrefix}${pContentHtml}</li>`;

                // Update current list state
                currentListState = {
                    isListItem: true,
                    level: listProperties.level,
                    numId: listProperties.numId,
                    abstractNumId: listProperties.abstractNumId
                };

            } else { // Current node is not a list item
                // Close any open lists
                while (currentListState.isListItem && currentListState.level >= 0) {
                    const currentAbstractNumDef = numberingDefinitions[currentListState.abstractNumId];
                    const currentNumFmt = currentAbstractNumDef?.levels[currentListState.level]?.fmt || 'decimal';
                    const currentListContainerTag = (currentNumFmt === 'bullet' || currentNumFmt === 'hybridMultilevel' || currentNumFmt === 'none') ? 'ul' : 'ol';
                    html += `</${currentListContainerTag}>`;
                    currentListState.level--;
                }
                // Reset list state
                currentListState = { isListItem: false, level: -1, numId: null, abstractNumId: null };

                // Add regular paragraph HTML
                // Only add <p> if there's content or if it's an explicit empty paragraph (e.g., <w:p/>)
                // FIX: Changed paragraphNode to node here.
                if (pContentHtml.trim() || node.childNodes.length === 0) { 
                    html += `<p style="${pInlineStyles.join(' ')}">${pContentHtml}</p>`;
                }
            }
        } else {
            // For other node types (tables, sdt, etc.), close any open lists
            while (currentListState.isListItem && currentListState.level >= 0) {
                const currentAbstractNumDef = numberingDefinitions[currentListState.abstractNumId];
                const currentNumFmt = currentAbstractNumDef?.levels[currentListState.level]?.fmt || 'decimal';
                const currentListContainerTag = (currentNumFmt === 'bullet' || currentNumFmt === 'hybridMultilevel' || currentNumFmt === 'none') ? 'ul' : 'ol';
                html += `</${currentListContainerTag}>`;
                currentListState.level--;
            }
            currentListState = { isListItem: false, level: -1, numId: null, abstractNumId: null }; // Ensure reset

            switch (node.localName) {
                case 'tbl':
                    html += await processTableNode(node, numberingDefinitions, zip, documentStyles);
                    break;
                case 'sdt':
                    const sdtContent = getChildElementByLocalName(node, 'sdtContent');
                    if (sdtContent) {
                        html += await processDocxNodes(sdtContent.childNodes, numberingDefinitions, zip, documentStyles);
                    }
                    break;
                case 'sectPr':
                    // Here you could parse page margins (w:pgMar) from sectPr
                    // and apply them to the body style or a wrapper div.
                    // For simplicity, we are using a fixed margin in the final HTML structure.
                    break;
                case 'bookmarkStart': // Ignore bookmarks
                case 'bookmarkEnd':
                    break;
                default:
                    // Process children of unknown nodes
                    html += await processDocxNodes(node.childNodes, numberingDefinitions, zip, documentStyles);
                    break;
            }
        }
    }
    // Close any remaining open lists at the end of the document
    while (currentListState.isListItem && currentListState.level >= 0) {
        const currentAbstractNumDef = numberingDefinitions[currentListState.abstractNumId];
        const currentNumFmt = currentAbstractNumDef?.levels[currentListState.level]?.fmt || 'decimal';
        const currentListContainerTag = (currentNumFmt === 'bullet' || currentNumFmt === 'hybridMultilevel' || currentNumFmt === 'none') ? 'ul' : 'ol';
        html += `</${currentListContainerTag}>`;
        currentListState.level--;
    }

    return html;
}

async function processParagraphNode(paragraphNode, numberingDefinitions, zip, documentStyles) {
    const pPr = getChildElementByLocalName(paragraphNode, 'pPr');
    let effectiveParagraphProps = {};
    let listProperties = { isListItem: false, level: -1, numId: null, abstractNumId: null };

    if (pPr) {
        // 1. Apply style from w:pStyle (if present)
        const paragraphStyleId = getChildElementByLocalName(pPr, 'pStyle')?.getAttribute('w:val');
        if (paragraphStyleId && documentStyles[paragraphStyleId]) {
            let currentStyle = documentStyles[paragraphStyleId];
            let styleChainProps = {}; // Accumulate properties through the basedOn chain
            while (currentStyle) {
                styleChainProps = mergeProperties(currentStyle.paragraphProps, styleChainProps); // New properties override older ones in chain
                if (currentStyle.basedOn && documentStyles[currentStyle.basedOn]) {
                    currentStyle = documentStyles[currentStyle.basedOn];
                } else {
                    currentStyle = null;
                }
            }
            effectiveParagraphProps = mergeProperties(effectiveParagraphProps, styleChainProps);
        }

        // 2. Apply direct paragraph properties (w:pPr) - these override style properties
        const jc = getChildElementByLocalName(pPr, 'jc');
        if (jc) effectiveParagraphProps.textAlign = jc.getAttribute('w:val') === 'both' ? 'justify' : jc.getAttribute('w:val');

        const spacing = getChildElementByLocalName(pPr, 'spacing');
        if (spacing) {
            if (spacing.getAttribute('w:before')) effectiveParagraphProps.marginTop = twipsToPx(spacing.getAttribute('w:before'));
            if (spacing.getAttribute('w:after')) effectiveParagraphProps.marginBottom = twipsToPx(spacing.getAttribute('w:after'));
            const line = spacing.getAttribute('w:line');
            const lineRule = spacing.getAttribute('w:lineRule');
            if (line) {
                // Word's line values are in twips. lineRule affects interpretation.
                // "auto" means a factor (e.g. 1.0, 1.5). 240 twips is approx 1 single line.
                // "atLeast" and "exact" mean fixed heights.
                if (lineRule === 'auto') {
                    // Common approximation: line value is in twips, 240 twips corresponds to 1 line for 12pt font.
                    // So, line / 240 gives a rough factor.
                    effectiveParagraphProps.lineHeight = parseFloat(line) / 240; 
                } else {
                    effectiveParagraphProps.lineHeight = twipsToPx(line); // fixed pixel height
                }
            }
        }

        const ind = getChildElementByLocalName(pPr, 'ind');
        if (ind) {
            if (ind.getAttribute('w:start')) effectiveParagraphProps.marginLeft = twipsToPx(ind.getAttribute('w:start'));
            if (ind.getAttribute('w:end')) effectiveParagraphProps.marginRight = twipsToPx(ind.getAttribute('w:end'));
            if (ind.getAttribute('w:firstLine')) effectiveParagraphProps.textIndent = twipsToPx(ind.getAttribute('w:firstLine'));
            if (ind.getAttribute('w:hanging')) {
                // For hanging indent, set negative text-indent
                const hangingVal = twipsToPx(ind.getAttribute('w:hanging'));
                effectiveParagraphProps.textIndent = -hangingVal;
                // And ensure margin-left is at least the hanging amount for the text block to clear
                if (!effectiveParagraphProps.marginLeft || effectiveParagraphProps.marginLeft < hangingVal) {
                    effectiveParagraphProps.marginLeft = hangingVal;
                }
            }
        }

        // Numbering properties (for lists) - these affect both paragraph styling and list structure
        const numPr = getChildElementByLocalName(pPr, 'numPr');
        if (numPr) {
            const ilvlNode = getChildElementByLocalName(numPr, 'ilvl');
            const numIdNode = getChildElementByLocalName(numPr, 'numId');

            if (ilvlNode && numIdNode) {
                const ilvl = ilvlNode.getAttribute('w:val');
                const numId = numIdNode.getAttribute('w:val');

                const numDef = numberingDefinitions[numId];
                if (numDef && numDef.type === 'num' && numDef.abstractNumId) {
                    const abstractNumDef = numberingDefinitions[numDef.abstractNumId];
                    if (abstractNumDef && abstractNumDef.type === 'abstract' && abstractNumDef.levels[ilvl]) {
                        listProperties.isListItem = true;
                        listProperties.level = parseInt(ilvl, 10);
                        listProperties.numId = numId;
                        listProperties.abstractNumId = numDef.abstractNumId;

                        // When a paragraph is a list item, its indentation is primarily driven by the list definition.
                        // We set padding-left on the <ul>/<ol> and then rely on the li/p structure.
                        // We should clear some conflicting paragraph styles that might interfere with list indentation.
                        delete effectiveParagraphProps.marginLeft;
                        delete effectiveParagraphProps.marginRight;
                        delete effectiveParagraphProps.textIndent;
                    }
                }
            }
        }
    }

    let pContentHtml = '';
    for (let i = 0; i < paragraphNode.childNodes.length; i++) {
        const childNode = paragraphNode.childNodes[i];
        if (childNode.nodeType === 1) {
            pContentHtml += await processRunNode(childNode, pPr, zip, documentStyles);
        }
    }
    
    // Special handling for tab leaders (e.g., in a Table of Contents)
    if (pContentHtml.includes('class="tab-leader"')) {
        effectiveParagraphProps.display = 'flex';
        effectiveParagraphProps.alignItems = 'baseline';
    }

    return {
        pContentHtml: pContentHtml,
        pInlineStyles: [propsToCssString(effectiveParagraphProps)], // Convert final props object to CSS string
        listProperties: listProperties
    };
}

async function processRunNode(runNode, pPr, zip, documentStyles) { // pPr is needed for tab leaders context
    let runHtml = '';
    let textContent = '';
    let effectiveRunProps = {};

    const rPr = getChildElementByLocalName(runNode, 'rPr'); // Run properties

    if (rPr) {
        // 1. Apply style from w:rStyle (if present)
        const characterStyleId = getChildElementByLocalName(rPr, 'rStyle')?.getAttribute('w:val');
        if (characterStyleId && documentStyles[characterStyleId]) {
            let currentStyle = documentStyles[characterStyleId];
            let styleChainProps = {};
            while (currentStyle) {
                styleChainProps = mergeProperties(currentStyle.runProps, styleChainProps);
                if (currentStyle.basedOn && documentStyles[currentStyle.basedOn]) {
                    currentStyle = documentStyles[currentStyle.basedOn];
                } else {
                    currentStyle = null;
                }
            }
            effectiveRunProps = mergeProperties(effectiveRunProps, styleChainProps);
        }

        // 2. Apply direct run properties (w:rPr) - these override style properties
        if (getChildElementByLocalName(rPr, 'b')) effectiveRunProps.fontWeight = 'bold';
        if (getChildElementByLocalName(rPr, 'i')) effectiveRunProps.fontStyle = 'italic';
        
        // Handle underline types
        const uNode = getChildElementByLocalName(rPr, 'u');
        if (uNode) {
            const uVal = uNode.getAttribute('w:val');
            if (uVal === 'single' || uVal === 'dash' || uVal === 'dot' || uVal === 'dashDotDotHeavy') { // Basic types
                effectiveRunProps.textDecoration = 'underline';
                if (uVal === 'dash') effectiveRunProps.textDecorationStyle = 'dashed';
                else if (uVal === 'dot') effectiveRunProps.textDecorationStyle = 'dotted';
                else if (uVal === 'double') effectiveRunProps.textDecorationStyle = 'double';
                // For 'dashDotDotHeavy' or more complex ones, CSS has limitations.
            } else if (uVal && uVal !== 'none') { // Catch other types, default to solid
                effectiveRunProps.textDecoration = 'underline';
                effectiveRunProps.textDecorationStyle = 'solid';
            }
        }

        const sz = getChildElementByLocalName(rPr, 'sz'); // Font size in half-points
        if (sz) effectiveRunProps.fontSize = halfPointsToPx(sz.getAttribute('w:val'));
        const color = getChildElementByLocalName(rPr, 'color');
        if (color && color.getAttribute('w:val') !== 'auto') effectiveRunProps.color = `#${color.getAttribute('w:val')}`;

        const rFonts = getChildElementByLocalName(rPr, 'rFonts'); // Font family
        if (rFonts) {
            const asciiFont = rFonts.getAttribute('w:ascii') || rFonts.getAttribute('w:hAnsi') || rFonts.getAttribute('w:cs') || rFonts.getAttribute('w:eastAsia');
            if (asciiFont) effectiveRunProps.fontFamily = `'${asciiFont}'`;
        }
        // Add more rPr properties like caps, smallCaps, strike, etc. if needed
    }

    for (let k = 0; k < runNode.childNodes.length; k++) {
        const childNode = runNode.childNodes[k];
        if (childNode.nodeType === 3) { // Text node
            textContent += childNode.nodeValue;
        } else if (childNode.nodeType === 1) { // Element node
            if (childNode.localName === 't') { // Actual text content
                textContent += childNode.textContent;
            } else if (childNode.localName === 'br') { // Line break
                runHtml += '<br/>';
            } else if (childNode.localName === 'tab') { // Tab character
                let hasDotLeader = false;
                if (pPr) { // Need paragraph properties to check for tab stops with leaders
                    const tabs = getChildElementByLocalName(pPr, 'tabs');
                    if (tabs) {
                        const tabStops = tabs.getElementsByTagName('w:tab');
                        for (let i = 0; i < tabStops.length; i++) {
                            // This logic is simplified; real tab processing is complex.
                            // Assuming any dot leader tab stop implies the current tab is a leader.
                            if (tabStops[i].getAttribute('w:leader') === 'dot') {
                                hasDotLeader = true;
                                break;
                            }
                        }
                    }
                }

                if (hasDotLeader) {
                    runHtml += `<span class="tab-leader" style="flex-grow: 1; border-bottom: 1px dotted black; margin: 0 5px;"></span>`;
                } else {
                    textContent += '\u00A0\u00A0\u00A0\u00A0'; // Represent a tab with multiple non-breaking spaces
                }
            } else if (childNode.localName === 'drawing' || childNode.localName === 'pict') {
                runHtml += await processImageNode(childNode, zip); // Delegate image processing
            } else if (childNode.localName === 'fldChar') { // Field character (e.g., for page numbers, TOC)
                const fldCharType = childNode.getAttribute('w:fldCharType');
                // Basic handling for TOC. Real TOC conversion is much more complex.
                if (fldCharType === 'begin' && childNode.nextSibling && childNode.nextSibling.localName === 'r' && getChildElementByLocalName(childNode.nextSibling, 'fldCode')) {
                    const fldCode = getChildElementByLocalName(childNode.nextSibling, 'fldCode').textContent;
                    if (fldCode.includes('TOC')) {
                        // This is a placeholder for TOC field. You might render a simple "Table of Contents" here,
                        // or generate actual TOC if you parse all headings.
                        runHtml += '<span class="toc-placeholder" style="font-weight: bold;">[Mục lục]</span>';
                    }
                } else if (fldCharType === 'end') {
                    // End of field
                }
            }
        }
    }

    if (textContent) {
        // Escape HTML entities to prevent XSS and display literal characters
        runHtml += `<span style="${propsToCssString(effectiveRunProps)}">${textContent.replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;')}</span>`;
    }
    return runHtml;
}

const EMU_PER_INCH = 914400; // 1 inch = 914400 EMUs
const EMU_TO_PX = PIXELS_PER_INCH / EMU_PER_INCH; // 96px / 914400 EMU

async function processImageNode(imageNode, zip) {
    let imageHtml = '';
    const blip = getChildElementByLocalName(imageNode, 'blip'); // For DrawingML
    const imageData = getChildElementByLocalName(imageNode, 'imageData'); // For VML (pict)

    let imageRelId = null;
    let width = null;
    let height = null;

    if (blip) { // DrawingML (modern Word)
        imageRelId = blip.getAttribute('r:embed');
        // Get dimensions from DrawingML's extent (cx, cy are in EMUs)
        const extent = getChildElementByLocalName(imageNode, 'extent');
        if (extent) {
            width = parseFloat(extent.getAttribute('cx')) * EMU_TO_PX;
            height = parseFloat(extent.getAttribute('cy')) * EMU_TO_PX;
        }
    } else if (imageData) { // VML (older Word or specific shapes)
        imageRelId = imageData.getAttribute('r:id');
        // Try to get dimensions from VML's shape style attribute
        const shape = getChildElementByLocalName(imageNode, 'shape');
        if (shape) {
            const style = shape.getAttribute('style');
            if (style) {
                const widthMatch = style.match(/width:([\d.]+)pt/);
                const heightMatch = style.match(/height:([\d.]+)pt/);
                if (widthMatch) width = parseFloat(widthMatch[1]) * (PIXELS_PER_INCH / POINTS_PER_INCH); // pt to px
                if (heightMatch) height = parseFloat(heightMatch[1]) * (PIXELS_PER_INCH / POINTS_PER_INCH); // pt to px
            }
        }
    }

    let imageFilePath = null;
    if (imageRelId) {
        const relationshipsXmlFile = zip.file('word/_rels/document.xml.rels');
        if (relationshipsXmlFile) {
            const relationshipsXmlString = await relationshipsXmlFile.async('string');
            const relParser = new DOMParser();
            const relDoc = relParser.parseFromString(relationshipsXmlString, 'text/xml');
            const relationships = relDoc.getElementsByTagName('Relationship');

            for (let j = 0; j < relationships.length; j++) {
                if (relationships[j].getAttribute('Id') === imageRelId) {
                    imageFilePath = relationships[j].getAttribute('Target');
                    break;
                }
            }
        }
    }

    if (imageFilePath) {
        // Normalize path (remove leading slashes if present, handle ../ for media folder)
        let normalizedImagePath = imageFilePath;
        if (normalizedImagePath.startsWith('../')) { // Example: ../media/image1.png
            normalizedImagePath = normalizedImagePath.replace('../', 'media/');
        } else if (normalizedImagePath.startsWith('/')) { // Example: /media/image1.png
            normalizedImagePath = normalizedImagePath.substring(1);
        } else if (!normalizedImagePath.startsWith('media/')) { // Ensure it's in media folder if not explicitly specified
             normalizedImagePath = `media/${normalizedImagePath}`;
        }
        // console.log(`Attempting to read image: word/${normalizedImagePath}`); // Debugging image path

        const imageFile = zip.file(`word/${normalizedImagePath}`);

        if (imageFile) {
            try {
                const imgBuffer = await imageFile.async('base64');
                const mimeType = getImageMimeType(normalizedImagePath); // Determine MIME type
                if (mimeType) {
                    const widthAttr = width ? `width="${width}px"` : '';
                    const heightAttr = height ? `height="${height}px"` : '';
                    imageHtml += `<img src="data:${mimeType};base64,${imgBuffer}" ${widthAttr} ${heightAttr} style="max-width: 100%; height: auto; display: block;"/>`;
                } else {
                    console.warn(`Could not determine MIME type for image: ${normalizedImagePath}`);
                }
            } catch (e) {
                console.error(`Error reading image file ${normalizedImagePath}:`, e);
            }
        } else {
            console.warn(`Image file not found in zip: word/${normalizedImagePath}. Check path and case sensitivity.`);
        }
    } else {
        console.warn('Image relationship ID not found or target path missing.');
    }
    return imageHtml;
}

function getImageMimeType(filename) {
    const ext = filename.split('.').pop().toLowerCase();
    switch (ext) {
        case 'png': return 'image/png';
        case 'jpg':
        case 'jpeg': return 'image/jpeg';
        case 'gif': return 'image/gif';
        case 'bmp': return 'image/bmp';
        case 'tiff': return 'image/tiff';
        case 'wmf': return 'image/x-wmf'; // Windows Metafile, often used in older docs
        case 'emf': return 'image/x-emf'; // Enhanced Metafile
        case 'svg': return 'image/svg+xml';
        default: return null;
    }
}

async function processTableNode(tableNode, numberingDefinitions, zip, documentStyles) {
    let effectiveTableProps = {};
    
    // Process table properties (tblPr) for table-wide borders, alignment, etc.
    const tblPr = getChildElementByLocalName(tableNode, 'tblPr');
    if (tblPr) {
        // Table width
        const tblW = getChildElementByLocalName(tblPr, 'tblW');
        if (tblW) {
            const widthType = tblW.getAttribute('w:type');
            const widthVal = tblW.getAttribute('w:w');
            if (widthType === 'dxa') { // Twentieths of a point (twips)
                effectiveTableProps.width = twipsToPx(widthVal);
            } else if (widthType === 'pct') { // Fiftieths of a percent
                effectiveTableProps.width = `${parseFloat(widthVal) / 50}%`;
            } else if (widthType === 'auto') {
                effectiveTableProps.width = 'auto';
            }
        }

        // Table alignment
        const jc = getChildElementByLocalName(tblPr, 'jc'); // Table alignment
        if (jc) {
            if (jc.getAttribute('w:val') === 'center') {
                effectiveTableProps.marginLeft = 'auto';
                effectiveTableProps.marginRight = 'auto';
            } else if (jc.getAttribute('w:val') === 'right') {
                effectiveTableProps.marginLeft = 'auto';
                effectiveTableProps.marginRight = '0';
            } else if (jc.getAttribute('w:val') === 'left') {
                effectiveTableProps.marginLeft = '0';
                effectiveTableProps.marginRight = 'auto';
            }
        }
        
        effectiveTableProps.borderCollapse = 'collapse'; // Always collapse borders for tables

        // Table borders (outer borders)
        const tblBorders = getChildElementByLocalName(tblPr, 'tblBorders');
        if (tblBorders) {
            ['top', 'bottom', 'left', 'right', 'insideH', 'insideV'].forEach(side => {
                const borderNode = getChildElementByLocalName(tblBorders, side);
                if (borderNode && borderNode.getAttribute('w:val') !== 'nil' && borderNode.getAttribute('w:val') !== 'none') {
                    const size = twipsToPx(borderNode.getAttribute('w:sz') || '4');
                    const color = borderNode.getAttribute('w:color') || '000000';
                    const type = borderNode.getAttribute('w:val') || 'single';
                    let borderStyle = 'solid';
                    if (type === 'dotted') borderStyle = 'dotted';
                    else if (type === 'dashed') borderStyle = 'dashed';
                    else if (type === 'double') borderStyle = 'double';
                    
                    if (side === 'insideH') {
                        effectiveTableProps.borderBottom = effectiveTableProps.borderBottom || `${size}px ${borderStyle} #${color}`; // Apply to rows
                    } else if (side === 'insideV') {
                        effectiveTableProps.borderRight = effectiveTableProps.borderRight || `${size}px ${borderStyle} #${color}`; // Apply to cells
                    } else {
                        effectiveTableProps[`border-${side}`] = `${size}px ${borderStyle} #${color}`;
                    }
                }
            });
        }
    }
    effectiveTableProps.marginBottom = effectiveTableProps.marginBottom || '1em'; // Default margin for table below

    let tableHtml = `<table style="${propsToCssString(effectiveTableProps)}">`;

    const rowNodes = tableNode.getElementsByTagName('w:tr');
    for (let i = 0; i < rowNodes.length; i++) {
        const trPr = getChildElementByLocalName(rowNodes[i], 'trPr');
        let effectiveRowProps = {};
        if (trPr) {
            const trHeight = getChildElementByLocalName(trPr, 'trHeight');
            if (trHeight && (trHeight.getAttribute('w:hRule') === 'exact' || trHeight.getAttribute('w:hRule') === 'atLeast')) {
                effectiveRowProps.height = twipsToPx(trHeight.getAttribute('w:val'));
            }
            // Apply horizontal inside border if defined at table level
            if (effectiveTableProps.borderBottom && !effectiveRowProps.borderBottom) {
                effectiveRowProps.borderBottom = effectiveTableProps.borderBottom;
            }
        }
        tableHtml += `<tr style="${propsToCssString(effectiveRowProps)}">`;

        const cellNodes = rowNodes[i].getElementsByTagName('w:tc');
        for (let j = 0; j < cellNodes.length; j++) {
            const cellNode = cellNodes[j];
            let effectiveCellProps = {};
            const cellAttrs = [];

            const tcPr = getChildElementByLocalName(cellNode, 'tcPr'); // Table Cell Properties
            if (tcPr) {
                // Colspan
                const gridSpan = getChildElementByLocalName(tcPr, 'gridSpan');
                if (gridSpan) cellAttrs.push(`colspan="${gridSpan.getAttribute('w:val')}"`);

                // Rowspan (vMerge)
                const vMerge = getChildElementByLocalName(tcPr, 'vMerge');
                const vMergeVal = vMerge ? vMerge.getAttribute('w:val') : null;
                if (vMerge && vMergeVal !== 'restart' && vMergeVal !== null) {
                    continue; // Skip this cell as it's merged with the cell above
                }
                // (Calculating actual rowspan from 'restart' is complex as it requires looking ahead/back)

                // Shading (background color)
                const shd = getChildElementByLocalName(tcPr, 'shd');
                if (shd && shd.getAttribute('w:fill') !== 'auto') effectiveCellProps.backgroundColor = `#${shd.getAttribute('w:fill')}`;

                // Vertical alignment
                const vAlign = getChildElementByLocalName(tcPr, 'vAlign');
                if (vAlign) effectiveCellProps.verticalAlign = vAlign.getAttribute('w:val');

                // Cell width
                const tcW = getChildElementByLocalName(tcPr, 'tcW');
                if (tcW) {
                    const widthType = tcW.getAttribute('w:type');
                    const widthVal = tcW.getAttribute('w:w');
                    if (widthType === 'dxa') { // Twentieths of a point
                        effectiveCellProps.width = twipsToPx(widthVal);
                    } else if (widthType === 'pct') { // Fiftieths of a percent
                        effectiveCellProps.width = `${parseFloat(widthVal) / 50}%`;
                    } else if (widthType === 'auto') {
                         effectiveCellProps.width = 'auto';
                    }
                }

                // Cell margins/padding (tcMar)
                const tcMar = getChildElementByLocalName(tcPr, 'tcMar');
                if (tcMar) {
                    ['top', 'bottom', 'left', 'right'].forEach(side => {
                        const marNode = getChildElementByLocalName(tcMar, side);
                        if (marNode && marNode.getAttribute('w:w')) {
                            effectiveCellProps[`padding-${side}`] = twipsToPx(marNode.getAttribute('w:w'));
                        }
                    });
                }


                // Cell borders (specific to cell, also accounts for insideH/insideV from table properties)
                const tcBorders = getChildElementByLocalName(tcPr, 'tcBorders');
                if (tcBorders) {
                    ['top', 'bottom', 'left', 'right'].forEach(side => {
                        const borderNode = getChildElementByLocalName(tcBorders, side);
                        if (borderNode && borderNode.getAttribute('w:val') !== 'nil' && borderNode.getAttribute('w:val') !== 'none') {
                            const size = twipsToPx(borderNode.getAttribute('w:sz') || '4');
                            const color = borderNode.getAttribute('w:color') || '000000';
                            const type = borderNode.getAttribute('w:val') || 'single';
                            let borderStyle = 'solid';
                            if (type === 'dotted') borderStyle = 'dotted';
                            else if (type === 'dashed') borderStyle = 'dashed';
                            else if (type === 'double') borderStyle = 'double';
                            effectiveCellProps[`border-${side}`] = `${size}px ${borderStyle} #${color}`;
                        } else {
                            // If border is 'nil' or not present, explicitly remove it if a default was set
                            effectiveCellProps[`border-${side}`] = 'none'; // Or inherit, depending on desired behavior
                        }
                    });
                }
            }
            // Apply vertical inside border if defined at table level
            if (effectiveTableProps.borderRight && !effectiveCellProps.borderRight) {
                effectiveCellProps.borderRight = effectiveTableProps.borderRight;
            }

            // Default cell padding and border if not overridden by tcPr
            effectiveCellProps.padding = effectiveCellProps.padding || '5px';
            effectiveCellProps.border = effectiveCellProps.border || '1px solid #ccc'; // Basic default if no specific borders

            let cellContentHtml = await processDocxNodes(cellNode.childNodes, numberingDefinitions, zip, documentStyles);
            tableHtml += `<td ${cellAttrs.join(' ')} style="${propsToCssString(effectiveCellProps)}">${cellContentHtml}</td>`;
        }
        tableHtml += '</tr>';
    }
    tableHtml += '</table>';
    return tableHtml;
}

function parseNumberingXml(numberingXmlString) {
    const parser = new DOMParser();
    const docXml = parser.parseFromString(numberingXmlString, 'text/xml');
    const numberingDefinitions = {};

    // First pass: Parse abstract numbering definitions (w:abstractNum)
    const abstractNums = docXml.getElementsByTagName('w:abstractNum');
    for (let i = 0; i < abstractNums.length; i++) {
        const abstractNumId = abstractNums[i].getAttribute('w:abstractNumId');
        if (abstractNumId === null) continue;

        const levels = {};
        const lvlNodes = abstractNums[i].getElementsByTagName('w:lvl');
        for (let j = 0; j < lvlNodes.length; j++) {
            const ilvl = lvlNodes[j].getAttribute('w:ilvl'); // Level ID
            if (ilvl === null) continue;

            const numFmtNode = getChildElementByLocalName(lvlNodes[j], 'numFmt');
            const lvlTextNode = getChildElementByLocalName(lvlNodes[j], 'lvlText'); // Bullet/number text (e.g., %1., •)
            const lvlJcNode = getChildElementByLocalName(lvlNodes[j], 'lvlJc'); // Level justification (left, center, right)
            const pPrNode = getChildElementByLocalName(lvlNodes[j], 'pPr'); // Paragraph properties for list item at this level

            let levelIndent = { left: 0, hanging: 0 };
            if (pPrNode) {
                const indNode = getChildElementByLocalName(pPrNode, 'ind');
                if (indNode) {
                    const start = indNode.getAttribute('w:start');
                    const hanging = indNode.getAttribute('w:hanging');
                    if (start) levelIndent.left = twipsToPx(start);
                    if (hanging) levelIndent.hanging = twipsToPx(hanging);
                }
            }

            levels[ilvl] = {
                fmt: numFmtNode ? numFmtNode.getAttribute('w:val') : 'decimal',
                lvlText: lvlTextNode ? lvlTextNode.getAttribute('w:val') : null,
                lvlJc: lvlJcNode ? lvlJcNode.getAttribute('w:val') : 'left',
                indent: levelIndent
            };
        }
        numberingDefinitions[abstractNumId] = { type: 'abstract', levels };
    }

    // Second pass: Parse numbering instances (w:num) that link to abstract definitions
    const nums = docXml.getElementsByTagName('w:num');
    for (let i = 0; i < nums.length; i++) {
        const numId = nums[i].getAttribute('w:numId');
        if (numId === null) continue;

        const abstractNumIdRef = getChildElementByLocalName(nums[i], 'abstractNumId');
        if (abstractNumIdRef) {
            const abstractNumIdVal = abstractNumIdRef.getAttribute('w:val');
            if (abstractNumIdVal !== null) {
                numberingDefinitions[numId] = { type: 'num', abstractNumId: abstractNumIdVal };
            }
        }
    }
    return numberingDefinitions;
}

// Hàm mới để phân tích styles.xml
async function parseStylesXml(stylesXmlString) {
    const parser = new DOMParser();
    const docXml = parser.parseFromString(stylesXmlString, 'text/xml');
    const styles = {}; // Store styles by styleId

    const styleNodes = docXml.getElementsByTagName('w:style');
    for (let i = 0; i < styleNodes.length; i++) {
        const styleId = styleNodes[i].getAttribute('w:styleId');
        const type = styleNodes[i].getAttribute('w:type'); // paragraph, character, table, numbering
        if (!styleId) continue;

        let styleProps = {
            type: type,
            isDefault: styleNodes[i].getAttribute('w:default') === '1',
            basedOn: getChildElementByLocalName(styleNodes[i], 'basedOn')?.getAttribute('w:val'),
            paragraphProps: {},
            runProps: {}
        };

        const pPr = getChildElementByLocalName(styleNodes[i], 'pPr');
        if (pPr) {
            const jc = getChildElementByLocalName(pPr, 'jc');
            if (jc) styleProps.paragraphProps.textAlign = jc.getAttribute('w:val') === 'both' ? 'justify' : jc.getAttribute('w:val');

            const spacing = getChildElementByLocalName(pPr, 'spacing');
            if (spacing) {
                if (spacing.getAttribute('w:before')) styleProps.paragraphProps.marginTop = twipsToPx(spacing.getAttribute('w:before'));
                if (spacing.getAttribute('w:after')) styleProps.paragraphProps.marginBottom = twipsToPx(spacing.getAttribute('w:after'));
                const line = spacing.getAttribute('w:line');
                const lineRule = spacing.getAttribute('w:lineRule');
                if (line) {
                    if (lineRule === 'auto') {
                        styleProps.paragraphProps.lineHeight = parseFloat(line) / 240;
                    } else {
                        styleProps.paragraphProps.lineHeight = twipsToPx(line);
                    }
                }
            }

            const ind = getChildElementByLocalName(pPr, 'ind');
            if (ind) {
                if (ind.getAttribute('w:start')) styleProps.paragraphProps.marginLeft = twipsToPx(ind.getAttribute('w:start'));
                if (ind.getAttribute('w:end')) styleProps.paragraphProps.marginRight = twipsToPx(ind.getAttribute('w:end'));
                if (ind.getAttribute('w:firstLine')) styleProps.paragraphProps.textIndent = twipsToPx(ind.getAttribute('w:firstLine'));
                if (ind.getAttribute('w:hanging')) {
                    const hangingVal = twipsToPx(ind.getAttribute('w:hanging'));
                    styleProps.paragraphProps.textIndent = -hangingVal;
                    if (!styleProps.paragraphProps.marginLeft || styleProps.paragraphProps.marginLeft < hangingVal) {
                        styleProps.paragraphProps.marginLeft = hangingVal;
                    }
                }
            }
            // Add more pPr properties as needed (borders, shading, etc.)
        }

        const rPr = getChildElementByLocalName(styleNodes[i], 'rPr');
        if (rPr) {
            if (getChildElementByLocalName(rPr, 'b')) styleProps.runProps.fontWeight = 'bold';
            if (getChildElementByLocalName(rPr, 'i')) styleProps.runProps.fontStyle = 'italic';
            
            const uNode = getChildElementByLocalName(rPr, 'u');
            if (uNode) {
                const uVal = uNode.getAttribute('w:val');
                if (uVal === 'single' || uVal === 'dash' || uVal === 'dot' || uVal === 'dashDotDotHeavy') {
                    styleProps.runProps.textDecoration = 'underline';
                    if (uVal === 'dash') styleProps.runProps.textDecorationStyle = 'dashed';
                    else if (uVal === 'dot') styleProps.runProps.textDecorationStyle = 'dotted';
                    else if (uVal === 'double') styleProps.runProps.textDecorationStyle = 'double';
                } else if (uVal && uVal !== 'none') {
                    styleProps.runProps.textDecoration = 'underline';
                    styleProps.runProps.textDecorationStyle = 'solid';
                }
            }

            const sz = getChildElementByLocalName(rPr, 'sz');
            if (sz) styleProps.runProps.fontSize = halfPointsToPx(sz.getAttribute('w:val'));
            const color = getChildElementByLocalName(rPr, 'color');
            if (color && color.getAttribute('w:val') !== 'auto') styleProps.runProps.color = `#${color.getAttribute('w:val')}`;
            const rFonts = getChildElementByLocalName(rPr, 'rFonts');
            if (rFonts) {
                const asciiFont = rFonts.getAttribute('w:ascii') || rFonts.getAttribute('w:hAnsi') || rFonts.getAttribute('w:cs') || rFonts.getAttribute('w:eastAsia');
                if (asciiFont) styleProps.runProps.fontFamily = `'${asciiFont}'`;
            }
            // Add more rPr properties as needed (underline type, caps, smallCaps, etc.)
        }
        styles[styleId] = styleProps;
    }
    return styles;
}


// Hàm chính để chuyển đổi DOCX Buffer sang HTML
async function convertDocxBufferToHtml(docxBuffer) {
    console.log(`Bắt đầu chuyển đổi DOCX Buffer sang HTML. Kích thước: ${docxBuffer.length} bytes`);
    const zip = await JSZip.loadAsync(docxBuffer);
    const documentXmlFile = zip.file('word/document.xml');
    if (!documentXmlFile) {
        throw new Error('Không tìm thấy document.xml trong file DOCX.');
    }
    const documentXmlString = await documentXmlFile.async('string');

    const numberingXmlFile = zip.file('word/numbering.xml');
    let numberingDefinitions = {};
    if (numberingXmlFile) {
        const numberingXmlString = await numberingXmlFile.async('string');
        numberingDefinitions = parseNumberingXml(numberingXmlString);
        // console.log('Parsed Numbering Definitions:', JSON.stringify(numberingDefinitions, null, 2)); // Debugging numbering
    } else {
        console.warn('numbering.xml not found. No list formatting will be applied.');
    }

    const stylesXmlFile = zip.file('word/styles.xml');
    let documentStyles = {};
    if (stylesXmlFile) {
        const stylesXmlString = await stylesXmlFile.async('string');
        documentStyles = await parseStylesXml(stylesXmlString);
        // console.log('Parsed Document Styles:', JSON.stringify(documentStyles, null, 2)); // Debugging styles
    } else {
        console.warn('styles.xml not found. Default Word styles will not be applied.');
    }

    const parser = new DOMParser();
    const docXml = parser.parseFromString(documentXmlString, 'text/xml');
    const body = getChildElementByLocalName(docXml.documentElement, 'body');

    if (!body) {
        throw new Error('Không tìm thấy phần tử w:body trong document.xml.');
    }

    // Pass documentStyles to processDocxNodes and its children
    const htmlOutput = await processDocxNodes(body.childNodes, numberingDefinitions, zip, documentStyles);
    console.log('Đã hoàn tất chuyển đổi DOCX sang HTML.');

    // Wrap the content with basic HTML structure and default CSS for print-like behavior
    return `
<!DOCTYPE html>
<html lang="en">
<head>
    <meta charset="UTF-8">
    <meta name="viewport" content="width=device-width, initial-scale=1.0">
    <title>Document Preview</title>
    <style>
        /* Basic CSS Reset/Normalization for better consistency */
        html, body, div, span, applet, object, iframe,
        h1, h2, h3, h4, h5, h6, p, blockquote, pre,
        a, abbr, acronym, address, big, cite, code,
        del, dfn, em, img, ins, kbd, q, s, samp,
        small, strike, strong, sub, sup, tt, var,
        b, u, i, center,
        dl, dt, dd, ol, ul, li,
        fieldset, form, label, legend,
        table, caption, tbody, tfoot, thead, tr, th, td,
        article, aside, canvas, details, embed,
        figure, figcaption, footer, header, hgroup,
        menu, nav, output, ruby, section, summary,
        time, mark, audio, video {
            margin: 0;
            padding: 0;
            border: 0;
            font-size: 100%;
            font: inherit;
            vertical-align: baseline;
            box-sizing: border-box; /* Include padding and border in the element's total width and height */
        }
        /* HTML5 display-role reset for older browsers */
        article, aside, details, figcaption, figure,
        footer, header, hgroup, menu, nav, section {
            display: block;
        }
        body {
            line-height: 1.2; /* A more common default line height */
            font-family: 'Times New Roman', serif; /* Common Word default, with serif fallback */
            margin: 1in; /* Default Word margin, parse from sectPr for precision */
            color: #000; /* Default text color */
            -webkit-print-color-adjust: exact; /* For better print fidelity of colors */
            color-adjust: exact;
        }
        ol, ul {
            list-style: none; /* Reset default list styles for custom handling */
        }
        blockquote, q {
            quotes: none;
        }
        blockquote:before, blockquote:after,
        q:before, q:after {
            content: '';
            content: none;
        }
        table {
            border-collapse: collapse;
            border-spacing: 0;
            width: 100%; /* Default table to 100% width unless specified by DOCX */
        }

        /* --- Custom Styles for DOCX Conversion --- */
        p {
            margin: 0;
            padding: 0;
            min-height: 1em; /* Ensure empty paragraphs take up space */
            text-indent: 0; /* Reset default text-indent */
        }
        /* Span elements within paragraphs might have specific styles */
        p > span {
            display: inline;
        }
        table {
            margin-bottom: 1em;
        }
        th, td {
            padding: 5px; /* Default cell padding if not explicitly set */
            border: 1px solid #ccc; /* Default cell border if not explicitly set */
            vertical-align: top; /* Default vertical align for cells */
        }
        img {
            max-width: 100%;
            height: auto;
            display: block; /* Ensure images don't have extra space below */
            page-break-inside: avoid; /* Try to keep images on one page for print */
        }

        /* List specific styling (basic) */
        ul, ol {
            /* These will be overridden by inline styles from numberingDefinitions */
            margin-left: 0; /* Base margin handled by parsing w:ind */
            padding-left: 0; /* Base padding for list items */
        }
        ul li, ol li {
            list-style-position: outside; /* Typically outside for Word lists */
            margin-left: 0; /* Reset */
            padding-left: 0; /* Reset */
        }
        /* Basic list types - will be overridden by DOCX numbering definitions */
        ul { list-style-type: disc; }
        ol { list-style-type: decimal; }

        /* Style for tab leaders */
        .tab-leader {
            flex-grow: 1;
            border-bottom: 1px dotted black;
            margin: 0 5px;
        }

        /* Print styles */
        @media print {
            body {
                margin: 0.5in; /* Smaller margins for print, or extract from sectPr */
            }
            table {
                break-inside: avoid; /* Try to keep tables on one page if possible */
            }
            img {
                page-break-inside: avoid;
            }
            /* You might need more specific print styles for headers/footers, page breaks etc. */
        }
    </style>
</head>
<body>
    ${htmlOutput}
</body>
</html>
    `;
}

// Export the main function
module.exports = {
    convertDocxBufferToHtml
};