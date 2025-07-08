const express = require('express');
const cors = require('cors');
const fs = require('fs');
const path = require('path');
const { google } = require('googleapis');
const stream = require('stream'); // Import the stream module

// Import các hàm Google Drive API từ index.js
const {
    loadCredentials,
    generateAuthUrl,
    exchangeCodeForTokens,
    getStoredAuthClient,
    listDriveFiles,
    downloadFileAsPdf
} = require('./googlediverAPI.js'); // Đảm bảo đường dẫn đúng

const generateDocx = require('./export-docx'); // Module để tạo DOCX từ template (Docxtemplater)
const extractVariablesRouter = require('./extractVariablesRouter'); // Express Router cho việc trích xuất biến
const scrapeRouter = require('./scrapeRouter'); // Express Router cho việc scraping

// Import hàm chuyển đổi tùy chỉnh của bạn (nếu vẫn cần chuyển đổi DOCX sang HTML)
const { convertDocxBufferToHtml } = require('./index'); // Đảm bảo đường dẫn đúng

const app = express();
app.use(cors()); // Global CORS handling
app.use(express.json());

// Đường dẫn đến file template DOCX (document.docx)
const DOCX_TEMPLATE_PATH = path.resolve(__dirname, 'document.docx');
// Đường dẫn để lưu file DOCX tạm thời sau khi điền dữ liệu
const TEMP_DOCX_OUTPUT_PATH = path.resolve(__dirname, 'output.docx');

// Biến toàn cục để lưu trữ client đã xác thực với Google Drive API
let googleDriveAuthClient = null;

// --- CẤU HÌNH REDIRECT URI ---
const REDIRECT_URI = 'http://localhost:4000/oauth2callback'; // Thay đổi cổng nếu server của bạn chạy trên cổng khác

// --- GOOGLE DRIVE API ENDPOINTS ---

// Endpoint để kích hoạt quá trình xác thực Google Drive API
app.get('/auth-drive', async (req, res) => {
    try {
        const credentials = loadCredentials();
        googleDriveAuthClient = await getStoredAuthClient(credentials);

        if (googleDriveAuthClient) {
            return res.status(200).send('Đã xác thực Google Drive API từ token đã lưu.');
        }

        const authUrl = generateAuthUrl(credentials, REDIRECT_URI);
        console.log('Đang chuyển hướng người dùng đến URL ủy quyền của Google:', authUrl);
        res.redirect(authUrl);
    } catch (error) {
        console.error('Lỗi xác thực Google Drive API:', error.message);
        res.status(500).send('Lỗi trong quá trình xác thực Google Drive API: ' + error.message);
    }
});

// Endpoint để xử lý callback từ Google sau khi ủy quyền
app.get('/oauth2callback', async (req, res) => {
    const code = req.query.code;
    if (!code) {
        return res.status(400).send('Không nhận được mã xác minh từ Google.');
    }

    try {
        const credentials = loadCredentials();
        googleDriveAuthClient = await exchangeCodeForTokens(credentials, code, REDIRECT_URI);
        res.status(200).send('Xác thực Google Drive API thành công! Bạn có thể đóng tab này.');
    } catch (error) {
        console.error('Lỗi khi trao đổi mã xác minh lấy token:', error.message);
        res.status(500).send('Lỗi trong quá trình xác thực: ' + error.message);
    }
});

// Middleware để kiểm tra xem đã xác thực Google Drive API chưa
const ensureDriveAuth = (req, res, next) => {
    if (!googleDriveAuthClient) {
        return res.status(401).send('Chưa xác thực Google Drive API. Vui lòng truy cập /auth-drive để bắt đầu quá trình.');
    }
    next();
};

// Endpoint để liệt kê các tệp trên Google Drive
app.get('/drive/files', ensureDriveAuth, async (req, res) => {
    try {
        const files = await listDriveFiles(googleDriveAuthClient);
        res.json(files);
    } catch (error) {
        console.error('Lỗi khi liệt kê tệp Google Drive:', error.message);
        res.status(500).send('Không thể liệt kê tệp Google Drive: ' + error.message);
    }
});

// Endpoint để tải xuống một tệp từ Google Drive dưới dạng PDF
app.get('/drive/download-pdf/:fileId', ensureDriveAuth, async (req, res) => {
    const fileId = req.params.fileId;
    if (!fileId) {
        return res.status(400).send('Vui lòng cung cấp ID tệp.');
    }

    try {
        const pdfStream = await downloadFileAsPdf(googleDriveAuthClient, fileId);
        res.setHeader('Content-Type', 'application/pdf');
        res.setHeader('Content-Disposition', `attachment; filename="${fileId}.pdf"`);
        pdfStream.pipe(res);
    } catch (error) {
        console.error(`Lỗi khi tải xuống tệp ${fileId} dưới dạng PDF:`, error.message);
        res.status(500).send('Không thể tải xuống tệp PDF: ' + error.message);
    }
});

// Endpoint để tải lên một file DOCX lên Google Drive
app.post('/drive/upload-docx', ensureDriveAuth, async (req, res) => {
    const { fileName, fileContentBase64 } = req.body;

    if (!fileName || !fileContentBase64) {
        return res.status(400).send('Vui lòng cung cấp tên tệp và nội dung tệp (base64).');
    }

    try {
        const fileBuffer = Buffer.from(fileContentBase64, 'base64');

        // CHUYỂN BUFFER THÀNH STREAM ĐỂ TẢI LÊN
        const bufferStream = new stream.PassThrough();
        bufferStream.end(fileBuffer); // Đẩy buffer vào stream

        const drive = google.drive({ version: 'v3', auth: googleDriveAuthClient });

        const fileMetadata = {
            name: fileName,
            mimeType: 'application/vnd.openxmlformats-officedocument.wordprocessingml.document',
            // parents: ['FOLDER_ID_HERE'] // Tùy chọn: ID của thư mục muốn tải lên
        };

        const media = {
            mimeType: 'application/vnd.openxmlformats-officedocument.wordprocessingml.document',
            body: bufferStream, // Sử dụng stream thay vì buffer trực tiếp
        };

        const uploadedFile = await drive.files.create({
            resource: fileMetadata,
            media: media,
            fields: 'id,webViewLink',
        });

        console.log('Tệp DOCX đã được tải lên Google Drive:', uploadedFile.data.name, uploadedFile.data.id);
        res.json({
            success: true,
            fileId: uploadedFile.data.id,
            webViewLink: uploadedFile.data.webViewLink,
            message: 'Tệp DOCX đã được tải lên Google Drive thành công.'
        });

    } catch (error) {
        console.error('Lỗi khi tải lên tệp DOCX lên Google Drive:', error.message);
        res.status(500).send('Không thể tải lên tệp DOCX lên Google Drive: ' + error.message);
    }
});


// --- CÁC API HIỆN CÓ CỦA BẠN ---

app.post('/generate-and-print-docx', async (req, res) => {
    const data = req.body;

    if (!fs.existsSync(DOCX_TEMPLATE_PATH)) {
        console.error(`Lỗi: Không tìm thấy file template DOCX tại ${DOCX_TEMPLATE_PATH}`);
        return res.status(500).json({ error: `File template 'document.docx' không tìm thấy. Vui lòng đặt nó vào thư mục gốc của backend.` });
    }

    let generatedDocxBuffer = null;

    try {
        generateDocx(data);

        if (!fs.existsSync(TEMP_DOCX_OUTPUT_PATH)) {
            throw new Error('File output.docx đã không được tạo bởi hàm generateDocx.');
        }

        generatedDocxBuffer = fs.readFileSync(TEMP_DOCX_OUTPUT_PATH);
        console.log('Đã đọc lại output.docx để chuyển đổi.');

        const htmlContent = await convertDocxBufferToHtml(generatedDocxBuffer);

        res.json({ success: true, htmlContent: htmlContent });

    } catch (err) {
        console.error('Lỗi trong quá trình tạo/chuyển đổi DOCX:', err.message);
        if (err.stack) {
            console.error(err.stack);
        }
        res.status(500).json({
            error: 'Đã xảy ra lỗi khi tạo hoặc chuyển đổi tài liệu.',
            details: err.message
        });
    } finally {
        if (fs.existsSync(TEMP_DOCX_OUTPUT_PATH)) {
            fs.unlink(TEMP_DOCX_OUTPUT_PATH, (unlinkErr) => {
                if (unlinkErr) console.error("Lỗi khi xóa output.docx tạm thời:", unlinkErr);
            });
        }
    }
});

// API mới: Tạo DOCX, tải lên Google Drive, và trả về URL để in qua Google Docs
app.post('/generate-upload-and-print-via-drive', ensureDriveAuth, async (req, res) => {
    const data = req.body;

    if (!fs.existsSync(DOCX_TEMPLATE_PATH)) {
        console.error(`Lỗi: Không tìm thấy file template DOCX tại ${DOCX_TEMPLATE_PATH}`);
        return res.status(500).json({ error: `File template 'document.docx' không tìm thấy. Vui lòng đặt nó vào thư mục gốc của backend.` });
    }

    let generatedDocxBuffer = null;

    try {
        generateDocx(data);

        if (!fs.existsSync(TEMP_DOCX_OUTPUT_PATH)) {
            throw new Error('File output.docx đã không được tạo bởi hàm generateDocx.');
        }
        generatedDocxBuffer = fs.readFileSync(TEMP_DOCX_OUTPUT_PATH);
        console.log('Đã tạo và đọc lại output.docx.');

        // CHUYỂN BUFFER THÀNH STREAM ĐỂ TẢI LÊN
        const bufferStream = new stream.PassThrough();
        bufferStream.end(generatedDocxBuffer); // Đẩy buffer vào stream

        const drive = google.drive({ version: 'v3', auth: googleDriveAuthClient });

        const fileMetadata = {
            name: `Generated_Document_${Date.now()}.docx`, // Tên file duy nhất
            mimeType: 'application/vnd.openxmlformats-officedocument.wordprocessingml.document',
            // parents: ['YOUR_GOOGLE_DRIVE_FOLDER_ID'] // Tùy chọn: ID của thư mục muốn tải lên
        };

        const media = {
            mimeType: 'application/vnd.openxmlformats-officedocument.wordprocessingml.document',
            body: bufferStream, // Sử dụng stream thay vì buffer trực tiếp
        };

        const uploadedFile = await drive.files.create({
            resource: fileMetadata,
            media: media,
            fields: 'id,webViewLink',
        });

        console.log('Tệp DOCX đã được tải lên Google Drive:', uploadedFile.data.name, uploadedFile.data.id);

        res.json({
            success: true,
            message: 'Tệp đã được tải lên Google Drive. Chuyển hướng người dùng đến link này để in.',
            googleDocsPrintUrl: uploadedFile.data.webViewLink
        });

    } catch (err) {
        console.error('Lỗi trong quá trình tạo/tải lên DOCX lên Google Drive:', err.message);
        if (err.stack) {
            console.error(err.stack);
        }
        res.status(500).json({
            error: 'Đã xảy ra lỗi khi tạo tài liệu hoặc tải lên Google Drive.',
            details: err.message
        });
    } finally {
        if (fs.existsSync(TEMP_DOCX_OUTPUT_PATH)) {
            fs.unlink(TEMP_DOCX_OUTPUT_PATH, (unlinkErr) => {
                if (unlinkErr) console.error("Lỗi khi xóa output.docx tạm thời:", unlinkErr);
            });
        }
    }
});


// API phân tích biến trong file .docx (dùng extractVariablesRouter)
app.use('/', extractVariablesRouter);

// Route cho scraping web
app.use('/', scrapeRouter);

// Khởi động server
const PORT = 4000; // Cổng của backend
app.listen(PORT, () => {
    console.log(`✅ Server đang chạy tại http://localhost:${PORT}`);
    console.log(`💡 Để xác thực Google Drive API, truy cập http://localhost:${PORT}/auth-drive lần đầu tiên và làm theo hướng dẫn.`);
    console.log(`🌐 Các API có sẵn:`);
    console.log(`   - GET /auth-drive: Bắt đầu quá trình xác thực Google Drive API (chuyển hướng).`);
    console.log(`   - GET /oauth2callback: Endpoint callback cho OAuth 2.0 (không truy cập trực tiếp).`);
    console.log(`   - GET /drive/files: Liệt kê các tệp trên Drive.`);
    console.log(`   - POST /drive/upload-docx: Tải lên một file DOCX (cần base64).`);
    console.log(`   - GET /drive/download-pdf/:fileId: Tải xuống file từ Drive dưới dạng PDF.`);
    console.log(`   - POST /generate-and-print-docx: Tạo DOCX, chuyển HTML (cũ).`);
    console.log(`   - POST /generate-upload-and-print-via-drive: Tạo DOCX, tải lên Drive, trả về URL Docs để in.`);
});
