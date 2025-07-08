const fs = require('fs');
const { google } = require('googleapis');

const SCOPES = ['https://www.googleapis.com/auth/drive.metadata.readonly', 'https://www.googleapis.com/auth/drive.file'];
const TOKEN_PATH = 'token.json';
const CREDENTIALS_PATH = 'credentials.json';

function loadCredentials() {
    try {
        const content = fs.readFileSync(CREDENTIALS_PATH);
        return JSON.parse(content);
    } catch (err) {
        console.error('Lỗi khi tải tệp credentials.json:', err.message);
        console.error('Vui lòng đảm bảo bạn đã tải tệp "credentials.json" từ Google Cloud Console và đặt nó vào thư mục gốc của dự án.');
        process.exit(1);
    }
}

/**
 * Tạo và trả về URL ủy quyền của Google.
 * @param {Object} credentials Thông tin xác thực từ tệp credentials.json.
 * @param {string} redirectUri URI chuyển hướng đã cấu hình trong Google Cloud Console.
 * @returns {string} URL ủy quyền của Google.
 */
function generateAuthUrl(credentials, redirectUri) {
    const { client_secret, client_id } = credentials.installed || credentials.web;
    const oAuth2Client = new google.auth.OAuth2(
        client_id, client_secret, redirectUri
    );
    return oAuth2Client.generateAuthUrl({
        access_type: 'offline',
        scope: SCOPES,
        prompt: 'consent' // Yêu cầu người dùng đồng ý lại để đảm bảo lấy refresh token
    });
}

/**
 * Trao đổi mã xác minh lấy token và lưu trữ.
 * @param {Object} credentials Thông tin xác thực từ tệp credentials.json.
 * @param {string} code Mã xác minh nhận được từ Google.
 * @param {string} redirectUri URI chuyển hướng đã cấu hình trong Google Cloud Console.
 * @returns {Promise<google.auth.OAuth2>} Client OAuth2 được xác thực.
 */
async function exchangeCodeForTokens(credentials, code, redirectUri) {
    const { client_secret, client_id } = credentials.installed || credentials.web;
    const oAuth2Client = new google.auth.OAuth2(
        client_id, client_secret, redirectUri // SỬA LỖI TẠI ĐÂY: Sử dụng client_secret
    );
    try {
        const { tokens } = await oAuth2Client.getToken(code);
        oAuth2Client.setCredentials(tokens);
        fs.writeFileSync(TOKEN_PATH, JSON.stringify(tokens));
        console.log('Token đã được lưu vào', TOKEN_PATH);
        return oAuth2Client;
    } catch (err) {
        console.error('Lỗi khi truy xuất access token:', err.message);
        throw new Error('Không thể trao đổi mã lấy token.');
    }
}

/**
 * Tải client đã xác thực từ token đã lưu.
 * @param {Object} credentials Thông tin xác thực từ tệp credentials.json.
 * @returns {Promise<google.auth.OAuth2|null>} Client OAuth2 được xác thực hoặc null nếu không có token.
 */
async function getStoredAuthClient(credentials) {
    const { client_secret, client_id, redirect_uris } = credentials.installed || credentials.web;
    const oAuth2Client = new google.auth.OAuth2(
        client_id, client_secret, redirect_uris[0]
    );
    if (fs.existsSync(TOKEN_PATH)) {
        try {
            const token = fs.readFileSync(TOKEN_PATH);
            oAuth2Client.setCredentials(JSON.parse(token));
            return oAuth2Client;
        } catch (err) {
            console.warn('Lỗi khi tải token từ tệp token.json:', err.message);
            return null; // Trả về null để báo hiệu cần ủy quyền lại
        }
    }
    return null; // Không có token
}

/**
 * Liệt kê 10 tệp đầu tiên trong Google Drive của người dùng.
 * @param {google.auth.OAuth2} auth Client được xác thực.
 */
async function listDriveFiles(auth) {
    const drive = google.drive({ version: 'v3', auth });
    try {
        const res = await drive.files.list({
            pageSize: 10, // Giới hạn 10 tệp đầu tiên
            fields: 'nextPageToken, files(id, name, mimeType)', // Chỉ lấy ID, tên và loại MIME
        });
        return res.data.files;
    } catch (err) {
        console.error('API Lỗi khi liệt kê tệp:', err.message);
        throw err;
    }
}

/**
 * Tải xuống một tài liệu Google Drive dưới dạng PDF.
 * @param {google.auth.OAuth2} auth Client được xác thực.
 * @param {string} fileId ID của tài liệu trong Google Drive.
 * @returns {Promise<stream.Readable>} Luồng dữ liệu của tệp PDF.
 */
async function downloadFileAsPdf(auth, fileId) {
    const drive = google.drive({ version: 'v3', auth });
    try {
        const response = await drive.files.export({
            fileId: fileId,
            mimeType: 'application/pdf',
        }, {
            responseType: 'stream', // Yêu cầu luồng dữ liệu
        });
        return response.data; // Trả về luồng dữ liệu
    } catch (err) {
        console.error('Lỗi khi tải xuống tệp dưới dạng PDF:', err.message);
        throw err;
    }
}

module.exports = {
    loadCredentials,
    generateAuthUrl,
    exchangeCodeForTokens,
    getStoredAuthClient,
    listDriveFiles,
    downloadFileAsPdf
};
