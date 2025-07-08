// scrapeRouter.js
const express = require('express');
const axios = require('axios');
const cheerio = require('cheerio');

const router = express.Router(); // Tạo một Express Router

// Endpoint API để thực hiện scraping
// Đường dẫn sẽ là /api/scrape-element khi được dùng trong server.js với app.use('/api', scrapeRouter);
router.get('/scrape-element', async (req, res) => {
    const targetUrl = req.query.url;
    const selector = req.query.selector;

    if (!targetUrl || !selector) {
        return res.status(400).json({ error: 'Thiếu tham số "url" hoặc "selector" trong query.' });
    }

    try {
        const { data: htmlContent } = await axios.get(targetUrl);
        const $ = cheerio.load(htmlContent);
        const elementText = $(selector).text().trim();

        if (elementText) {
            res.json({ 
                success: true, 
                data: elementText, 
                sourceUrl: targetUrl, 
                usedSelector: selector 
            });
        } else {
            res.status(404).json({ 
                success: false, 
                message: `Không tìm thấy phần tử nào với selector: "${selector}" trên trang: ${targetUrl}` 
            });
        }

    } catch (error) {
        console.error('Lỗi khi scraping:', error.message);
        res.status(500).json({ 
            success: false, 
            error: 'Đã xảy ra lỗi khi lấy dữ liệu từ trang web mục tiêu.', 
            details: error.message 
        });
    }
});

module.exports = router; // Xuất router này để server.js có thể sử dụng