const express = require('express');
const { google } = require('googleapis');
const puppeteer = require('puppeteer');
const archiver = require('archiver');
const ExcelJS = require('exceljs');
const sizeOf = require('image-size');
const path = require('path');
const fs = require('fs');
const cors = require('cors');

const app = express();
app.use(cors());
app.use(express.json());
app.use(express.static('public'));

// 캡처 저장 폴더
const CAPTURES_DIR = path.join(__dirname, 'captures');
if (!fs.existsSync(CAPTURES_DIR)) {
    fs.mkdirSync(CAPTURES_DIR, { recursive: true });
}

// OAuth2 클라이언트 생성
async function getAuthClient() {
    let credentials, token;

    // 환경변수 또는 파일에서 읽기
    if (process.env.GOOGLE_CREDENTIALS) {
        // 줄바꿈만 제거 후 파싱
        const credStr = process.env.GOOGLE_CREDENTIALS.replace(/[\r\n]+/g, '').trim();
        const tokenStr = process.env.GOOGLE_TOKEN.replace(/[\r\n]+/g, '').trim();
        credentials = JSON.parse(credStr);
        token = JSON.parse(tokenStr);
    } else {
        // 로컬 파일 사용 (개발용)
        const CREDENTIALS_PATH = path.join(__dirname, '..', 'credentials.json');
        const TOKEN_PATH = path.join(__dirname, '..', 'token.json', 'Google.Apis.Auth.OAuth2.Responses.TokenResponse-user');
        credentials = JSON.parse(fs.readFileSync(CREDENTIALS_PATH));
        token = JSON.parse(fs.readFileSync(TOKEN_PATH));
    }

    const { client_id, client_secret, redirect_uris } = credentials.installed;
    const oauth2Client = new google.auth.OAuth2(client_id, client_secret, redirect_uris[0]);

    oauth2Client.setCredentials({
        access_token: token.access_token,
        refresh_token: token.refresh_token,
        token_type: token.token_type,
        expiry_date: token.Issued ? new Date(token.Issued).getTime() + (token.expires_in * 1000) : token.expiry_date
    });

    return oauth2Client;
}

// 스프레드시트 데이터 가져오기
app.post('/api/fetch-sheet', async (req, res) => {
    try {
        const { spreadsheetId, sheetName, range } = req.body;

        const auth = await getAuthClient();
        const sheets = google.sheets({ version: 'v4', auth });

        const fullRange = sheetName ? `${sheetName}!${range}` : range;

        const response = await sheets.spreadsheets.values.get({
            spreadsheetId,
            range: fullRange,
        });

        const rows = response.data.values || [];

        // 데이터 파싱 (날짜, 이름, 링크, 제목)
        const data = rows.map((row, index) => ({
            index: index + 1,
            date: row[0] || '',
            name: row[1] || '',
            link: row[2] || '',
            title: row[3] || ''
        })).filter(item => item.link && item.link.startsWith('http'));

        res.json({ success: true, data });
    } catch (error) {
        console.error('스프레드시트 조회 오류:', error);
        res.status(500).json({ success: false, error: error.message });
    }
});

// 크롬 스타일 주소창 HTML 생성
function createAddressBarHTML(url) {
    let displayUrl = url;
    try {
        const urlObj = new URL(url);
        displayUrl = urlObj.host + urlObj.pathname + urlObj.search;
    } catch (e) {}

    return `
    <div style="width:100%;height:40px;background:#dee1e6;display:flex;align-items:flex-end;padding:0 8px;font-family:'Segoe UI',Arial,sans-serif;box-sizing:border-box;">
        <div style="display:flex;align-items:center;height:32px;background:white;border-radius:8px 8px 0 0;padding:0 12px;min-width:180px;max-width:220px;">
            <span style="font-size:12px;color:#202124;white-space:nowrap;overflow:hidden;text-overflow:ellipsis;flex:1;">네이버 블로그</span>
            <span style="margin-left:8px;color:#5f6368;font-size:12px;">×</span>
        </div>
    </div>
    <div style="width:100%;height:44px;background:white;display:flex;align-items:center;padding:0 12px;font-family:'Segoe UI',Arial,sans-serif;box-sizing:border-box;border-bottom:1px solid #dadce0;">
        <div style="display:flex;gap:6px;margin-right:10px;">
            <span style="color:#5f6368;font-size:16px;">←</span>
            <span style="color:#c4c4c4;font-size:16px;">→</span>
            <span style="color:#5f6368;font-size:16px;">↻</span>
        </div>
        <div style="flex:1;height:30px;background:#f1f3f4;border-radius:15px;display:flex;align-items:center;padding:0 14px;max-width:700px;">
            <span style="font-size:13px;color:#202124;white-space:nowrap;overflow:hidden;text-overflow:ellipsis;">${displayUrl}</span>
        </div>
        <div style="display:flex;gap:12px;margin-left:10px;color:#5f6368;">
            <span style="font-size:16px;">☆</span>
            <span style="font-size:16px;">⋮</span>
        </div>
    </div>
    `;
}

// 블로그 스크린샷 캡처
app.post('/api/capture', async (req, res) => {
    const { items } = req.body;

    if (!items || items.length === 0) {
        return res.status(400).json({ success: false, error: '캡처할 항목이 없습니다.' });
    }

    // 캡처 세션 ID 생성
    const sessionId = Date.now().toString();
    const sessionDir = path.join(CAPTURES_DIR, sessionId);
    fs.mkdirSync(sessionDir, { recursive: true });

    let browser;
    const results = [];

    try {
        browser = await puppeteer.launch({
            headless: true,
            executablePath: process.env.PUPPETEER_EXECUTABLE_PATH || undefined,
            args: ['--no-sandbox', '--disable-setuid-sandbox', '--window-size=1200,900']
        });

        for (let i = 0; i < items.length; i++) {
            const item = items[i];
            const page = await browser.newPage();
            await page.setViewport({ width: 1200, height: 900 });

            try {
                console.log(`캡처 중 (${i + 1}/${items.length}): ${item.link}`);

                await page.goto(item.link, {
                    waitUntil: 'networkidle2',
                    timeout: 30000
                });

                // 네이버 블로그 제목 추출 및 제목 영역까지의 높이 계산
                let blogTitle = '';
                let titleBottom = 400; // 기본값

                try {
                    // 네이버 블로그 제목 선택자들
                    const titleSelectors = [
                        '.se-title-text',
                        '.se-module-text.se-title-text',
                        '.se_title .se_textView',
                        '.se_title',
                        '.htitle',
                        '.tit_h3',
                        '.pcol1 .itemSubjectBol498'
                    ];

                    for (const selector of titleSelectors) {
                        const titleElement = await page.$(selector);
                        if (titleElement) {
                            const titleData = await page.evaluate(el => {
                                const rect = el.getBoundingClientRect();
                                return {
                                    text: el.textContent,
                                    bottom: rect.bottom
                                };
                            }, titleElement);

                            if (titleData.text && titleData.text.trim()) {
                                blogTitle = titleData.text.trim();
                                titleBottom = Math.min(titleData.bottom + 50, 600); // 제목 아래 50px 여유
                                break;
                            }
                        }
                    }

                    // title 태그에서 시도
                    if (!blogTitle) {
                        blogTitle = await page.title();
                        if (blogTitle.includes(' : 네이버 블로그')) {
                            blogTitle = blogTitle.replace(' : 네이버 블로그', '');
                        }
                    }
                } catch (e) {
                    blogTitle = item.title || '제목 없음';
                }

                // 파일명 생성 (특수문자 제거)
                const safeName = item.name.replace(/[<>:"/\\|?*]/g, '_');
                const safeDate = item.date.replace(/[<>:"/\\|?*]/g, '_');
                const filename = `${safeDate}_${safeName}_${i + 1}.png`;
                const filepath = path.join(sessionDir, filename);

                // 블로그 콘텐츠 스크린샷 (제목까지만)
                const contentScreenshot = await page.screenshot({
                    clip: {
                        x: 0,
                        y: 0,
                        width: 1200,
                        height: titleBottom
                    },
                    encoding: 'base64'
                });

                // 주소창 + 콘텐츠를 합친 이미지 생성
                const combinedPage = await browser.newPage();
                const addressBarHeight = 84;
                const totalHeight = addressBarHeight + titleBottom;

                await combinedPage.setViewport({ width: 1200, height: totalHeight });
                await combinedPage.setContent(`
                    <!DOCTYPE html>
                    <html>
                    <head>
                        <style>
                            * { margin: 0; padding: 0; }
                            body { width: 1200px; height: ${totalHeight}px; }
                        </style>
                    </head>
                    <body>
                        ${createAddressBarHTML(item.link)}
                        <img src="data:image/png;base64,${contentScreenshot}" style="width: 1200px; display: block;">
                    </body>
                    </html>
                `, { waitUntil: 'networkidle0' });

                await combinedPage.screenshot({
                    path: filepath,
                    fullPage: true
                });

                await combinedPage.close();

                results.push({
                    ...item,
                    blogTitle,
                    filename,
                    success: true
                });

            } catch (error) {
                console.error(`캡처 실패: ${item.link}`, error.message);
                results.push({
                    ...item,
                    error: error.message,
                    success: false
                });
            } finally {
                await page.close();
            }
        }

    } catch (error) {
        console.error('브라우저 오류:', error);
        return res.status(500).json({ success: false, error: error.message });
    } finally {
        if (browser) await browser.close();
    }

    res.json({
        success: true,
        sessionId,
        results,
        totalCount: items.length,
        successCount: results.filter(r => r.success).length
    });
});

// 엑셀 다운로드 (스크린샷 포함)
app.post('/api/download-excel/:sessionId', async (req, res) => {
    const { sessionId } = req.params;
    const { results } = req.body;
    const sessionDir = path.join(CAPTURES_DIR, sessionId);

    if (!fs.existsSync(sessionDir)) {
        return res.status(404).json({ success: false, error: '세션을 찾을 수 없습니다.' });
    }

    try {
        const workbook = new ExcelJS.Workbook();
        const worksheet = workbook.addWorksheet('캡처 결과');

        // 헤더 설정
        worksheet.columns = [
            { header: '번호', key: 'index', width: 8 },
            { header: '날짜', key: 'date', width: 12 },
            { header: '이름', key: 'name', width: 15 },
            { header: '링크', key: 'link', width: 50 },
            { header: '제목', key: 'title', width: 50 },
            { header: '스크린샷', key: 'screenshot', width: 55 },
            { header: '상태', key: 'status', width: 10 }
        ];

        // 헤더 스타일
        worksheet.getRow(1).font = { bold: true, color: { argb: 'FFFFFFFF' } };
        worksheet.getRow(1).fill = {
            type: 'pattern',
            pattern: 'solid',
            fgColor: { argb: 'FF4F46E5' }
        };
        worksheet.getRow(1).alignment = { horizontal: 'center', vertical: 'middle' };
        worksheet.getRow(1).height = 25;

        // 데이터 추가
        for (let i = 0; i < results.length; i++) {
            const item = results[i];
            const rowNum = i + 2;

            const row = worksheet.addRow({
                index: item.index,
                date: item.date,
                name: item.name,
                link: item.link,
                title: item.blogTitle || '',
                status: item.success ? '완료' : '실패'
            });

            // 모든 셀 중앙 정렬
            row.eachCell((cell) => {
                cell.alignment = { horizontal: 'center', vertical: 'middle', wrapText: true };
            });

            // 스크린샷 추가 (6번째 열 = 스크린샷 열)
            if (item.success && item.filename) {
                const imagePath = path.join(sessionDir, item.filename);
                if (fs.existsSync(imagePath)) {
                    // 이미지 크기 읽기
                    const dimensions = sizeOf(imagePath);

                    // 원본 비율 유지하며 너비 380px에 맞춤
                    const targetWidth = 380;
                    const aspectRatio = dimensions.height / dimensions.width;
                    const targetHeight = targetWidth * aspectRatio;

                    // 행 높이 설정 (이미지 높이 + 여백)
                    worksheet.getRow(rowNum).height = (targetHeight / 1.33) + 10;

                    const imageId = workbook.addImage({
                        filename: imagePath,
                        extension: 'png'
                    });

                    worksheet.addImage(imageId, {
                        tl: { col: 5, row: rowNum - 1 + 0.05 },
                        ext: { width: targetWidth, height: targetHeight }
                    });
                }
            } else {
                worksheet.getRow(rowNum).height = 30;
            }
        }

        // 엑셀 파일 전송
        res.setHeader('Content-Type', 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet');
        res.setHeader('Content-Disposition', `attachment; filename=captures_${sessionId}.xlsx`);

        await workbook.xlsx.write(res);
        res.end();

    } catch (error) {
        console.error('엑셀 생성 오류:', error);
        res.status(500).json({ success: false, error: error.message });
    }
});

// 개별 이미지 다운로드
app.get('/api/image/:sessionId/:filename', (req, res) => {
    const { sessionId, filename } = req.params;
    const filepath = path.join(CAPTURES_DIR, sessionId, filename);

    if (!fs.existsSync(filepath)) {
        return res.status(404).json({ success: false, error: '이미지를 찾을 수 없습니다.' });
    }

    res.sendFile(filepath);
});

const PORT = process.env.PORT || 3000;
app.listen(PORT, () => {
    console.log(`서버 시작: http://localhost:${PORT}`);
});
