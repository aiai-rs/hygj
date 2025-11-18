const TelegramBot = require('grammatical-telegram-bot'); // 轻量快捷
const fs = require('fs');
const path = require('path');
const { execSync } = require('child_process');

// 这两个库是核心（用 puppeteer 无头浏览器渲染表格和文本）
const puppeteer = require('puppeteer');
const exceljs = require('exceljs'); // 读取 xlsx
const { fileURLToPath } = require('url');

const TOKEN = process.env.BOT_TOKEN || '你的TOKEN放这里';

const bot = new TelegramBot(TOKEN, { polling: true });

let browser; // 全局复用浏览器实例，省内存

async function launchBrowser() {
    if (!browser) {
        browser = await puppeteer.launch({
            headless: true,
            args: ['--no-sandbox', '--disable-setuid-sandbox', '--disable-gpu', '--disable-dev-shm-usage']
        });
    }
    return browser;
}

// ========= Excel → 图片 =========
async function excelToImage(filePath) {
    const workbook = new exceljs.Workbook();
    await workbook.xlsx.readFile(filePath);
    const worksheet = workbook.worksheets[0]; // 只取第一个 sheet

    let html = `<table border="1" style="border-collapse:collapse;font-family:Arial,sans-serif;font-size:14px;">`;
    worksheet.eachRow({ includeEmpty: true }, (row, rowNumber) => {
        html += '<tr>';
        row.eachCell({ includeEmpty: true }, (cell) => {
            html += `<td style="padding:8px;min-width:80px;text-align:center;">${cell.value ?? ''}</td>`;
        });
        html += '</tr>';
    });
    html += '</table>';

    const page = await browser.newPage();
    await page.setContent(html, { waitUntil: 'networkidle0' });
    await page.setViewport({ width: 1920, height: 1080 });

    const imgBuffer = await page.screenshot({ fullPage: true });
    await page.close();
    const outPath = filePath.replace(/\.(xlsx|xls)$/i, '.png');
    fs.writeFileSync(outPath, imgBuffer);
    return outPath;
}

// ========= TXT → 图片 =========
async function txtToImage(filePath) {
    const content = fs.readFileSync(filePath, 'utf-8');
    const escaped = content.replace(/`/g, '\\`');

    const html = `
    <!DOCTYPE html>
    <html><body>
    <pre style="font-family: 'Courier New', monospace; font-size: 16px; line-height: 1.4; margin:20px;">
    ${escaped}
    </pre>
    </body></html>`;

    const page = await browser.newPage();
    await page.setContent(html);
    const pre = await page.$('pre');
    const bounding = await pre.boundingBox();

    const imgBuffer = await page.screenshot({
        clip: {
            x: bounding.x - 20,
            y: bounding.y - 20,
            width: bounding.width + 40,
            height: bounding.height + 40
        },
        omitBackground: true
    });
    await page.close();

    const outPath = filePath.replace(/\.txt$/i, '.png');
    fs.writeFileSync(outPath, imgBuffer);
    return outPath;
}

// ========= Bot 逻辑 =========
bot.on('document', async (msg) => {
    const fileName = msg.document.file_name.toLowerCase();
    if (!fileName.endsWith('.xlsx') && !fileName.endsWith('.xls') && !fileName.endsWith('.txt')) {
        return bot.sendMessage(msg.chat.id, '请上传 .xlsx 或 .txt 文件哦～');
    }

    await bot.sendMessage(msg.chat.id, '收到！正在转换，请稍等 5-15 秒...');

    try {
        const file = await bot.getFile(msg.document.file_id);
        const fileUrl = `https://api.telegram.org/file/bot${TOKEN}/${file.file_path}`;
        const localPath = path.join('/tmp', msg.document.file_name);
        const response = await fetch(fileUrl);
        await require('stream').pipeline(response.body, fs.createWriteStream(localPath));

        let imgPath;
        if (fileName.endsWith('.xlsx') || fileName.endsWith('.xls')) {
            imgPath = await excelToImage(localPath);
        } else if (fileName.endsWith('.txt')) {
            imgPath = await txtToImage(localPath);
        }

        await bot.sendPhoto(msg.chat.id, fs.readFileSync(imgPath), {
            caption: `已转换：${msg.document.file_name}`
        });

        // 清理
        fs.unlinkSync(localPath);
        fs.unlinkSync(imgPath);
    } catch (err) {
        console.error(err);
        bot.sendMessage(msg.chat.id, '转换失败了：' + err.message);
    }
});

bot.onText(/\/start/, (msg) => {
    bot.sendMessage(msg.chat.id, '把 .xlsx 或 .txt 文件发给我，我会帮你转成高清图片～\n支持超大表格和长文本！');
});

(async () => {
    await launchBrowser();
    console.log('Node.js 版文件转图片 Bot 已启动');
})();
