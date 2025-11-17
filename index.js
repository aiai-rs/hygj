const { Telegraf } = require('telegraf');
const xlsx = require('xlsx');
const PptxGenJS = require('pptxgenjs');
const puppeteer = require('puppeteer');
const sharp = require('sharp');
const fs = require('fs-extra');
const path = require('path');
const fetch = require('node-fetch');
const express = require('express');

const BOT_TOKEN = process.env.BOT_TOKEN;
if (!BOT_TOKEN) {
  console.error('BOT_TOKEN not set!');
  process.exit(1);
}

const bot = new Telegraf(BOT_TOKEN);

// 临时目录（处理后立即删除）
const TEMP_DIR = path.join(__dirname, 'temp');
fs.ensureDirSync(TEMP_DIR);

// 支持格式
const SUPPORTED_FORMATS = ['.xlsx', '.xls', '.xlsm', '.csv', '.pptx', '.pot', '.swf'];

// 处理文件消息
bot.on('document', async (ctx) => {
  const file = ctx.message.document;
  if (!file) return;

  const ext = path.extname(file.file_name).toLowerCase();
  if (!SUPPORTED_FORMATS.includes(ext)) {
    return ctx.reply('不支持的格式。只支持: xlsx, xls, xlsm, csv, pptx, pot, swf');
  }

  ctx.reply('处理中...'); // 加进度提示

  try {
    // 下载文件
    const fileUrl = await ctx.telegram.getFileLink(file.file_id);
    const filePath = path.join(TEMP_DIR, `input${ext}`);
    const response = await fetch(fileUrl);
    await fs.writeFile(filePath, response.buffer);

    let images = [];
    if (['.xlsx', '.xls', '.xlsm', '.csv'].includes(ext)) {
      images = await convertSpreadsheetToImages(filePath);
    } else if (['.pptx', '.pot'].includes(ext)) {
      images = await convertPptxToImages(filePath);
    } else if (ext === '.swf') {
      images = await convertSwfToImages(filePath);
    }

    // 发送图片
    for (const imgPath of images) {
      await ctx.replyWithPhoto({ source: imgPath });
      // 立即删除图片
      fs.unlinkSync(imgPath);
    }

    // 删除原文件
    fs.unlinkSync(filePath);

    ctx.reply(`转换完成！发送了 ${images.length} 张图片。`);
  } catch (error) {
    console.error(error);
    ctx.reply('转换失败: ' + error.message);
    // 清理临时文件
    cleanupTemp();
  }
});

// Excel/CSV 转换为图片
async function convertSpreadsheetToImages(filePath) {
  let workbook;
  if (path.extname(filePath) === '.csv') {
    const csvData = fs.readFileSync(filePath, 'utf8');
    workbook = xlsx.utils.book_new();
    const ws = xlsx.utils.aoa_to_sheet([csvData.split('\n').map(row => row.split(','))]);
    xlsx.utils.book_append_sheet(workbook, ws, 'Sheet1');
  } else {
    workbook = xlsx.readFile(filePath);
  }

  const images = [];
  for (const sheetName of workbook.SheetNames) {
    const ws = workbook.Sheets[sheetName];
    const html = generateTableHtml(ws);
    const imgPath = path.join(TEMP_DIR, `${sheetName}.png`);
    await renderHtmlToImage(html, imgPath);
    images.push(imgPath);
  }
  return images;
}

function generateTableHtml(ws) {
  const htmlTable = xlsx.utils.sheet_to_html(ws);
  return `
    <html>
      <head>
        <style>
          body { margin: 0; font-family: Arial; }
          table { border-collapse: collapse; width: 100%; }
          td, th { border: 1px solid #ddd; padding: 8px; text-align: left; }
          th { background-color: #f2f2f2; }
        </style>
      </head>
      <body>${htmlTable}</body>
    </html>`;
}

// PPTX转换为图片
async function convertPptxToImages(filePath) {
  const images = [];
  const browser = await puppeteer.launch({ headless: true, args: ['--no-sandbox', '--disable-setuid-sandbox'] }); // Render 兼容
  const page = await browser.newPage();
  const text = 'PPTX预览（简化模式）';
  await page.setContent(`<div style="font-size:24px;padding:20px;">${text}</div>`);
  const imgPath = path.join(TEMP_DIR, 'slide1.png');
  await page.screenshot({ path: imgPath, fullPage: true });
  images.push(imgPath);
  await browser.close();
  return images;
}

// SWF转换为图片
async function convertSwfToImages(filePath) {
  const images = [];
  const browser = await puppeteer.launch({ headless: true, args: ['--no-sandbox', '--disable-setuid-sandbox'] });
  const page = await browser.newPage();
  const html = `<html><body><embed src="file://${filePath}" type="application/x-shockwave-flash" width="800" height="600"></body></html>`;
  await page.setContent(html, { waitUntil: 'networkidle0' });
  const imgPath = path.join(TEMP_DIR, 'swf_frame.png');
  await page.screenshot({ path: imgPath });
  images.push(imgPath);
  await browser.close();
  return images;
}

// HTML渲染为图片
async function renderHtmlToImage(html, outputPath) {
  const browser = await puppeteer.launch({ headless: true, args: ['--no-sandbox', '--disable-setuid-sandbox'] }); // Render 安全
  const page = await browser.newPage();
  await page.setContent(html);
  await page.screenshot({ path: outputPath, fullPage: true });
  await browser.close();
  await sharp(outputPath).png().toFile(outputPath);
}

// 清理临时文件
function cleanupTemp() {
  fs.emptyDirSync(TEMP_DIR);
}

// Express 服务器
const app = express();
app.use(bot.webhookCallback('/bot')); // Webhook 端点

const PORT = process.env.PORT || 10000;
app.get('/', (req, res) => res.send('Bot is running!'));
app.listen(PORT, async () => {
  console.log(`Web server listening on port ${PORT}`);
  
  // 设置 Webhook（用 Render URL）
  const WEBHOOK_URL = process.env.RENDER_URL || 'https://hygj.onrender.com'; // 你的服务 URL
  const WEBHOOK_PATH = `/bot${BOT_TOKEN}`; // 安全路径
  await bot.telegram.setWebhook(`${WEBHOOK_URL}${WEBHOOK_PATH}`);
  console.log(`Webhook set to ${WEBHOOK_URL}${WEBHOOK_PATH}`);
  
  // 启动 Bot（webhook 模式）
  await bot.launch();
  console.log('Bot started with webhook!');
});

// 优雅关闭
process.once('SIGINT', () => {
  bot.stop('SIGINT');
  process.exit(0);
});
process.once('SIGTERM', () => {
  bot.stop('SIGTERM');
  process.exit(0);
});
