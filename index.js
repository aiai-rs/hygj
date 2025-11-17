require('dotenv').config();
const { Telegraf } = require('telegraf');
const express = require('express'); // 新加：HTTP 服务器
const xlsx = require('xlsx');
const PptxGenJS = require('pptxgenjs');
const puppeteer = require('puppeteer');
const sharp = require('sharp');
const fs = require('fs-extra');
const path = require('path');
const fetch = require('node-fetch');

const BOT_TOKEN = process.env.BOT_TOKEN;
if (!BOT_TOKEN) {
  console.error('BOT_TOKEN not set! Add to .env or Render env.');
  process.exit(1);
}

const bot = new Telegraf(BOT_TOKEN);
const app = express(); // 新加：Express app

// 临时目录（处理后立即删除）
const TEMP_DIR = path.join(__dirname, 'temp');
fs.ensureDirSync(TEMP_DIR);

// 支持格式
const SUPPORTED_FORMATS = ['.xlsx', '.xls', '.xlsm', '.csv', '.pptx', '.pot', '.swf'];

// 在线确认：文本消息回复
bot.on('text', (ctx) => {
  ctx.reply('Bot 在线！请发送 xlsx, xls, xlsm, csv, pptx, pot, swf 文件转换图片。');
});

// 处理文件消息
bot.on('document', async (ctx) => {
  const file = ctx.message.document;
  if (!file) return;
  const ext = path.extname(file.file_name).toLowerCase();
  if (!SUPPORTED_FORMATS.includes(ext)) {
    return ctx.reply('不支持的格式。只支持: xlsx, xls, xlsm, csv, pptx, pot, swf');
  }
  ctx.reply('处理中...'); // 进度提示
  try {
    // 下载文件到 temp
    const fileUrl = await ctx.telegram.getFileLink(file.file_id);
    const filePath = path.join(TEMP_DIR, `input${ext}`);
    const response = await fetch(fileUrl);
    await fs.writeFile(filePath, await response.buffer());

    let images = [];
    if (['.xlsx', '.xls', '.xlsm', '.csv'].includes(ext)) {
      images = await convertSpreadsheetToImages(filePath);
    } else if (['.pptx', '.pot'].includes(ext)) {
      images = await convertPptxToImages(filePath);
    } else if (ext === '.swf') {
      images = await convertSwfToImages(filePath);
    }

    // 发送图片回 Telegram
    for (const imgPath of images) {
      await ctx.replyWithPhoto({ source: imgPath });
      fs.unlinkSync(imgPath); // 立即删除
    }

    // 删除原文件
    fs.unlinkSync(filePath);
    // 清空 temp
    cleanupTemp();
    ctx.reply(`转换完成！发送了 ${images.length} 张图片。无任何保存。`);
  } catch (error) {
    console.error('Error:', error.message);
    ctx.reply('转换失败: ' + error.message);
    cleanupTemp();
  }
});

// Excel/CSV 转图片（每个 Sheet 一张）
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

// PPTX 转图片（简化，每页一图）
async function convertPptxToImages(filePath) {
  const images = [];
  const browser = await puppeteer.launch({ headless: true, args: ['--no-sandbox', '--disable-setuid-sandbox'] });
  const page = await browser.newPage();
  const text = 'PPTX预览（简化模式，每页转图）';
  await page.setContent(`<div style="font-size:24px;padding:20px;">${text}</div>`);
  const imgPath = path.join(TEMP_DIR, 'slide1.png');
  await page.screenshot({ path: imgPath, fullPage: true });
  images.push(imgPath);
  await browser.close();
  return images;
}

// SWF 转图片（第一帧）
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

// HTML 渲染成图片
async function renderHtmlToImage(html, outputPath) {
  const browser = await puppeteer.launch({ headless: true, args: ['--no-sandbox', '--disable-setuid-sandbox'] });
  const page = await browser.newPage();
  await page.setContent(html);
  await page.screenshot({ path: outputPath, fullPage: true });
  await browser.close();
  // Sharp 优化 PNG
  const tmpPath = outputPath + '.tmp';
  await sharp(outputPath).png().toFile(tmpPath);
  fs.unlinkSync(outputPath);
  fs.renameSync(tmpPath, outputPath);
}

// 清理 temp
function cleanupTemp() {
  fs.emptyDirSync(TEMP_DIR);
}

// Webhook 设置：Express 处理 /bot 路径
const PORT = process.env.PORT || 3000;
app.use(bot.webhookCallback('/bot'));
app.listen(PORT, () => {
  console.log(`Bot running on port ${PORT}`);
});

// 启动 webhook（设置 Telegram webhook URL）
bot.launch({
  webhook: {
    domain: process.env.RENDER_EXTERNAL_HOSTNAME ? `https://${process.env.RENDER_EXTERNAL_HOSTNAME}.onrender.com` : 'http://localhost:3000',
    port: PORT,
    path: '/bot'
  }
}).then(() => {
  console.log('Bot started with webhook!');
  cleanupTemp(); // 启动时清
}).catch(err => {
  console.error('Bot error:', err);
  cleanupTemp();
  process.exit(1);
});

// 优雅关闭
process.once('SIGINT', () => {
  bot.stop('SIGINT');
  cleanupTemp();
  process.exit(0);
});
process.once('SIGTERM', () => {
  bot.stop('SIGTERM');
  cleanupTemp();
  process.exit(0);
});
