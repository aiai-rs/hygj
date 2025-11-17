const { Telegraf } = require('telegraf');
const xlsx = require('xlsx');
const PptxGenJS = require('pptxgenjs');
const puppeteer = require('puppeteer');
const sharp = require('sharp');
const fs = require('fs-extra');
const path = require('path');
const fetch = require('node-fetch');
const express = require('express'); // 新增

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

// 处理文件消息（保持原样）
bot.on('document', async (ctx) => {
  const file = ctx.message.document;
  if (!file) return;

  const ext = path.extname(file.file_name).toLowerCase();
  if (!SUPPORTED_FORMATS.includes(ext)) {
    return ctx.reply('不支持的格式。只支持: xlsx, xls, xlsm, csv, pptx, pot, swf');
  }

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

// Excel/CSV 转换为图片（保持原样，包括 generateTableHtml）
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
  const htmlTable = xlsx.utils.sheet_to_html(ws); // 优化：用内置函数，更准确
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

// PPTX转换为图片（保持）
async function convertPptxToImages(filePath) {
  const images = [];
  const browser = await puppeteer.launch({ headless: true });
  const page = await browser.newPage();
  const text = 'PPTX预览（简化模式）';
  await page.setContent(`<div style="font-size:24px;padding:20px;">${text}</div>`);
  const imgPath = path.join(TEMP_DIR, 'slide1.png');
  await page.screenshot({ path: imgPath, fullPage: true });
  images.push(imgPath);
  await browser.close();
  return images;
}

// SWF转换为图片（保持）
async function convertSwfToImages(filePath) {
  const images = [];
  const browser = await puppeteer.launch({ headless: true });
  const page = await browser.newPage();
  const html = `<html><body><embed src="file://${filePath}" type="application/x-shockwave-flash" width="800" height="600"></body></html>`;
  await page.setContent(html, { waitUntil: 'networkidle0' });
  const imgPath = path.join(TEMP_DIR, 'swf_frame.png');
  await page.screenshot({ path: imgPath });
  images.push(imgPath);
  await browser.close();
  return images;
}

// HTML渲染为图片（保持）
async function renderHtmlToImage(html, outputPath) {
  const browser = await puppeteer.launch({ headless: true });
  const page = await browser.newPage();
  await page.setContent(html);
  await page.screenshot({ path: outputPath, fullPage: true });
  await browser.close();
  await sharp(outputPath).png().toFile(outputPath);
}

// 清理临时文件（保持）
function cleanupTemp() {
  fs.emptyDirSync(TEMP_DIR);
}

// 启动Bot（保持）
bot.launch().then(() => {
  console.log('Bot started!');
  cleanupTemp();
}).catch(err => {
  console.error(err);
  cleanupTemp();
  process.exit(1);
});

// 新增：Express 服务器监听端口（绕过 Render 端口检查）
const app = express();
const PORT = process.env.PORT || 10000;
app.get('/', (req, res) => res.send('Bot is running!'));
app.listen(PORT, () => {
  console.log(`Web server listening on port ${PORT}`);
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
