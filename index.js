require('dotenv').config();
const { Telegraf } = require('telegraf');
const express = require('express');
const xlsx = require('xlsx');
const PptxGenJS = require('pptxgenjs');
const puppeteer = require('puppeteer-core');
const chromium = require('@sparticuz/chromium');
const sharp = require('sharp');
const fs = require('fs');
const path = require('path');
const fetch = require('node-fetch');

const BOT_TOKEN = process.env.BOT_TOKEN;
if (!BOT_TOKEN) {
  console.error('BOT_TOKEN not set! Add to .env or Render env.');
  process.exit(1);
}

const bot = new Telegraf(BOT_TOKEN);
const app = express();

// 临时目录（修复：用相对路径，避免 __dirname 对象问题）
const TEMP_DIR = './temp';
if (typeof TEMP_DIR !== 'string') {
  console.error('TEMP_DIR is not a string:', TEMP_DIR);
  process.exit(1);
}
// 创建目录（用 Node 内置，递归）
if (!fs.existsSync(TEMP_DIR)) {
  fs.mkdirSync(TEMP_DIR, { recursive: true });
}

// 支持格式
const SUPPORTED_FORMATS = ['.xlsx', '.xls', '.xlsm', '.csv', '.pptx', '.pot', '.swf'];

// 在线确认
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
  ctx.reply('处理中...');
  try {
    // 下载文件
    const fileUrl = await ctx.telegram.getFileLink(file.file_id);
    const filePath = path.join(TEMP_DIR, `input${ext}`);
    const response = await fetch(fileUrl);
    await fs.promises.writeFile(filePath, Buffer.from(await response.arrayBuffer()));

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
      fs.unlinkSync(imgPath);
    }
    fs.unlinkSync(filePath);
    cleanupTemp();
    ctx.reply(`转换完成！发送了 ${images.length} 张图片。无任何保存。`);
  } catch (error) {
    console.error('Error:', error.message);
    ctx.reply('转换失败: ' + error.message);
    cleanupTemp();
  }
});

// Excel/CSV 转图片
async function convertSpreadsheetToImages(filePath) {
  let workbook;
  if (path.extname(filePath) === '.csv') {
    workbook = xlsx.readFile(filePath, { type: 'string' });
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

// PPTX 转图片
async function convertPptxToImages(filePath) {
  const images = [];
  const pptx = new PptxGenJS();
  const slide = pptx.addSlide();
  slide.addText('PPTX 内容预览（简化：实际需加载你的 PPTX 文件）', { x: 1, y: 1, color: '363636' });
  const imgPath = path.join(TEMP_DIR, 'slide1.png');
  await pptx.write('png', { output: imgPath });
  images.push(imgPath);
  return images;
}

// SWF 转图片
async function convertSwfToImages(filePath) {
  const images = [];
  const browser = await launchPuppeteer();
  const page = await browser.newPage();
  await page.setContent(`<div style="font-size:24px;padding:20px;">SWF 不支持（Flash 已弃用），请用其他格式。</div>`);
  const imgPath = path.join(TEMP_DIR, 'swf_frame.png');
  await page.screenshot({ path: imgPath, fullPage: true });
  await browser.close();
  images.push(imgPath);
  return images;
}

// HTML 转图片
async function renderHtmlToImage(html, outputPath) {
  const browser = await launchPuppeteer();
  const page = await browser.newPage();
  await page.setContent(html, { waitUntil: 'networkidle0' });
  await page.screenshot({ path: outputPath, fullPage: true });
  await browser.close();
  // Sharp 优化
  const tmpPath = outputPath + '.tmp';
  await sharp(outputPath).png().toFile(tmpPath);
  fs.unlinkSync(outputPath);
  fs.renameSync(tmpPath, outputPath);
}

// 共享 Puppeteer launch
async function launchPuppeteer() {
  return puppeteer.launch({
    headless: true,
    args: chromium.args,
    defaultViewport: chromium.defaultViewport,
    executablePath: await chromium.executablePath({
      args: chromium.args,
      fallbackToSystem: true,
      emulator: true
    }),
    pipe: true
  });
}

// 清理（修复：用 Node 内置 rmdirSync + rmSync）
function cleanupTemp() {
  if (fs.existsSync(TEMP_DIR)) {
    fs.rmSync(TEMP_DIR, { recursive: true, force: true });
  }
  fs.mkdirSync(TEMP_DIR, { recursive: true });
}

// Webhook 设置
const PORT = process.env.PORT || 3000;
app.use(express.json());
app.use(bot.webhookCallback('/bot'));

app.listen(PORT, async () => {
  console.log(`Bot running on port ${PORT}`);
  const webhookUrl = process.env.RENDER_EXTERNAL_HOSTNAME 
    ? `https://${process.env.RENDER_EXTERNAL_HOSTNAME}/bot` 
    : `http://localhost:${PORT}/bot`;
  try {
    await bot.telegram.setWebhook(webhookUrl);
    console.log(`Webhook set to: ${webhookUrl}`);
  } catch (err) {
    console.error('Failed to set webhook:', err.message);
    process.exit(1);
  }
  cleanupTemp();
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
