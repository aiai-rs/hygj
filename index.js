require('dotenv').config();
const { Telegraf } = require('telegraf');
const express = require('express');
const xlsx = require('xlsx');
const PptxGenJS = require('pptxgenjs');
const puppeteer = require('puppeteer-core');
const chromium = require('@sparticuz/chromium');
const sharp = require('sharp');
const path = require('path');
const fetch = require('node-fetch');

const BOT_TOKEN = process.env.BOT_TOKEN;
if (!BOT_TOKEN) {
  console.error('BOT_TOKEN not set! Add to .env or Render env.');
  process.exit(1);
}

const bot = new Telegraf(BOT_TOKEN);
const app = express();

// 支持格式
const SUPPORTED_FORMATS = ['.xlsx', '.xls', '.xlsm', '.csv', '.pptx', '.pot', '.swf'];

// 在线确认
bot.on('text', (ctx) => {
  ctx.reply('Bot 在线！请发送 xlsx, xls, xlsm, csv, pptx, pot, swf 文件转换图片。');
});

// 处理文件消息（全内存，无目录）
bot.on('document', async (ctx) => {
  const file = ctx.message.document;
  if (!file) return;
  const ext = path.extname(file.file_name).toLowerCase();
  console.log('Debug: Processing file ext:', ext);  // 调试1
  if (!SUPPORTED_FORMATS.includes(ext)) {
    return ctx.reply('不支持的格式。只支持: xlsx, xls, xlsm, csv, pptx, pot, swf');
  }
  ctx.reply('处理中...');
  try {
    // 下载到 Buffer
    const fileUrl = await ctx.telegram.getFileLink(file.file_id);
    const response = await fetch(fileUrl);
    const fileBuffer = Buffer.from(await response.arrayBuffer());
    console.log('Debug: File buffer size:', fileBuffer.length);  // 调试2

    let imageBuffers = [];
    if (['.xlsx', '.xls', '.xlsm', '.csv'].includes(ext)) {
      imageBuffers = await convertSpreadsheetToBuffers(fileBuffer, ext);
    } else if (['.pptx', '.pot'].includes(ext)) {
      imageBuffers = await convertPptxToBuffers(fileBuffer);
    } else if (ext === '.swf') {
      imageBuffers = await convertSwfToBuffers(fileBuffer);
    }

    console.log('Debug: Generated', imageBuffers.length, 'images');  // 调试3

    // 发送 Buffer 图片
    for (const imgBuffer of imageBuffers) {
      await ctx.replyWithPhoto({
        source: imgBuffer,
        filename: `converted${Date.now()}.png`
      });
    }
    ctx.reply(`转换完成！发送了 ${imageBuffers.length} 张图片。无任何保存。`);
  } catch (error) {
    console.error('Error:', error.message);
    ctx.reply('转换失败: ' + error.message);
  }
});

// Excel/CSV 转 Buffer 图片
async function convertSpreadsheetToBuffers(fileBuffer, ext) {
  console.log('Debug: Converting spreadsheet, ext:', ext);  // 调试4
  let workbook = xlsx.read(fileBuffer, { type: 'buffer' });
  const imageBuffers = [];
  for (const sheetName of workbook.SheetNames) {
    const ws = workbook.Sheets[sheetName];
    const html = generateTableHtml(ws);
    const imgBuffer = await renderHtmlToBuffer(html);
    imageBuffers.push(imgBuffer);
  }
  return imageBuffers;
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

// PPTX 转 Buffer
async function convertPptxToBuffers(fileBuffer) {
  console.log('Debug: Converting PPTX (simplified)');  // 调试5
  const pptx = new PptxGenJS();
  const slide = pptx.addSlide();
  slide.addText('PPTX 内容预览（简化：实际需加载 Buffer）', { x: 1, y: 1, color: '363636' });
  const buffer = await pptx.write('nodebuffer', { compression: false });
  return [await sharp(buffer).png().toBuffer()];
}

// SWF 转 Buffer
async function convertSwfToBuffers(fileBuffer) {
  console.log('Debug: Converting SWF (prompt)');  // 调试6
  const browser = await launchPuppeteer();
  const page = await browser.newPage();
  await page.setContent(`<div style="font-size:24px;padding:20px;">SWF 不支持（Flash 已弃用），请用其他格式。</div>`);
  const buffer = await page.screenshot({ encoding: 'binary', fullPage: true });
  await browser.close();
  return [await sharp(Buffer.from(buffer)).png().toBuffer()];
}

// HTML 转 Buffer
async function renderHtmlToBuffer(html) {
  console.log('Debug: Rendering HTML to buffer');  // 调试7
  const browser = await launchPuppeteer();
  const page = await browser.newPage();
  await page.setContent(html, { waitUntil: 'networkidle0' });
  const screenshotBuffer = await page.screenshot({ encoding: 'binary', fullPage: true });
  await browser.close();
  return sharp(screenshotBuffer).png().toBuffer();
}

// Puppeteer launch
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
