const { Bot } = require("grammy");
const fs = require("fs");
const path = require("path");
const puppeteer = require("puppeteer");

const TOKEN = process.env.BOT_TOKEN;
if (!TOKEN) process.exit(1);

const bot = new Bot(TOKEN);
let browser;

async function launchBrowser() {
    if (browser) return browser;
    
    browser = await puppeteer.launch({
        product: "chrome",
        headless: "new",
        args: [
            "--no-sandbox",
            "--disable-setuid-sandbox",
            "--disable-dev-shm-usage",
            "--disable-gpu",
            "--no-zygote",
            "--single-process"
        ]
    });
    return browser;
}

// Excel → 图片
async function excelToImage(filePath) {
    const ExcelJS = require("exceljs");
    const workbook = new ExcelJS.Workbook();
    await workbook.xlsx.readFile(filePath);
    const ws = workbook.getWorksheet(1);

    let rows = [];
    ws.eachRow((row) => {
        rows.push(row.values);
    });

    let html = `<table border="1" style="border-collapse:collapse;font-size:15px;font-family:Arial;">`;
    rows.forEach((row, i) => {
        html += "<tr>";
        row.forEach((cell, j) => {
            if (i === 0) html += `<th style="background:#eee;padding:8px">${cell || ""}</th>`;
            else html += `<td style="padding:8px;text-align:center">${cell || ""}</td>`;
        });
        html += "</tr>";
    });
    html += "</table>";

    const page = await browser.newPage();
    await page.setContent(html);
    const buffer = await page.screenshot({ fullPage: true });
    await page.close();
    
    const outPath = filePath.replace(/\.(xlsx|xls)$/i, ".png");
    fs.writeFileSync(outPath, buffer);
    return outPath;
}

// TXT → 图片
async function txtToImage(filePath) {
    const text = fs.readFileSync(filePath, "utf-8")
        .replace(/&/g, "&amp;")
        .replace(/</g, "&lt;")
        .replace(/>/g, "&gt;");

    const html = `<pre style="font-family:'Courier New',monospace;font-size:16px;line-height:1.5;margin:30px;">${text}</pre>`;

    const page = await browser.newPage();
    await page.setContent(html);
    const buffer = await page.screenshot({ fullPage: true });
    await page.close();

    const outPath = filePath.replace(/\.txt$/i, ".png");
    fs.writeFileSync(outPath, buffer);
    return outPath;
}

bot.command("start", (ctx) => ctx.reply("发 xlsx 或 txt 文件给我，立即转高清图！"));

bot.on("message:document", async (ctx) => {
    const file = ctx.message.document;
    const name = file.file_name?.toLowerCase() || "";

    if (!name.endsWith(".xlsx") && !name.endsWith(".xls") && !name.endsWith(".txt")) {
        return ctx.reply("只支持 .xlsx / .xls / .txt 文件");
    }

    await ctx.reply("正在转换…");

    try {
        await launchBrowser();

        const url = await ctx.getFile().then(f => f.getUrl());
        const localPath = `/tmp/${file.file_name}`;
        const res = await fetch(url);
        fs.writeFileSync(localPath, Buffer.from(await res.arrayBuffer()));

        const imgPath = name.endsWith(".txt") ? await txtToImage(localPath) : await excelToImage(localPath);

        await ctx.replyWithPhoto({ source: fs.readFileSync(imgPath) }, { caption: file.file_name });

        // 立即删除
        fs.unlinkSync(localPath);
        fs.unlinkSync(imgPath);

    } catch (e) {
        console.error(e);
        await ctx.reply("出错：" + e.message);
    }
});

(async () => {
    await launchBrowser();
    console.log("Bot 已启动");
    bot.start();
})();
