const { Bot } = require("grammy");
const fs = require("fs");
const path = require("path");
const puppeteer = require("puppeteer");

const bot = new Bot(process.env.BOT_TOKEN || "");

let browser;
async function getBrowser() {
    if (!browser) {
        browser = await puppeteer.launch({
            headless: true,
            args: ["--no-sandbox", "--disable-setuid-sandbox"]
        });
    }
    return browser;
}

async function fileToImage(filePath, isTxt) {
    const content = isTxt 
        ? `<pre style="font-family:monospace;font-size:16px;margin:30px;">${fs.readFileSync(filePath, "utf-8")
            .replace(/&/g, "&amp;").replace(/</g, "&lt;").replace(/>/g, "&gt;")}</pre>`
        : await excelToHtml(filePath);

    const page = await browser.newPage();
    await page.setContent(isTxt ? content : `<body style="margin:0">${content}</body>`);
    const buffer = await page.screenshot({ fullPage: true });
    await page.close();
    return buffer;
}

async function excelToHtml(filePath) {
    const ExcelJS = require("exceljs");
    const wb = new ExcelJS.Workbook();
    await wb.xlsx.readFile(filePath);
    const ws = wb.worksheets[0];
    let html = `<table border="1" style="border-collapse:collapse;font-size:15px;">`;
    ws.eachRow((row, i) => {
        html += "<tr>";
        row.eachCell({ includeEmpty: true }, cell => {
            const val = cell.value ?? "";
            html += i === 1 
                ? `<th style="background:#f0f0f0;padding:8px">${val}</th>`
                : `<td style="padding:8px;text-align:center">${val}</td>`;
        });
        html += "</tr>";
    });
    html += "</table>";
    return html;
}

bot.start();
bot.command("start", ctx => ctx.reply("发 xlsx 或 txt 给我，转图片"));

bot.on("message:document", async ctx => {
    const doc = ctx.message.document;
    const name = doc.file_name.toLowerCase();
    if (!name.endsWith(".xlsx") && !name.endsWith(".xls") && !name.endsWith(".txt")) 
        return ctx.reply("只支持 xlsx/xls/txt");

    await ctx.reply("转换中…");

    try {
        await getBrowser();
        const url = await ctx.getFile().then(f => f.getUrl());
        const local = `/tmp/${doc.file_name}`;
        const res = await fetch(url);
        fs.writeFileSync(local, Buffer.from(await res.arrayBuffer()));

        const img = await fileToImage(local, name.endsWith(".txt"));
        await ctx.replyWithPhoto({ source: img }, { caption: doc.file_name });

        fs.unlinkSync(local);
    } catch (e) {
        ctx.reply("失败：" + e.message);
        console.error(e);
    }
});

console.log("Bot 启动成功！");
