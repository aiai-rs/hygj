const { Bot } = require("grammy");
const fs = require("fs");
const puppeteer = require("puppeteer");

const bot = new Bot(process.env.BOT_TOKEN);

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

async function excelToImage(path) {
    const ExcelJS = require("exceljs");
    const wb = new ExcelJS.Workbook();
    await wb.xlsx.readFile(path);
    const ws = wb.worksheets[0];

    let html = `<table border="1" style="border-collapse:collapse;font-size:15px;font-family:Arial;">`;
    ws.eachRow({ includeEmpty: true }, (row, i) => {
        html += "<tr>";
        row.eachCell({ includeEmpty: true }, cell => {
            const val = cell.value ?? "";
            html += i === 1 ? `<th style="background:#f0f0f0;padding:10px">${val}</th>` : `<td style="padding:8px;text-align:center">${val}</td>`;
        });
        html += "</tr>";
    });
    html += "</table>";

    const page = await browser.newPage();
    await page.setContent(html);
    const buffer = await page.screenshot({ fullPage: true });
    await page.close();
    return buffer;
}

async function txtToImage(path) {
    const text = fs.readFileSync(path, "utf-8").replace(/&/g, "&amp;").replace(/</g, "&lt;").replace(/>/g, "&gt;");
    const html = `<pre style="font-family:monospace;font-size:16px;margin:40px;line-height:1.5;">${text}</pre>`;

    const page = await browser.newPage();
    await page.setContent(html);
    const buffer = await page.screenshot({ fullPage: true });
    await page.close();
    return buffer;
}

bot.command("start", ctx => ctx.reply("发 .xlsx 或 .txt 给我，立即转高清图片！"));

bot.on("message:document", async ctx => {
    const file = ctx.message.document;
    const name = (file.file_name || "").toLowerCase();

    if (!name.endsWith(".xlsx") && !name.endsWith(".xls") && !name.endsWith(".txt")) {
        return ctx.reply("只支持 xlsx/xls/txt 文件");
    }

    await ctx.reply("转换中…");

    try {
        await getBrowser();

        const url = await ctx.getFile().then(f => f.getUrl());
        const localPath = `/tmp/${file.file_name}`;
        const res = await fetch(url);
        fs.writeFileSync(localPath, Buffer.from(await res.arrayBuffer()));

        const img = name.endsWith(".txt") ? await txtToImage(localPath) : await excelToImage(localPath);

        await ctx.replyWithPhoto({ source: img }, { caption: `已转换：${file.file_name}` });

        fs.unlinkSync(localPath);
    } catch (e) {
        console.error(e);
        ctx.reply("失败：" + e.message);
    }
});

bot.start({ drop_pending_updates: true });
console.log("Bot 已启动成功！");
