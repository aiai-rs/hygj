const { Bot } = require("grammy");
const fs = require("fs");
const path = require("path");
const puppeteer = require("puppeteer");

const bot = new Bot(process.env.BOT_TOKEN);

// 关键：告诉 puppeteer 去 Render 真正的缓存目录找 Chrome
process.env.PUPPETEER_CACHE_DIR = "/opt/render/.cache/puppeteer";

let browser;
async function getBrowser() {
    if (!browser) {
        browser = await puppeteer.launch({
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
    }
    return browser;
}

async function toImage(filePath, isTxt) {
    const page = await browser.newPage();

    if (isTxt) {
        const text = fs.readFileSync(filePath, "utf-8")
            .replace(/&/g, "&amp;").replace(/</g, "&lt;").replace(/>/g, "&gt;");
        await page.setContent(`<pre style="font-family:monospace;font-size:16px;margin:40px;">${text}</pre>`);
    } else {
        const ExcelJS = require("exceljs");
        const wb = new ExcelJS.Workbook();
        await wb.xlsx.readFile(filePath);
        const ws = wb.worksheets[0];
        let html = `<table border="1" style="border-collapse:collapse;font-size:15px;font-family:Arial;">`;
        ws.eachRow((row, i) => {
            html += "<tr>";
            row.eachCell({ includeEmpty: true }, cell => {
                const val = cell.value ?? "";
                html += i === 1 
                    ? `<th style="background:#f0f0f0;padding:10px">${val}</th>`
                    : `<td style="padding:8px; text-align:center">${val}</td>`;
            });
            html += "</tr>";
        });
        html += "</table>";
        await page.setContent(html);
    }

    const buffer = await page.screenshot({ fullPage: true });
    await page.close();
    return buffer;
}

bot.command("start", ctx => ctx.reply("发 .xlsx / .xls / .txt 给我，立即转高清图片！"));

bot.on("message:document", async ctx => {
    const doc = ctx.message.document;
    const name = (doc.file_name || "").toLowerCase();

    if (!name.endsWith(".xlsx") && !name.endsWith(".xls") && !name.endsWith(".txt")) {
        return ctx.reply("只支持 xlsx、xls、txt 文件哦～");
    }

    await ctx.reply("正在转换，请稍等 5-20 秒...");

    try {
        await getBrowser();

        const fileUrl = await ctx.getFile().then(f => f.getUrl());
        const localPath = `/tmp/${doc.file_name}`;
        const res = await fetch(fileUrl);
        fs.writeFileSync(localPath, Buffer.from(await res.arrayBuffer()));

        const imgBuffer = await toImage(localPath, name.endsWith(".txt"));

        await ctx.replyWithPhoto(
            { source: imgBuffer },
            { caption: `已转换：${doc.file_name}` }
        );

        // 立即删除，零残留
        fs.unlinkSync(localPath);

    } catch (e) {
        console.error(e);
        await ctx.reply("转换失败：\n" + e.message);
    }
});

// 关键：丢弃旧更新，避免 409 冲突
bot.start({ drop_pending_updates: true });

console.log("Bot 已启动，准备接收文件！");
