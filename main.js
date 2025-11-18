const { Bot } = require("grammy");
const fs = require("fs");
const path = require("path");
const puppeteer = require("puppeteer");

const TOKEN = process.env.BOT_TOKEN;
if (!TOKEN) {
    console.error("错误：请在 Render 环境变量中设置 BOT_TOKEN");
    process.exit(1);
}

const bot = new Bot(TOKEN);
let browser;

async function getBrowser() {
    if (!browser) {
        browser = await puppeteer.launch({
            product: "chrome",
            headless: "new",
            args: [
                "--no-sandbox",
                "--disable-setuid-sandbox",
                "--disable-dev-shm-usage",
                "--disable-gpu",
                "--no-zygote",
                "--single-process",
                "--disable-extensions"
            ]
        });
    }
    return browser;
}

// Excel → 图片
async function excelToImage(filePath) {
    const ExcelJS = require("exceljs");
    const workbook = new ExcelJS.Workbook();
    await workbook.xlsx.readFile(filePath);
    const ws = workbook.worksheets[0];

    let html = `<table border="1" style="border-collapse:collapse;font-family:Arial,sans-serif;font-size:15px;">`;
    ws.eachRow((row, rn) => {
        html += "<tr>";
        row.eachCell({ includeEmpty: true }, (cell) => {
            const val = cell.value ?? "";
            if (rn === 1) html += `<th style="background:#f0f0f0;padding:8px">${val}</th>`;
            else html += `<td style="padding:8px;text-align:center">${val}</td>`;
        });
        html += "</tr>";
    });
    html += "</table>";

    const page = await browser.newPage();
    await page.setContent(html, { waitUntil: "networkidle0" });
    const imgBuffer = await page.screenshot({ fullPage: true });
    await page.close();

    const out = filePath.replace(/\.(xlsx|xls)$/i, ".png");
    fs.writeFileSync(out, imgBuffer);
    return out;
}

// TXT → 图片
async function txtToImage(filePath) {
    const text = fs.readFileSync(filePath, "utf-8")
        .replace(/&/g, "&amp;").replace(/</g, "&lt;").replace(/>/g, "&gt;");

    const html = `
    <!DOCTYPE html><html><body style="margin:30px;background:white;">
    <pre style="font-family:'Courier New',monospace;font-size:16px;line-height:1.5;">${text}</pre>
    </body></html>`;

    const page = await browser.newPage();
    await page.setContent(html);
    const pre = await page.$("pre");
    const box = await pre.boundingBox();

    const imgBuffer = await page.screenshot({
        clip: {
            x: box.x - 30,
            y: box.y - 30,
            width: box.width + 60,
            height: box.height + 60
        }
    });
    await page.close();

    const out = filePath.replace(/\.txt$/i, ".png");
    fs.writeFileSync(out, imgBuffer);
    return out;
}

// Bot 主逻辑
bot.command("start", (ctx) => ctx.reply(
    "发 .xlsx、.xls 或 .txt 文件给我，我马上转成高清图片发回给你！\n支持超大表格和长文本～"
));

bot.on("message:document", async (ctx) => {
    const file = ctx.message.document;
    const name = file.file_name.toLowerCase();

    if (!name.endsWith(".xlsx") && !name.endsWith(".xls") && !name.endsWith(".txt")) {
        return ctx.reply("只支持 .xlsx、.xls 和 .txt 文件哦～");
    }

    await ctx.reply("正在转换，请稍等 5-20 秒...");

    try {
        await getBrowser(); // 确保浏览器已启动

        const url = await ctx.getFile().then(f => f.getUrl());
        const localPath = path.join("/tmp", file.file_name);

        // 下载文件
        const res = await fetch(url);
        const buffer = Buffer.from(await res.arrayBuffer());
        fs.writeFileSync(localPath, buffer);

        let imgPath;
        if (name.endsWith(".xlsx") || name.endsWith(".xls")) {
            imgPath = await excelToImage(localPath);
        } else {
            imgPath = await txtToImage(localPath);
        }

        await ctx.replyWithPhoto(
            { source: fs.readFileSync(imgPath) },
            { caption: `转换完成：${file.file_name}` }
        );

        // 立即删除所有临时文件（零残留）
        [localPath, imgPath].forEach(p => fs.existsSync(p) && fs.unlinkSync(p));

    } catch (e) {
        console.error(e);
        await ctx.reply("转换失败了：\n" + e.message);
    }
});

// 启动
(async () => {
    await getBrowser();
    console.log("Bot 已启动，等待文件...");
    bot.start();
})();
