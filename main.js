const { Bot } = require("grammy");
const fs = require("fs");
const path = require("path");
const puppeteer = require("puppeteer");
const ExcelJS = require("exceljs");

const TOKEN = process.env.BOT_TOKEN;
if (!TOKEN) {
    console.error("请在环境变量中设置 BOT_TOKEN！");
    process.exit(1);
}

const bot = new Bot(TOKEN);

// 全局复用浏览器，节省内存和启动时间
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
                "--single-process",
                "--no-zygote"
            ],
            executablePath: process.env.PUPPETEER_EXECUTABLE_PATH || undefined
        });
    }
    return browser;
}

// Excel → 图片
async function excelToImage(filePath) {
    const workbook = new ExcelJS.Workbook();
    await workbook.xlsx.readFile(filePath);
    const worksheet = workbook.worksheets[0];

    let html = `
    <style>
        table { border-collapse: collapse; font-family: Arial, sans-serif; font-size: 14px; }
        td, th { border: 1px solid #333; padding: 8px 12px; text-align: center; }
        th { background: #f0f0f0; }
    </style>
    <table>`;

    worksheet.eachRow((row, rowNumber) => {
        html += "<tr>";
        row.eachCell({ includeEmpty: true }, (cell) => {
            const value = cell.value || "";
            if (rowNumber === 1) {
                html += `<th>${value}</th>`;
            } else {
                html += `<td>${value}</td>`;
            }
        });
        html += "</tr>";
    });
    html += "</table>";

    const page = await browser.newPage();
    await page.setContent(html, { waitUntil: "networkidle0" });
    const body = await page.$("body");
    const bounding = await body.boundingBox();

    const screenshot = await page.screenshot({
        clip: {
            x: bounding.x,
            y: bounding.y,
            width: bounding.width + 40,
            height: bounding.height + 40
        }
    });
    await page.close();

    const outPath = filePath.replace(/\.(xlsx|xls)$/i, ".png");
    fs.writeFileSync(outPath, screenshot);
    return outPath;
}

// TXT → 图片
async function txtToImage(filePath) {
    const content = fs.readFileSync(filePath, "utf-8");
    const escaped = content
        .replace(/&/g, "&amp;")
        .replace(/</g, "&lt;")
        .replace(/>/g, "&gt;");

    const html = `
    <!DOCTYPE html>
    <html>
    <body style="margin:30px; background:white;">
    <pre style="font-family:'Courier New',monospace; font-size:16px; line-height:1.5;">${escaped}</pre>
    </body>
    </html>`;

    const page = await browser.newPage();
    await page.setContent(html);
    await page.setViewport({ width: 1200, height: 800 });

    const pre = await page.$("pre");
    const box = await pre.boundingBox();

    const screenshot = await page.screenshot({
        clip: {
            x: box.x - 20,
            y: box.y - 20,
            width: box.width + 40,
            height: box.height + 40
        }
    });
    await page.close();

    const outPath = filePath.replace(/\.txt$/i, ".png");
    fs.writeFileSync(outPath, screenshot);
    return outPath;
}

// ========== Bot 逻辑 ==========
bot.command("start", (ctx) =>
    ctx.reply("把 .xlsx 或 .txt 文件发给我，我会立刻转成高清图片发回给你～\n支持超大表格和长文本！")
);

bot.on("message:document", async (ctx) => {
    const file = ctx.message.document;
    const fileName = file.file_name.toLowerCase();

    if (!fileName.endsWith(".xlsx") && !fileName.endsWith(".xls") && !fileName.endsWith(".txt")) {
        return ctx.reply("请上传 .xlsx、.xls 或 .txt 文件哦～");
    }

    await ctx.reply("收到！正在转换，请稍等 5-20 秒...");

    try {
        const browserInstance = await getBrowser();
        const fileLink = await ctx.getFile().then(f => f.getUrl());
        const localPath = path.join("/tmp", file.file_name);

        // 下载文件
        const response = await fetch(fileLink);
        const buffer = await response.arrayBuffer();
        fs.writeFileSync(localPath, Buffer.from(buffer));

        let imgPath;
        if (fileName.endsWith(".xlsx") || fileName.endsWith(".xls")) {
            imgPath = await excelToImage(localPath);
        } else {
            imgPath = await txtToImage(localPath);
        }

        await ctx.replyWithPhoto({ source: imgPath }, { caption: `已转换：${file.file_name}` });

        // 清理临时文件
        [localPath, imgPath].forEach(p => fs.existsSync(p) && fs.unlinkSync(p));

    } catch (err) {
        console.error(err);
        await ctx.reply("转换失败了：\n" + err.message);
    }
});

(async () => {
    await getBrowser(); // 提前启动浏览器
    console.log("Bot 已启动，等待文件...");
    await bot.start();
})();
