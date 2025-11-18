const { Bot } = require("grammy");
const fs = require("fs");
const path = require("path");
const puppeteer = require("puppeteer");

// 你的 Bot Token 从环境变量读取
const bot = new Bot(process.env.BOT_TOKEN || "");

let browser;

// 获取浏览器实例（官方镜像自带 Chrome）
async function getBrowser() {
    if (!browser) {
        browser = await puppeteer.launch({
            headless: true,
            args: ["--no-sandbox", "--disable-setuid-sandbox", "--disable-dev-shm-usage"]
        });
    }
    return browser;
}

// Excel → HTML → 截图
async function excelToImage(filePath) {
    const ExcelJS = require("exceljs");
    const workbook = new ExcelJS.Workbook();
    await workbook.xlsx.readFile(filePath);
    const worksheet = workbook.worksheets[0];

    let html = `<table border="1" style="border-collapse:collapse;font-family:Arial,sans-serif;font-size:15px;">`;
    worksheet.eachRow({ includeEmpty: true }, (row, rowNumber) => {
        html += "<tr>";
        row.eachCell({ includeEmpty: true }, (cell) => {
            const value = cell.value ?? "";
            if (rowNumber === 1) {
                html += `<th style="background:#f0f0f0;padding:10px;min-width:100px">${value}</th>`;
            } else {
                html += `<td style="padding:8px;text-align:center;min-width:100px">${value}</td>`;
            }
        });
        html += "</tr>";
    });
    html += "</table>";

    const page = await browser.newPage();
    await page.setContent(html, { waitUntil: "networkidle0" });
    const buffer = await page.screenshot({ fullPage: true });
    await page.close();
    return buffer;
}

// TXT → 图片
async function txtToImage(filePath) {
    const content = fs.readFileSync(filePath, "utf-8")
        .replace(/&/g, "&amp;")
        .replace(/</g, "&lt;")
        .replace(/>/g, "&gt;");

    const html = `
    <!DOCTYPE html>
    <html><body style="margin:40px;background:white;">
    <pre style="font-family:'Courier New',monospace;font-size:16px;line-height:1.5;white-space:pre-wrap;word-wrap:break-word;">
    ${content}
    </pre>
    </body></html>`;

    const page = await browser.newPage();
    await page.setContent(html, { waitUntil: "networkidle0" });
    const buffer = await page.screenshot({ fullPage: true });
    await page.close();
    return buffer;
}

// ============ Telegram Bot 逻辑 ============
bot.command("start", (ctx) => {
    ctx.reply("把 .xlsx、.xls 或 .txt 文件发给我，我会立即转成高清图片发回给你！\n支持超大表格和长文本～");
});

bot.on("message:document", async (ctx) => {
    const doc = ctx.message.document;
    const fileName = doc.file_name?.toLowerCase() || "";

    if (!fileName.endsWith(".xlsx") && !fileName.endsWith(".xls") && !fileName.endsWith(".txt")) {
        return ctx.reply("只支持 .xlsx、.xls 和 .txt 文件哦～");
    }

    await ctx.reply("收到！正在转换，请稍等 5-25 秒...");

    try {
        await getBrowser(); // 确保浏览器已启动

        // 下载文件
        const fileUrl = await ctx.getFile().then(f => f.getUrl());
        const localPath = path.join("/tmp", doc.file_name);
        const response = await fetch(fileUrl);
        const buffer = Buffer.from(await response.arrayBuffer());
        fs.writeFileSync(localPath, buffer);

        // 转换
        const imgBuffer = fileName.endsWith(".txt") 
            ? await txtToImage(localPath)
            : await excelToImage(localPath);

        // 发回图片
        await ctx.replyWithPhoto(
            { source: imgBuffer },
            { caption: `已转换：${doc.file_name}` }
        );

        // 立即删除临时文件（零残留）
        fs.unlinkSync(localPath);

    } catch (error) {
        console.error("转换错误:", error);
        await ctx.reply(`转换失败了：\n${error.message}`);
    }
});

// 启动 Bot（丢弃旧更新，避免 409 冲突）
bot.start({
    drop_pending_updates: true
});

console.log("xlsx/txt 转图片 Bot 已启动成功！");
