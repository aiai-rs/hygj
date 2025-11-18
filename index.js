const { Bot } = require("grammy");
const fs = require("fs");
const puppeteer = require("puppeteer");

const bot = new Bot(process.env.BOT_TOKEN);

let browser;
async function getBrowser() {
    if (browser) return browser;

    console.log("正在启动 Chrome...");
    browser = await puppeteer.launch({
        headless: true,
        args: [
            "--no-sandbox",
            "--disable-setuid-sandbox",
            "--disable-dev-shm-usage",
            "--disable-gpu",
            "--disable-extensions",
            "--no-zygote",
            "--single-process",
            "--disable-background-timer-throttling",
            "--disable-renderer-backgrounding",
            "--disable-features=TranslateUI,BlinkGenPropertyTrees"
        ],
        ignoreHTTPSErrors: true,
        dumpio: false
    });
    console.log("Chrome 启动成功！");
    return browser;
}

// 转换函数（不变，略... 同上个版本的 excelToImage 和 txtToImage）

// Bot 逻辑（同上个版本，一字不动）

// 关键：加一个延时启动，让容器网络先稳下来
setTimeout(async () => {
    try {
        await getBrowser(); // 提前启动一次 Chrome，避免第一次使用时超时
    } catch (e) {
        console.error("预启动 Chrome 失败，但会自动重试:", e.message);
    }
}, 8000);

bot.start({ drop_pending_updates: true });
console.log("Bot 已启动成功！正在等待文件...");
