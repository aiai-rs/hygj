# 官方 puppeteer 专用镜像，自带最新 Chrome，Render 免费层完美支持
FROM ghcr.io/puppeteer/puppeteer:23.5.0

ENV NODE_ENV=production

WORKDIR /app

# 先复制 package.json 安装依赖
COPY package*.json ./
RUN npm install --omit=dev

# 复制所有代码
COPY . .

# 启动命令
CMD ["node", "index.js"]
