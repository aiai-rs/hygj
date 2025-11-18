# 使用官方 puppeteer 专用镜像（自带 Chrome + 所有依赖）
FROM ghcr.io/puppeteer/puppeteer:23.5.0

ENV NODE_ENV=production

WORKDIR /app

# 复制 package.json 先装依赖
COPY package*.json ./
RUN npm install --omit=dev

# 复制代码
COPY . .

# 启动
CMD ["node", "main.js"]
