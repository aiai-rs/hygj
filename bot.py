import telebot
import os
from flask import Flask, request, abort
from file_to_image import convert_file
import tempfile

# 从环境变量读取Token（安全！）
BOT_TOKEN = os.getenv('BOT_TOKEN')
if not BOT_TOKEN:
    raise ValueError("BOT_TOKEN 环境变量未设置！请在Render Dashboard中添加。")

bot = telebot.TeleBot(BOT_TOKEN)
app = Flask(__name__)  # Flask app，用于 webhook

@bot.message_handler(commands=['start'])
def start_message(message):
    bot.reply_to(message, "上传文件（xlsx/csv/pptx/swf），我帮你转成图片！")

@bot.message_handler(content_types=['document'])
def handle_document(message):
    try:
        file_info = bot.get_file(message.document.file_id)
        downloaded_file = bot.download_file(file_info.file_path)
        
        # 创建临时文件（带后缀）
        suffix = os.path.splitext(message.document.file_name)[1] if message.document.file_name else ''
        with tempfile.NamedTemporaryFile(delete=False, suffix=suffix) as temp_file:
            temp_file.write(downloaded_file)
            temp_file_path = temp_file.name
        
        # 转换
        bot.reply_to(message, "正在转换...")
        images = convert_file(temp_file_path)
        
        # 发送图片
        if images:
            for img_path in images:
                with open(img_path, 'rb') as photo:
                    bot.send_photo(message.chat.id, photo, caption=os.path.basename(img_path))
            bot.reply_to(message, f"转换完成！共 {len(images)} 张图片。")
        else:
            bot.reply_to(message, "转换失败或不支持格式！支持：xlsx/xls/xlsm/csv/pptx/pot/swf")
        
        # 清理：删除临时文件（无记录）
        if os.path.exists(temp_file_path):
            os.unlink(temp_file_path)
        
    except Exception as e:
        bot.reply_to(message, f"错误: {str(e)}")
        print(f"Bot错误: {e}")  # Render 日志输出

# Webhook 端点：接收 Telegram 更新
@app.route('/webhook', methods=['POST'])
def webhook():
    if request.headers.get('content-type') == 'application/json':
        json_string = request.get_data().decode('utf-8')
        update = telebot.types.Update.de_json(json_string)
        bot.process_new_updates([update])
        return ''
    else:
        abort(403)

# 启动 Flask app（Render 用 gunicorn 运行此文件）
if __name__ == "__main__":
    # 本地测试：python bot.py
    port = int(os.environ.get('PORT', 5000))
    app.run(host='0.0.0.0', port=port, debug=False)
    
    # 注意：部署后，手动设置 webhook（见上方 curl 命令）
    # 示例（本地或一次性脚本）：
    # webhook_url = 'https://your-service.onrender.com/webhook'
    # bot.remove_webhook()  # 移除旧 polling
    # bot.set_webhook(url=webhook_url)
