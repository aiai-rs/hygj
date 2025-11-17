import telebot
import os
from file_to_image import convert_file
import tempfile

# 从环境变量读取Token（安全！）
BOT_TOKEN = os.getenv('BOT_TOKEN')
if not BOT_TOKEN:
    raise ValueError("BOT_TOKEN 环境变量未设置！请在Render Dashboard中添加。")

bot = telebot.TeleBot(BOT_TOKEN)

@bot.message_handler(commands=['start'])
def start_message(message):
    bot.reply_to(message, "上传文件（xlsx/csv/pptx/swf），我帮你转成图片！")

@bot.message_handler(content_types=['document'])
def handle_document(message):
    try:
        file_info = bot.get_file(message.document.file_id)
        downloaded_file = bot.download_file(file_info.file_path)
        
        # 创建临时文件
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
        print(f"Bot错误: {e}")  # Render日志

if __name__ == "__main__":
    print("Bot启动中...")
    bot.polling(none_stop=True)