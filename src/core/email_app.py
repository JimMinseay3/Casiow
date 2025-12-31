import sys
import os
import multiprocessing

# 添加src目录到Python路径
src_dir = os.path.join(os.path.dirname(os.path.abspath(__file__)), '..', '..')
sys.path.insert(0, src_dir)

import time
from src.config.config import EmailConfig
from src.core.email_sender import EmailSender
from src.utils.logger import setup_logger
from src.ui.recipient_window import RecipientInputWindow
import tkinter as tk
import tkinter.messagebox as messagebox

def main():
    # 初始化配置，直接使用默认账户
    selected_account = "default"
    logger = setup_logger()
    logger.info(f"使用默认邮箱账户: {selected_account}")
    config = EmailConfig(selected_account)
    
    # 创建主窗口并启动持续运行模式
    logger.info("邮件发送系统已启动，窗口将保持打开状态...")
    logger.info("点击'确认并发送'按钮将任务添加到发送池，系统会自动处理发送任务。")
    root = tk.Tk()
    app = RecipientInputWindow(root, selected_account)
    
    # 启动发送池处理
    app.send_pool.start_processing()
    
    # 进入主事件循环（窗口保持打开）
    try:
        root.mainloop()
    except KeyboardInterrupt:
        logger.info("收到退出信号，正在关闭...")
    
    # 停止发送池处理
    app.send_pool.stop_processing()
    logger.info("邮件发送系统已关闭。")
    