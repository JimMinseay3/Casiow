import time
import os
from config import EmailConfig
from email_sender import EmailSender
from logger import setup_logger
from recipient_window import get_email_info
import tkinter.messagebox as messagebox

def main():
    # 初始化配置
    config = EmailConfig()
    
    # 获取邮件信息（窗口保持打开，用户点击确认后返回）
    print("请在弹出窗口中输入邮件信息，点击确认开始发送...")
    email_info = get_email_info()
    
    # 检查是否有有效信息
    if not email_info["recipients"] or not email_info["attachments"]:
        print("未获取到有效信息，程序结束")
        return
    
    # 设置附件
    config.set_attachments(email_info["attachments"])
    
    # 初始化日志
    logger = setup_logger()
    logger.info("开始批量发送邮件...")
    logger.info(f"收件人数量: {len(email_info['recipients'])}")
    logger.info(f"附件数量: {len(config.ATTACHMENTS)}")
    
    # 打印附件清单
    logger.info("附件清单:")
    for attachment in config.ATTACHMENTS:
        logger.info(f"- {os.path.basename(attachment)}")
    
    # 初始化发送器
    email_sender = EmailSender(config)
    
    # 批量发送
    total_success = 0
    total_fail = 0
    
    for i, recipient in enumerate(email_info["recipients"]):
        logger.info(f"\n正在为 {recipient['email']} 发送 ({i+1}/{len(email_info['recipients'])})...")
        success, fail = email_sender.send_with_separate_attachments(recipient)
        total_success += success
        total_fail += fail

    # 发送结果提示（弹窗）
    result_msg = f"""发送完成！
    成功: {total_success} 封
    失败: {total_fail} 封"""
    messagebox.showinfo("发送结果", result_msg)
    
    # 发送完成后，窗口仍保持打开，用户可继续操作或关闭

if __name__ == "__main__":
    main()