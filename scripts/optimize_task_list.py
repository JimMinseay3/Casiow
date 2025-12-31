#!/usr/bin/env python
# -*- coding: utf-8 -*-
"""
优化任务列表显示效果的测试脚本
"""

import sys
import os
import tkinter as tk
from tkinter import ttk, messagebox
from datetime import datetime, timedelta

# 添加项目根目录到Python路径
project_root = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
sys.path.insert(0, project_root)
sys.path.insert(0, os.path.join(project_root, 'src'))

from src.ui.recipient_window import RecipientInputWindow
from src.core.send_pool import SendPool, SendPoolItem

def create_sample_tasks(app):
    """创建示例任务来测试不同的显示效果"""
    test_recipients = [
        {"email": "user1@example.com", "title": "重要通知1", "content": "这是第一封重要通知邮件的内容"},
        {"email": "user2@example.com", "title": "重要通知2", "content": "这是第二封重要通知邮件的内容"},
        {"email": "user3@example.com", "title": "重要通知3", "content": "这是第三封重要通知邮件的内容"},
        {"email": "user4@example.com", "title": "重要通知4", "content": "这是第四封重要通知邮件的内容"},
        {"email": "user5@example.com", "title": "重要通知5", "content": "这是第五封重要通知邮件的内容"}
    ]
    
    test_attachments = ["document1.pdf", "image1.png", "report1.docx"]
    
    test_account_config = {
        'SMTP_SERVER': 'smtp.example.com',
        'SMTP_PORT': 587,
        'SENDER': 'sender@example.com',
        'PASSWORD': 'password',
        'SENDER_NAME': 'Test Sender',
        'SEND_INTERVAL': 1
    }
    
    # 创建不同状态的任务
    # 1. 待处理任务
    item_id1 = app.send_pool.add_item(test_recipients[:2], test_attachments[:1], test_account_config)
    
    # 2. 发送中任务
    item_id2 = app.send_pool.add_item(test_recipients[2:], test_attachments, test_account_config)
    item2 = app.send_pool.get_item(item_id2)
    item2.status = "sending"
    item2.start_time = datetime.now() - timedelta(minutes=5)
    
    # 3. 已完成任务
    item_id3 = app.send_pool.add_item(test_recipients[:3], test_attachments[:2], test_account_config)
    item3 = app.send_pool.get_item(item_id3)
    item3.status = "completed"
    item3.start_time = datetime.now() - timedelta(minutes=15)
    item3.end_time = datetime.now() - timedelta(minutes=10)
    item3.success_count = 3
    item3.fail_count = 0
    
    # 4. 失败任务（带错误信息）
    item_id4 = app.send_pool.add_item(test_recipients[3:], test_attachments, test_account_config)
    item4 = app.send_pool.get_item(item_id4)
    item4.status = "failed"
    item4.start_time = datetime.now() - timedelta(minutes=8)
    item4.end_time = datetime.now() - timedelta(minutes=7)
    item4.success_count = 2
    item4.fail_count = 3
    item4.error_message = "部分邮件发送失败，共3个错误"
    item4.error_messages = [
        {
            "recipient": "user4@example.com",
            "error": "SMTP服务器连接超时",
            "timestamp": (datetime.now() - timedelta(minutes=7, seconds=30)).isoformat()
        },
        {
            "recipient": "user5@example.com",
            "error": "附件大小超过限制",
            "timestamp": (datetime.now() - timedelta(minutes=7, seconds=15)).isoformat()
        },
        {
            "recipient": "user6@example.com",
            "error": "收件人邮箱不存在",
            "timestamp": (datetime.now() - timedelta(minutes=7)).isoformat()
        }
    ]
    
    print(f"已添加测试任务: {item_id1}, {item_id2}, {item_id3}, {item_id4}")

def main():
    root = tk.Tk()
    app = RecipientInputWindow(root)
    
    # 创建示例任务
    create_sample_tasks(app)
    
    # 刷新界面显示
    app.refresh_pool_status()
    
    # 切换到发送池状态标签页
    app.notebook.select(app.pool_frame)
    
    root.mainloop()

if __name__ == "__main__":
    main()