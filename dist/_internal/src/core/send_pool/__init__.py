#!/usr/bin/env python
# -*- coding: utf-8 -*-
"""
发送池模块，用于管理邮件发送任务队列
"""

import threading
import time
import uuid
from datetime import datetime
from typing import List, Dict, Optional


class SendPoolItem:
    """发送池中的单个任务项"""
    
    def __init__(self, recipients: List[Dict], attachments: List[str], account_config: Dict):
        self.id = int(time.time() * 1000000) % 1000000000  # 生成唯一ID
        self.recipients = recipients
        self.attachments = attachments
        self.account_config = account_config
        self.status = "pending"  # pending, sending, completed, failed
        self.created_time = datetime.now()
        self.start_time: Optional[datetime] = None
        self.end_time: Optional[datetime] = None
        self.success_count = 0
        self.fail_count = 0
        self.error_message: Optional[str] = None
        self.error_messages: List[Dict] = []


class SendPool:
    """发送池类，用于管理邮件发送任务队列"""
    
    def __init__(self):
        self.items: Dict[int, SendPoolItem] = {}
        self.lock = threading.Lock()
        self.processing = False
        self.processing_thread: Optional[threading.Thread] = None
        
    def add_item(self, recipients: List[Dict], attachments: List[str], account_config: Dict) -> int:
        """添加一个新的发送任务到池中"""
        with self.lock:
            item = SendPoolItem(recipients, attachments, account_config)
            self.items[item.id] = item
            return item.id
    
    def remove_item(self, item_id: int) -> bool:
        """从池中移除一个任务"""
        with self.lock:
            if item_id in self.items:
                del self.items[item_id]
                return True
            return False
    
    def get_item(self, item_id: int) -> Optional[SendPoolItem]:
        """获取指定ID的任务"""
        with self.lock:
            return self.items.get(item_id)
    
    def get_all_items(self) -> List[SendPoolItem]:
        """获取所有任务"""
        with self.lock:
            return list(self.items.values())
    
    def start_processing(self):
        """启动发送池处理线程"""
        if not self.processing:
            self.processing = True
            self.processing_thread = threading.Thread(target=self._process_items, daemon=True)
            self.processing_thread.start()
    
    def stop_processing(self):
        """停止发送池处理线程"""
        self.processing = False
        if self.processing_thread and self.processing_thread.is_alive():
            self.processing_thread.join()
    
    def _process_items(self):
        """处理发送池中的任务（在单独的线程中运行）"""
        while self.processing:
            # 查找待处理的任务
            pending_items = [item for item in self.get_all_items() if item.status == "pending"]
            
            for item in pending_items:
                # 更新任务状态
                item.status = "sending"
                item.start_time = datetime.now()
                
                # 实际发送邮件
                try:
                    # 导入所需的模块
                    from src.config.config import EmailConfig
                    from src.core.email_sender import EmailSender
                    
                    # 创建临时配置
                    temp_account = "temp_account"
                    config = EmailConfig()
                    config.EMAIL_ACCOUNTS[temp_account] = {
                        'SMTP_SERVER': item.account_config.get('SMTP_SERVER', ''),
                        'SMTP_PORT': item.account_config.get('SMTP_PORT', 587),
                        'SENDER': item.account_config.get('SENDER', ''),
                        'PASSWORD': item.account_config.get('PASSWORD', ''),
                        'SENDER_NAME': item.account_config.get('SENDER_NAME', '')
                    }
                    config.CURRENT_ACCOUNT = temp_account
                    config.SEND_INTERVAL = item.account_config.get('SEND_INTERVAL', 1)
                    config.set_attachments(item.attachments)
                    
                    # 创建发送器
                    email_sender = EmailSender(config)
                    
                    # 发送邮件
                    item.success_count = 0
                    item.fail_count = 0
                    item.error_messages = []
                    
                    for recipient in item.recipients:
                        if self.processing:  # 检查是否需要停止处理
                            if email_sender.send(recipient):
                                item.success_count += 1
                            else:
                                item.fail_count += 1
                                item.error_messages.append({
                                    'email': recipient['email'],
                                    'error': f"发送失败"
                                })
                            # 添加发送间隔
                            time.sleep(config.SEND_INTERVAL)
                        else:
                            break
                    
                    if item.fail_count == 0:
                        item.status = "completed"
                    else:
                        item.status = "failed"
                    
                except Exception as e:
                    item.status = "failed"
                    item.error_message = str(e)
                    item.error_messages.append({
                        'email': 'unknown',
                        'error': str(e)
                    })
                
                # 更新任务结束时间
                item.end_time = datetime.now()
                
                # 如果不再需要处理更多任务，则退出循环
                if not self.processing:
                    break
            
            # 如果没有待处理的任务，短暂休眠以避免过度占用CPU
            if not pending_items:
                time.sleep(1)


# 示例使用方法：
# send_pool = SendPool()
# item_id = send_pool.add_item(
#     recipients=[{"email": "user@example.com", "title": "Test", "content": "Test content"}],
#     attachments=[],
#     account_config={"SMTP_SERVER": "smtp.example.com", "SMTP_PORT": 587}
# )
# send_pool.start_processing()