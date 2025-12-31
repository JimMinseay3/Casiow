import smtplib
import time
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from email.mime.application import MIMEApplication
from email.header import Header
import os

class EmailSender:
    def __init__(self, config):
        self.config = config
    
    def send(self, recipient_info):
        """发送单封邮件（包含所有附件）"""
        try:
            # 创建邮件对象
            msg = MIMEMultipart()
            
            # 正确设置 From 头部（修复 RFC5322 格式问题）
            sender_name = Header(self.config.SENDER_NAME, 'utf-8').encode()
            msg['From'] = f"{sender_name} <{self.config.SENDER}>"
            
            msg['To'] = recipient_info['email']
            msg['Subject'] = recipient_info['title']
            
            # 添加邮件正文
            msg.attach(MIMEText(recipient_info['content'], 'plain', 'utf-8'))
            
            # 添加所有附件
            for file_path in self.config.ATTACHMENTS:
                if not os.path.exists(file_path):
                    print(f"附件不存在，跳过: {file_path}")
                    continue
                
                try:
                    file_name = os.path.basename(file_path)
                    with open(file_path, 'rb') as file:
                        attachment = MIMEApplication(file.read(), _subtype="octet-stream")
                        attachment.add_header('Content-Disposition', 'attachment', filename=file_name)
                        msg.attach(attachment)
                    
                    print(f"已添加附件: {file_name}")
                except Exception as e:
                    print(f"添加附件失败: {file_path}, 错误: {e}")
            
            # 发送邮件
            with smtplib.SMTP_SSL(self.config.SMTP_SERVER, self.config.SMTP_PORT) as server:
                # 启用调试模式（可选）
                # server.set_debuglevel(1)
                
                server.login(self.config.SENDER, self.config.PASSWORD)
                server.sendmail(self.config.SENDER, recipient_info['email'], msg.as_string())
            
            print(f"邮件已成功发送至: {recipient_info['email']}")
            return True
            
        except Exception as e:
            print(f"发送邮件失败: {e}")
            return False
    
    def send_with_separate_attachments(self, recipient_info):
        """为每个附件单独发送一封邮件（带重试机制）"""
        success_count = 0
        fail_count = 0
        total_attachments = len(self.config.ATTACHMENTS)
        
        for i, attachment in enumerate(self.config.ATTACHMENTS):
            file_name = os.path.basename(attachment)
            print(f"\n发送附件 {i+1}/{total_attachments}: {file_name} 到 {recipient_info['email']}")
            
            for attempt in range(3):  # 最多重试3次
                try:
                    # 创建邮件对象
                    msg = MIMEMultipart()
                    
                    # 正确设置 From 头部（修复 RFC5322 格式问题）
                    sender_name = Header(self.config.SENDER_NAME, 'utf-8').encode()
                    msg['From'] = f"{sender_name} <{self.config.SENDER}>"
                    
                    msg['To'] = recipient_info['email']
                    msg['Subject'] = f"{recipient_info['title']} - {file_name}"
                    
                    # 添加邮件正文
                    content = f"{recipient_info['content']}\n\n附件: {file_name}"
                    msg.attach(MIMEText(content, 'plain', 'utf-8'))
                    
                    # 添加附件
                    if not os.path.exists(attachment):
                        raise FileNotFoundError(f"附件不存在: {attachment}")
                    
                    with open(attachment, 'rb') as file:
                        part = MIMEApplication(file.read(), _subtype="octet-stream")
                        part.add_header('Content-Disposition', 'attachment', filename=file_name)
                        msg.attach(part)
                    
                    # 发送邮件
                    with smtplib.SMTP_SSL(self.config.SMTP_SERVER, self.config.SMTP_PORT) as server:
                        # 启用调试模式（可选）
                        # server.set_debuglevel(1)
                        
                        server.login(self.config.SENDER, self.config.PASSWORD)
                        server.sendmail(self.config.SENDER, recipient_info['email'], msg.as_string())
                    
                    print(f"尝试 {attempt+1}/3 成功 ✔️")
                    success_count += 1
                    break  # 成功后跳出重试循环
                    
                except FileNotFoundError as e:
                    print(f"错误: {e}")
                    fail_count += 1
                    break  # 文件不存在，无需重试
                    
                except smtplib.SMTPException as e:
                    error_code = str(e).split()[0] if str(e) else "未知错误"
                    print(f"尝试 {attempt+1}/3 SMTP错误 ({error_code}): {e}")
                    
                    # 特定错误码处理
                    if "550" in error_code or "553" in error_code:
                        print("可能是发件人权限问题或附件限制，跳过此附件")
                        fail_count += 1
                        break
                    
                    if attempt < 2:  # 不是最后一次尝试
                        print(f"将在 {self.config.SEND_INTERVAL*2} 秒后重试...")
                        time.sleep(self.config.SEND_INTERVAL*2)
                    else:
                        print("已达到最大重试次数 ❌")
                        fail_count += 1
                        
                except Exception as e:
                    print(f"尝试 {attempt+1}/3 其他错误: {e}")
                    if attempt < 2:
                        print(f"将在 {self.config.SEND_INTERVAL*2} 秒后重试...")
                        time.sleep(self.config.SEND_INTERVAL*2)
                    else:
                        print("已达到最大重试次数 ❌")
                        fail_count += 1
            
            # 正常发送间隔
            time.sleep(self.config.SEND_INTERVAL)
        
        return success_count, fail_count