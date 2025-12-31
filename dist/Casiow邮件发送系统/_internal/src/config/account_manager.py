import json
import os
import sys
sys.path.append(os.path.dirname(os.path.dirname(os.path.abspath(__file__))))

from config.config import EmailConfig

class AccountManager:
    def __init__(self, config_file=None):
        """
        初始化账户管理器
        
        Args:
            config_file (str): 存储邮箱账户配置的文件名
        """
        if config_file is None:
            # 优先检查 config 目录下的配置文件
            self.config_file = os.path.join("config", "email_accounts.json")
            if not os.path.exists(self.config_file):
                # 兼容旧路径
                self.config_file = "email_accounts.json"
        else:
            self.config_file = config_file
            
        self.accounts = self.load_accounts()
    
    def load_accounts(self):
        """
        从配置文件加载邮箱账户信息
        
        Returns:
            dict: 邮箱账户信息字典
        """
        # 如果配置文件不存在，创建默认配置
        if not os.path.exists(self.config_file):
            default_accounts = EmailConfig.EMAIL_ACCOUNTS.copy()
            self.save_accounts(default_accounts)
            return default_accounts
        
        # 从配置文件加载账户信息
        try:
            with open(self.config_file, 'r', encoding='utf-8') as f:
                return json.load(f)
        except Exception as e:
            print(f"加载邮箱账户配置失败: {e}")
            return EmailConfig.EMAIL_ACCOUNTS.copy()
    
    def save_accounts(self, accounts=None):
        """
        保存邮箱账户信息到配置文件
        
        Args:
            accounts (dict, optional): 要保存的账户信息，默认为当前账户
        """
        if accounts is None:
            accounts = self.accounts
            
        try:
            with open(self.config_file, 'w', encoding='utf-8') as f:
                json.dump(accounts, f, ensure_ascii=False, indent=2)
        except Exception as e:
            print(f"保存邮箱账户配置失败: {e}")
    
    def get_accounts(self):
        """
        获取所有邮箱账户
        
        Returns:
            dict: 所有邮箱账户信息
        """
        return self.accounts.copy()
    
    def get_account_names(self):
        """
        获取所有邮箱账户名称
        
        Returns:
            list: 邮箱账户名称列表
        """
        return list(self.accounts.keys())
    
    def add_account(self, name, smtp_server, smtp_port, sender, password, sender_name):
        """
        添加新的邮箱账户
        
        Args:
            name (str): 账户名称
            smtp_server (str): SMTP服务器地址
            smtp_port (int): SMTP端口号
            sender (str): 发件人邮箱地址
            password (str): 邮箱密码或授权码
            sender_name (str): 发件人显示名称
            
        Returns:
            bool: 添加成功返回True，否则返回False
        """
        if name in self.accounts:
            print(f"账户 '{name}' 已存在")
            return False
            
        self.accounts[name] = {
            'SMTP_SERVER': smtp_server,
            'SMTP_PORT': smtp_port,
            'SENDER': sender,
            'PASSWORD': password,
            'SENDER_NAME': sender_name
        }
        
        self.save_accounts()
        return True
    
    def update_account(self, name, smtp_server=None, smtp_port=None, sender=None, 
                      password=None, sender_name=None):
        """
        更新邮箱账户信息
        
        Args:
            name (str): 账户名称
            smtp_server (str, optional): SMTP服务器地址
            smtp_port (int, optional): SMTP端口号
            sender (str, optional): 发件人邮箱地址
            password (str, optional): 邮箱密码或授权码
            sender_name (str, optional): 发件人显示名称
            
        Returns:
            bool: 更新成功返回True，否则返回False
        """
        if name not in self.accounts:
            print(f"账户 '{name}' 不存在")
            return False
            
        account = self.accounts[name]
        
        if smtp_server is not None:
            account['SMTP_SERVER'] = smtp_server
        if smtp_port is not None:
            account['SMTP_PORT'] = smtp_port
        if sender is not None:
            account['SENDER'] = sender
        if password is not None:
            account['PASSWORD'] = password
        if sender_name is not None:
            account['SENDER_NAME'] = sender_name
            
        self.save_accounts()
        return True
    
    def delete_account(self, name):
        """
        删除邮箱账户
        
        Args:
            name (str): 要删除的账户名称
            
        Returns:
            bool: 删除成功返回True，否则返回False
        """
        if name not in self.accounts:
            print(f"账户 '{name}' 不存在")
            return False
            
        if len(self.accounts) <= 1:
            print("不能删除最后一个账户")
            return False
            
        del self.accounts[name]
        self.save_accounts()
        return True
    
    def get_account(self, name):
        """
        获取指定邮箱账户信息
        
        Args:
            name (str): 账户名称
            
        Returns:
            dict or None: 账户信息字典，不存在则返回None
        """
        return self.accounts.get(name)