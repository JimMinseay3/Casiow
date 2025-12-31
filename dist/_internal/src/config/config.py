import json
import os

class EmailConfig:
    # 默认邮箱配置（包含所有账户）
    DEFAULT_EMAIL_ACCOUNTS = {
        'wecom': {
            'SMTP_SERVER': 'smtp.exmail.qq.com',
            'SMTP_PORT': 465,
            'SENDER': 'jiaoyizhuli@ysgjzbfwyxgs1.wecom.work',
            'PASSWORD': 'XLju85xYtKNQJQtg',  # 授权码
            'SENDER_NAME': '交易助理'
        },
        'hualiang': {
            'SMTP_SERVER': 'mail.hualiang.com.hk',
            'SMTP_PORT': 25,
            'SENDER': 'hlicl@hualiang.com.hk',
            'PASSWORD': 'Hc@3619',  # 授权码
            'SENDER_NAME': 'hlicl@hualiang.com.hk'
        },
        'YS': {
            'SMTP_SERVER': 'smtp.exmail.qq.com',
            'SMTP_PORT': 465,
            'SENDER': 'jiaoyizhuliyanshengpinzh@ys-ics.net',
            'PASSWORD': 'oCPQKusLvhQwnPZW',  # 授权码
            'SENDER_NAME': '交易助理-衍生品智库'
        }
    }
    
    def __init__(self, account_name=None):
        """初始化配置，可选择邮箱账户"""
        # 加载邮箱账户配置
        self.EMAIL_ACCOUNTS = self.load_email_accounts()
        
        # 当前使用的邮箱账户
        self.CURRENT_ACCOUNT = account_name if account_name and account_name in self.EMAIL_ACCOUNTS else next(iter(self.EMAIL_ACCOUNTS))
        self._attachments = []
    
    def load_email_accounts(self):
        """从外部配置文件加载邮箱账户信息"""
        config_file = "email_accounts.json"
        if os.path.exists(config_file):
            try:
                with open(config_file, 'r', encoding='utf-8') as f:
                    return json.load(f)
            except Exception as e:
                print(f"加载邮箱账户配置失败，使用默认配置: {e}")
                return self.DEFAULT_EMAIL_ACCOUNTS.copy()
        else:
            # 如果配置文件不存在，创建默认配置文件
            try:
                with open(config_file, 'w', encoding='utf-8') as f:
                    json.dump(self.DEFAULT_EMAIL_ACCOUNTS, f, ensure_ascii=False, indent=2)
            except Exception as e:
                print(f"创建默认邮箱账户配置文件失败: {e}")
            return self.DEFAULT_EMAIL_ACCOUNTS.copy()
    
    SEND_INTERVAL = 2  # 秒
    
    @property
    def SMTP_SERVER(self):
        """获取当前账户的SMTP服务器"""
        return self.EMAIL_ACCOUNTS[self.CURRENT_ACCOUNT]['SMTP_SERVER']
    
    @property
    def SMTP_PORT(self):
        """获取当前账户的SMTP端口"""
        return self.EMAIL_ACCOUNTS[self.CURRENT_ACCOUNT]['SMTP_PORT']
    
    @property
    def SENDER(self):
        """获取当前账户的发件人邮箱"""
        return self.EMAIL_ACCOUNTS[self.CURRENT_ACCOUNT]['SENDER']
    
    @property
    def PASSWORD(self):
        """获取当前账户的密码/授权码"""
        return self.EMAIL_ACCOUNTS[self.CURRENT_ACCOUNT]['PASSWORD']
    
    @property
    def SENDER_NAME(self):
        """获取当前账户的发件人名称"""
        return self.EMAIL_ACCOUNTS[self.CURRENT_ACCOUNT]['SENDER_NAME']
    
    @property
    def ATTACHMENTS(self):
        """从 GUI 获取附件列表"""
        return self._attachments
    
    def set_attachments(self, attachments):
        """设置附件列表（从 GUI 获取后调用）"""
        self._attachments = attachments
    
    def switch_account(self, account_name):
        """切换邮箱账户"""
        if account_name in self.EMAIL_ACCOUNTS:
            self.CURRENT_ACCOUNT = account_name
            return True
        return False
    
    def get_available_accounts(self):
        """获取所有可用的邮箱账户名称"""
        return list(self.EMAIL_ACCOUNTS.keys())