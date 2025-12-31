class EmailConfig:
    SMTP_SERVER = 'smtp.exmail.qq.com'
    SMTP_PORT = 465
    SENDER = 'jiaoyizhuliyanshengpinzh@ys-ics.net'
    PASSWORD = 'oCPQKusLvhQwnPZW'  # 替换为实际授权码
    SENDER_NAME = '交易助理-衍生品智库'
    SEND_INTERVAL = 2  # 秒
    
    @property
    def ATTACHMENTS(self):
        """从 GUI 获取附件列表"""
        if not hasattr(self, '_attachments'):
            return []
        return self._attachments
    
    def set_attachments(self, attachments):
        """设置附件列表（从 GUI 获取后调用）"""
        self._attachments = attachments