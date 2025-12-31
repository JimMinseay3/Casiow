import logging
import os
from datetime import datetime

class DaySeparatorHandler(logging.FileHandler):
    def __init__(self, filename, mode='a', encoding=None, delay=False):
        super().__init__(filename, mode, encoding, delay)
        self.last_date = None
        self._check_and_add_separator()

    def _check_and_add_separator(self):
        """检查日期并在必要时添加分隔符"""
        current_date = datetime.now().strftime('%Y-%m-%d')
        
        # 如果是新的一天（或者第一次初始化）
        if self.last_date is None:
            # 尝试从文件读取最后一行日期
            if os.path.exists(self.baseFilename) and os.path.getsize(self.baseFilename) > 0:
                with open(self.baseFilename, 'rb') as f:
                    try:
                        f.seek(-100, os.SEEK_END)
                    except IOError:
                        pass
                    last_lines = f.read().decode('utf-8', errors='ignore')
                    if last_lines:
                        # 简单的日期匹配
                        import re
                        dates = re.findall(r'\d{4}-\d{2}-\d{2}', last_lines)
                        if dates:
                            self.last_date = dates[-1]

        if self.last_date and self.last_date != current_date:
            with open(self.baseFilename, 'a', encoding=self.encoding) as f:
                f.write('\n\n')
        
        self.last_date = current_date

    def emit(self, record):
        self._check_and_add_separator()
        super().emit(record)

def setup_logger():
    logger = logging.getLogger('email_sender')
    # 防止重复添加处理器
    if logger.handlers:
        return logger
        
    logger.setLevel(logging.INFO)
    
    # 确保 logs 目录存在
    log_dir = 'logs'
    if not os.path.exists(log_dir):
        os.makedirs(log_dir)
    
    log_path = os.path.join(log_dir, 'email_sender.log')
    
    # 使用自定义的带日期分隔的文件处理器
    file_handler = DaySeparatorHandler(log_path, encoding='utf-8')
    file_handler.setLevel(logging.INFO)
    
    # 创建控制台处理器
    console_handler = logging.StreamHandler()
    console_handler.setLevel(logging.INFO)
    
    # 创建格式化器并添加到处理器
    formatter = logging.Formatter('%(asctime)s - %(levelname)s - %(message)s')
    file_handler.setFormatter(formatter)
    console_handler.setFormatter(formatter)
    
    # 添加处理器到logger
    logger.addHandler(file_handler)
    logger.addHandler(console_handler)
    
    return logger    