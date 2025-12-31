# Casiow 邮件发送系统

一个基于Python的批量邮件发送工具，支持通过GUI界面管理收件人和附件，可选择单封邮件包含所有附件或为每个附件单独发送邮件。

## 功能特点
- 图形化界面管理收件人和附件
- 支持保存常用收件人信息
- 两种发送模式：单封邮件包含所有附件/每个附件单独发送
- 邮件发送重试机制和错误处理
- 详细发送日志记录

## 安装说明

### 环境要求
- Python 3.8+ 
- Windows系统

### 安装步骤
1. 克隆仓库
```bash
git clone https://github.com/yourusername/casiow.git
cd casiow
```

2. 创建并激活虚拟环境
```bash
python -m venv .venv
.venv\Scripts\activate  # Windows系统
```

3. 安装依赖
```bash
pip install -r requirements.txt
```

## 使用方法
1. 运行主程序
```bash
python main.py
```

2. 在弹出的GUI界面中：
   - 点击"添加附件"选择需要发送的文件
   - 输入收件人邮箱、邮件主题和正文
   - 点击"添加收件人"将其加入发送列表
   - 点击"确认"开始发送邮件

3. 发送完成后会显示发送结果统计

## 项目结构
```
├── config.py          # 邮件配置信息
├── email_sender.py    # 邮件发送核心逻辑
├── recipient_window.py # GUI界面实现
├── main.py            # 程序入口
├── logger.py          # 日志配置
└── requirements.txt   # 项目依赖
```

## 注意事项
- 请确保SMTP服务器配置正确（在config.py中修改）
- 敏感信息（如邮箱密码）建议通过环境变量或配置文件管理
- 大批量发送邮件时请注意邮件服务商的发送限制
- recipients.xlsx文件用于批量导入收件人（需包含email, title, content列）

## 许可证
[MIT](LICENSE)