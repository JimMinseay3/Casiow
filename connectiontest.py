import smtplib

# 配置信息
smtp_server = 'smtp.exmail.qq.com'
smtp_port = 465
sender = 'jiaoyizhuliyanshengpinzh@ys-ics.net'
password = 'oCPQKusLvhQwnPZW'

try:
    with smtplib.SMTP_SSL(smtp_server, smtp_port) as server:
        server.login(sender, password)
        print("SMTP连接成功！授权码有效。")
except Exception as e:
    print(f"连接失败：{e}")
    print("请检查邮箱地址或授权码是否正确。")