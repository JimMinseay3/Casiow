import smtplib
import sys
import os

# 添加src目录到Python路径
sys.path.insert(0, os.path.join(os.path.dirname(__file__), '..', '..'))

from src.config.config import EmailConfig

def test_connection(account_name='default'):
    """测试指定账户的SMTP连接"""
    print(f"正在测试邮箱账户: {account_name}")
    
    # 初始化配置
    config = EmailConfig(account_name)
    
    try:
        print(f"SMTP服务器: {config.SMTP_SERVER}")
        print(f"SMTP端口: {config.SMTP_PORT}")
        print(f"发件人邮箱: {config.SENDER}")
        
        # 创建SMTP连接
        with smtplib.SMTP_SSL(config.SMTP_SERVER, config.SMTP_PORT) as server:
            # 登录
            server.login(config.SENDER, config.PASSWORD)
            print("✓ SMTP连接成功！授权码有效。")
            return True
            
    except Exception as e:
        print(f"✗ 连接失败：{e}")
        print("请检查邮箱地址或授权码是否正确。")
        return False

def main():
    """主函数"""
    # 测试默认账户
    print("测试默认邮箱账户:")
    print("=" * 40)
    test_connection('default')
    
    print("\n")
    
    # 测试华亮账户
    print("测试华亮邮箱账户:")
    print("=" * 40)
    test_connection('hualiang')

if __name__ == "__main__":
    main()