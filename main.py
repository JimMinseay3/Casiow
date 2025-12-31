import sys
import os

if __name__ == "__main__":
    # 将src目录添加到Python路径中
    sys.path.insert(0, os.path.join(os.path.dirname(__file__), 'src'))

    # 导入主程序
    from core import email_app

    # 运行应用程序
    email_app.main()