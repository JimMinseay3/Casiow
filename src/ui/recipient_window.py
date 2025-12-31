import tkinter as tk
from tkinter import ttk, messagebox, filedialog
import json
import os
from datetime import datetime
import _tkinter  # 添加导入以捕获 TclError
import re
from tkinter import simpledialog
import threading
import queue

# 导入EmailConfig
from src.config.config import EmailConfig
from src.config.account_manager import AccountManager




from src.utils.logger import setup_logger

class RecipientInputWindow:
    def __init__(self, root, email_account="default"):
        self.root = root
        self.root.title("Casiow")
        self.logger = setup_logger()
        
        # 设置窗口尺寸为1280x720
        self.root.geometry("1280x720")
        self.root.resizable(True, True)
        
        # 绑定窗口关闭事件（二次确认）
        self.root.protocol("WM_DELETE_WINDOW", self.on_close)
        
        self.recipients = []
        self.attachments = []  
        # 确保 data 目录存在
        data_dir = "data"
        if not os.path.exists(data_dir):
            os.makedirs(data_dir)
        self.storage_file = os.path.join(data_dir, "email_recipients.json")
        self.saved_recipients = {} # 更改为字典，存储分组信息
        self.send_completed = False  # 标记是否已发送过
        
        # 初始化文件队列，用于处理附件选择的异步操作
        self.file_queue = queue.Queue()
        
        # 初始化邮箱账户配置
        self.config = EmailConfig(email_account)
        
        # 初始化发送池
        from src.core.send_pool import SendPool
        self.send_pool = SendPool()
        
        self.create_widgets()
        self.load_saved_recipients()
        
        # 开始检查文件队列
        self.check_file_queue()

    def create_widgets(self):
        # 配置主窗口的网格布局
        self.root.grid_rowconfigure(0, weight=0)  # 邮箱账户选择行
        self.root.grid_rowconfigure(1, weight=1)  # 主内容区域
        self.root.grid_columnconfigure(0, weight=1)  # 主内容列
        
        # 添加邮箱账户选择框（放在最上方）
        frame_account = ttk.Frame(self.root)
        frame_account.grid(row=0, column=0, padx=10, pady=10, sticky="ew")
        frame_account.grid_columnconfigure(2, weight=1)  # 让组合框可以拉伸
        
        ttk.Label(frame_account, text="选择邮箱账户:", font=("Arial", 10, "bold")).grid(row=0, column=0, padx=(0, 10), sticky="w")
        
        # 获取可用账户列表
        self.account_manager = AccountManager()
        self.available_accounts = self.account_manager.get_account_names()
        
        self.account_var = tk.StringVar(value=self.available_accounts[0] if self.available_accounts else "default")
        self.account_combo = ttk.Combobox(frame_account, textvariable=self.account_var, values=self.available_accounts, state="readonly", width=20)
        self.account_combo.grid(row=0, column=1, padx=(0, 10), sticky="w")
        self.account_combo.bind("<<ComboboxSelected>>", self.on_account_change)
        
        # 账户管理按钮
        ttk.Button(frame_account, text="管理账户", command=self.manage_accounts).grid(row=0, column=3, padx=(10, 0), sticky="w")
        
        # 新建实例按钮

        
        # 占位标签，用于撑开空间
        ttk.Label(frame_account, text="").grid(row=0, column=2, sticky="ew")
        
        # 主内容区域从第1行开始
        
        # 创建主内容框架（使用Notebook标签页）
        self.notebook = ttk.Notebook(self.root)
        self.notebook.grid(row=1, column=0, sticky="nsew", padx=10, pady=(0, 10))
        
        # 创建主内容页面
        self.main_frame = ttk.Frame(self.notebook)
        self.notebook.add(self.main_frame, text="邮件发送")
        
        # 创建发送池状态页面
        self.pool_frame = ttk.Frame(self.notebook)
        self.notebook.add(self.pool_frame, text="发送池状态")
        
        # 创建发送池状态界面组件
        self.create_pool_status_widgets()
        
        self.main_frame.grid_rowconfigure(0, weight=1)
        self.main_frame.grid_columnconfigure(0, weight=2)  # 联系人管理区域
        self.main_frame.grid_columnconfigure(1, weight=1)  # 附件管理区域
        self.main_frame.grid_columnconfigure(2, weight=3)  # 邮件内容区域
        
        # 左侧联系人管理区域
        contact_frame = ttk.LabelFrame(self.main_frame, text="联系人管理")
        contact_frame.grid(row=0, column=0, sticky="nsew", padx=(0, 5), pady=0)
        contact_frame.grid_rowconfigure(2, weight=1)  # 让树形视图可以拉伸
        contact_frame.grid_columnconfigure(0, weight=1)
        contact_frame.grid_columnconfigure(1, weight=0)  # 滚动条列不拉伸
        
        # 搜索框
        search_frame = ttk.Frame(contact_frame)
        search_frame.grid(row=0, column=0, columnspan=2, sticky="ew", padx=5, pady=(5, 0))
        search_frame.grid_columnconfigure(1, weight=1)
        
        ttk.Label(search_frame, text="搜索:").grid(row=0, column=0, sticky="w", padx=(0, 5), pady=2)
        self.search_var = tk.StringVar()
        self.search_entry = ttk.Entry(search_frame, textvariable=self.search_var, width=30)
        self.search_entry.grid(row=0, column=1, sticky="ew", padx=(0, 5), pady=2)
        self.search_var.trace_add("write", self.on_search_change)
        
        # 按钮框架
        contact_button_frame = ttk.Frame(contact_frame)
        contact_button_frame.grid(row=1, column=0, columnspan=2, sticky="ew", padx=5, pady=0)
        contact_button_frame.grid_columnconfigure(0, weight=1)
        contact_button_frame.grid_columnconfigure(1, weight=1)
        contact_button_frame.grid_columnconfigure(2, weight=1)
        
        # 新建分组按钮
        ttk.Button(contact_button_frame, text="新建分组", command=self.create_new_group).grid(row=0, column=0, sticky="ew", padx=(0, 2), pady=0)
        
        # 新建联系人按钮
        ttk.Button(contact_button_frame, text="新建联系人", command=self.create_new_recipient).grid(row=0, column=1, sticky="ew", padx=(2, 2), pady=0)
        
        # 删除选中按钮
        ttk.Button(contact_button_frame, text="删除选中", command=self.delete_saved_recipient).grid(row=0, column=2, sticky="ew", padx=(2, 0), pady=0)
        
        # 联系人树形视图
        self.contact_tree = ttk.Treeview(contact_frame)
        self.contact_tree["columns"] = ("邮箱", "备注")
        self.contact_tree.column("#0", width=150, minwidth=150)
        self.contact_tree.column("邮箱", width=150, minwidth=150)
        self.contact_tree.column("备注", width=100, minwidth=100)
        
        self.contact_tree.heading("#0", text="分组")
        self.contact_tree.heading("邮箱", text="邮箱")
        self.contact_tree.heading("备注", text="备注")
        
        # 绑定分组列标题点击事件
        self.contact_tree.bind("<Button-1>", self.on_group_header_click)
        
        contact_tree_scrollbar = ttk.Scrollbar(contact_frame, orient="vertical", command=self.contact_tree.yview)
        self.contact_tree.configure(yscrollcommand=contact_tree_scrollbar.set)
        self.contact_tree.grid(row=2, column=0, sticky="nsew", padx=(5, 0), pady=0)
        contact_tree_scrollbar.grid(row=2, column=1, sticky="ns", padx=(0, 5), pady=0)
        
        self.contact_tree.bind("<<TreeviewSelect>>", self.on_contact_tree_select)
        self.contact_tree.bind("<Double-1>", self.on_contact_tree_double_click)
        self.contact_tree.bind("<Button-3>", self.on_contact_tree_right_click)  # 右键菜单
        
        # 中间附件管理区域
        attachment_frame = ttk.LabelFrame(self.main_frame, text="附件管理")
        attachment_frame.grid(row=0, column=1, sticky="nsew", padx=(0, 5), pady=0)
        attachment_frame.grid_rowconfigure(1, weight=1)
        attachment_frame.grid_columnconfigure(0, weight=1)
        attachment_frame.grid_columnconfigure(1, weight=0)  # 滚动条列不拉伸
        
        # 添加附件按钮
        ttk.Button(attachment_frame, text="添加附件", command=self.add_attachments).grid(row=0, column=0, columnspan=2, sticky="ew", padx=5, pady=5)
        
        # 附件预览列表
        self.attach_preview = tk.Listbox(attachment_frame)
        attach_preview_scrollbar = ttk.Scrollbar(attachment_frame, orient="vertical", command=self.attach_preview.yview)
        self.attach_preview.configure(yscrollcommand=attach_preview_scrollbar.set)
        self.attach_preview.grid(row=1, column=0, sticky="nsew", padx=(5, 0), pady=5)
        attach_preview_scrollbar.grid(row=1, column=1, sticky="ns", padx=(0, 5), pady=5)
        
        # 附件统计标签
        self.attach_label = ttk.Label(attachment_frame, text="已选附件: 0 个")
        self.attach_label.grid(row=2, column=0, columnspan=2, sticky="w", padx=5, pady=(0, 5))
        
        # 右侧邮件内容区域
        email_frame = ttk.LabelFrame(self.main_frame, text="邮件内容")
        email_frame.grid(row=0, column=2, sticky="nsew", padx=(0, 5), pady=0)
        email_frame.grid_rowconfigure(2, weight=1)  # 内容输入区域可以拉伸
        email_frame.grid_rowconfigure(4, weight=1)  # 待发送列表可以拉伸
        email_frame.grid_columnconfigure(1, weight=1)
        email_frame.grid_columnconfigure(2, weight=0)  # 滚动条列不拉伸
        
        # 收件人输入
        ttk.Label(email_frame, text="收件人:").grid(row=0, column=0, sticky="w", padx=(5, 0), pady=5)
        self.email_var = tk.StringVar()
        self.email_entry = ttk.Entry(email_frame, textvariable=self.email_var, width=50)
        self.email_entry.grid(row=0, column=1, columnspan=2, sticky="ew", padx=(0, 5), pady=5)
        
        # 主题输入
        ttk.Label(email_frame, text="主题:").grid(row=1, column=0, sticky="w", padx=(5, 0), pady=5)
        self.title_var = tk.StringVar()
        self.title_entry = ttk.Entry(email_frame, textvariable=self.title_var, width=50)
        self.title_entry.grid(row=1, column=1, columnspan=2, sticky="ew", padx=(0, 5), pady=5)
        
        # 内容输入
        ttk.Label(email_frame, text="内容:").grid(row=2, column=0, sticky="nw", padx=(5, 0), pady=5)
        self.content_text = tk.Text(email_frame, height=10, width=50)
        content_scrollbar = ttk.Scrollbar(email_frame, orient="vertical", command=self.content_text.yview)
        self.content_text.configure(yscrollcommand=content_scrollbar.set)
        self.content_text.grid(row=2, column=1, sticky="nsew", padx=(0, 0), pady=5)
        content_scrollbar.grid(row=2, column=2, sticky="ns", padx=(0, 5), pady=5)
        
        # 操作按钮
        button_frame = ttk.Frame(email_frame)
        button_frame.grid(row=3, column=0, columnspan=3, sticky="ew", padx=5, pady=5)
        button_frame.grid_columnconfigure(0, weight=1)
        button_frame.grid_columnconfigure(1, weight=1)
        button_frame.grid_columnconfigure(2, weight=1)
        
        ttk.Button(button_frame, text="清空内容", command=self.clear_content).grid(row=0, column=0, padx=5, pady=5)
        ttk.Button(button_frame, text="添加到已保存", command=self.add_recipient_to_saved_list).grid(row=0, column=1, padx=5, pady=5)
        ttk.Button(button_frame, text="添加到待发送", command=self.add_recipient_to_send_list).grid(row=0, column=2, padx=5, pady=5)
        
        # 待发送收件人列表
        recipient_frame = ttk.LabelFrame(email_frame, text="待发送收件人列表")
        recipient_frame.grid(row=4, column=0, columnspan=3, sticky="nsew", padx=5, pady=(0, 5))
        recipient_frame.grid_rowconfigure(0, weight=1)
        recipient_frame.grid_columnconfigure(0, weight=1)
        recipient_frame.grid_columnconfigure(1, weight=0)  # 滚动条列不拉伸
        
        self.recipient_tree = ttk.Treeview(recipient_frame, columns=("邮箱", "主题"), show="headings")
        self.recipient_tree.heading("邮箱", text="邮箱")
        self.recipient_tree.heading("主题", text="主题")
        self.recipient_tree.column("邮箱", width=200)
        self.recipient_tree.column("主题", width=200)
        
        recipient_tree_scrollbar = ttk.Scrollbar(recipient_frame, orient="vertical", command=self.recipient_tree.yview)
        self.recipient_tree.configure(yscrollcommand=recipient_tree_scrollbar.set)
        self.recipient_tree.grid(row=0, column=0, sticky="nsew", padx=(5, 0), pady=5)
        recipient_tree_scrollbar.grid(row=0, column=1, sticky="ns", padx=(0, 5), pady=5)
        
        # 删除选中按钮
        ttk.Button(recipient_frame, text="删除选中", command=self.delete_selected).grid(row=1, column=0, sticky="w", padx=5, pady=(0, 5))
        
        # 底部按钮区域
        bottom_frame = ttk.Frame(self.root)
        bottom_frame.grid(row=2, column=0, sticky="ew", padx=10, pady=10)
        bottom_frame.grid_columnconfigure(0, weight=1)
        bottom_frame.grid_columnconfigure(1, weight=1)
        
        ttk.Button(bottom_frame, text="确认并发送", command=self.confirm_and_close).grid(row=0, column=0, padx=5, pady=5)
        ttk.Button(bottom_frame, text="取消", command=self.on_close).grid(row=0, column=1, padx=5, pady=5)

    def create_pool_status_widgets(self):
        """创建发送池状态界面"""
        # 配置发送池框架的网格布局
        self.pool_frame.grid_rowconfigure(0, weight=0)  # 控制按钮行
        self.pool_frame.grid_rowconfigure(1, weight=0)  # 筛选器行
        self.pool_frame.grid_rowconfigure(2, weight=1)  # 任务列表行
        self.pool_frame.grid_columnconfigure(0, weight=1)  # 主内容列
        
        # 控制按钮区域
        control_frame = ttk.Frame(self.pool_frame)
        control_frame.grid(row=0, column=0, sticky="ew", padx=10, pady=10)
        control_frame.grid_columnconfigure(3, weight=1)  # 让中间空白区域可以拉伸
        
        ttk.Button(control_frame, text="刷新", command=self.refresh_pool_status).pack(side=tk.LEFT, padx=(0, 5))
        ttk.Button(control_frame, text="清空已完成", command=self.clear_completed_tasks).pack(side=tk.LEFT, padx=(0, 5))
        ttk.Button(control_frame, text="全部清除", command=self.clear_all_tasks).pack(side=tk.LEFT, padx=(0, 5))
        
        # 添加导出按钮
        ttk.Button(control_frame, text="导出任务", command=self.export_tasks).pack(side=tk.LEFT, padx=(0, 5))
        
        # 任务统计标签
        self.pool_stats_label = ttk.Label(control_frame, text="任务总数: 0 | 待处理: 0 | 发送中: 0 | 已完成: 0 | 失败: 0")
        self.pool_stats_label.pack(side=tk.RIGHT)
        
        # 筛选器区域
        filter_frame = ttk.Frame(self.pool_frame)
        filter_frame.grid(row=1, column=0, sticky="ew", padx=10, pady=(0, 5))
        
        ttk.Label(filter_frame, text="状态筛选:").pack(side=tk.LEFT, padx=(0, 5))
        self.filter_var = tk.StringVar(value="all")
        filter_combo = ttk.Combobox(filter_frame, textvariable=self.filter_var, 
                                   values=["all", "pending", "sending", "completed", "failed"],
                                   state="readonly", width=10)
        filter_combo.pack(side=tk.LEFT, padx=(0, 10))
        filter_combo.bind("<<ComboboxSelected>>", lambda e: self.refresh_pool_status())
        
        # 任务列表区域
        list_frame = ttk.LabelFrame(self.pool_frame, text="发送任务列表")
        list_frame.grid(row=2, column=0, sticky="nsew", padx=10, pady=(0, 10))
        list_frame.grid_rowconfigure(0, weight=1)
        list_frame.grid_columnconfigure(0, weight=1)
        list_frame.grid_columnconfigure(1, weight=0)  # 滚动条列不拉伸
        
        # 创建任务列表Treeview
        self.pool_tree = ttk.Treeview(list_frame, columns=("ID", "状态", "收件人数", "附件数", "创建时间", "开始时间", "结束时间", "成功数", "失败数"), show="headings")
        
        # 设置列标题和宽度
        self.pool_tree.heading("ID", text="任务ID")
        self.pool_tree.heading("状态", text="状态")
        self.pool_tree.heading("收件人数", text="收件人数")
        self.pool_tree.heading("附件数", text="附件数")
        self.pool_tree.heading("创建时间", text="创建时间")
        self.pool_tree.heading("开始时间", text="开始时间")
        self.pool_tree.heading("结束时间", text="结束时间")
        self.pool_tree.heading("成功数", text="成功数")
        self.pool_tree.heading("失败数", text="失败数")
        
        # 设置列宽和对齐方式
        self.pool_tree.column("ID", width=100, anchor="center")
        self.pool_tree.column("状态", width=80, anchor="center")
        self.pool_tree.column("收件人数", width=80, anchor="center")
        self.pool_tree.column("附件数", width=80, anchor="center")
        self.pool_tree.column("创建时间", width=150, anchor="center")
        self.pool_tree.column("开始时间", width=150, anchor="center")
        self.pool_tree.column("结束时间", width=150, anchor="center")
        self.pool_tree.column("成功数", width=80, anchor="center")
        self.pool_tree.column("失败数", width=80, anchor="center")
        
        # 添加滚动条
        pool_tree_scrollbar = ttk.Scrollbar(list_frame, orient="vertical", command=self.pool_tree.yview)
        self.pool_tree.configure(yscrollcommand=pool_tree_scrollbar.set)
        self.pool_tree.grid(row=0, column=0, sticky="nsew", padx=(5, 0), pady=5)
        pool_tree_scrollbar.grid(row=0, column=1, sticky="ns", padx=(0, 5), pady=5)
        
        # 双击事件绑定
        self.pool_tree.bind("<Double-1>", self.on_pool_tree_double_click)
        
        # 右键菜单
        self.pool_tree.bind("<Button-3>", self.on_pool_tree_right_click)
        
        # 初始刷新
        self.refresh_pool_status()

    def refresh_pool_status(self):
        """刷新发送池状态"""
        try:
            # 获取所有任务
            items = self.send_pool.get_all_items()
            
            # 获取筛选条件
            filter_status = self.filter_var.get()
            
            # 清空现有数据
            for item in self.pool_tree.get_children():
                self.pool_tree.delete(item)
            
            # 统计变量
            total_count = len(items)
            pending_count = 0
            sending_count = 0
            completed_count = 0
            failed_count = 0
            
            # 添加任务到列表
            for item in items:
                # 更新统计
                if item.status == "pending":
                    pending_count += 1
                elif item.status == "sending":
                    sending_count += 1
                elif item.status == "completed":
                    completed_count += 1
                elif item.status == "failed":
                    failed_count += 1
                
                # 应用筛选条件
                if filter_status != "all" and item.status != filter_status:
                    continue
                
                # 格式化时间
                created_time = item.created_time.strftime("%Y-%m-%d %H:%M:%S") if item.created_time else ""
                start_time = item.start_time.strftime("%Y-%m-%d %H:%M:%S") if item.start_time else ""
                end_time = item.end_time.strftime("%Y-%m-%d %H:%M:%S") if item.end_time else ""
                
                # 插入到Treeview
                self.pool_tree.insert("", tk.END, values=(
                    item.id,
                    item.status,
                    len(item.recipients),
                    len(item.attachments),
                    created_time,
                    start_time,
                    end_time,
                    item.success_count,
                    item.fail_count
                ))
            
            # 更新统计标签
            self.pool_stats_label.config(text=f"任务总数: {total_count} | 待处理: {pending_count} | 发送中: {sending_count} | 已完成: {completed_count} | 失败: {failed_count}")
            
        except Exception as e:
            print(f"刷新发送池状态失败: {e}")

    def export_tasks(self):
        """导出任务列表到文件"""
        try:
            import json
            from datetime import datetime
            
            # 获取所有任务
            items = self.send_pool.get_all_items()
            
            # 转换为可序列化的格式
            export_data = []
            for item in items:
                export_data.append({
                    "id": item.id,
                    "status": item.status,
                    "recipients_count": len(item.recipients),
                    "attachments_count": len(item.attachments),
                    "created_time": item.created_time.isoformat() if item.created_time else None,
                    "start_time": item.start_time.isoformat() if item.start_time else None,
                    "end_time": item.end_time.isoformat() if item.end_time else None,
                    "success_count": item.success_count,
                    "fail_count": item.fail_count
                })
            
            # 生成文件名
            timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
            data_dir = "data"
            if not os.path.exists(data_dir):
                os.makedirs(data_dir)
            filename = os.path.join(data_dir, f"send_pool_tasks_{timestamp}.json")
            
            # 写入文件
            with open(filename, 'w', encoding='utf-8') as f:
                json.dump(export_data, f, ensure_ascii=False, indent=2)
            
            messagebox.showinfo("导出成功", f"任务列表已导出到文件: {filename}")
            
        except Exception as e:
            messagebox.showerror("导出失败", f"导出任务列表失败: {e}")

    def clear_completed_tasks(self):
        """清空已完成的任务"""
        try:
            items = self.send_pool.get_all_items()
            removed_count = 0
            
            for item in items:
                if item.status in ["completed", "failed"]:
                    if self.send_pool.remove_item(item.id):
                        removed_count += 1
            
            # 刷新显示
            self.refresh_pool_status()
            messagebox.showinfo("提示", f"已清除 {removed_count} 个已完成/失败的任务")
            
        except Exception as e:
            messagebox.showerror("错误", f"清空任务失败: {e}")

    def clear_all_tasks(self):
        """清空所有任务"""
        try:
            if messagebox.askyesno("确认", "确定要清空所有发送任务吗？这将删除所有待处理、发送中、已完成和失败的任务。"):
                items = self.send_pool.get_all_items()
                removed_count = 0
                
                for item in items:
                    if self.send_pool.remove_item(item.id):
                        removed_count += 1
                
                # 刷新显示
                self.refresh_pool_status()
                messagebox.showinfo("提示", f"已清除 {removed_count} 个任务")
                
        except Exception as e:
            messagebox.showerror("错误", f"清空任务失败: {e}")

    def on_pool_tree_double_click(self, event):
        """处理发送池任务列表的双击事件"""
        selected_item = self.pool_tree.focus()
        if not selected_item:
            return
            
        # 获取选中项的值
        values = self.pool_tree.item(selected_item, 'values')
        if not values:
            return
            
        task_id = values[0]
        self.show_task_detail(task_id)

    def on_pool_tree_right_click(self, event):
        """处理发送池任务列表的右键点击事件"""
        # 选择被点击的项目
        item = self.pool_tree.identify_row(event.y)
        if item:
            self.pool_tree.selection_set(item)
            
        selected_item = self.pool_tree.focus()
        if not selected_item:
            return
            
        # 获取选中项的值
        values = self.pool_tree.item(selected_item, 'values')
        if not values:
            return
            
        task_id = values[0]
        
        # 创建右键菜单
        context_menu = tk.Menu(self.root, tearoff=0)
        context_menu.add_command(label="查看详情", command=lambda: self.show_task_detail(task_id))
        context_menu.add_command(label="删除任务", command=lambda: self.delete_task(task_id))
        
        # 显示菜单
        context_menu.post(event.x_root, event.y_root)

    def show_task_detail(self, task_id):
        """显示任务详情"""
        try:
            item = self.send_pool.get_item(int(task_id))
            if not item:
                messagebox.showwarning("警告", "任务不存在或已被删除")
                return
            
            # 创建详情对话框
            detail_window = tk.Toplevel(self.root)
            detail_window.title(f"任务详情 - {task_id}")
            detail_window.geometry("800x600")
            detail_window.resizable(True, True)
            detail_window.grab_set()  # 模态对话框
            
            # 居中显示
            detail_window.transient(self.root)
            detail_window.update_idletasks()
            x = (detail_window.winfo_screenwidth() // 2) - (detail_window.winfo_width() // 2)
            y = (detail_window.winfo_screenheight() // 2) - (detail_window.winfo_height() // 2)
            detail_window.geometry(f"+{x}+{y}")
            
            # 创建笔记本控件用于标签页
            notebook = ttk.Notebook(detail_window)
            notebook.pack(fill=tk.BOTH, expand=True, padx=10, pady=10)
            
            # 基本信息标签页
            basic_frame = ttk.Frame(notebook)
            notebook.add(basic_frame, text="基本信息")
            
            basic_text = tk.Text(basic_frame, wrap=tk.WORD)
            basic_scrollbar = ttk.Scrollbar(basic_frame, orient="vertical", command=basic_text.yview)
            basic_text.configure(yscrollcommand=basic_scrollbar.set)
            
            basic_text.grid(row=0, column=0, sticky="nsew")
            basic_scrollbar.grid(row=0, column=1, sticky="ns")
            basic_frame.grid_rowconfigure(0, weight=1)
            basic_frame.grid_columnconfigure(0, weight=1)
            
            # 计算成功率和失败率
            total_recipients = len(item.recipients)
            success_rate = round(item.success_count / total_recipients * 100, 2) if total_recipients > 0 else 0
            fail_rate = round(item.fail_count / total_recipients * 100, 2) if total_recipients > 0 else 0
            
            # 构建基本详细信息文本
            basic_details = f"""任务ID: {item.id}
状态: {item.status}
创建时间: {item.created_time.strftime('%Y-%m-%d %H:%M:%S') if item.created_time else ''}
开始时间: {item.start_time.strftime('%Y-%m-%d %H:%M:%S') if item.start_time else ''}
结束时间: {item.end_time.strftime('%Y-%m-%d %H:%M:%S') if item.end_time else ''}
成功数量: {item.success_count}
失败数量: {item.fail_count}
成功率: {success_rate}%
失败率: {fail_rate}%

收件人总数: {total_recipients}
附件总数: {len(item.attachments)}
"""
            
            # 插入基本信息文本
            basic_text.insert(tk.END, basic_details)
            basic_text.config(state=tk.DISABLED)  # 设置为只读
            
            # 收件人列表标签页
            recipients_frame = ttk.Frame(notebook)
            notebook.add(recipients_frame, text=f"收件人列表 ({len(item.recipients)})")
            
            recipients_tree = ttk.Treeview(recipients_frame, columns=("邮箱", "主题"), show="headings")
            recipients_tree.heading("邮箱", text="邮箱")
            recipients_tree.heading("主题", text="主题")
            recipients_tree.column("邮箱", width=300)
            recipients_tree.column("主题", width=300)
            
            recipients_scrollbar = ttk.Scrollbar(recipients_frame, orient="vertical", command=recipients_tree.yview)
            recipients_tree.configure(yscrollcommand=recipients_scrollbar.set)
            recipients_tree.grid(row=0, column=0, sticky="nsew")
            recipients_scrollbar.grid(row=0, column=1, sticky="ns")
            recipients_frame.grid_rowconfigure(0, weight=1)
            recipients_frame.grid_columnconfigure(0, weight=1)
            
            # 添加收件人数据
            for recipient in item.recipients:
                recipients_tree.insert("", tk.END, values=(
                    recipient.get('email', ''),
                    recipient.get('title', '')
                ))
            
            # 附件列表标签页
            attachments_frame = ttk.Frame(notebook)
            notebook.add(attachments_frame, text=f"附件列表 ({len(item.attachments)})")
            
            attachments_listbox = tk.Listbox(attachments_frame)
            attachments_scrollbar = ttk.Scrollbar(attachments_frame, orient="vertical", command=attachments_listbox.yview)
            attachments_listbox.configure(yscrollcommand=attachments_scrollbar.set)
            attachments_listbox.grid(row=0, column=0, sticky="nsew")
            attachments_scrollbar.grid(row=0, column=1, sticky="ns")
            attachments_frame.grid_rowconfigure(0, weight=1)
            attachments_frame.grid_columnconfigure(0, weight=1)
            
            # 添加附件数据
            for attachment in item.attachments:
                attachments_listbox.insert(tk.END, attachment)
            
            # 错误信息标签页（如果有错误）
            if item.error_messages:
                errors_frame = ttk.Frame(notebook)
                notebook.add(errors_frame, text=f"错误信息 ({len(item.error_messages)})")
                
                # 创建Treeview来显示错误信息
                errors_tree = ttk.Treeview(errors_frame, columns=("收件人", "错误信息", "时间"), show="headings")
                errors_tree.heading("收件人", text="收件人")
                errors_tree.heading("错误信息", text="错误信息")
                errors_tree.heading("时间", text="时间")
                errors_tree.column("收件人", width=200)
                errors_tree.column("错误信息", width=300)
                errors_tree.column("时间", width=150)
                
                errors_scrollbar = ttk.Scrollbar(errors_frame, orient="vertical", command=errors_tree.yview)
                errors_tree.configure(yscrollcommand=errors_scrollbar.set)
                errors_tree.grid(row=0, column=0, sticky="nsew")
                errors_scrollbar.grid(row=0, column=1, sticky="ns")
                errors_frame.grid_rowconfigure(0, weight=1)
                errors_frame.grid_columnconfigure(0, weight=1)
                
                # 添加错误信息到Treeview
                for error in item.error_messages:
                    # 解析时间戳
                    timestamp = error.get("timestamp", "")
                    if timestamp:
                        try:
                            # 尝试格式化时间戳
                            dt = datetime.fromisoformat(timestamp.replace("Z", "+00:00"))
                            formatted_time = dt.strftime("%Y-%m-%d %H:%M:%S")
                        except:
                            formatted_time = timestamp
                    else:
                        formatted_time = ""
                    
                    errors_tree.insert("", tk.END, values=(
                        error.get("recipient", ""),
                        error.get("error", ""),
                        formatted_time
                    ))
            
        except Exception as e:
            messagebox.showerror("错误", f"显示任务详情失败: {e}")

    def delete_task(self, task_id):
        """删除任务"""
        try:
            if messagebox.askyesno("确认删除", f"确定要删除任务 {task_id} 吗？"):
                if self.send_pool.remove_item(int(task_id)):
                    self.refresh_pool_status()
                    messagebox.showinfo("成功", f"任务 {task_id} 已删除")
                else:
                    messagebox.showerror("错误", f"删除任务 {task_id} 失败")
        except Exception as e:
            messagebox.showerror("错误", f"删除任务失败: {e}")

    def add_attachments(self):
        """添加附件（支持多选、去重、更新预览）"""
        try:
            # 使用tkinter标准文件选择器
            new_files = filedialog.askopenfilenames(
                title="选择附件",
                filetypes=[
                    ("所有文件", "*.*"),
                    ("文本文件", "*.txt"),
                    ("PDF文件", "*.pdf"),
                    ("图片文件", "*.jpg *.jpeg *.png *.gif"),
                    ("Word文档", "*.doc *.docx"),
                    ("Excel文件", "*.xls *.xlsx")
                ]
            )
            
            # 处理选中的文件
            self._process_selected_files(new_files)
        except Exception as e:
            messagebox.showerror("错误", f"选择文件时发生错误: {str(e)}")

    def check_file_queue(self):
        """检查文件队列并处理结果"""
        try:
            while True:
                item = self.file_queue.get_nowait()
                action, data = item
                
                if action == "files_selected":
                    self._process_selected_files(data)
                elif action == "error":
                    messagebox.showerror("错误", f"选择文件时发生错误: {data}")
        except queue.Empty:
            pass
        
        # 继续定期检查队列，但使用更高效的间隔
        self.root.after(50, self.check_file_queue)

    def _process_selected_files(self, new_files):
        """处理选中的文件"""
        if new_files:
            # 去重逻辑：用集合保证唯一性
            current_files = set(self.attachments)
            
            # 先收集所有有效的文件（使用列表推导式提高效率）
            valid_files = [file for file in new_files 
                          if file not in current_files 
                          and os.path.exists(file) 
                          and os.access(file, os.R_OK)]
            
            # 批量添加有效文件
            if valid_files:
                self.attachments.extend(valid_files)
                
                # 更新附件预览和统计标签（批量更新）
                self.update_attach_preview()
                self.attach_label.config(
                    text=f"已选附件: {len(self.attachments)} 个"
                )
            else:
                messagebox.showinfo("提示", "没有添加新的附件（可能已存在或无法访问）")
        else:
            # 用户取消选择，不做任何操作
            pass

    def update_attach_preview(self):
        """更新附件预览列表"""
        # 批量更新预览列表，避免频繁的UI操作
        self.attach_preview.delete(0, tk.END)
        # 预先计算所有文件名以减少UI线程工作量
        file_names = [os.path.basename(file) for file in self.attachments]
        # 批量插入文件名
        for name in file_names:
            self.attach_preview.insert(tk.END, name)

    def load_saved_recipients(self):
        """从本地 JSON 文件加载历史收件人"""
        try:
            if os.path.exists(self.storage_file):
                with open(self.storage_file, 'r', encoding='utf-8') as f:
                    saved_data = json.load(f)
                    self.saved_recipients = saved_data.get('groups', {}) # 从 'groups' 键加载
            else:
                # 如果文件不存在，则初始化为空字典
                self.saved_recipients = {}
            
            # 确保"未分组"分组存在
            if "未分组" not in self.saved_recipients:
                self.saved_recipients["未分组"] = []
                
            self.populate_contact_tree() # 调用新方法填充 Treeview
        except Exception as e:
            self.logger.error(f"加载保存的收件人失败: {e}")
            # 出现异常时确保"未分组"分组存在
            if "未分组" not in self.saved_recipients:
                self.saved_recipients["未分组"] = []
            self.populate_contact_tree()
            self.saved_recipients = []

    def save_recipient(self, email, title, content):
        """将收件人信息保存到本地 JSON 文件"""
        try:
            # 检查是否已存在该收件人，存在则更新，不存在则新增
            updated = False
            for i, recipient in enumerate(self.saved_recipients):
                if recipient['email'] == email:
                    self.saved_recipients[i] = {
                        'email': email,
                        'title': title,
                        'content': content,
                        'last_used': datetime.now().strftime("%Y-%m-%d %H:%M:%S")
                    }
                    updated = True
                    break
            if not updated:
                self.saved_recipients.append({
                    'email': email,
                    'title': title,
                    'content': content,
                    'last_used': datetime.now().strftime("%Y-%m-%d %H:%M:%S")
                })
            
            # 写入文件
            with open(self.storage_file, 'w', encoding='utf-8') as f:
                json.dump(
                    {'groups': self.saved_recipients}, # 更改为 'groups' 键
                    f,
                    ensure_ascii=False,
                    indent=2
                )
            self.populate_contact_tree() # 刷新 Treeview
            return True
        except Exception as e:
            self.logger.error(f"保存收件人失败: {e}")
            return False

    def populate_contact_tree(self, search_term=""):
        """填充联系人 Treeview，显示分组和联系人，支持搜索功能和分组排序"""
        for item in self.contact_tree.get_children():
            self.contact_tree.delete(item)

        # 获取所有分组名称并根据排序规则排序
        group_names = list(self.saved_recipients.keys())
        
        # 根据排序状态对分组进行排序
        if hasattr(self, 'group_sort_order'):
            if self.group_sort_order == 'asc':
                group_names.sort()
            elif self.group_sort_order == 'desc':
                group_names.sort(reverse=True)
            # 默认情况下保持原有顺序（按创建时间）

        # 如果有搜索词，则过滤联系人
        if search_term:
            # 创建一个临时的过滤后的数据结构
            filtered_recipients = {}
            for group_name, recipients in self.saved_recipients.items():
                filtered_recipients[group_name] = []
                for recipient in recipients:
                    # 检查邮箱或备注是否包含搜索词
                    if (search_term in recipient['email'].lower() or 
                        search_term in recipient['content'].lower()):
                        filtered_recipients[group_name].append(recipient)
            
            # 只显示包含匹配联系人的分组（按排序顺序）
            for group_name in group_names:
                if group_name in filtered_recipients:
                    recipients = filtered_recipients[group_name]
                    if recipients:  # 只有当分组中有匹配的联系人时才显示该分组
                        group_id = self.contact_tree.insert('', 'end', text=group_name, open=True, tags=('group',))
                        for recipient in recipients:
                            # 使用content作为备注字段
                            self.contact_tree.insert(group_id, 'end', text=recipient['email'], values=(recipient['email'], recipient['content']), tags=('recipient',))
        else:
            # 无搜索词时显示所有联系人（按排序顺序）
            for group_name in group_names:
                if group_name in self.saved_recipients:
                    recipients = self.saved_recipients[group_name]
                    group_id = self.contact_tree.insert('', 'end', text=group_name, open=True, tags=('group',))
                    for recipient in recipients:
                        # 使用content作为备注字段
                        self.contact_tree.insert(group_id, 'end', text=recipient['email'], values=(recipient['email'], recipient['content']), tags=('recipient',))

    def on_contact_tree_select(self, event):
        """处理联系人 Treeview 的选择事件"""
        selected_item = self.contact_tree.focus()
        if selected_item:
            item_tags = self.contact_tree.item(selected_item, 'tags')
            if 'recipient' in item_tags:
                values = self.contact_tree.item(selected_item, 'values')
                if values:
                    self.email_var.set(values[0])
                    # 不再自动设置主题和内容输入框
                    self.content_text.delete(1.0, tk.END)
            elif 'group' in item_tags:
                # 如果选择的是分组，清空输入框
                self.email_var.set("")
                self.title_var.set("")
                self.content_text.delete(1.0, tk.END)

    def on_contact_tree_double_click(self, event):
        """处理联系人 Treeview 的双击事件，实现双击备注可修改功能"""
        selected_item = self.contact_tree.focus()
        if selected_item:
            item_tags = self.contact_tree.item(selected_item, 'tags')
            if 'recipient' in item_tags:
                # 获取当前选中行的值
                values = self.contact_tree.item(selected_item, 'values')
                if values:
                    # 弹出输入框让用户修改备注
                    current_note = values[1]  # 备注在第二列
                    new_note = simpledialog.askstring("修改备注", "请输入新的备注:", initialvalue=current_note, parent=self.root)
                    
                    if new_note is not None:  # 用户点击了确定按钮
                        # 更新联系人数据
                        email = self.contact_tree.item(selected_item, 'text')
                        parent_id = self.contact_tree.parent(selected_item)
                        group_name = self.contact_tree.item(parent_id, 'text') if parent_id else "未分组"
                        
                        # 在保存的数据中找到对应的联系人并更新备注
                        if group_name in self.saved_recipients:
                            for i, recipient in enumerate(self.saved_recipients[group_name]):
                                if recipient['email'] == email:
                                    self.saved_recipients[group_name][i]['content'] = new_note
                                    self.saved_recipients[group_name][i]['title'] = new_note
                                    break
                        
                        # 保存到文件并刷新显示
                        self.save_recipient_data()
                        self.populate_contact_tree()
                        
                        # 如果当前选中的联系人就是被修改的联系人，更新输入框
                        if self.email_var.get() == email:
                            self.title_var.set(new_note)
                            self.content_text.delete(1.0, tk.END)
                            self.content_text.insert(tk.END, new_note)

    def on_contact_tree_right_click(self, event):
        """处理联系人 Treeview 的右键点击事件，显示右键菜单"""
        # 选择被点击的项目
        item = self.contact_tree.identify_row(event.y)
        if item:
            self.contact_tree.selection_set(item)
            
        selected_item = self.contact_tree.focus()
        if not selected_item:
            return
            
        item_tags = self.contact_tree.item(selected_item, 'tags')
        
        # 创建右键菜单
        context_menu = tk.Menu(self.root, tearoff=0)
        
        if 'recipient' in item_tags:
            # 联系人右键菜单
            context_menu.add_command(label="修改备注", command=lambda: self.on_contact_tree_double_click(event))
            
            # 添加"移动到"子菜单
            move_menu = tk.Menu(context_menu, tearoff=0)
            # 获取所有分组名称（除了当前分组）
            current_parent = self.contact_tree.parent(selected_item)
            current_group = self.contact_tree.item(current_parent, 'text') if current_parent else "未分组"
            
            for group_name in self.saved_recipients.keys():
                if group_name != current_group:
                    move_menu.add_command(label=group_name, command=lambda g=group_name: self.move_recipient_to_group(selected_item, g))
            
            context_menu.add_cascade(label="移动到", menu=move_menu)
            context_menu.add_separator()
            context_menu.add_command(label="删除", command=self.delete_saved_recipient)
        elif 'group' in item_tags:
            # 分组右键菜单
            context_menu.add_command(label="重命名分组", command=lambda: self.rename_group(selected_item))
            context_menu.add_command(label="删除分组", command=self.delete_saved_recipient)
        
        # 显示菜单
        try:
            context_menu.post(event.x_root, event.y_root)
        except:
            pass

    def move_recipient_to_group(self, recipient_item, target_group):
        """将联系人移动到指定分组"""
        # 获取联系人信息
        email = self.contact_tree.item(recipient_item, 'text')
        values = self.contact_tree.item(recipient_item, 'values')
        
        if not values:
            return
            
        # 获取源分组
        source_parent = self.contact_tree.parent(recipient_item)
        source_group = self.contact_tree.item(source_parent, 'text') if source_parent else "未分组"
        
        # 确保目标分组存在
        if target_group not in self.saved_recipients:
            self.saved_recipients[target_group] = []
        
        # 从源分组移除
        if source_group in self.saved_recipients:
            self.saved_recipients[source_group] = [
                r for r in self.saved_recipients[source_group] 
                if r['email'] != email
            ]
        
        # 创建新的联系人数据
        new_recipient = {
            'email': email,
            'title': values[1],  # 备注
            'content': values[1],  # 备注
            'last_used': datetime.now().strftime("%Y-%m-%d %H:%M:%S")
        }
        
        # 添加到目标分组
        self.saved_recipients[target_group].append(new_recipient)
        
        # 保存并刷新
        self.save_recipient_data()
        self.populate_contact_tree()
        messagebox.showinfo("成功", f"联系人已移动到 '{target_group}' 分组！")

    def rename_group(self, group_item):
        """重命名分组"""
        current_name = self.contact_tree.item(group_item, 'text')
        
        # 不能重命名"未分组"
        if current_name == "未分组":
            messagebox.showwarning("警告", "不能重命名'未分组'！")
            return
        
        new_name = simpledialog.askstring("重命名分组", "请输入新的分组名称:", initialvalue=current_name, parent=self.root)
        if not new_name or new_name == current_name:
            return
            
        new_name = new_name.strip()
        if not new_name:
            messagebox.showwarning("警告", "分组名称不能为空！")
            return
            
        if new_name in self.saved_recipients:
            messagebox.showwarning("警告", f"分组 '{new_name}' 已存在！")
            return
        
        # 更新数据结构
        self.saved_recipients[new_name] = self.saved_recipients.pop(current_name)
        
        # 保存并刷新
        self.save_recipient_data()
        self.populate_contact_tree()
        messagebox.showinfo("成功", f"分组已重命名为 '{new_name}'！")

    def on_search_change(self, *args):
        """处理搜索框内容变化事件，实现动态搜索功能"""
        search_term = self.search_var.get().lower().strip()
        self.populate_contact_tree(search_term)

    def create_new_group(self):
        """创建新的分组"""
        # 弹出输入框让用户输入新的分组名
        group_name = simpledialog.askstring("新建分组", "请输入新的分组名:", parent=self.root)
        
        if group_name:
            group_name = group_name.strip()
            if group_name and group_name not in self.saved_recipients:
                # 添加新的分组到数据结构中
                self.saved_recipients[group_name] = []
                # 写入文件
                self.save_recipient_data()
                # 刷新联系人显示
                self.populate_contact_tree()
                messagebox.showinfo("成功", f"新分组 '{group_name}' 创建成功！")
            elif group_name in self.saved_recipients:
                messagebox.showwarning("警告", f"分组 '{group_name}' 已存在！")
            else:
                messagebox.showwarning("警告", "分组名不能为空！")

    def on_group_header_click(self, event):
        """处理分组列标题点击事件，实现排序切换"""
        # 确定点击的是哪个列
        region = self.contact_tree.identify("region", event.x, event.y)
        if region == "heading":
            column = self.contact_tree.identify_column(event.x)
            # 只处理第一列（分组列）
            if column == "#1":
                # 切换排序状态
                if not hasattr(self, 'group_sort_order'):
                    self.group_sort_order = 'asc'  # 默认升序
                elif self.group_sort_order == 'asc':
                    self.group_sort_order = 'desc'  # 降序
                else:
                    self.group_sort_order = 'default'  # 默认（按创建时间）
                
                # 重新加载数据并排序
                self.populate_contact_tree()
                
                # 更新列标题显示
                if self.group_sort_order == 'asc':
                    self.contact_tree.heading("#1", text="分组 ↑")
                elif self.group_sort_order == 'desc':
                    self.contact_tree.heading("#1", text="分组 ↓")
                else:
                    self.contact_tree.heading("#1", text="分组")

    def create_new_recipient(self):
        """创建新的联系人，默认归入未分组"""
        # 确保"未分组"分组存在
        if "未分组" not in self.saved_recipients:
            self.saved_recipients["未分组"] = []
        
        # 弹出输入框获取联系人信息
        email = simpledialog.askstring("新建联系人", "请输入邮箱地址:", parent=self.root)
        if not email:
            return
            
        email = email.strip()
        if not re.match(r"^[a-zA-Z0-9._%+-]+@[a-zA-Z0-9.-]+\.[a-zA-Z]{2,}$", email):
            messagebox.showwarning("警告", "请输入有效的邮箱地址！")
            return
            
        note = simpledialog.askstring("新建联系人", "请输入备注:", parent=self.root)
        if note is None:  # 用户点击了取消
            return
            
        note = note.strip() if note else ""
        
        # 检查邮箱是否已存在
        for group_name, recipients in self.saved_recipients.items():
            for recipient in recipients:
                if recipient['email'] == email:
                    messagebox.showwarning("警告", f"邮箱 '{email}' 已存在于分组 '{group_name}' 中！")
                    return
        
        # 添加到"未分组"
        self.saved_recipients["未分组"].append({
            'email': email,
            'title': note,
            'content': note,
            'last_used': datetime.now().strftime("%Y-%m-%d %H:%M:%S")
        })
        
        self.save_recipient_data()
        self.populate_contact_tree()
        messagebox.showinfo("成功", f"联系人 '{email}' 已添加到'未分组'！")

    def save_recipient_data(self):
        """将当前保存的收件人数据写入文件"""
        try:
            with open(self.storage_file, 'w', encoding='utf-8') as f:
                json.dump(
                    {'groups': self.saved_recipients},
                    f,
                    ensure_ascii=False,
                    indent=2
                )
        except Exception as e:
            print(f"保存收件人数据失败: {e}")

    def add_recipient_to_saved_list(self):
        """添加收件人到已添加列表并保存到当前选中的分组"""
        email = self.email_var.get().strip()
        title = self.title_var.get().strip()
        content = self.content_text.get(1.0, tk.END).strip()

        if not email:
            messagebox.showwarning("警告", "收件人邮箱不能为空！")
            return
        if not re.match(r"[^@]+@[^@]+\.[^@]+", email):
            messagebox.showwarning("警告", "请输入有效的邮箱地址！")
            return

        # 验证邮箱格式
        if not re.match(r"^[a-zA-Z0-9._%+-]+@[a-zA-Z0-9.-]+\.[a-zA-Z]{2,}$", email):
            messagebox.showwarning("警告", "请输入有效的邮箱地址！")
            return

        # 获取当前选中的分组
        selected_item = self.contact_tree.focus()
        group_name = "未分组" # 默认分组
        if selected_item:
            # 检查选中的是否是分组节点
            if self.contact_tree.parent(selected_item) == '':
                group_name = self.contact_tree.item(selected_item, 'text')
            else:
                # 如果选中了联系人，则获取其父分组
                parent_item = self.contact_tree.parent(selected_item)
                group_name = self.contact_tree.item(parent_item, 'text')
        else:
            messagebox.showwarning("警告", "请先选择一个分组或联系人！")
            return
        
        if group_name not in self.saved_recipients:
            self.saved_recipients[group_name] = []

        # 检查是否已存在该收件人，存在则更新，不存在则新增
        updated = False
        for i, recipient in enumerate(self.saved_recipients[group_name]):
            if recipient['email'] == email:
                self.saved_recipients[group_name][i] = {
                    'email': email,
                    'title': title,
                    'content': content,
                    'last_used': datetime.now().strftime("%Y-%m-%d %H:%M:%S")
                }
                updated = True
                break
        if not updated:
            self.saved_recipients[group_name].append({
                'email': email,
                'title': title,
                'content': content,
                'last_used': datetime.now().strftime("%Y-%m-%d %H:%M:%S")
            })
        
        self.save_recipient_data() # 保存到文件
        self.populate_contact_tree() # 刷新联系人 Treeview
        messagebox.showinfo("成功", "联系人已添加/更新！")

    def delete_saved_recipient(self):
        """从当前选中的分组中删除收件人或分组"""
        selected_item = self.contact_tree.focus()
        if not selected_item:
            messagebox.showwarning("警告", "请选择要删除的联系人或分组！")
            return

        item_tags = self.contact_tree.item(selected_item, 'tags')
        if 'group' in item_tags:
            group_name = self.contact_tree.item(selected_item, 'text')
            if messagebox.askyesno("确认删除分组", f"确定要删除分组 '{group_name}' 及其所有联系人吗？"):
                del self.saved_recipients[group_name]
                self.save_recipient_data()
                self.populate_contact_tree()
        elif 'recipient' in item_tags:
            email_to_delete = self.contact_tree.item(selected_item, 'text')
            parent_id = self.contact_tree.parent(selected_item)
            group_name = self.contact_tree.item(parent_id, 'text') if parent_id else "未分组"

            if messagebox.askyesno("确认删除联系人", f"确定要删除联系人 '{email_to_delete}' 吗？"):
                if group_name in self.saved_recipients:
                    self.saved_recipients[group_name] = [r for r in self.saved_recipients[group_name] if r['email'] != email_to_delete]
                    self.save_recipient_data()
                    self.populate_contact_tree()
        else:
            messagebox.showwarning("警告", "请选择有效的联系人或分组进行删除！")

    def select_saved_recipient(self):
        """从联系人 Treeview 中加载选中的联系人信息到输入框"""
        # 确保在调用此方法前，已经通过 on_contact_tree_select 更新了输入框
        # 此方法主要用于处理“选择”按钮的点击，如果用户直接点击联系人，则由 on_contact_tree_select 处理
        selected_item = self.contact_tree.focus()
        if selected_item:
            item_tags = self.contact_tree.item(selected_item, 'tags')
            if 'recipient' in item_tags:
                values = self.contact_tree.item(selected_item, 'values')
                if values:
                    self.email_var.set(values[0])
                    # 不再自动设置主题和内容输入框
                    self.content_text.delete(1.0, tk.END)
            else:
                messagebox.showwarning("警告", "请选择一个具体的联系人！")
        else:
            messagebox.showwarning("警告", "请选择一个联系人！")

    def update_recipient_tree(self):
        """更新已添加收件人列表"""
        self.recipient_tree.delete(*self.recipient_tree.get_children())
        for recipient in self.recipients:
            self.recipient_tree.insert("", "end", values=(recipient['email'], recipient['title']))

    def clear_content(self):
        """清空邮件正文"""
        self.content_text.delete(1.0, tk.END)

    def add_recipient_to_send_list(self):
        """将当前输入框中的收件人信息添加到待发送列表"""
        email = self.email_var.get().strip()
        title = self.title_var.get().strip()
        content = self.content_text.get(1.0, tk.END).strip()

        if not email:
            messagebox.showwarning("警告", "收件人邮箱不能为空！")
            return

        # 验证邮箱格式
        if not re.match(r"^[a-zA-Z0-9._%+-]+@[a-zA-Z0-9.-]+\.[a-zA-Z]{2,}$", email):
            messagebox.showwarning("警告", "请输入有效的邮箱地址！")
            return

        # 检查是否已存在于待发送列表，存在则更新，不存在则新增
        updated = False
        for i, r in enumerate(self.recipients):
            if r['email'] == email:
                self.recipients[i] = {'email': email, 'title': title, 'content': content}
                updated = True
                break
        if not updated:
            self.recipients.append({'email': email, 'title': title, 'content': content})

        self.update_recipient_tree()
        messagebox.showinfo("成功", "收件人已添加到待发送列表！")

    def delete_selected(self):
        """从已添加收件人列表中删除选中项"""
        selected_items = self.recipient_tree.selection()
        if not selected_items:
            messagebox.showwarning("警告", "请选择要删除的收件人！")
            return

        for item in selected_items:
            email_to_delete = self.recipient_tree.item(item, 'values')[0]
            self.recipients = [r for r in self.recipients if r['email'] != email_to_delete]
            self.recipient_tree.delete(item)
        messagebox.showinfo("成功", "选中收件人已删除！")

    def confirm_and_close(self):
        """确认并发送邮件"""
        if not self.recipients:
            messagebox.showwarning("警告", "请添加至少一个收件人！")
            return
        
        # 获取当前选择的账户配置
        selected_account = self.account_var.get()
        account_config = self.account_manager.get_account(selected_account)
        
        try:
            # 导入发送器
            from src.core.email_sender import EmailSender
            from src.config.config import EmailConfig
            
            # 创建配置对象
            config = EmailConfig()
            
            # 临时设置配置信息
            temp_account = "temp_account"
            config.EMAIL_ACCOUNTS[temp_account] = {
                'SMTP_SERVER': account_config.get('SMTP_SERVER', ''),
                'SMTP_PORT': account_config.get('SMTP_PORT', 587),
                'SENDER': account_config.get('SENDER', ''),
                'PASSWORD': account_config.get('PASSWORD', ''),
                'SENDER_NAME': account_config.get('SENDER_NAME', '')
            }
            config.CURRENT_ACCOUNT = temp_account
            
            # 设置发送间隔
            config.SEND_INTERVAL = account_config.get('SEND_INTERVAL', 1)
            config.set_attachments(self.attachments)
            
            # 将任务添加到发送池
            item_id = self.send_pool.add_item(
                recipients=self.recipients,
                attachments=self.attachments,
                account_config={
                    'SMTP_SERVER': account_config.get('SMTP_SERVER', ''),
                    'SMTP_PORT': account_config.get('SMTP_PORT', 587),
                    'SENDER': account_config.get('SENDER', ''),
                    'PASSWORD': account_config.get('PASSWORD', ''),
                    'SENDER_NAME': account_config.get('SENDER_NAME', ''),
                    'SEND_INTERVAL': account_config.get('SEND_INTERVAL', 1)
                }
            )
            
            # 显示结果消息
            messagebox.showinfo("任务添加成功", f"发送任务已添加到发送池！任务ID: {item_id}")
            
            # 刷新发送池状态显示
            self.refresh_pool_status()
            
            # 清空当前输入，准备下一次发送
            self.recipients.clear()
            self.attachments.clear()
            self.update_recipient_tree()
            self.update_attach_preview()
            self.attach_label.config(text="已选附件: 0 个")
            
            # 标记已发送
            self.send_completed = True
            
        except Exception as e:
            messagebox.showerror("发送错误", f"发送邮件时发生错误: {str(e)}")

    def get_email_info(self):
        """获取收件人信息和附件列表"""
        self.root.mainloop()
        # 在持续运行模式下，这个方法不会返回直到窗口被显式关闭
        return {
            "recipients": self.recipients,
            "attachments": self.attachments,
            "account": self.account_var.get()  # 返回当前选择的账户
        }

    def on_close(self):
        """关闭窗口时二次确认"""
        # 检查是否有正在进行的任务
        pending_tasks = [item for item in self.send_pool.get_all_items() if item.status in ["pending", "sending"]]
        
        if pending_tasks:
            message = f"有 {len(pending_tasks)} 个任务正在发送池中处理，确定要退出吗？"
            if not messagebox.askyesno("退出确认", message):
                return
        
        # 如果已经发送过或者没有待处理任务，直接退出
        self.root.destroy()

    def manage_accounts(self):
        """管理邮箱账户"""
        # 创建账户管理窗口
        account_window = tk.Toplevel(self.root)
        account_window.title("邮箱账户管理")
        account_window.geometry("600x400")
        account_window.resizable(True, True)
        
        # 创建账户列表框架
        list_frame = ttk.Frame(account_window)
        list_frame.pack(fill=tk.BOTH, expand=True, padx=10, pady=10)
        
        # 创建账户列表
        columns = ("名称", "SMTP服务器", "端口", "发件人邮箱", "发件人名称")
        self.account_tree = ttk.Treeview(list_frame, columns=columns, show="headings", height=10)
        
        # 设置列标题
        for col in columns:
            self.account_tree.heading(col, text=col)
            self.account_tree.column(col, width=100)
        
        # 添加滚动条
        scrollbar = ttk.Scrollbar(list_frame, orient=tk.VERTICAL, command=self.account_tree.yview)
        self.account_tree.configure(yscrollcommand=scrollbar.set)
        
        # 布局
        self.account_tree.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        scrollbar.pack(side=tk.RIGHT, fill=tk.Y)
        
        # 加载账户数据
        self.load_account_data()
        
        # 创建按钮框架
        button_frame = ttk.Frame(account_window)
        button_frame.pack(fill=tk.X, padx=10, pady=(0, 10))
        
        # 添加按钮
        ttk.Button(button_frame, text="添加账户", command=self.add_account).pack(side=tk.LEFT, padx=(0, 5))
        ttk.Button(button_frame, text="编辑账户", command=self.edit_account).pack(side=tk.LEFT, padx=(0, 5))
        ttk.Button(button_frame, text="删除账户", command=self.delete_account).pack(side=tk.LEFT, padx=(0, 5))
        ttk.Button(button_frame, text="关闭", command=account_window.destroy).pack(side=tk.RIGHT)
        
        # 双击编辑
        self.account_tree.bind("<Double-1>", lambda event: self.edit_account())
    
    def load_account_data(self):
        """加载账户数据到表格"""
        # 清空现有数据
        for item in self.account_tree.get_children():
            self.account_tree.delete(item)
        
        # 添加账户数据
        accounts = self.account_manager.get_accounts()
        for name, info in accounts.items():
            self.account_tree.insert("", tk.END, values=(
                name,
                info['SMTP_SERVER'],
                info['SMTP_PORT'],
                info['SENDER'],
                info['SENDER_NAME']
            ))
    
    def add_account(self):
        """添加新账户"""
        self.show_account_dialog("添加账户")
    
    def edit_account(self):
        """编辑选中的账户"""
        selected = self.account_tree.selection()
        if not selected:
            messagebox.showwarning("警告", "请先选择一个账户！")
            return
        
        item = selected[0]
        values = self.account_tree.item(item, 'values')
        account_name = values[0]
        
        self.show_account_dialog("编辑账户", account_name)
    
    def delete_account(self):
        """删除选中的账户"""
        selected = self.account_tree.selection()
        if not selected:
            messagebox.showwarning("警告", "请先选择一个账户！")
            return
        
        item = selected[0]
        values = self.account_tree.item(item, 'values')
        account_name = values[0]
        
        if len(self.account_manager.get_account_names()) <= 1:
            messagebox.showwarning("警告", "不能删除最后一个账户！")
            return
        
        if messagebox.askyesno("确认删除", f"确定要删除账户 '{account_name}' 吗？"):
            if self.account_manager.delete_account(account_name):
                messagebox.showinfo("成功", f"账户 '{account_name}' 已删除！")
                self.load_account_data()
                # 更新主窗口的账户列表
                self.update_account_list()
            else:
                messagebox.showerror("错误", f"删除账户 '{account_name}' 失败！")
    
    def show_account_dialog(self, title, account_name=None):
        """显示账户编辑对话框"""
        # 创建对话框
        dialog = tk.Toplevel(self.root)
        dialog.title(title)
        dialog.geometry("400x300")
        dialog.resizable(False, False)
        dialog.grab_set()  # 模态对话框
        
        # 居中显示
        dialog.transient(self.root)
        dialog.update_idletasks()
        x = (dialog.winfo_screenwidth() // 2) - (dialog.winfo_width() // 2)
        y = (dialog.winfo_screenheight() // 2) - (dialog.winfo_height() // 2)
        dialog.geometry(f"+{x}+{y}")
        
        # 获取账户信息（如果是编辑）
        account_info = None
        if account_name:
            account_info = self.account_manager.get_account(account_name)
        
        # 创建表单
        form_frame = ttk.Frame(dialog)
        form_frame.pack(fill=tk.BOTH, expand=True, padx=20, pady=20)
        
        # 表单字段
        fields = [
            ("账户名称:", "name"),
            ("SMTP服务器:", "smtp_server"),
            ("SMTP端口:", "smtp_port"),
            ("发件人邮箱:", "sender"),
            ("密码/授权码:", "password"),
            ("发件人名称:", "sender_name")
        ]
        
        entries = {}
        initial_values = {
            "name": account_name if account_name else "",
            "smtp_server": account_info['SMTP_SERVER'] if account_info else "",
            "smtp_port": str(account_info['SMTP_PORT']) if account_info else "",
            "sender": account_info['SENDER'] if account_info else "",
            "password": account_info['PASSWORD'] if account_info else "",
            "sender_name": account_info['SENDER_NAME'] if account_info else ""
        }
        
        for i, (label_text, field_name) in enumerate(fields):
            ttk.Label(form_frame, text=label_text).grid(row=i, column=0, sticky="w", pady=5)
            entry = ttk.Entry(form_frame, width=30, show="*" if field_name == "password" else "")
            entry.grid(row=i, column=1, sticky="ew", pady=5, padx=(10, 0))
            entry.insert(0, initial_values[field_name])
            entries[field_name] = entry
        
        form_frame.columnconfigure(1, weight=1)
        
        # 按钮框架
        button_frame = ttk.Frame(dialog)
        button_frame.pack(fill=tk.X, padx=20, pady=(0, 20))
        
        def save_account():
            # 获取表单数据
            name = entries["name"].get().strip()
            smtp_server = entries["smtp_server"].get().strip()
            smtp_port = entries["smtp_port"].get().strip()
            sender = entries["sender"].get().strip()
            password = entries["password"].get()
            sender_name = entries["sender_name"].get().strip()
            
            # 验证必填字段
            if not name or not smtp_server or not smtp_port or not sender or not password or not sender_name:
                messagebox.showwarning("警告", "请填写所有字段！")
                return
            
            # 验证端口号
            try:
                smtp_port = int(smtp_port)
            except ValueError:
                messagebox.showwarning("警告", "SMTP端口必须是数字！")
                return
            
            # 保存账户
            if account_name:  # 编辑现有账户
                if self.account_manager.update_account(name, smtp_server, smtp_port, sender, password, sender_name):
                    messagebox.showinfo("成功", "账户信息已更新！")
                    dialog.destroy()
                    self.load_account_data()
                    # 更新主窗口的账户列表
                    self.update_account_list()
                else:
                    messagebox.showerror("错误", "更新账户信息失败！")
            else:  # 添加新账户
                if self.account_manager.add_account(name, smtp_server, smtp_port, sender, password, sender_name):
                    messagebox.showinfo("成功", "新账户已添加！")
                    dialog.destroy()
                    self.load_account_data()
                    # 更新主窗口的账户列表
                    self.update_account_list()
                else:
                    messagebox.showerror("错误", "添加新账户失败！")
        
        # 保存和取消按钮
        ttk.Button(button_frame, text="保存", command=save_account).pack(side=tk.RIGHT, padx=(10, 0))
        ttk.Button(button_frame, text="取消", command=dialog.destroy).pack(side=tk.RIGHT)
    
    def update_account_list(self):
        """更新主窗口的账户列表"""
        # 重新加载账户列表
        self.available_accounts = self.account_manager.get_account_names()
        self.account_combo['values'] = self.available_accounts
        
        # 如果当前选中的账户已被删除，选择第一个账户
        current_account = self.account_var.get()
        if current_account not in self.available_accounts:
            if self.available_accounts:
                self.account_var.set(self.available_accounts[0])
                self.on_account_change()

    def on_account_change(self, event=None):
        """账户选择变化时的处理函数"""
        selected_account = self.account_var.get()
        print(f"切换到账户: {selected_account}")
        # 这里可以添加账户切换时需要执行的其他逻辑



    



