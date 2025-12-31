import tkinter as tk
from tkinter import ttk, messagebox, filedialog
import json
import os
from datetime import datetime
import _tkinter  # 添加导入以捕获 TclError


class RecipientInputWindow:
    def __init__(self, root):
        self.root = root
        self.root.title("Casiow")
        self.root.geometry("800x700")
        self.root.resizable(True, True)
        
        # 绑定窗口关闭事件（二次确认）
        self.root.protocol("WM_DELETE_WINDOW", self.on_close)
        
        self.recipients = []
        self.attachments = []  
        self.storage_file = "email_recipients.json"
        self.saved_recipients = []
        self.send_completed = False  # 标记是否已发送过
        
        self.create_widgets()
        self.load_saved_recipients()

    def create_widgets(self):
        # 整体网格配置（与之前一致）
        self.root.grid_rowconfigure(4, weight=1)  
        self.root.grid_columnconfigure(0, weight=1)
        
        # 1. 附件选择区域（与之前一致）
        frame_attach = ttk.LabelFrame(self.root, text="附件选择")
        frame_attach.grid(row=0, column=0, padx=10, pady=5, sticky="ew")
        
        ttk.Button(frame_attach, text="添加附件", command=self.add_attachments).grid(row=0, column=0, padx=5, pady=5)
        self.attach_label = ttk.Label(frame_attach, text="已选附件: 0 个")
        self.attach_label.grid(row=0, column=1, padx=5, pady=5, sticky="ew")
        frame_attach.grid_columnconfigure(1, weight=1)

        self.attach_preview = tk.Listbox(frame_attach, height=3)
        self.attach_preview.grid(row=1, column=0, columnspan=2, padx=5, pady=5, sticky="ew")
        frame_attach.grid_rowconfigure(1, weight=1)

        # 2. 快速选择收件人（与之前一致）
        frame_quick = ttk.LabelFrame(self.root, text="快速选择收件人")
        frame_quick.grid(row=1, column=0, padx=10, pady=5, sticky="ew")
        
        self.recipient_listbox = tk.Listbox(frame_quick, height=3)
        self.recipient_listbox.grid(row=0, column=0, padx=5, pady=5, sticky="ew")
        frame_quick.grid_columnconfigure(0, weight=1)

        scroll_quick = ttk.Scrollbar(frame_quick, orient="vertical", command=self.recipient_listbox.yview)
        scroll_quick.grid(row=0, column=1, sticky="ns")
        self.recipient_listbox.config(yscrollcommand=scroll_quick.set)

        frame_quick_buttons = ttk.Frame(frame_quick)
        frame_quick_buttons.grid(row=0, column=2, padx=5, pady=5, sticky="ns")
        
        ttk.Button(frame_quick_buttons, text="选择", command=self.select_saved_recipient).pack(pady=2, fill="x")
        ttk.Button(frame_quick_buttons, text="删除", command=self.delete_saved_recipient).pack(pady=2, fill="x")

        # 3. 收件人信息（与之前一致）
        frame_info = ttk.LabelFrame(self.root, text="收件人信息")
        frame_info.grid(row=2, column=0, padx=10, pady=5, sticky="ew")
        frame_info.grid_columnconfigure(1, weight=1)

        ttk.Label(frame_info, text="收件人邮箱:").grid(row=0, column=0, padx=5, pady=5, sticky="w")
        self.email_var = tk.StringVar()
        ttk.Entry(frame_info, textvariable=self.email_var).grid(row=0, column=1, padx=5, pady=5, sticky="ew")

        ttk.Label(frame_info, text="邮件主题:").grid(row=1, column=0, padx=5, pady=5, sticky="w")
        self.title_var = tk.StringVar()
        ttk.Entry(frame_info, textvariable=self.title_var).grid(row=1, column=1, padx=5, pady=5, sticky="ew")

        ttk.Label(frame_info, text="邮件正文:").grid(row=2, column=0, padx=5, pady=5, sticky="nw")
        self.content_text = tk.Text(frame_info, height=5)
        self.content_text.grid(row=2, column=1, padx=5, pady=5, sticky="nsew")
        frame_info.grid_rowconfigure(2, weight=1)

        scroll_content = ttk.Scrollbar(frame_info, orient="vertical", command=self.content_text.yview)
        scroll_content.grid(row=2, column=2, sticky="ns")
        self.content_text.config(yscrollcommand=scroll_content.set)

        frame_info_buttons = ttk.Frame(frame_info)
        frame_info_buttons.grid(row=3, column=0, columnspan=3, padx=5, pady=5, sticky="ew")
        
        ttk.Button(frame_info_buttons, text="添加收件人", command=self.add_recipient).grid(row=0, column=0, padx=5)
        ttk.Button(frame_info_buttons, text="清空正文", command=self.clear_content).grid(row=0, column=1, padx=5)

        # 4. 已添加收件人（与之前一致）
        ttk.Label(self.root, text="已添加的收件人:").grid(row=3, column=0, padx=10, pady=(10, 0), sticky="w")
        
        style = ttk.Style()
        style.configure("Treeview", background="#ffffff", foreground="black", fieldbackground="#ffffff")
        
        self.recipient_tree = ttk.Treeview(self.root, columns=["email", "title"], show="headings")
        self.recipient_tree.grid(row=4, column=0, padx=10, pady=5, sticky="nsew")
        self.root.grid_rowconfigure(4, weight=1)

        scroll_tree = ttk.Scrollbar(self.root, orient="vertical", command=self.recipient_tree.yview)
        scroll_tree.grid(row=4, column=1, sticky="ns")
        self.recipient_tree.config(yscrollcommand=scroll_tree.set)

        self.recipient_tree.heading("email", text="邮箱地址")
        self.recipient_tree.heading("title", text="邮件主题")
        self.recipient_tree.column("email", width=350)
        self.recipient_tree.column("title", width=350)

        ttk.Button(self.root, text="删除选中项", command=self.delete_selected).grid(row=5, column=0, padx=10, pady=5, sticky="w")

        # 5. 确认按钮（与之前一致）
        ttk.Button(self.root, text="确认", command=self.confirm).grid(row=6, column=0, pady=10)

    # 以下方法与之前一致（省略重复代码）
    def add_attachments(self): ...
    def update_attach_preview(self): ...
    def load_saved_recipients(self): ...
    def save_recipient(self, email, title, content): ...
    def select_saved_recipient(self): ...
    def delete_saved_recipient(self): ...
    def add_recipient(self): ...
    def clear_content(self): ...
    def delete_selected(self): ...

    def confirm(self):
        """点击确认后立即触发发送（不关闭窗口）"""
        if not self.recipients:
            messagebox.showerror("错误", "请至少添加一个收件人")
            return
        if not self.attachments:
            messagebox.showerror("错误", "请选择至少一个附件")
            return
        
        # 直接确认发送，无需额外弹窗（或保留确认弹窗）
        self.send_completed = True  # 标记为已发送
        # 这里不关闭窗口，主程序会继续执行发送逻辑

    def on_close(self):
        """关闭窗口时二次确认"""
        if self.send_completed:
            # 已发送过
            if messagebox.askyesno("退出确认", "确定要退出窗口吗？"):
                self.root.destroy()
        else:
            # 未发送过
            if messagebox.askyesno("退出确认", "尚未发送邮件，确定要退出吗？"):
                self.root.destroy()

def get_email_info():
    root = tk.Tk()
    app = RecipientInputWindow(root)
    
    while True:
        try:
            root.update()  # 刷新窗口
            # 当用户点击确认后，返回信息
            if app.send_completed:
                return {
                    "recipients": app.recipients,
                    "attachments": app.attachments
                }
            # 窗口被关闭时退出（增加try-except捕获）
            if not app.root.winfo_exists():
                return {
                    "recipients": [],
                    "attachments": []
                }
        except _tkinter.TclError:
            # 窗口已销毁时，直接退出循环
            return {
                "recipients": [],
                "attachments": []
            }