"""
文件分卷压缩邮件工具
作者：gorkrr@126.com
版本：1.0
功能：
1. 调用7-Zip分卷压缩并加密文件（2MB/卷）
2. 通过SMTP协议发送分卷到指定邮箱
3. 支持打包为独立EXE
"""

import os
import glob
import tkinter as tk
from email.mime.text import MIMEText
from tkinter import ttk, filedialog, messagebox
import subprocess
import smtplib
from email.mime.multipart import MIMEMultipart
from email.mime.base import MIMEBase
from email import encoders


class CompressionApp:
    """主应用程序类"""

    def __init__(self, root):
        """初始化GUI界面"""
        self.root = root
        self.root.title("文件分卷邮寄工具 v1.0")
        self.root.resizable(False, False) # 什么意思

        # 配置默认参数
        self.config = {
            '7z_path': 'C:\\Program Files\\7-Zip\\7z.exe',  # 7-Zip安装路径
            'smtp_server': 'smtp.example.com',  # SMTP服务器
            'smtp_port': 587,  # SMTP端口
            'sender_email': 'your-email@example.com',  # 发件邮箱
            'sender_password': 'your-password'  # 邮箱密码/授权码
        }

        self._create_widgets()
        self._layout()

    def _create_widgets(self):
        """创建界面组件"""
        # 文件选择
        self.lbl_file = ttk.Label(self.root, text="选择文件:")
        self.entry_file = ttk.Entry(self.root, width=40)
        self.btn_browse = ttk.Button(self.root, text="浏览", command=self._browse_file)

        # 邮件配置
        self.lbl_email = ttk.Label(self.root, text="收件邮箱:")
        self.entry_email = ttk.Entry(self.root, width=40)

        # 安全配置
        self.lbl_password = ttk.Label(self.root, text="压缩密码:")
        self.entry_password = ttk.Entry(self.root, show="*", width=40)

        # 操作按钮
        self.btn_start = ttk.Button(self.root, text="开始处理", command=self._start_process)

        # 进度信息
        self.progress = ttk.Progressbar(self.root, orient=tk.HORIZONTAL, mode='determinate')
        self.lbl_status = ttk.Label(self.root, text="就绪")

    def _layout(self):
        """布局管理"""
        # 第一行：文件选择
        self.lbl_file.grid(row=0, column=0, padx=5, pady=5, sticky=tk.W)
        self.entry_file.grid(row=0, column=1, padx=5, pady=5)
        self.btn_browse.grid(row=0, column=2, padx=5, pady=5)

        # 第二行：收件邮箱
        self.lbl_email.grid(row=1, column=0, padx=5, pady=5, sticky=tk.W)
        self.entry_email.grid(row=1, column=1, columnspan=2, padx=5, pady=5, sticky=tk.EW)

        # 第三行：压缩密码
        self.lbl_password.grid(row=2, column=0, padx=5, pady=5, sticky=tk.W)
        self.entry_password.grid(row=2, column=1, columnspan=2, padx=5, pady=5, sticky=tk.EW)

        # 第四行：操作按钮
        self.btn_start.grid(row=3, column=1, padx=5, pady=10)

        # 第五行：进度条和状态
        self.progress.grid(row=4, column=0, columnspan=3, padx=5, pady=5, sticky=tk.EW)
        self.lbl_status.grid(row=5, column=0, columnspan=3, padx=5, pady=5)

    def _browse_file(self):
        """打开文件选择对话框"""
        file_path = filedialog.askopenfilename()
        if file_path:
            self.entry_file.delete(0, tk.END)
            self.entry_file.insert(0, file_path)

    def _start_process(self):
        """主处理流程"""
        # 获取输入参数
        file_path = self.entry_file.get()
        password = self.entry_password.get()
        target_email = self.entry_email.get()

        # 输入验证
        if not all([file_path, password, target_email]):
            messagebox.showerror("错误", "请填写所有必填字段！")
            return

        if not os.path.exists(file_path):
            messagebox.showerror("错误", "文件不存在！")
            return

        try:
            # 执行分卷压缩
            self._update_status("正在压缩文件...", 20)
            archive_name = self._compress_file(file_path, password)

            # 发送分卷文件
            self._update_status("正在发送邮件...", 50)
            parts = glob.glob(f"{archive_name}.*")
            total = len(parts)

            for idx, part in enumerate(parts, 1):
                self._send_email(part, target_email, password)
                progress = 50 + int((idx / total) * 50)
                self._update_status(f"发送中 ({idx}/{total})", progress)

            messagebox.showinfo("完成", "所有分卷已发送成功！")
            self._reset_ui()

        except Exception as e:
            messagebox.showerror("错误", f"处理失败: {str(e)}")
            self._reset_ui()

    def _compress_file(self, file_path, password):
        """
        调用7-Zip进行分卷压缩
        :param file_path: 要压缩的文件路径
        :param password: 压缩密码
        :return: 生成的压缩包前缀
        """
        archive_name = "output.7z"

        try:
            subprocess.run(
                [
                    self.config['7z_path'],
                    'a',  # 添加文件
                    '-v2m',  # 分卷大小2MB
                    f'-p{password}',  # 设置密码
                    '-mx=9',  # 最大压缩率
                    archive_name,  # 输出文件名
                    file_path  # 要压缩的文件
                ],
                check=True,
                stdout=subprocess.DEVNULL,
                stderr=subprocess.DEVNULL
            )
            return archive_name
        except subprocess.CalledProcessError as e:
            raise RuntimeError(f"7-Zip压缩失败: {str(e)}")

    def _send_email(self, attachment, target_email, password):
        """
        通过SMTP发送邮件
        :param attachment: 附件路径
        :param target_email: 收件邮箱
        :param password: 压缩密码
        """
        msg = MIMEMultipart()
        msg['From'] = self.config['sender_email']
        msg['To'] = target_email
        msg['Subject'] = f"分卷文件: {os.path.basename(attachment)}"

        # 构建邮件正文
        body = f"""
        请查收分卷压缩文件

        文件信息：
        - 分卷名称：{os.path.basename(attachment)}
        - 解压密码：{password}
        - 需要所有分卷文件才能解压

        本邮件由自动分卷工具发送
        """
        msg.attach(MIMEText(body, 'plain'))

        # 添加附件
        with open(attachment, 'rb') as f:
            part = MIMEBase('application', 'octet-stream')
            part.set_payload(f.read())
            encoders.encode_base64(part)
            part.add_header(
                'Content-Disposition',
                f'attachment; filename="{os.path.basename(attachment)}"'
            )
            msg.attach(part)

        # 发送邮件
        try:
            with smtplib.SMTP(
                    self.config['smtp_server'],
                    self.config['smtp_port']
            ) as server:
                server.starttls()
                server.login(
                    self.config['sender_email'],
                    self.config['sender_password']
                )
                server.send_message(msg)
        except Exception as e:
            raise RuntimeError(f"邮件发送失败: {str(e)}")

    def _update_status(self, text, value=None):
        """更新界面状态"""
        self.lbl_status.config(text=text)
        if value is not None:
            self.progress['value'] = value
        self.root.update_idletasks()

    def _reset_ui(self):
        """重置界面状态"""
        self._update_status("就绪", 0)


if __name__ == "__main__":
    root = tk.Tk()
    app = CompressionApp(root)
    root.mainloop()