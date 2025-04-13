"""
文件分卷压缩邮件工具（优化版）
作者：AI助手
版本：1.1
更新内容：
1. 压缩包名自动匹配源文件名
2. 邮件主题显示原始文件名和分卷编号
3. 发送后自动清理分卷文件
4. 预填充默认邮箱和密码
"""

import os
import glob
import tkinter as tk
from tkinter import ttk, filedialog, messagebox
import subprocess
import smtplib
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from email.mime.base import MIMEBase
from email import encoders


class CompressionApp:
    """主应用程序类"""

    def __init__(self, root):
        """初始化界面和配置"""
        self.root = root
        self.root.title("文件分卷邮寄工具 v1.1")
        self.root.resizable(False, False)

        # 配置参数
        self.config = {
            '7z_path': 'C:\\Program Files\\7-Zip\\7z.exe',  # 7-Zip安装路径
            'smtp_server': 'smtp.126.com',  # SMTP服务器
            'smtp_port': 25,  # SMTP端口
            'sender_email': 'gorkrr@126.com',  # 发件邮箱
            'sender_password': 'NLcQ3hhdTTEe3Bhf',  # 邮箱密码/授权码
            'receive_email': "gorkrr@126.com", # 默认接受邮箱
            'compress_password': "zxy123" # 默认压缩密码
        }


        self._create_widgets()
        self._layout()
        self._set_defaults()

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
        components = [
            (self.lbl_file, 0, 0), (self.entry_file, 0, 1), (self.btn_browse, 0, 2),
            (self.lbl_email, 1, 0), (self.entry_email, 1, 1, 2),
            (self.lbl_password, 2, 0), (self.entry_password, 2, 1, 2),
            (self.btn_start, 3, 1),
            (self.progress, 4, 0, 3), (self.lbl_status, 5, 0, 3)
        ]

        for comp in components:
            if len(comp) == 3:
                comp[0].grid(row=comp[1], column=comp[2], padx=5, pady=5, sticky=tk.W)
            else:
                comp[0].grid(
                    row=comp[1], column=comp[2],
                    columnspan=comp[3], padx=5, pady=5, sticky=tk.EW
                )

    def _set_defaults(self):
        """设置界面默认值"""
        self.entry_email.insert(0, self.config['receive_email'])
        self.entry_password.insert(0, self.config['compress_password'])

    def _browse_file(self):
        """选择文件"""
        if file_path := filedialog.askopenfilename():
            self.entry_file.delete(0, tk.END)
            self.entry_file.insert(0, file_path)

    def _start_process(self):
        """处理流程控制器"""
        # 获取输入参数
        params = {
            'file_path': self.entry_file.get(),
            'password': self.entry_password.get(),
            'target_email': self.entry_email.get()
        }

        # 输入验证
        if error := self._validate_input(params):
            messagebox.showerror("输入错误", error)
            return

        try:
            # 阶段1：压缩文件
            self._update_status("正在压缩文件...", 20)
            archive_name, original_name = self._compress_file(**params)

            # 阶段2：发送邮件
            self._update_status("正在准备发送...", 40)
            parts = glob.glob(f"{archive_name}.*")

            # 阶段3：逐个发送
            total_parts = len(parts)
            for idx, part in enumerate(parts, 1):
                self._send_email(part, params['target_email'], params['password'],
                                 original_name, idx, total_parts)
                progress = 40 + int((idx / total_parts) * 60)
                self._update_status(f"发送中 ({idx}/{total_parts})", progress)

            # 阶段4：清理文件
            self._cleanup_files(archive_name)
            messagebox.showinfo("完成", "所有分卷已发送并清理完成！")

        except Exception as e:
            messagebox.showerror("处理错误", str(e))
        finally:
            self._reset_ui()

    def _validate_input(self, params):
        """输入验证"""
        if not all(params.values()):
            return "所有字段必须填写！"
        if not os.path.exists(params['file_path']):
            return "文件不存在！"
        return ""

    def _compress_file(self, file_path, password, target_email):
        """执行7-Zip压缩"""
        original_name = os.path.basename(file_path)
        base_name = os.path.splitext(original_name)[0]
        archive_name = f"{base_name}.7z"

        try:
            subprocess.run(
                [
                    self.config['7z_path'],
                    'a', '-v2m', f'-p{password}', '-mx9',
                    archive_name, file_path
                ],
                check=True,
                stdout=subprocess.DEVNULL,
                stderr=subprocess.DEVNULL
            )
            return archive_name, original_name
        except subprocess.CalledProcessError as e:
            raise RuntimeError(f"压缩失败: {e.stderr or '未知错误'}")

    def _send_email(self, part_path, target_email, password,
                    original_name, part_num, total_parts):
        """发送单个分卷邮件"""
        msg = MIMEMultipart()
        msg['From'] = self.config['sender_email']
        msg['To'] = target_email
        msg['Subject'] = f"{original_name} - 分卷({part_num}/{total_parts})"

        # 构建正文
        body = f"""
        文件名：{original_name}
        分卷信息：{part_num}/{total_parts}
        解压密码：{password}
        """
        msg.attach(MIMEText(body, 'plain', 'utf-8'))

        # 添加附件
        with open(part_path, 'rb') as f:
            part = MIMEBase('application', 'octet-stream')
            part.set_payload(f.read())
            encoders.encode_base64(part)
            part.add_header('Content-Disposition',
                            f'attachment; filename="{os.path.basename(part_path)}"')
            msg.attach(part)

        # SMTP发送
        try:
            with smtplib.SMTP(self.config['smtp_server'], self.config['smtp_port']) as server:
                server.starttls()
                server.login(self.config['sender_email'], self.config['sender_password'])
                server.send_message(msg)
        except Exception as e:
            raise RuntimeError(f"邮件发送失败: {str(e)}")

    def _cleanup_files(self, archive_base):
        """清理分卷文件"""
        for f in glob.glob(f"{archive_base}.*"):
            os.remove(f)

    def _update_status(self, text, value=None):
        """更新界面状态"""
        self.lbl_status.config(text=text)
        if value is not None:
            self.progress['value'] = value
        self.root.update()

    def _reset_ui(self):
        """重置界面状态"""
        self.progress['value'] = 0
        self.lbl_status.config(text="就绪")


if __name__ == "__main__":
    root = tk.Tk()
    app = CompressionApp(root)
    root.mainloop()