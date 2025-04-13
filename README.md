# 分卷压缩发送

1、分卷压缩并设置密码 2、顺序发送到邮箱

# 界面操作流程

1. 点击"浏览"选择文件
2. 输入收件邮箱和压缩密码
3. 点击"开始处理"启动流程
4. 查看进度条和状态提示

# 配置方法
```python
# 配置默认参数
self.config = {
    '7z_path': 'C:\\Program Files\\7-Zip\\7z.exe',  # 7-Zip安装路径
    'smtp_server': 'smtp.126.com',  # SMTP服务器
    'smtp_port': 25,  # SMTP端口
    'sender_email': 'gorkrr@126.com',  # 发件邮箱
    'sender_password': 'NLcQ3hhdTTEe3Bhf',  # 邮箱密码/授权码
}
```


# 打包方法：
```shell
pip install pyinstaller

pyinstaller --onefile --windowed build.spec
```