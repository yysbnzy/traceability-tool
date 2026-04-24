# -*- coding: utf-8 -*-
"""
Send usage manual via email
"""

import smtplib
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from email.mime.base import MIMEBase
from email import encoders
import os

# Email configuration
sender_email = "1239778532@qq.com"
receiver_email = "1239778532@qq.com"
password = "dlqmyuibjzuvibec"
smtp_server = "smtp.qq.com"
smtp_port = 587

# Attachment file
attachment_path = r"C:\Users\Administrator\.openclaw\workspace\projects\traceability-tool\dist\溯源分析工具_v3.1_使用说明.docx"

# Create message
msg = MIMEMultipart()
msg['From'] = sender_email
msg['To'] = receiver_email
msg['Subject'] = "溯源分析工具 v3.1 最终版 - 使用说明文档"

# Email body
body = """您好！

附件是用例需求溯源分析工具 v3.1 最终版的使用说明文档（Word格式）。

文档包含内容：
1. 工具简介与适用场景
2. 界面布局详解（5个区域）
3. 配置项详细说明（用例文档/拼接/需求文档）
4. 输入数据格式要求（含配套模板说明）
5. 输出报告5个Sheet详解
6. 典型使用场景示例（4个场景）
7. 常见问题与解决方案（5个Q&A）

配套文件清单：
- 溯源分析工具_v3_1_最终版.exe（主程序，双击运行）
- 用例需求溯源输入模板.xlsx（参考格式）
- 用例需求溯源输入模板_output.xlsx（示例输出）

如有问题请随时联系。

祝工作顺利！
"""

msg.attach(MIMEText(body, 'plain', 'utf-8'))

# Add attachment
if os.path.exists(attachment_path):
    with open(attachment_path, "rb") as attachment:
        part = MIMEBase('application', 'octet-stream')
        part.set_payload(attachment.read())

    encoders.encode_base64(part)
    part.add_header(
        'Content-Disposition',
        f'attachment; filename= {os.path.basename(attachment_path)}'
    )
    msg.attach(part)
    
    # Send email
    try:
        server = smtplib.SMTP(smtp_server, smtp_port)
        server.starttls()
        server.login(sender_email, password)
        text = msg.as_string()
        server.sendmail(sender_email, receiver_email, text)
        server.quit()
        print("[OK] 邮件发送成功！")
        print(f"  收件人: {receiver_email}")
        print(f"  附件: {os.path.basename(attachment_path)} ({os.path.getsize(attachment_path)/1024:.1f} KB)")
    except Exception as e:
        print(f"[ERROR] 发送失败: {e}")
else:
    print(f"[ERROR] 附件文件不存在: {attachment_path}")
