# -*- coding: utf-8 -*-
"""Send Traceability Tool v2.1 via QQ Email with ZIP"""

import smtplib
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from email.mime.base import MIMEBase
from email import encoders
import os
import socket

# Email configuration
sender_email = "1239778532@qq.com"
receiver_email = "1239778532@qq.com"
password = "dlqmyuibjzuvibec"
smtp_server = "smtp.qq.com"
smtp_port = 587

zip_path = r"C:\Users\Administrator\.openclaw\workspace\projects\traceability-tool\dist\TraceabilityTool_v2_1.zip"
readme_path = r"C:\Users\Administrator\.openclaw\workspace\projects\traceability-tool\README_v2.1.md"

# Check file size
zip_size = os.path.getsize(zip_path)
print("ZIP size: %.2f MB" % (zip_size/1024/1024))

# Create message
msg = MIMEMultipart()
msg['From'] = sender_email
msg['To'] = receiver_email
msg['Subject'] = "双向溯源分析工具 v2.1 - ZIP压缩版 (24MB)"

body = """你好！

附件是双向溯源分析工具 v2.1 版本（ZIP压缩）。

【体积优化】
- 去掉pandas依赖，体积从56MB降到24MB

【更新内容】
1. A/B列：用例->需求（合并单元格竖向展开）
2. D/E列：需求->用例（合并单元格竖向展开）
3. G/I列：用例->需求（逗号分隔参考版）
4. K/L列：需求->用例（逗号分隔参考版）
5. 孤儿用例/需求自动归类至异常项Sheet

【解压后使用】
1. 解压ZIP得到 TraceabilityTool_v2_1.exe
2. 准备Excel（两列：用例ID、需求ID）
3. 运行程序，选择输入文件
4. 输出文件自动生成（同目录，文件名加 _output.xlsx）

如有问题请反馈！
"""

msg.attach(MIMEText(body, 'plain', 'utf-8'))

# Attach ZIP
print("Attaching ZIP...")
with open(zip_path, "rb") as f:
    part = MIMEBase('application', 'zip')
    part.set_payload(f.read())
encoders.encode_base64(part)
part.add_header('Content-Disposition', 'attachment; filename="TraceabilityTool_v2_1.zip"')
msg.attach(part)

# Attach README
with open(readme_path, "rb") as f:
    part = MIMEBase('application', 'octet-stream')
    part.set_payload(f.read())
encoders.encode_base64(part)
part.add_header('Content-Disposition', 'attachment; filename="README_v2.1.md"')
msg.attach(part)

# Send email with longer timeout
try:
    print("Connecting to SMTP server (timeout=60s)...")
    socket.setdefaulttimeout(60)
    server = smtplib.SMTP(smtp_server, smtp_port, timeout=60)
    server.starttls()
    print("Logging in...")
    server.login(sender_email, password)
    print("Sending email (this may take a while)...")
    server.sendmail(sender_email, receiver_email, msg.as_string())
    server.quit()
    print("[OK] Email sent successfully!")
    print("  To: %s" % receiver_email)
    print("  ZIP size: %.2f MB" % (zip_size/1024/1024))
except Exception as e:
    print("[ERROR] Failed: %s" % e)
    import traceback
    traceback.print_exc()
