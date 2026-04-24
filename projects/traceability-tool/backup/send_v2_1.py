# -*- coding: utf-8 -*-
"""Send Traceability Tool v2.1 via QQ Email"""

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

exe_path = r"C:\Users\Administrator\.openclaw\workspace\projects\traceability-tool\dist\TraceabilityTool_v2_1_final.exe"
readme_path = r"C:\Users\Administrator\.openclaw\workspace\projects\traceability-tool\README_v2.1.md"

# Check file size
exe_size = os.path.getsize(exe_path)
print("EXE size: %.2f MB" % (exe_size/1024/1024))

# Create message
msg = MIMEMultipart()
msg['From'] = sender_email
msg['To'] = receiver_email
msg['Subject'] = "双向溯源分析工具 v2.1 - 合并单元格版 (23MB)"

body = """你好！

附件是双向溯源分析工具 v2.1 版本。

【体积优化】
- 去掉pandas依赖，体积从56MB降到24MB
- 和之前v3.1版本大小一致

【更新内容】
1. A/B列：用例->需求（合并单元格竖向展开）
2. D/E列：需求->用例（合并单元格竖向展开）
3. G/I列：用例->需求（逗号分隔参考版）
4. K/L列：需求->用例（逗号分隔参考版）
5. 孤儿用例/需求自动归类至异常项Sheet

【使用方法】
1. 准备Excel（两列：用例ID、需求ID）
2. 运行程序，选择输入文件
3. 可选设置前缀过滤
4. 点击"开始分析"
5. 输出文件自动生成（同目录，文件名加 _output.xlsx）

如有问题请反馈！
"""

msg.attach(MIMEText(body, 'plain', 'utf-8'))

# Attach EXE (23MB < 50MB limit)
print("Attaching EXE...")
with open(exe_path, "rb") as f:
    part = MIMEBase('application', 'octet-stream')
    part.set_payload(f.read())
encoders.encode_base64(part)
part.add_header('Content-Disposition', 'attachment; filename="TraceabilityTool_v2_1.exe"')
msg.attach(part)

# Attach README
with open(readme_path, "rb") as f:
    part = MIMEBase('application', 'octet-stream')
    part.set_payload(f.read())
encoders.encode_base64(part)
part.add_header('Content-Disposition', 'attachment; filename="README_v2.1.md"')
msg.attach(part)

# Send email
try:
    print("Connecting to SMTP server...")
    server = smtplib.SMTP(smtp_server, smtp_port)
    server.starttls()
    print("Logging in...")
    server.login(sender_email, password)
    print("Sending email...")
    server.sendmail(sender_email, receiver_email, msg.as_string())
    server.quit()
    print("[OK] Email sent successfully!")
    print("  To: %s" % receiver_email)
    print("  Attachment: TraceabilityTool_v2_1.exe (%.2f MB)" % (exe_size/1024/1024))
except Exception as e:
    print("[ERROR] Failed: %s" % e)
