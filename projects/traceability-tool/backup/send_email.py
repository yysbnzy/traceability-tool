# -*- coding: utf-8 -*-
"""
Send email to QQ mailbox
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

# Create message
msg = MIMEMultipart()
msg['From'] = sender_email
msg['To'] = receiver_email
msg['Subject'] = "Traceability Tool - Fixed Version v3"

# Email body
body = """Hello!

The attachment is the fixed traceability analysis tool.

Fixes:
1. Fixed orphan requirement detection logic
2. Correctly identifies orphan requirements like SWRD-CyberSecurity-008
3. Supports bidirectional traceability analysis

Usage:
1. Place EXE file and test case document in same directory
2. Double-click run_analysis_fixed_v3.exe
3. Analysis result Excel file will be generated automatically

Output files:
- Bidirectional traceability table
- Abnormal analysis: orphan test cases (120) + orphan requirements (1)
- Statistics summary

Data validation passed!

"""

msg.attach(MIMEText(body, 'plain', 'utf-8'))

# Attachment file
attachment_path = r"C:\Users\Administrator\.openclaw\workspace\projects\traceability-tool\traceability_tool_release.zip"

# Add attachment
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
    print("[OK] Email sent successfully!")
    print(f"  To: {receiver_email}")
    print(f"  Attachment: {os.path.basename(attachment_path)} ({os.path.getsize(attachment_path)/1024/1024:.2f} MB)")
except Exception as e:
    print(f"[ERROR] Failed: {e}")
