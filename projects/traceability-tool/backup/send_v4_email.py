# -*- coding: utf-8 -*-
"""
Send v4 EXE to QQ email
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
msg['Subject'] = "Traceability Tool v4 - Fixed"

# Email body
body = """Hello!

This is v4 of the traceability analysis tool with bug fixes.

Fixes in v4:
1. Fixed hardcoded file path issue
2. Auto-detects Excel files in current directory
3. Supports drag-and-drop operation
4. Shows friendly error messages

Usage:
1. Place EXE and Excel file in same folder
2. Double-click EXE to run
3. Or drag Excel file onto EXE
4. Analysis results will be saved in same folder

The program will automatically find and process the Excel file.

"""

msg.attach(MIMEText(body, 'plain', 'utf-8'))

# Attachment file
attachment_path = r"C:\Users\Administrator\.openclaw\workspace\projects\traceability-tool\dist\run_analysis_fixed_v4.exe"

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
