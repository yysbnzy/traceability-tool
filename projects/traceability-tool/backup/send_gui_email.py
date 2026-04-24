# -*- coding: utf-8 -*-
"""
Send GUI version EXE to QQ email
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
msg['Subject'] = "Traceability Tool - GUI Version"

# Email body
body = """Hello!

This is the GUI version of the traceability analysis tool.

Features:
1. Select case document (required) - Excel file with case and requirement IDs
2. Select requirement document (optional) - for validation
3. Configure column mappings
4. Configure prefix/suffix for ID formatting
5. Run analysis and generate report

Usage:
1. Double-click EXE to open GUI
2. Click "Browse..." to select case document
3. (Optional) Click "Browse..." to select requirement document
4. Configure columns if needed
5. Click "Start Analysis"
6. Select output location for the report

Note: This is the original GUI version with full functionality.

"""

msg.attach(MIMEText(body, 'plain', 'utf-8'))

# Attachment file
attachment_path = r"C:\Users\Administrator\.openclaw\workspace\projects\traceability-tool\dist\traceability_v2_lite.exe"

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
