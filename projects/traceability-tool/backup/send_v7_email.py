# -*- coding: utf-8 -*-
"""
Send GUI v7 EXE to QQ email
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
msg['Subject'] = "Traceability Tool GUI v7 - Complete Version"

# Email body
body = """Hello!

This is the complete GUI version v7 of the traceability analysis tool.

Features:
1. Select case document with A-Z column dropdown for case ID and requirement ID
2. Optional requirement document selection with A-Z column dropdown
3. Prefix recognition filter for both case and requirement IDs
4. Field concatenation: prefix + original ID + suffix
5. Complete analysis with orphan case/requirement detection

Output Excel with 5 sheets:
1. Bidirectional traceability table
2. Input source (raw data)
3. Exception analysis (orphan cases, orphan requirements, prefix mismatches)
4. Missing test cases for requirements
5. Statistics summary

Please test with your document and let me know if any issues.

"""

msg.attach(MIMEText(body, 'plain', 'utf-8'))

# Attachment file
attachment_path = r"C:\Users\Administrator\.openclaw\workspace\projects\traceability-tool\dist\traceability_tool_gui_v7.exe"

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
