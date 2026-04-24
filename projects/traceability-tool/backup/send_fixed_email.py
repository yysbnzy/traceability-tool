# -*- coding: utf-8 -*-
"""Send fixed version EXE to QQ email"""
import smtplib
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from email.mime.base import MIMEBase
from email import encoders
import os

sender_email = "1239778532@qq.com"
receiver_email = "1239778532@qq.com"
password = "dlqmyuibjzuvibec"
smtp_server = "smtp.qq.com"
smtp_port = 587

msg = MIMEMultipart()
msg['From'] = sender_email
msg['To'] = receiver_email
msg['Subject'] = "Traceability Tool - FIXED VERSION (Orphan Requirement Detected)"

body = """Hello!

This is the FIXED version of the traceability analysis tool.

BUG FIXED:
- Previous version skipped rows when case_id was empty, even if requirement_id existed
- Now correctly detects orphan requirements (empty case_id + existing req_id)

CORRECT LOGIC:
- Skip row only when BOTH case_id AND requirement_id are empty
- Case_id empty + Req_id exists = ORPHAN REQUIREMENT (detected)
- Case_id exists + Req_id empty = ORPHAN CASE (detected)

TEST RESULT with your document:
- Orphan requirement found: SWRD-CyberSecurity-008
- Orphan cases: 120
- Total cases: 165

Output Excel includes:
1. Bidirectional traceability
2. Input source (includes orphan data)
3. Exception analysis
4. Missing requirements
5. Statistics

Please test and confirm!
"""

msg.attach(MIMEText(body, 'plain', 'utf-8'))

attachment_path = r"C:\Users\Administrator\.openclaw\workspace\projects\traceability-tool\dist\traceability_tool_fixed.exe"

with open(attachment_path, "rb") as attachment:
    part = MIMEBase('application', 'octet-stream')
    part.set_payload(attachment.read())

encoders.encode_base64(part)
part.add_header('Content-Disposition', f'attachment; filename= {os.path.basename(attachment_path)}')
msg.attach(part)

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