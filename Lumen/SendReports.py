import smtplib
import os
from email.message import EmailMessage

EMAIL_ADDRESS = os.gentenv('EMAIL_ADDRESS')
EMAIL_PASSWORD = os.getenv('EMAIL_PASSWORD')
TO_EMAIL = os.getenv('TO_EMAIL')

# Create Email
msg = EmailMessage()
msg['Subject'] = 'Daily Report?'
msg['From'] = EMAIL_ADDRESS
msg['To'] = TO_EMAIL
msg.set_content('Report Attached')

# Attach File
report_path = 'Reports/report.xlsx'  # Change this to your report file path
if not os.path.isfile(report_path):
    raise FileNotFoundError(f"The report file {report_path} does not exist.")

with open(report_path, 'rb') as f:
    file_data = f.read()
    file_name = os.path.basename(report_path)

msg.add_attachment(file_data, maintype='application', subtype='octet-stream', filename=file_name)

# Send email
with smtplib.SMTP_SSL('smtp.gmail.com', 465) as smtp:
    smtp.login(EMAIL_ADDRESS, EMAIL_PASSWORD)
    smtp.send_message(msg)

print('Email sent successfully.')