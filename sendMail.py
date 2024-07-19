import smtplib
from email.message import EmailMessage
import os

class OutlookEmailSender:
    def __init__(self, email, password):
        self.email = email
        self.password = password
        self.smtp_server = 'smtp.gmail.com'
        self.smtp_port = 587

    def send_custom_email(self, from_email, to_email, cc_email, bcc_email, subject, body):
        self._send_email_internal(from_email, to_email, cc_email, bcc_email, subject, body, [])

    def send_custom_email_with_attachments(self, from_email, to_email, cc_email, bcc_email, subject, body, attachments):
        self._send_email_internal(from_email, to_email, cc_email, bcc_email, subject, body, attachments)

    def _send_email_internal(self, from_email, to_email, cc_email, bcc_email, subject, body, attachments):
        try:
            msg = EmailMessage()
            msg['Subject'] = subject
            msg['From'] = from_email
            msg['To'] = to_email
            msg['Cc'] = cc_email
            msg['Bcc'] = bcc_email
            msg.set_content(body)

            for file in attachments:
                with open(file, 'rb') as f:
                    file_data = f.read()
                    file_name = os.path.basename(file)
                    msg.add_attachment(file_data, maintype='application', subtype='octet-stream', filename=file_name)

            with smtplib.SMTP(self.smtp_server, self.smtp_port) as server:
                server.starttls()
                server.login(self.email, self.password)
                server.send_message(msg)

            print("Email sent successfully!")
        except Exception as e:
            print(f"Failed to send email: {e}")

# Example usage:
# if __name__ == "__main__":
#     email_sender = OutlookEmailSender('your_email@outlook.com', 'your_password')
#     email_sender.send_email('recipient@example.com', 'Test Subject', 'This is a test email.')
#     email_sender.send_email_with_attachments('recipient@example.com', 'Test Subject with Attachments', 'This email has attachments.', ['path/to/attachment1', 'path/to/attachment2'])
#     email_sender.send_custom_email('your_email@outlook.com', 'recipient@example.com', 'cc@example.com', 'bcc@example.com', 'Custom Subject', 'This is a custom email.')
#     email_sender.send_custom_email_with_attachments('your_email@outlook.com', 'recipient@example.com', 'cc@example.com', 'bcc@example.com', 'Custom Subject with Attachments', 'This email has custom attachments.', ['path/to/attachment1', 'path/to/attachment2'])
