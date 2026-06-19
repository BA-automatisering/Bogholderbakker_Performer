"""Handles all about sending emails"""
from dataclasses import dataclass
from typing import List
import smtplib
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from email.mime.base import MIMEBase
from email import encoders
import os


@dataclass
class Email:
    """Stuff for emails is defined"""
    subject: str
    body: str
    recipients: List[str]
    cc: List[str] | None = None
    attachment_path: str | None = None


class EmailSender:
    """Klasse til at sende emails"""
    def __init__(self, smtp_server, smtp_port, sender_email, test_recipient: str | None = None):
        self.smtp_server = smtp_server
        self.smtp_port = smtp_port
        self.sender_email = sender_email
        self.test_recipient = test_recipient

    def send_email(self, mail):
        """Sender email"""
        recipients = [self.test_recipient] if self.test_recipient else mail.recipients
        msg = MIMEMultipart()
        msg['Fra'] = self.sender_email
        msg['Til'] = ", ".join(recipients)
        msg['Subject'] = mail.subject

        # Tilføj e-mailteksten
        msg.attach(MIMEText(mail.body, 'plain'))

        # Vedhæft Excel-fil
        if mail.attachment_path:
            with open(mail.attachment_path, "rb") as attachment:
                part = MIMEBase("application", "octet-stream")
                part.set_payload(attachment.read())
            encoders.encode_base64(part)
            part.add_header(
                "Content-Disposition",
                f"attachment; filename= {os.path.basename(mail.attachment_path)}",
            )
            msg.attach(part)

        # Log på SMTP-serveren og send e-mailen
        with smtplib.SMTP(self.smtp_server, self.smtp_port) as server:
            server.starttls()
            server.sendmail(self.sender_email, recipients, msg.as_string())
            print(f"Email er sendt til {recipients} (original: {mail.recipients})")


def build_emails(file_entries, orchestrator_connection, filnavn):
    """Build Email objects from file metadata created by FileManager."""
    emails = []

    if not file_entries:
        orchestrator_connection.log_trace("Ingen fil at sende!!")
        return emails

    for entry in file_entries:
        subject = filnavn
        body = """
            Stamdata kontrol er gennemført. Kontrollisten er vedhæftet denne e-mail
        """
        

        emails.append(
            Email(
                subject=subject,
                body=body,
                recipients=[entry["recipient"]],
                attachment_path=entry["attachment_path"]
            )
        )

    return emails