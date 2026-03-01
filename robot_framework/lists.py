import smtplib
from email.message import EmailMessage
import base64
import traceback
from io import BytesIO
import datetime

from OpenOrchestrator.orchestrator_connection.connection import OrchestratorConnection
from robot_framework import config
from robot_framework import globals


def send_manuelliste(process_name):
    
    orchestrator_connection.log_trace("Send manuelliste started...")
    x = datetime.datetime.now()
    print((x.strftime("%d-%b-%Y")))
    
    to_address = "lejp@aarhus.dk"
    msg = EmailMessage()
    msg['to'] = to_address
    msg['from'] = config.LIST_SENDER
    msg['subject'] = f"Bogholderbakker {x.strftime("%d-%b-%Y")}: {process_name}"
    
    n = 0
    header = process_name+"  "+x.strftime("%d-%b-%Y")
    body = ""
    while n < len(globals.manuelliste ):
        body = body + "Fakturanr. "+globals.manuelliste[n]["Fakturanr"]+": "+globals.manuelliste[n]["Beskrivelse"]+"<br>"
        n += 1
    
    html_message = f"""
    <html>
        <body>
            <b>{header}</b>
            <p>{body}</p>
        </body>
    </html>
    """
    msg.set_content("Please enable HTML to view this message.")
    msg.add_alternative(html_message, subtype='html')

    # Send message
    with smtplib.SMTP(config.SMTP_SERVER, config.SMTP_PORT) as smtp:
        smtp.starttls()
        smtp.send_message(msg)
        
    orchestrator_connection.log_trace("Send manuelliste ended...")
    