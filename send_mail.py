# Utility Imports
import os
import yaml
import logging
import smtplib
from tqdm import tqdm
import logging.config
from dotenv import load_dotenv
load_dotenv()

# Data Imports
import numpy as np
import pandas as pd

# Email Imports
import base64
from email import encoders
from email.mime.text import MIMEText
from email.mime.base import MIMEBase
from email.mime.multipart import MIMEMultipart


# Configure logging to write to a file with YAML format
logging_config = {
    'version': 1,
    'handlers': {
        'yaml_file_handler': {
            'class': 'logging.FileHandler',
            'filename': 'email_log.yaml',
            'formatter': 'yaml_formatter',
        },
    },
    'formatters': {
        'yaml_formatter': {
            'format': '{asctime} - {levelname} - {message}',
            'style': '{',
        },
    },
    'root': {
        'level': 'INFO',
        'handlers': ['yaml_file_handler'],
    },
}
with open('logging_config.yaml', 'w') as config_file:
    yaml.safe_dump(logging_config, config_file, default_flow_style=False)
logging.config.dictConfig(logging_config)
logger = logging.getLogger(__name__)

# Credentials
EMAIL_ADDRESS = os.getenv("EMAIL_ADDRESS")
EMAIL_PASSWORD = os.getenv("EMAIL_PASSWORD")
SMTP_SERVER = os.getenv("SMTP_SERVER")
SMTP_PORT = os.getenv("SMTP_PORT")


class EmailSender:
    def __init__(self, smtp_server, smtp_port, email_address, email_password):
        """
        Initializes an instance of the EmailSender class with the given SMTP server, port, email address, 
        and email password.

        Args:
            smtp_server (str): The SMTP server to use for sending emails.
            smtp_port (int): The port number of the SMTP server.
            email_address (str): The email address to use as the sender's address.
            email_password (str): The password for the email address.
        """
        self.smtp_server = smtp_server
        self.smtp_port = smtp_port
        self.email_address = email_address
        self.email_password = email_password

    def send_email(self, recipient_name, recipient_email, attachment_path):
        """
        Sends an email to the specified recipient with a certificate attachment.

        Args:
            recipient_name (str): The name of the recipient.
            recipient_email (str): The email address of the recipient.
            attachment_path (str): The path to the certificate attachment file.

        Raises:
            Exception: If an error occurs while sending the email.
            FileNotFoundError: Throw a error is the file does not exists
            smtplib.SMTPException: Throw an error if the SMTP server does not connect


        Returns:
            None
        """
        text = f'''Your Email in normal text format.'''.format(recipient_name=recipient_name)
        
        encoded = base64.b64encode(open("logo.png", "rb").read()).decode()
        html = f"""\Your Email in HTML Format"""
        
        try:
            # Create email message
            logger.info(f"Creating email message for {recipient_email}")
            msg = MIMEMultipart("alternative")
            msg["Subject"] = "Python Workshop Certificate"
            msg["From"] = self.email_address
            msg["To"] = recipient_email
            part1 = MIMEText(text, "plain")
            part2 = MIMEText(html, "html")
            msg.attach(part1)
            msg.attach(part2)

            # Attach certificate
            with open(attachment_path, 'rb') as attachment:
                pdf_part = MIMEBase("application", "octet-stream")
                pdf_part.set_payload(attachment.read())
            encoders.encode_base64(pdf_part)
            pdf_part.add_header(
                "Content-Disposition",
                f"attachment; filename={recipient_name}.pdf",
            )
            msg.attach(pdf_part)

            # Send email
            logger.info(f"Sending email to {recipient_email}")
            with smtplib.SMTP(self.smtp_server, self.smtp_port) as server:
                server.starttls()
                server.login(self.email_address, self.email_password)
                server.sendmail(from_addr=self.email_address, to_addrs=recipient_email, msg=msg.as_string())
                logger.info(f"Email sent successfully to {recipient_email}")
        except FileNotFoundError:
            logger.error(f"Attachment file {attachment_path} not found")
        except smtplib.SMTPException as e:
            logger.error(
                f"Error occurred while sending email to {recipient_email}: {e}")
        except Exception as e:
            logger.error(
                f"Unexpected error occurred while sending email to {recipient_email}: {e}")
            raise


def main():
    """
    Executes the main function of the program.

    This function initializes the email sender with the provided SMTP server, port, email address, 
    and email password. It then loads the test data from a CSV file named "test_data.csv" and extracts the 
    email addresses and names of the recipients. 

    Next, it initializes a progress bar using the tqdm library to track the progress of sending emails 
    to each recipient. The function then iterates over the recipients, sending an email to each one with 
    an attachment. The attachment path is constructed using the recipient's name and the file extension ".pdf".

    Parameters:
    None

    Returns:
    None
    """
    try:
        # Credentials
        smtp_server = SMTP_SERVER
        smtp_port = SMTP_PORT
        email_address = EMAIL_ADDRESS
        email_password = EMAIL_PASSWORD

        # Initialize email sender
        email_sender = EmailSender(
            smtp_server, smtp_port, email_address, email_password)

        # # Load test data
        # test_data = pd.read_csv("test_data.csv")
        # test_emails = np.array(test_data['Email'])
        # test_names = np.array(test_data['Name'])

        # # Initialize progress bar with the main loop
        # with tqdm(total=len(test_emails), position=0) as pbar:
        #     # Send email to each recipient
        #     for recipient_name, recipient_email in zip(test_names, test_emails):
        #         attachment_path = f'test/{recipient_name}.pdf'
        #         email_sender.send_email(
        #             recipient_name, recipient_email, attachment_path)
        #         pbar.update(1)

        # Load actual data
        actual_data = pd.read_csv("student_data.csv")
        actual_emails = np.array(actual_data['Email'])
        actual_names = np.array(actual_data['Name'])

        # Initialize progress bar with the main loop
        with tqdm(total=len(actual_emails), position=0) as pbar:
            # Send email to each recipient
            for recipient_name, recipient_email in zip(actual_names, actual_emails):
                attachment_path = f'Certificates/{recipient_name}.pdf'
                email_sender.send_email(
                    recipient_name, recipient_email, attachment_path)
                pbar.update(1)

        
    except Exception as e:
        logger.error(f"An error occurred: {e}")


if __name__ == "__main__":
    main()
    # print the log file
    with open('email_log.yaml', 'r') as f:
        print(f.read())
