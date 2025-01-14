import smtplib
from email.mime.text import MIMEText
from email.mime.multipart import MIMEMultipart
from email.mime.application import MIMEApplication
from datetime import datetime
import os
import configparser
import sys
import platform
import subprocess
import ssl
import logging
from pathlib import Path

# Setup base paths
BASE_DIR = Path(__file__).resolve().parent.parent
CONFIG_DIR = BASE_DIR / 'config'
LOGS_DIR = BASE_DIR / 'logs'

# Ensure directories exist
LOGS_DIR.mkdir(exist_ok=True)

# Setup logging
logging.basicConfig(
    filename=LOGS_DIR / 'report_automation.log',
    level=logging.INFO,
    format='%(asctime)s - %(message)s'
)

def get_generate_report_command():
    system = platform.system().lower()
    return ["npx", "ts-node", ".\\prisma\\scripts\\generateReport.ts"] if system == "windows" else ["npx", "ts-node", "/Users/rupinajay/Developer/Autoapplys.com/prisma/scripts/generateReport.ts"]

def generate_excel_report():
    try:
        command = get_generate_report_command()
        logging.info(f"Executing command: {' '.join(command)}")
        
        process = subprocess.run(
            command,
            check=True,
            stdout=subprocess.PIPE,
            stderr=subprocess.PIPE,
            text=True
        )
        
        if process.returncode == 0:
            logging.info("Excel report generated successfully")
            return True
        else:
            logging.error(f"Error generating report: {process.stderr}")
            return False
            
    except Exception as e:
        logging.error(f"Error generating report: {str(e)}")
        return False

def load_config():
    """
    Load email configuration from config file
    """
    config_file = CONFIG_DIR / 'email_config.ini'
    if not config_file.exists():
        raise FileNotFoundError(f"Config file not found at: {config_file}")
        
    config = configparser.ConfigParser()
    config.read(config_file)
    return config['Email']

def send_excel_report(file_path, config):
    try:
        if not os.path.exists(file_path):
            raise FileNotFoundError(f"Excel file not found at: {file_path}")

        msg = MIMEMultipart()
        current_date = datetime.now().strftime('%Y-%m-%d')
        
        msg['Subject'] = f'User Report - {current_date}'
        msg['From'] = config['sender_email']
        msg['To'] = config['recipients']
        
        body = f'Please find attached the user report generated on {current_date}.'
        msg.attach(MIMEText(body, 'plain'))
        
        with open(file_path, 'rb') as file:
            excel_attachment = MIMEApplication(file.read(), _subtype='xlsx')
            excel_attachment.add_header(
                'Content-Disposition', 
                'attachment', 
                filename=os.path.basename(file_path)
            )
            msg.attach(excel_attachment)
        
        context = ssl.create_default_context()
        
        with smtplib.SMTP_SSL(config['smtp_server'], int(config['smtp_port']), context=context) as server:
            server.set_debuglevel(0)
            server.login(config['sender_email'], config['email_password'])
            server.send_message(msg)
        
        logging.info(f"Email sent successfully to {config['recipients']} with attachment: {file_path}")
        return True
        
    except Exception as e:
        logging.error(f"Error sending email: {str(e)}")
        return False

def verify_godaddy_connection(config):
    try:
        logging.info("Verifying GoDaddy SMTP connection...")
        context = ssl.create_default_context()
        with smtplib.SMTP_SSL(config['smtp_server'], int(config['smtp_port']), context=context) as server:
            server.set_debuglevel(0)
            server.login(config['sender_email'], config['email_password'])
            logging.info("Connection test successful!")
            return True
    except Exception as e:
        logging.error(f"Connection test failed: {str(e)}")
        return False

def main():
    try:
        config = load_config()
        
        if not verify_godaddy_connection(config):
            logging.error("Failed to verify GoDaddy SMTP connection")
            sys.exit(1)
            
        if not generate_excel_report():
            logging.error("Failed to generate Excel report")
            sys.exit(1)
            
        excel_file = "user_report.xlsx"
        
        success = send_excel_report(excel_file, config)
        
        if not success:
            sys.exit(1)
            
    except Exception as e:
        logging.error(f"Error in main process: {str(e)}")
        sys.exit(1)

if __name__ == "__main__":
    main()