import smtplib
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from email.mime.base import MIMEBase
from email import encoders
import os
import datetime
import pandas as pd

current_day = datetime.date.today()
current_day = current_day.strftime('%d-%m-%Y')

# This configuration works only on Lotus's local network or using Zscaler VPN
smtp_host = 'mailrelay.global.lotuss.org'
smtp_port = 25


sender_email_address = 'AR-Lotuss@lotuss.com'   # email ผู้ส่ง
dev_team = []
user_team = ['pattamaporn.karuna@lotuss.com','paramita.ritidet@lotuss.com'] # email ผู้รับ
cc_team = []


class EmailSender:
    def __init__(self, smtp_host, smtp_port):
        self.smtp_host = smtp_host
        self.smtp_port = smtp_port

    def send_email(self, sender_email_address, recipient_emails, subject, message_text, attachment_files=None, cc_emails=None):
        try:
            message = MIMEMultipart()
            message['Subject'] = subject
            message['From'] = sender_email_address
            message['To'] = ', '.join(recipient_emails)
            if cc_emails != None:
                message['Cc'] = ', '.join(cc_emails)

            message.attach(MIMEText(message_text, 'html'))
            
            if attachment_files is not None:
                for attachment_file in attachment_files:
                    if attachment_file:
                        with open(attachment_file, 'rb') as attachment:
                            part = MIMEBase('application', 'octet-stream')
                            part.set_payload(attachment.read())
                            encoders.encode_base64(part)
                            backslash = "/"
                            part.add_header('Content-Disposition', f'attachment; filename= {attachment_file.split(backslash)[-1]}')
                            message.attach(part)

            with smtplib.SMTP(self.smtp_host, self.smtp_port) as server:
                server.send_message(message)

        except Exception as e:
            print(f"An error occurred while sending the email: {e}")

def sender_email(message_text,subject = f"Notification today | {current_day}",attached_file_path = None):
    attachment_file = None
    email_sender = EmailSender(smtp_host, smtp_port)
    recipient_emails = user_team + dev_team

    if(attached_file_path is not None):
        try:
            attachment_file_list = [attached_file_path + file for file in os.listdir(attached_file_path)]
            print(f'Attached file list : {attachment_file_list}')
            attachment_file = attachment_file_list
        except Exception as e:
            print(e)
            attachment_file = None
    email_sender.send_email(sender_email_address, recipient_emails, subject, message_text, attachment_file, cc_team)


def send_invoice_to_all_stores(invoice_folder_path='invoicepdf/', master_file_path='masterfile/ADM01.xlsx'):
    try:
        # Load master data
        master_df = pd.read_excel(master_file_path, dtype=str)

        # Clean the dataframe: fill NaNs to prevent crash on missing data
        master_df = master_df.fillna('')

        # Get all store folders under invoicepdf
        store_folders = [folder for folder in os.listdir(invoice_folder_path) if os.path.isdir(os.path.join(invoice_folder_path, folder))]

        print(f"Found folders: {store_folders}")

        for store_id in store_folders:
            try:
                store_path = os.path.join(invoice_folder_path, store_id)
                # ✅ Check if folder is empty
                if not os.listdir(store_path):
                    print(f"[ℹ️] Folder {store_id} is empty. Skipping email.")
                    continue

                # Lookup store info
                store_info = master_df[master_df['One Login ID'] == store_id]

                if store_info.empty:
                    print(f"[⚠️] Store ID {store_id} not found in master file.")
                    continue

                email = store_info.iloc[0]['Email']
                store_name = store_info.iloc[0]['Description'] if 'Description' in store_info.columns else 'Store'

                if not email:
                    print(f"[⚠️] Missing email for Store ID {store_id}.")
                    continue

                # Build subject and message
                subject = f"ใบแจ้งหนี้ สาขา {store_name}"
                message_text = f"""
                <html>
                <body>
                    <p>เรียน ผู้จัดการสาขา</p>
                    <p>กรุณา พิมพ์ใบแจ้งหนี้ ตามไฟล์แนบ และนำส่งร้านค้าเช่า</p>
                    <p><i>อีเมลจากระบบอัติโนมัติ กรุณาอย่าตอบกลับ</i></p>
                </body>
                </html>
                """

                # Check for PDF files to attach
                attachments = [os.path.join(store_path, file) for file in os.listdir(store_path) if file.lower().endswith('.pdf')]
                if not attachments:
                    print(f"[⚠️] No PDF files found in folder {store_id}. Skipping email.")
                    continue

                # Send email
                email_sender = EmailSender(smtp_host, smtp_port)
                email_sender.send_email(
                    sender_email_address=sender_email_address,
                    recipient_emails=[email],
                    subject=subject,
                    message_text=message_text,
                    attachment_files=attachments,
                    cc_emails=[]
                )
                print(f"[✅] Email sent to {email} for Store ID {store_id}.")

            except Exception as store_error:
                print(f"[❌] Failed to process Store ID {store_id}: {store_error}")

    except Exception as main_error:
        print(f"[‼️] Critical error: {main_error}")


if __name__ == "__main__":
    send_invoice_to_all_stores()