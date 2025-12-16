#!/usr/bin/env python
# coding: utf-8

import smtplib
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from email.mime.base import MIMEBase
from email import encoders
import pandas as pd
import os
import time

def send_email(subject, message, recipient, attachment_path, email_config):
    try:
        msg = MIMEMultipart()
        msg['From'] = email_config['sender_email']
        msg['To'] = recipient
        msg['Subject'] = subject

        msg.attach(MIMEText(message, 'plain'))
        # Attach the file
        with open(attachment_path, "rb") as attachment:
            part = MIMEBase('application', 'octet-stream')
            part.set_payload(attachment.read())
            encoders.encode_base64(part)
            part.add_header(
                'Content-Disposition',
                f'attachment; filename= {os.path.basename(attachment_path)}'
            )
            msg.attach(part)
        
        server = smtplib.SMTP(email_config['smtp_server'], email_config['smtp_port'])
        server.starttls()
        server.login(email_config['sender_email'], email_config['password'])
        server.send_message(msg)
        server.quit()
        
    except Exception as e:
        print(f"Failed to send email to {recipient}. Error: {e}")

def load_data_from_excel(file_path):
    try:
        # Load student scores directly from sheet x (index starts at 0, so this is 1st sheet)
        df_scores = pd.read_excel(file_path, sheet_name=0)
        return df_scores
    except Exception as e:
        print(f"Error loading Excel data: {e}")
        return None

def get_certificate_path(student, folder_path):
    try:
        # this one loops through all pdf files that were splitted from pdf splitter, so it's possible to use index here
        filename = f"{folder_path}/folder/filename_{student['student_id']}.pdf" 
        if os.path.exists(filename):
            return filename
        else:
            print(f"Certificate file for student {student['student_id']} does not exist.")
            return None
    except Exception as e:
        print(f"Failed to get certificate file for student {student['student_id']}. Error: {e}")
        return None

def main():
    ######################################
    file_path = 'some_file.xlsx'
    folder_path = 'some_folder_of_your_choice'
    ######################################
    
    if not os.path.exists(folder_path):
        os.makedirs(folder_path)
    
    # Load student data only (no config)
    data = load_data_from_excel(file_path)

    if data is not None:
        # create smtp email and password using google by yourself, and put information here
        email_config_base = {
            'sender_email': 'youremail@email.com',
            'password': 'google-generated-password',
            'smtp_server': 'smtp.gmail.com',
            'smtp_port': 587
        }
        
        for index, student in data.iterrows():
            start_time = time.time()  # Start timing
            message = ("Any message") # Add your message here

            attachment_path = get_certificate_path(student, folder_path)
            if attachment_path:
                # send_email(subject, message, recipient, attachment_path, email_config)
                send_email("Any topic", message, student['email'], attachment_path, email_config_base)
                duration = time.time() - start_time
                print(f"Email successfully sent to {student['email']} (Count = {index+1}), Duration: {duration:.2f} seconds")


if __name__ == "__main__":
    main()
