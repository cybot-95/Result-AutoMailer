# -*- coding: utf-8 -*-
import calendar
import datetime
import glob
import imaplib
import os
import smtplib
import time
from email.mime.application import MIMEApplication
from email.mime.image import MIMEImage
from email.mime.base import MIMEBase
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from email import encoders

import numpy as np
import pandas as pd

import secret

exam = input("Examination Name: ")


def get_students():
    student_list = pd.read_csv("student_list.csv")
    return student_list


def send_email(student_list):
    retry_list = []

    students = student_list

    # prepare server
    from_email = secret.gmail_name
    server = smtplib.SMTP("smtp.gmail.com", 587)
    server.starttls()
    server.login(secret.gmail_address, secret.gmail_password)

    for index, student in students.iterrows():
        to_email = student["Email"]  # activate when ready

        message = MIMEMultipart("alternative")
        message["From"] = from_email
        message["To"] = to_email
        message["reply-to"] = from_email
        message["Subject"] = f"{exam} Examination Marks"

        # Formatted text to be inserted
        name = f'{student["Name"]}'

        body = (
                """\
                Dear Parent Your Ward """
                + name
                + "'s" + f""",<br><br>
                The updated marks of {exam} Examination conducted is attached below.
                """
        )

        message.attach(MIMEText(body, "html"))
        # The body and the attachments for the mail
        attach_file_name = f'{name}.xlsx'
        attach_file = open(attach_file_name, 'rb')  # Open the file as binary mode
        payload = MIMEBase('application', 'vnd.ms-excel')
        payload.set_payload(attach_file.read())
        attach_file.close()
        encoders.encode_base64(payload)  # encode the attachment
        # add payload header with filename
        payload.add_header('Content-Decomposition', 'attachment', filename=attach_file_name)
        message.attach(payload)
        email_text = message.as_string()
        try:
            server.sendmail(from_email, to_email, email_text)
        except Exception as e:
            retry_list.append(student)
            continue

    return retry_list


def send_error(e):
    from_email = secret.gmail_name
    to_email = secret.admin_address
    cc_email = secret.gmail_address

    server = smtplib.SMTP("smtp.gmail.com", 587)
    server.starttls()
    server.login(secret.gmail_address, secret.gmail_password)

    today = datetime.datetime.now().strftime("%x")

    message = MIMEMultipart()
    message["From"] = from_email
    message["To"] = to_email
    message["Cc"] = cc_email
    message[
        "Subject"
    ] = "ERROR: An error occurred with the reminder email system ({today})."

    body = (
        "There was an error while trying to email students.\n\n"
        "Here is the error text:\n\n"
        f"{repr(e)}"
    )

    message.attach(MIMEText(body))
    email_text = message.as_string()

    server.sendmail(from_email, [to_email] + [cc_email], email_text)


def send_success():
    from_email = secret.gmail_name
    to_email = secret.gmail_address

    server = smtplib.SMTP("smtp.gmail.com", 587)
    server.starttls()
    server.login(secret.gmail_address, secret.gmail_password)

    today = datetime.datetime.now().strftime("%x")

    message = MIMEMultipart()
    message["From"] = from_email
    message["To"] = to_email
    message[
        "Subject"
    ] = f"The Reminder Email System finished executing successfully ({today})."

    body = "No errors occurred while sending emails.\n"

    message.attach(MIMEText(body))
    email_text = message.as_string()

    server.sendmail(from_email, to_email, email_text)


def main():
    listt = get_students()
    send_email(listt)


if __name__ == "__main__":
    try:
        main()
    except Exception as e:
        send_error(e)
        print("There has been an error while emailing the participants.\n\n")
        raise
    else:
        send_success()
