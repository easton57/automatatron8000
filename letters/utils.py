import os
import random
import logging
import datetime
from smtplib import SMTP
from datetime import date
from email.mime.text import MIMEText
from email.mime.multipart import MIMEMultipart
from email.mime.application import MIMEApplication
import keyring
import pandas as pd
from pyad import pyad


# Time/Date variables to get the current date for the logfile name.
now = datetime.datetime.now()
current_time = now.strftime("%H:%M:%S")
logging.basicConfig(filename=f'log_dir', level=logging.DEBUG)


def find_job_title(title):
    """Function to determine which set of password reset instructions to give the user"""
    thin_client = ['Closer', 'Fund', 'Final Docs', 'Underwriter', 'Loan Processor',
                   'Loan Coordinator', 'Closing', "Servicing", "Compliance"]
    citrix_laptop = ['Loan Officer', 'Manager', 'Vice President']
    domain = ['Loan Officer Marketing', 'Technical Support', 'Mortgage Systems',
              'Systems Administrator']

    all_titles = thin_client + citrix_laptop + domain

    for item in all_titles:
        if item in thin_client and (title in item or item in title):
            device_type = 'Thin Client Users – Once you have logged into your device, you will ' \
                         'want to press & hold these keys in this order: Ctrl + Alt + Del'\
                         '\n\n'\
                         'Note: You will not use/press the '\
                         '“+” key for the instructions above, it serves merely to indicate you ' \
                         'will want to press the following keys all together.'
            return device_type

        elif item in citrix_laptop and (title in item or item in title):
            device_type = 'Laptop Users (Citrix) – At the top of your desktop is a gray bar with ' \
                          'an arrow pointing downwards, upon clicking the bar it will present ' \
                          'you some options. To change your password, we are looking for the ' \
                          'button labeled “Ctrl + Alt + Del”, upon clicking this button, it will ' \
                          'bring you to a new screen with a list of options, one of which is ' \
                          '“Change a password”. '
            return device_type

        elif item in domain and (title in item or item in title):
            device_type = 'Non-Citrix Users - Once you have logged into your device, you will ' \
                          'want to press & hold these keys in this order: Ctrl + Alt + Del\n\n'\
                          'Note: You will not use/press the “+” key for the instructions above, ' \
                          'it serves merely to indicate you will want to press the following ' \
                          'keys all together.'
            return device_type
        else:
            pass
    else:
        # Sets the password reset instructions to the below string
        # if the title has not been seen before.
        device_type = "-WARNING- UNABLE TO DETERMINE PASSWORD INSTRUCTIONS BASED OFF OF TITLE, " \
                      "PLEASE POPULATE THIS FIELD!"
        logging.warning(f"{current_time} - FAILED TO DETERMINE PASSWORD CHANGE INSTRUCTIONS "
                        f"FOR USER.")
        return device_type


def log_cleanup():
    """Function to perform log clean up and only keep 32 items in the logging directory"""
    log_path = "log_dir"
    log_list = []
    today = date.today()
    delta = datetime.timedelta(days=32)
    alpha = today - delta
    target_date = alpha.strftime("%Y-%m-%d")

    # Getting how many logs are in the folder
    logging.info(f"{current_time} - Starting log clean up...")
    for log in os.listdir(log_path):
        if log.endswith(".log"):
            log_list.append(log)
        else:
            continue

    # Checks to see if we have 32 or more logs and if we do it cleans up the oldest.
    if len(log_list) >= 32:
        oldest_log = sorted([log_path + file for file in os.listdir(log_path)],
                            key=os.path.getmtime)[0]
        old_logname = oldest_log.split("\\")[-1]

        if target_date == old_logname.replace(".log", ""):
            logging.info(f"{current_time} - Attempting deletion of log {oldest_log}...")
            os.remove(oldest_log)
        else:
            logging.warning(f"{current_time} - Unable to locate the log with the target date of "
                            f"{target_date} which is marked for deletion...")
    else:
        logging.warning(f"{current_time} - Unable to locate the log with the target date of "
                        f"{target_date} which is marked for deletion...")


def password_generator():
    """Generates random passwords for our users"""
    df = pd.read_csv('./Dictionary.csv')
    dictionary = df['Words']
    complies = False
    word1 = random.choice(dictionary)
    word2 = random.choice(dictionary)
    word3 = random.choice(dictionary)

    password = f'{word1} {word2} {word3}'

    while not complies:
        if len(password) > 25 or len(password) <= 15:
            password = f'{random.choice(dictionary)} {random.choice(dictionary)} ' \
                       f'{random.choice(dictionary)}'
            print(password)
            print(len(password))
        else:
            complies = True
            print(f"{password} COMPLIES WITH PASSWORD POLICY")
            print(len(password))

    return password


def reset_user_password(user):
    """Function to reset a users password"""
    try:
        pyad.set_defaults(ldap_server="domain.local", username="domain\svcPDFFiller",
                          password=keyring.get_password("system", "svcDocxFiller"))

        password = password_generator()

        user = pyad.aduser.ADUser.from_cn(user)
        user.set_password(f"{password}")

        print(f"RESET {user}")
        return password

    except Exception as error:
        print("Failed to reset password with the following error:")
        logging.warning(error)
        print(error)


def send_log(users):
    """Sends the logging and success email at the end of runtime"""
    msg = MIMEMultipart('related')
    msg['From'] = "svcDocxFiller@domain.com"
    msg['Subject'] = f"New Hire Letter Generation Log for {date.today()}"

    if len(users) != 0:
        response = "Letters have been generated for the following users:<br>"
        breaks = "<br><br>"
    else:
        response = ""
        breaks = ""

    new_line = '<br>'
    msg_body = MIMEText("<p>Hello there,</p>"
                        "<p>Attached is the log for the programs last run.</p>"
                        f"<p>{response}{(new_line.join(users))}{breaks}"
                        "<br>Please report any strange behavior to "
                        "<strong>Easton Seidel</strong>.</p>"
                        "<p>Thanks,<br>svcDocxFiller</p>", "html")
    msg.attach(msg_body)

    path = "log_dir"
    log = f"{date.today()}.log"
    with open(path + log, "rb") as file:
        part = MIMEApplication(
            file.read(),
            Name=log
        )
    part['Content-Disposition'] = 'attachment; filename="%s"' % log
    msg.attach(part)

    smtp_server = SMTP("smtp_ip:25")

    try:
        receiver = "domain-servicedesk@domain.com"
        # receiver = "eseidel@domain.com"
        msg['To'] = receiver
        smtp_server.sendmail(msg["From"], receiver, msg.as_string())
    except Exception as error:
        print(f"Failed to send post run email: {error}")

    smtp_server.quit()
